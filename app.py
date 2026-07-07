"""
Multi-Marketplace Stock Validation App (Lazada / Shopee / TikTok / Zalora)
---------------------------------------------------------------------------
Upload StockValidation CSVs + marketplace stock files, get back a single
colour-coded Excel workbook with one tab-pair per marketplace plus a
combined summary dashboard.

Run locally:
    streamlit run app.py

Deploy on Streamlit Community Cloud:
    1. Push this folder to a GitHub repo.
    2. On share.streamlit.io, "New app" -> point at the repo / app.py.
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Multi-Marketplace Stock Validation", layout="wide")

# ----------------------------------------------------------------------------
# Constants / styling
# ----------------------------------------------------------------------------
NAVY = "1F3864"
GREEN_FILL, GREEN_FONT = "C6EFCE", "375623"
RED_FILL, RED_FONT = "FFC7CE", "9C0006"
ORANGE_FILL, ORANGE_FONT = "FFEB9C", "7D4800"
GREY_FILL, GREY_FONT = "D9D9D9", "595959"
ALT_ROW_FILL = "F2F7FB"
THIN_GREY = Side(style="thin", color="BFBFBF")

NOT_FOUND_LABELS = {
    "Lazada": "NOT IN S&P",
    "Shopee": "NOT IN SHOPEE",
    "TikTok": "NOT IN TIKTOK",
    "Zalora": "NOT IN ZALORA",
}

MARKETPLACE_ORDER = ["Lazada", "Shopee", "TikTok", "Zalora"]

# ----------------------------------------------------------------------------
# File-type detection
# ----------------------------------------------------------------------------
def classify_file(filename: str):
    """Return (marketplace, file_role) based on filename heuristics."""
    name = filename.lower()

    if "pricestock" in name:
        return "Lazada", "stock_price"
    if "mass_update_sales_info" in name:
        return "Shopee", "mass_update"
    if "tiktoksellercenter_batchedit" in name or "batchedit" in name:
        return "TikTok", "batch_edit"
    if "sellerstocktemplate" in name:
        return "Zalora", "stock_file"
    if "sellerstatustemplate" in name:
        return "Zalora", "status_file"
    if "sohbysku" in name:
        return "Warehouse", "soh"
    if name.startswith("all") and name.endswith(".csv"):
        return "Warehouse", "all_report"

    # StockValidation CSVs - identify marketplace by keyword in name
    if "stockvalidation" in name or name.endswith(".csv"):
        for mp in MARKETPLACE_ORDER:
            if mp.lower() in name:
                return mp, "stock_validation"

    return None, None


# ----------------------------------------------------------------------------
# Parsers per marketplace stock file
# ----------------------------------------------------------------------------
def fix_shopee_xlsx(file_bytes: bytes) -> bytes:
    """Patch the activePane XML bug present in Shopee's exported xlsx."""
    src = io.BytesIO(file_bytes)
    dst = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                text = data.decode("utf-8", errors="ignore")
                for old, new in [
                    ("bottom_left", "bottomLeft"),
                    ("top_left", "topLeft"),
                    ("bottom_right", "bottomRight"),
                    ("top_right", "topRight"),
                ]:
                    text = text.replace(f'activePane="{old}"', f'activePane="{new}"')
                    text = text.replace(f'pane="{old}"', f'pane="{new}"')
                data = text.encode("utf-8")
            zout.writestr(item, data)
    dst.seek(0)
    return dst.read()


def parse_lazada_stock_price(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), header=0, skiprows=3)
    df = df.iloc[:, :15]
    df.columns = [
        "Product ID", "catId", "Product Name", "currencyCode", "sku.skuId",
        "status", "Shop SKU", "SellerSKU", "Quantity", "Price", "SpecialPrice",
        "SpecialPrice Start", "SpecialPrice End", "Variations Combo", "md5key",
    ]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    return df.groupby("SellerSKU")["Quantity"].sum().to_dict()


def parse_shopee_mass_update(file_bytes: bytes) -> dict:
    fixed = fix_shopee_xlsx(file_bytes)
    df = pd.read_excel(io.BytesIO(fixed), header=2, skiprows=[3, 4])
    df = df.iloc[:, :10]
    df.columns = [
        "Product ID", "Category", "Product Name", "Parent SKU", "SKU",
        "Price", "GTIN", "Quantity", "Minimum Purchase Quantity", "Fail Reason",
    ]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    return df.groupby("SKU")["Quantity"].sum().to_dict()


def parse_tiktok_batch_edit(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), header=2, skiprows=[3, 4])
    df = df.iloc[:, :8]
    df.columns = [
        "Product ID", "Category", "Product Name", "SKU ID",
        "Variation Option", "Price", "Quantity", "Seller SKU",
    ]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    return df.groupby("Seller SKU")["Quantity"].sum().to_dict()


def parse_zalora_stock_file(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    df = df.iloc[:, :4]
    df.columns = ["SellerSku", "ShopSku", "Quantity", "Name"]
    df["SellerSku"] = df["SellerSku"].astype(str).str.strip()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    return df.groupby("SellerSku")["Quantity"].sum().to_dict()


def parse_zalora_status_file(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    df = df.iloc[:, :4]
    df.columns = ["SellerSku", "ShopSku", "Name", "Status"]
    df["SellerSku"] = df["SellerSku"].astype(str).str.strip()
    return df.set_index("SellerSku")["Status"].astype(str).str.strip().str.lower().to_dict()


def parse_soh(file_bytes: bytes) -> dict:
    """SOHbySKU is an XML-formatted .xls file. Data rows start at index 7."""
    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    root = ET.fromstring(file_bytes)
    rows = root.findall(".//ss:Row", ns)
    lookup = {}
    for row in rows[7:]:
        cells = row.findall("ss:Cell", ns)
        values = []
        for cell in cells:
            data_el = cell.find("ss:Data", ns)
            values.append(data_el.text if data_el is not None else None)
        if len(values) > 14 and values[6]:
            sku = str(values[6]).strip()
            try:
                qty = int(float(values[14]))
            except (TypeError, ValueError):
                qty = 0
            lookup[sku] = lookup.get(sku, 0) + qty
    return lookup


def parse_all_report(file_bytes: bytes) -> dict:
    df = pd.read_csv(io.BytesIO(file_bytes))
    df.columns = [c.strip() for c in df.columns]
    qty_col = df.columns[14] if len(df.columns) > 14 else None
    if "sellerSKU" not in df.columns or qty_col is None:
        return {}
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0).astype(int)
    return df.groupby("sellerSKU")[qty_col].sum().to_dict()


def parse_stock_validation_csv(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes))
    df.columns = [c.strip() for c in df.columns]
    df["Expected Stock"] = pd.to_numeric(df["Expected Stock"], errors="coerce").fillna(0).astype(int)
    df["Seller SKU"] = df["Seller SKU"].astype(str).str.strip()
    return df


# ----------------------------------------------------------------------------
# Remark logic (identical across marketplaces)
# ----------------------------------------------------------------------------
def get_status_remark(expected, sp_quantity, not_found_label):
    if pd.isna(sp_quantity):
        return "Mismatch", not_found_label
    exp = int(expected)
    sp = int(sp_quantity)
    if exp == 0 and sp > 0:
        return "Mismatch", "UPDATE 0"
    if exp != sp:
        return "Mismatch", "IMPACT"
    return "Match", "TRUE"


def build_marketplace_df(stockval_df: pd.DataFrame, sp_lookup: dict, marketplace: str,
                          status_lookup: dict = None):
    df = stockval_df.copy()
    df["SP_Quantity"] = df["Seller SKU"].map(sp_lookup)
    not_found_label = NOT_FOUND_LABELS[marketplace]

    statuses, remarks = [], []
    for _, row in df.iterrows():
        status, remark = get_status_remark(row["Expected Stock"], row["SP_Quantity"], not_found_label)
        statuses.append(status)
        remarks.append(remark)
    df["Status"] = statuses
    df["Remark"] = remarks

    if marketplace == "Zalora" and status_lookup:
        df["Zalora_Status"] = df["Seller SKU"].map(status_lookup).fillna("—")
        cols = list(df.columns)
        cols.remove("Zalora_Status")
        insert_at = cols.index("SP_Quantity") + 1
        cols.insert(insert_at, "Zalora_Status")
        df = df[cols]

    return df


# ----------------------------------------------------------------------------
# Excel writer
# ----------------------------------------------------------------------------
def style_header(ws, ncols, row=1, height=36):
    ws.row_dimensions[row].height = height
    fill = PatternFill("solid", fgColor=NAVY)
    font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN_GREY)


def write_df_sheet(wb, sheet_name, df, remark_col="Remark", status_col="Status"):
    ws = wb.create_sheet(sheet_name[:31])
    ws.append(list(df.columns))
    style_header(ws, len(df.columns))

    remark_idx = df.columns.get_loc(remark_col) + 1 if remark_col in df.columns else None
    status_idx = df.columns.get_loc(status_col) + 1 if status_col in df.columns else None

    color_map = {
        "TRUE": (GREEN_FILL, GREEN_FONT),
        "Match": (GREEN_FILL, GREEN_FONT),
        "IMPACT": (RED_FILL, RED_FONT),
        "Mismatch": (RED_FILL, RED_FONT),
        "UPDATE 0": (ORANGE_FILL, ORANGE_FONT),
    }
    for label in NOT_FOUND_LABELS.values():
        color_map[label] = (GREY_FILL, GREY_FONT)

    for r_i, row in enumerate(df.itertuples(index=False), start=2):
        for c_i, value in enumerate(row, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.border = Border(top=THIN_GREY, bottom=THIN_GREY, left=THIN_GREY, right=THIN_GREY)
            if r_i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=ALT_ROW_FILL)
        remark_val = df.iloc[r_i - 2][remark_col] if remark_col in df.columns else None
        fill_color, font_color = color_map.get(remark_val, (None, None))
        if fill_color:
            for idx in (remark_idx, status_idx):
                if idx:
                    cell = ws.cell(row=r_i, column=idx)
                    cell.fill = PatternFill("solid", fgColor=fill_color)
                    cell.font = Font(name="Arial", size=10, bold=True, color=font_color)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    for i, col in enumerate(df.columns, start=1):
        width = max(12, min(30, int(df[col].astype(str).str.len().max() or 10) + 4))
        if col == "Zalora_Status":
            width = 16
        ws.column_dimensions[get_column_letter(i)].width = width
    return ws


def add_summary_block(ws, start_row, title, counts):
    fill = PatternFill("solid", fgColor=NAVY)
    ws.cell(row=start_row, column=1, value=title).font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    ws.cell(row=start_row, column=1).fill = fill
    for c in range(2, 7):
        ws.cell(row=start_row, column=c).fill = fill

    labels = ["Total", "TRUE", "IMPACT", "UPDATE 0", "NOT FOUND", "Total Mismatches"]
    colors = [None, (GREEN_FILL, GREEN_FONT), (RED_FILL, RED_FONT), (ORANGE_FILL, ORANGE_FONT),
              (GREY_FILL, GREY_FONT), (RED_FILL, RED_FONT)]
    values = [
        counts["total"], counts["true"], counts["impact"],
        counts["update0"], counts["not_found"], counts["mismatches"],
    ]
    for i, (label, val, color) in enumerate(zip(labels, values, colors)):
        r = start_row + 1
        c = i + 1
        ws.cell(row=r, column=c, value=label).font = Font(name="Arial", size=10, bold=True)
        cell = ws.cell(row=r + 1, column=c, value=val)
        cell.font = Font(name="Arial", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center")
        if color:
            cell.fill = PatternFill("solid", fgColor=color[0])
            cell.font = Font(name="Arial", size=11, bold=True, color=color[1])
    return start_row + 4


def compute_counts(df, not_found_label):
    total = len(df)
    true_ct = (df["Remark"] == "TRUE").sum()
    impact_ct = (df["Remark"] == "IMPACT").sum()
    update0_ct = (df["Remark"] == "UPDATE 0").sum()
    nf_ct = (df["Remark"] == not_found_label).sum()
    mismatches = (df["Status"] == "Mismatch").sum()
    return dict(total=total, true=true_ct, impact=impact_ct, update0=update0_ct,
                not_found=nf_ct, mismatches=mismatches)


def build_workbook(marketplace_data: dict, warehouse_df: pd.DataFrame = None, brand_name="Shop"):
    wb = Workbook()
    wb.remove(wb.active)
    summary_ws = wb.create_sheet("Summary")
    summary_ws.column_dimensions["A"].width = 20
    for col in "BCDEF":
        summary_ws.column_dimensions[col].width = 16

    summary_ws.cell(row=1, column=1,
                     value=f"📊  Multi-Marketplace Stock Validation — {brand_name}").font = Font(
        name="Arial", size=14, bold=True)
    row_cursor = 3

    for marketplace in MARKETPLACE_ORDER:
        if marketplace not in marketplace_data:
            continue
        df = marketplace_data[marketplace]
        not_found_label = NOT_FOUND_LABELS[marketplace]
        counts = compute_counts(df, not_found_label)
        row_cursor = add_summary_block(summary_ws, row_cursor, marketplace.upper(), counts)

        write_df_sheet(wb, f"{marketplace} - StockVal", df)
        mism_df = df[df["Status"] == "Mismatch"]
        write_df_sheet(wb, f"{marketplace} - Mismatches", mism_df)

    if warehouse_df is not None:
        counts = compute_counts(warehouse_df, "NOT FOUND")
        row_cursor = add_summary_block(summary_ws, row_cursor, "SOH vs ALL", counts)
        write_df_sheet(wb, "SOH vs ALL", warehouse_df)
        mism_df = warehouse_df[warehouse_df["Status"] == "Mismatch"]
        write_df_sheet(wb, "SOH vs ALL Mismatches", mism_df)

    return wb


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
st.title("📊 Multi-Marketplace Stock Validation")
st.caption("Lazada · Shopee · TikTok · Zalora — upload any subset of files, get one colour-coded workbook.")

with st.expander("ℹ️ What files can I upload?", expanded=False):
    st.markdown(
        """
- **StockValidation CSVs** — one per marketplace (filename should contain the marketplace name, e.g. `stockValidation-lazada.csv`)
- **Lazada**: `pricestock...xlsx` (Stock & Price export)
- **Shopee**: `mass_update_sales_info...xlsx`
- **TikTok**: `Tiktoksellercenter_batchedit...xlsx`
- **Zalora**: `SellerStockTemplate...xlsx` (+ optional `SellerStatusTemplate...xlsx`)
- **Warehouse (optional)**: `SOHbySKU...xls` and `ALL...csv`

Only upload the marketplaces you want validated — any subset works.
        """
    )

brand_name = st.text_input("Brand / Shop name (shown on the summary tab)", value="My Shop")

uploaded_files = st.file_uploader(
    "Upload your files",
    accept_multiple_files=True,
    type=["csv", "xlsx", "xls"],
)

if uploaded_files:
    detected = []
    buckets = {"stock_validation": {}, "stock_price": None, "mass_update": None,
               "batch_edit": None, "stock_file": None, "status_file": None,
               "soh": None, "all_report": None}

    for f in uploaded_files:
        mp, role = classify_file(f.name)
        detected.append((f.name, mp, role))
        content = f.read()
        if role == "stock_validation":
            buckets["stock_validation"][mp] = parse_stock_validation_csv(content)
        elif role == "stock_price":
            buckets["stock_price"] = content
        elif role == "mass_update":
            buckets["mass_update"] = content
        elif role == "batch_edit":
            buckets["batch_edit"] = content
        elif role == "stock_file":
            buckets["stock_file"] = content
        elif role == "status_file":
            buckets["status_file"] = content
        elif role == "soh":
            buckets["soh"] = content
        elif role == "all_report":
            buckets["all_report"] = content

    st.subheader("Detected files")
    det_df = pd.DataFrame(detected, columns=["Filename", "Marketplace", "Role"])
    st.dataframe(det_df, use_container_width=True, hide_index=True)

    unclassified = det_df[det_df["Marketplace"].isna()]
    if len(unclassified):
        st.warning(
            "Some files could not be auto-classified and were ignored: "
            + ", ".join(unclassified["Filename"].tolist())
        )

    if st.button("🚀 Run Validation", type="primary"):
        marketplace_data = {}

        if "Lazada" in buckets["stock_validation"] and buckets["stock_price"]:
            sp_lookup = parse_lazada_stock_price(buckets["stock_price"])
            marketplace_data["Lazada"] = build_marketplace_df(
                buckets["stock_validation"]["Lazada"], sp_lookup, "Lazada")

        if "Shopee" in buckets["stock_validation"] and buckets["mass_update"]:
            sp_lookup = parse_shopee_mass_update(buckets["mass_update"])
            marketplace_data["Shopee"] = build_marketplace_df(
                buckets["stock_validation"]["Shopee"], sp_lookup, "Shopee")

        if "TikTok" in buckets["stock_validation"] and buckets["batch_edit"]:
            sp_lookup = parse_tiktok_batch_edit(buckets["batch_edit"])
            marketplace_data["TikTok"] = build_marketplace_df(
                buckets["stock_validation"]["TikTok"], sp_lookup, "TikTok")

        if "Zalora" in buckets["stock_validation"] and buckets["stock_file"]:
            sp_lookup = parse_zalora_stock_file(buckets["stock_file"])
            status_lookup = parse_zalora_status_file(buckets["status_file"]) if buckets["status_file"] else None
            marketplace_data["Zalora"] = build_marketplace_df(
                buckets["stock_validation"]["Zalora"], sp_lookup, "Zalora", status_lookup)

        warehouse_df = None
        if buckets["soh"] and buckets["all_report"]:
            soh_lookup = parse_soh(buckets["soh"])
            all_lookup = parse_all_report(buckets["all_report"])
            rows = []
            all_skus = set(soh_lookup) | set(all_lookup)
            for sku in all_skus:
                exp = soh_lookup.get(sku)
                act = all_lookup.get(sku)
                if exp is None:
                    continue
                status, remark = get_status_remark(exp, act, "NOT FOUND")
                rows.append({"Seller SKU": sku, "Expected Stock": exp, "SP_Quantity": act,
                             "Status": status, "Remark": remark})
            if rows:
                warehouse_df = pd.DataFrame(rows)

        if not marketplace_data and warehouse_df is None:
            st.error(
                "No complete marketplace pair was found (need both a StockValidation CSV "
                "AND the matching stock file for at least one marketplace)."
            )
        else:
            wb = build_workbook(marketplace_data, warehouse_df, brand_name)
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

            suffix = "Multi" if len(marketplace_data) > 1 else (list(marketplace_data)[0] if marketplace_data else "Warehouse")
            filename = f"Stock_Validation_{suffix}_{datetime.now().strftime('%Y%m%d')}.xlsx"

            st.success("Validation complete!")
            cols = st.columns(len(marketplace_data) or 1)
            for i, (mp, df) in enumerate(marketplace_data.items()):
                with cols[i]:
                    mismatches = (df["Status"] == "Mismatch").sum()
                    st.metric(f"{mp} mismatches", int(mismatches), delta=None)

            st.download_button(
                "⬇️ Download Excel Report",
                data=buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Upload StockValidation CSVs plus the matching marketplace stock files to begin.")
