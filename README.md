# Multi-Marketplace Stock Validation

A Streamlit dashboard that validates `Expected Stock` from StockValidation
reports against live marketplace stock files for **Lazada**, **Shopee**,
**TikTok**, **Zalora**, and **Shopify** — producing a single colour-coded
Excel workbook with a combined summary dashboard, plus a Product Master
gating step for SOH / DTC / Warehouse reports.

## Features

- Sidebar file uploads per marketplace (no ZIP required — one file per slot)
- Product Master validation gate before stock comparison (SOH, DTC, Warehouse)
- Bidirectional mismatch detection (missing in marketplace report **and**
  missing in the validation file)
- Status Validation / Downloads / Saved Reports dashboard tabs
- Reports persist to a local `Reports/` folder between runs

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## Push to GitHub

```bash
git init
git add .
git commit -m "Multi-marketplace stock validation dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

## Deploy to Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Click **New app**.
3. Select your repo, branch (`main`), and main file path (`app.py`).
4. Click **Deploy** — Streamlit installs `requirements.txt` automatically.

## Project structure

```
.
├── app.py                    # Full app — parsing, validation, and dashboard UI
├── requirements.txt          # Python dependencies
├── .streamlit/config.toml    # Raises the upload size limit
├── .gitignore
└── README.md
```

## File format notes

| Upload | Format |
|---|---|
| Lazada MP file | `pricestock...xlsx` Stock & Price export |
| Shopee MP file | `mass_update_sales_info...xlsx` export |
| TikTok Active/Inactive Batch Edit | `Tiktoksellercenter_batchedit...xlsx` (both required) |
| Zalora MP file | `SellerStockTemplate...xlsx` |
| Shopify MP file | Standard Shopify "Export inventory" CSV |
| Inventory files | StockValidation CSV per marketplace (`Seller SKU`, `Expected Stock`) |
| Product Master | Any file with SKU + Name columns |
| SOH | `SOHbySKU...xls` warehouse export |
