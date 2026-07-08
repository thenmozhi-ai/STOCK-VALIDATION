# Multi-Marketplace Stock Validation

Validates `Expected Stock` from StockValidation reports against live marketplace
stock files for **Lazada**, **Shopee**, **TikTok**, and/or **Zalora** — outputs a
single colour-coded Excel workbook with a combined summary dashboard.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## Deploy to GitHub

```bash
git init
git add .
git commit -m "Multi-marketplace stock validation app"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

## Deploy to Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Click **New app**.
3. Select your repo, branch (`main`), and main file path (`app.py`).
4. Click **Deploy**. Streamlit will install `requirements.txt` automatically.

## What to upload in the app

| File | Marketplace | Notes |
|------|-------------|-------|
| `stockValidation-lazada.csv` / similar | Lazada | filename must contain "lazada" |
| `stockValidation-shopee.csv` / similar | Shopee | filename must contain "shopee" |
| `stockValidation-tiktok.csv` / similar | TikTok | filename must contain "tiktok" |
| `stockValidation-zalora.csv` / similar | Zalora | filename must contain "zalora" |
| `pricestock...xlsx` | Lazada | Stock & Price export |
| `mass_update_sales_info...xlsx` | Shopee | Mass Update export |
| `Tiktoksellercenter_batchedit...xlsx` | TikTok | Batch Edit export |
| `SellerStockTemplate...xlsx` | Zalora | Stock file |
| `SellerStatusTemplate...xlsx` | Zalora | Optional — active/inactive status |
| `SOHbySKU...xls` | Warehouse | Optional Step 1 comparison |
| `ALL...csv` (filename starts with `ALL`) | Warehouse | Optional Step 1 comparison |

You can upload any subset — the app only builds tabs for marketplaces where
both the StockValidation CSV *and* the matching stock file are present.

## Output

A single `.xlsx` workbook:

- **Summary** — KPI dashboard, one block per marketplace
- **`{Marketplace} - StockVal`** — full comparison with Status/Remark columns
- **`{Marketplace} - Mismatches`** — mismatch rows only
- **`SOH vs ALL`** / **`SOH vs ALL Mismatches`** — if warehouse files were uploaded

Colour coding: green = match, red = impact/mismatch, orange = "update to 0",
grey = SKU not found on that marketplace.
