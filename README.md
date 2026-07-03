# Stock Validation — ALL Marketplaces

A Streamlit app that validates `Expected Stock` from StockValidation reports against live
marketplace stock files for **Lazada, Shopee, TikTok, and Zalora**, and outputs a single
colour-coded Excel workbook (`Stock_Validation_All_MP.xlsx`).

Works with any subset of the four marketplaces — upload just what you have.

## Files it recognizes

| Keyword in filename | Marketplace | File type |
|---|---|---|
| `stockValidation-lazada` / `lazada` (.csv) | Lazada | StockValidation |
| `stockValidation-shopee` / `shopee` (.csv) | Shopee | StockValidation |
| `stockValidation-tiktok` / `tiktok` (.csv) | TikTok | StockValidation |
| `stockValidation-zalora` / `zalora` (.csv) | Zalora | StockValidation |
| `pricestock` | Lazada | Stock & Price (.xlsx) |
| `mass_update_sales_info` | Shopee | Mass Update (.xlsx or .zip) |
| `Tiktoksellercenter_batchedit` / `batchedit` | TikTok | Batch Edit (.xlsx) |
| `SellerStockTemplate` | Zalora | Stock file (required) |
| `SellerStatusTemplate` | Zalora | Status file (optional) |

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## Deploy on GitHub + Streamlit Community Cloud

1. Create a new GitHub repo and push these files:
   ```bash
   git init
   git add app.py requirements.txt README.md
   git commit -m "Stock validation ALL MP app"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app**, select your repo/branch, and set **Main file path** to `app.py`.
4. Click **Deploy**. Streamlit Cloud installs `requirements.txt` automatically and gives you a
   public URL you can share.

## Notes

- Shopee exports occasionally trip an `activePane` XML bug on read — this app patches it
  automatically before parsing.
- The optional SOH/ALL warehouse comparison is intentionally **not** auto-run — the app flags
  when both files are present but leaves the check to manual verification, since filename
  patterns alone don't guarantee the two files share the same SKU catalogue.
- Duplicate SKUs are summed automatically before comparison.
