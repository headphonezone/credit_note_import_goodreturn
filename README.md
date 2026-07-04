# HPZ Credit Note Generator

Streamlit app that reads the RMS Google Sheet, fetches rate/GST data from
Unicommerce, and produces a Tally-ready Excel file for Credit Note (Goods Return) import.

## How it works

```
Google Sheet (Return Good rows, date-filtered)
        ↓
Unicommerce API 1: saleOrder/search  →  internal sale order code
Unicommerce API 2: saleOrder/get     →  shippingPackageCode
Unicommerce API 3: invoice/details/get → rate, GST %, COD charges
        ↓
Excel (Tally Accounting Voucher format)
        ↓
TallyPrime  →  Data → Import → Excel
```

## Deploy on Streamlit Community Cloud (free)

1. Push this folder to a GitHub repo (can be private).
2. Go to https://share.streamlit.io → New app → pick the repo.
3. Set main file: `app.py`
4. Go to **Settings → Secrets** and paste:

```toml
UC_PASSWORD = "Yashu@123"
SHEET_ID    = "your_google_sheet_spreadsheet_id"
GCP_SA_JSON = '''{ ... paste full service account JSON ... }'''
```

5. Click Deploy. That's it — free, always-on.

## GCP Service Account setup (one-time)

1. GCP Console → IAM & Admin → Service Accounts → Create.
2. Name it `streamlit-sheets-reader` (or anything).
3. No roles needed at project level.
4. After creation → Keys tab → Add Key → JSON → download.
5. In Google Sheet → Share → paste the service account email → Viewer.
6. Enable APIs: Google Sheets API, Google Drive API (in GCP Console).

## Local run

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.template .streamlit/secrets.toml
# fill in secrets.toml
streamlit run app.py
```

## Tally import steps

1. Open TallyPrime → Company: Headphone Zone Pvt Ltd Chennai.
2. Go to **Data → Import → Excel**.
3. Select the downloaded `.xlsx` file.
4. Tally auto-maps columns (they match the standard template headers).
5. Import. Done.

## Notes

- One sheet row = 1 qty = 1 credit note line item.
- Multiple sheet rows with the same Order No → merged into one credit note with multiple items.
- Item name in Excel = Model Name from Google Sheet (Tally must have this as a stock item).
- Rate and GST come from Unicommerce invoice — not from the sheet.
- COD Charges reversed only if the original order was COD (detected from Unicommerce).
- CGST+SGST for intra-state (TN); IGST for inter-state (auto-detected from Unicommerce).
