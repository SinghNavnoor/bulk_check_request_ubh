# UBH Check Request Generator

A Streamlit web app for the UBH team. Upload the monthly CSV, generate combined check request forms, and download the DOCX — in three steps. No data is stored on the server.

---

## What it does

1. **Upload** — Drop the monthly CSV file (exported from HubSpot)
2. **Generate** — The app fills in the Word template for each row and merges them into one combined document
3. **Download** — Click Download DOCX to save the file to your computer

When the browser tab is closed, all data is wiped. Nothing is saved on the server.

---

## How to run locally

**Requirements:** Python 3.10+, Microsoft Word (for local PDF preview — not needed for the web app)

```bash
# Install dependencies
pip3 install -r requirements.txt

# Run the app
python3 -m streamlit run app.py
```

The app opens at `http://localhost:8501`. Log in with the team credentials.

For local runs, create `.streamlit/secrets.toml` with:
```toml
USERNAME = ******
PASSWORD = ******
```

This file is gitignored and never committed.

---

## How to deploy (Streamlit Community Cloud)

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
2. Click **Create app** → **Deploy a public app from GitHub**
3. Set repository: `SinghNavnoor/bulk_check_request_ubh`, branch: `main`, file: `app.py`
4. Click **Advanced settings** → paste into Secrets:
   ```toml
   USERNAME = ******
   PASSWORD = ******
   ```

5. Click **Deploy** — first deploy takes 3–5 minutes (LibreOffice installs automatically)

To change the password: update the secret in the Streamlit Cloud dashboard → Reboot app. No code change needed.

---

## Project structure

```
├── app.py                         # Streamlit web app
├── csv_to_word_forms.py           # Core engine — fills template, merges forms
├── csv_to_pdf_forms.py            # CLI-only wrapper (not used by web app)
├── Form Template/
│   └── Rapid Rehousing Program Check Request Form - Template.docx
├── requirements.txt               # Python dependencies
├── packages.txt                   # apt packages (libreoffice) for Streamlit Cloud
├── .streamlit/
│   ├── config.toml                # Theme and server settings
│   └── secrets.toml               # Local credentials (gitignored)
└── .gitignore
```

---

## Privacy

- Uploaded CSV files are processed in a temporary directory that is deleted immediately after the document is generated
- No client data is written to a database or stored on disk beyond the lifetime of one request
- Credentials are stored in Streamlit's secrets manager, never in the repository
- HTTPS is enforced by Streamlit Community Cloud

---

## CSV format

The app expects a CSV exported from HubSpot with columns including:

| Column | Example |
|---|---|
| Client Name | Client X |
| Payment Date (Today's Date) | 5/20/2020 |
| Program (Sync) | Housing Program |
| Check Type | Monthly Rent |
| Type of Assistance | Rent |
| UBH Amount | 2000 |
| Client Rent Amount | 0 |
| Check Payable to (Sync) | Y LLC |
| Landlord Address Sync | 1234 New Street, Long Beach, CA |
| Month (Month the Assistance is being paid for) | July |
| Payment Month - Calc | June |
| Payment Year - Calc | 2026 |

Column headers are matched flexibly (case-insensitive, tolerates typos like "Assitance").
