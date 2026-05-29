# QBO Cash Import Streamlit App

A Streamlit app for converting PayPal cash sales reports into a QuickBooks Online import CSV.

## What this app does

- Accepts PayPal "POS Raw Data" and "POS Receipts" Excel reports via file upload
- Filters cash payments from the receipts report
- Merges cash receipts with itemized sales lines
- Builds a QuickBooks Online import-ready CSV
- Provides a download button for the generated CSV

## How to run locally

1. Install requirements

   ```bash
   pip install -r requirements.txt
   ```

2. Start the app

   ```bash
   streamlit run QBO-cash-import.py
   ```

3. Open the local Streamlit URL shown in the terminal

## Supported file format

- Upload `.xlsx` files only
- Expected report names are:
  - "PayPal-POS-Raw-Data-Report-START-END.xlsx"
  - "PayPal-POS-Receipts-Report-START-END.xlsx"
- The app reads Excel contents with `openpyxl`
- Output file in CSV format required by QuickBooks Online (QBO)

## Notes

- The app now runs directly from `QBO-cash-import.py`
- `streamlit_app.py` is no longer used
- `./.streamlit/config.toml` is included to configure Streamlit for deployment
- The app theme now matches the JCFC Boosters short.io branding from `go.jcfcboosters.org`
- If you plan to deploy on Streamlit Cloud, the same file is the app entrypoint

## Deploying to Streamlit Community Cloud

1. Go to https://streamlit.io/cloud and sign in with GitHub.
2. Create a new app.
3. Select repo: `jcfcboosters/jcfc-treasurer`.
4. Choose branch: `main`.
5. Set the main file path to: `QBO-cash-import.py`.
6. Deploy and use the generated public URL.

If the app shows an error, check the deployment logs for missing dependencies and confirm `requirements.txt` includes `streamlit`, `pandas`, and `openpyxl`.


