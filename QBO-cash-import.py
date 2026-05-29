import os
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PayPal Cash to QuickBooks", layout="wide")


def safe_read_excel(source, skiprows):
    try:
        df = pd.read_excel(source, skiprows=skiprows, engine="openpyxl")
    except ImportError as exc:
        raise ImportError(
            "The 'openpyxl' package is required to read .xlsx files. "
            "Install it with `pip install openpyxl` or add it to requirements.txt."
        ) from exc
    return df


def parse_reports(itemized_source, receipts_source):
    items_df = safe_read_excel(itemized_source, skiprows=6)
    items_df.columns = items_df.columns.str.strip()

    receipts_df = safe_read_excel(receipts_source, skiprows=16)
    receipts_df.columns = receipts_df.columns.str.strip()

    return items_df, receipts_df


def build_qbo_dataframe(items_df: pd.DataFrame, receipts_df: pd.DataFrame):
    cash_receipts = receipts_df[
        receipts_df['Payment method'].astype(str).str.lower() == 'cash'
    ].copy()

    if cash_receipts.empty:
        return None

    merged_df = pd.merge(
        items_df,
        cash_receipts[['Receipt number', 'Payment method']],
        on='Receipt number',
        how='inner'
    )

    if merged_df.empty:
        return None

    merged_df['Variant'] = merged_df['Variant'].fillna('').astype(str)
    merged_df['Line_Description'] = merged_df.apply(
        lambda r: f"{r['Name']} ({r['Variant']})" if r['Variant'] != '' else r['Name'],
        axis=1
    )

    qbo_df = pd.DataFrame({
        'Sales Receipt No.': merged_df['Receipt number'],
        'Transaction Date': merged_df['Date'],
        'Customer': "Cash Customer",
        'Product/Service': merged_df['SKU'],
        'Description': merged_df['Line_Description'],
        'Quantity': merged_df['Quantity'],
        'Rate': merged_df['Price (USD)'],
        'Deposit To': "Undeposited Funds",
        'Payment Method': "Cash"
    })

    qbo_df = qbo_df.dropna(subset=['Sales Receipt No.', 'Product/Service'])
    return qbo_df


def run_streamlit_app():
    st.title("PayPal Cash to QuickBooks")
    st.markdown(
        "Use the controls below to upload your PayPal Excel reports. "
        "This app generates a QuickBooks Online import CSV for cash sales."
    )

    today = date.today()
    start_date = st.date_input("Report START date", value=today)
    end_date = st.date_input("Report END date", value=today)

    st.markdown("### Upload PayPal report files")
    itemized_file = st.file_uploader(
        "PayPal POS Raw Data Report",
        type=["xlsx"],
        help="Upload the itemized sales report from PayPal.",
        key="itemized_file"
    )
    receipts_file = st.file_uploader(
        "PayPal POS Receipts Report",
        type=["xlsx"],
        help="Upload the receipts payment report from PayPal.",
        key="receipts_file"
    )

    if itemized_file is not None:
        st.write(f"- Uploaded itemized report: `{itemized_file.name}`")
    if receipts_file is not None:
        st.write(f"- Uploaded receipts report: `{receipts_file.name}`")

    if st.button("Generate QuickBooks CSV"):
        if start_date > end_date:
            st.error("Start date must be the same as or earlier than end date.")
            return

        if itemized_file is None or receipts_file is None:
            st.error("Please upload both the itemized report and the receipts report.")
            return

        try:
            with st.spinner("Parsing PayPal reports..."):
                items_df, receipts_df = parse_reports(itemized_file, receipts_file)
                qbo_df = build_qbo_dataframe(items_df, receipts_df)

            if qbo_df is None or qbo_df.empty:
                st.warning(
                    "No cash transactions were found in the receipts report, "
                    "or the receipt numbers did not match between the two reports."
                )
                return

            csv_bytes = qbo_df.to_csv(index=False).encode('utf-8')
            file_name = f"qbo_cash_import_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.csv"

            st.success("CSV generated successfully.")
            st.download_button(
                label="Download QuickBooks CSV",
                data=csv_bytes,
                file_name=file_name,
                mime="text/csv"
            )
            st.dataframe(qbo_df.head(20))
            st.info("The generated CSV is ready for QuickBooks Online Advanced Batch Import.")

        except ImportError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"An unexpected processing error occurred: {exc}")


if __name__ == "__main__":
    run_streamlit_app()
