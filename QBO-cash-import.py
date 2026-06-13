import streamlit as pd_stream
import pandas as pd
import io

pd_stream.set_page_config(page_title="PayPal POS to QBO Cash Merger", layout="centered")

pd_stream.title("💰 PayPal POS to QuickBooks Cash Merger")
pd_stream.markdown("""
This utility extracts itemized **Cash Sales** from your PayPal Zettle Excel exports 
and formats them perfectly for QuickBooks Online Advanced Batch Import.
""")
pd_stream.markdown("---")

pd_stream.subheader("Step 1: Upload PayPal Excel Exports")
itemized_file = pd_stream.file_uploader("Upload 'PayPal-POS-Raw-Data-Report' (.xlsx)", type=["xlsx"])
receipts_file = pd_stream.file_uploader("Upload 'PayPal-POS-Receipts-Report' (.xlsx)", type=["xlsx"])

def load_excel_with_flexible_headers(uploaded_file, target_column, aliases=None):
    """
    Reads the entire Excel sheet, locates the exact row containing the target_column,
    strips the metadata above it, and perfectly aligns the column headers.
    """
    if aliases is None:
        aliases = []

    target_names = {target_column.lower(), *(alias.lower() for alias in aliases)}

    # Read the file completely without headers or skipping anything
    df = pd.read_excel(uploaded_file, header=None)
    
    header_row_idx = None
    # Look at every row to see where our target column header lives
    for idx, row in df.iterrows():
        row_values = row.map(lambda x: '' if pd.isna(x) else str(x).strip().lower()).tolist()
        for value in row_values:
            if not value:
                continue
            if any(target_name == value or target_name in value or value in target_name for target_name in target_names):
                header_row_idx = idx
                break
        if header_row_idx is not None:
            break
            
    if header_row_idx is None:
        raise ValueError(f"Could not find the expected column '{target_column}' anywhere in the file.")
        
    # Set the discovered row as the actual columns
    df.columns = df.iloc[header_row_idx].astype(str).str.strip()
    
    # Drop the metadata rows above the headers, and drop the header row itself from the data
    df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
    
    # Force clean all header column names to prevent string mismatch errors
    df.columns = df.columns.str.strip()
    return df


def standardize_dataframe_columns(df):
    """
    Clean column names for consistent downstream processing.
    """
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df


def find_column(df, candidates):
    """
    Find the first candidate column name that exists in the dataframe.
    """
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    return None


def build_qbo_output_df(
    merged_df,
    items_receipt_col,
    sku_col,
    date_col,
    quantity_col,
    price_col,
    name_col,
    variant_col=None,
):
    """Build the final QBO output dataframe with formatted dates and line totals."""
    merged_df = merged_df.copy()

    if variant_col is None:
        merged_df['variant'] = ''
        variant_col = 'variant'

    merged_df[variant_col] = merged_df[variant_col].fillna('').astype(str)
    merged_df['line_description'] = merged_df.apply(
        lambda r: f"{r[name_col]} ({r[variant_col]})" if r[variant_col] != '' else r[name_col],
        axis=1,
    )

    qbo_df = pd.DataFrame()
    qbo_df['Sales Receipt No.'] = merged_df[items_receipt_col]
    qbo_df['Transaction Date'] = pd.to_datetime(
        merged_df[date_col],
        errors='coerce',
    ).dt.strftime('%Y-%m-%d')
    qbo_df['Customer'] = 'Cash Customer'
    qbo_df['Product/Service'] = merged_df[sku_col]
    qbo_df['Description'] = merged_df['line_description']
    qbo_df['Quantity'] = merged_df[quantity_col]
    qbo_df['Rate'] = merged_df[price_col]
    qbo_df['Total'] = (
        pd.to_numeric(qbo_df['Quantity'], errors='coerce') *
        pd.to_numeric(qbo_df['Rate'], errors='coerce')
    )
    qbo_df['Deposit To'] = 'Undeposited Funds'
    qbo_df['Payment Method'] = 'Cash'
    return qbo_df

if itemized_file and receipts_file:
    pd_stream.markdown("---")
    pd_stream.subheader("Step 2: Verify and Process")
    
    try:
        # Load files completely dynamically regardless of empty rows or metadata length
        receipt_aliases = ['Receipt no.', 'Receipt #', 'Receipt']
        items_df = standardize_dataframe_columns(load_excel_with_flexible_headers(itemized_file, 'Receipt number', aliases=receipt_aliases))
        receipts_df = standardize_dataframe_columns(load_excel_with_flexible_headers(receipts_file, 'Receipt number', aliases=receipt_aliases))

        payment_col = find_column(receipts_df, ['payment method', 'payment type', 'payment'])
        sku_col = find_column(items_df, ['sku', 'item number', 'item code', 'item', 'product', 'product/service'])
        receipt_candidates = ['receipt number', 'receipt no.', 'receipt #', 'receipt']
        items_receipt_col = find_column(items_df, receipt_candidates)
        receipts_receipt_col = find_column(receipts_df, receipt_candidates)

        if payment_col is None:
            pd_stream.error("❌ The receipts file is missing the 'Payment method' column. Check if files were swapped.")
        elif sku_col is None:
            pd_stream.error("❌ The itemized sales file is missing the 'SKU' column. Check if files were swapped.")
        elif items_receipt_col is None or receipts_receipt_col is None:
            pd_stream.error("❌ The receipt number column is missing from one of the uploaded files.")
        else:
            if items_receipt_col != receipts_receipt_col:
                receipts_df = receipts_df.rename(columns={receipts_receipt_col: items_receipt_col})

            cash_mask = receipts_df[payment_col].astype(str).str.lower().str.contains('cash', na=False)
            cash_receipts = receipts_df[cash_mask].copy()
            
            if cash_receipts.empty:
                pd_stream.error("❌ No cash transactions found in the receipts file.")
                unique_payments = receipts_df[payment_col].dropna().unique()
                pd_stream.info(f"The unique payment values discovered were: {list(unique_payments)}")
            else:
                # Merge the dynamically parsed datasets together on the cleaned receipt keys
                merged_df = pd.merge(
                    items_df,
                    cash_receipts[[items_receipt_col, payment_col]],
                    on=items_receipt_col,
                    how='inner'
                )

                if merged_df.empty:
                    pd_stream.warning("⚠️ Found cash sales in the payment log, but couldn't link them to inventory rows. Verify both reports cover the exact same date ranges.")
                else:
                    name_col = find_column(merged_df, ['name', 'item name', 'product name', 'description'])
                    variant_col = find_column(merged_df, ['variant', 'item variant', 'option'])
                    date_col = find_column(merged_df, ['date', 'transaction date', 'sale date', 'order date'])
                    quantity_col = find_column(merged_df, ['quantity', 'qty', 'count'])
                    price_col = find_column(merged_df, ['price (usd)', 'price', 'amount', 'rate', 'unit price', 'sale amount'])

                    if name_col is None or date_col is None or quantity_col is None or price_col is None:
                        pd_stream.error("❌ The itemized sales file is missing one or more required columns: Name, Date, Quantity, or Price.")
                    else:
                        qbo_df = build_qbo_output_df(
                            merged_df,
                            items_receipt_col,
                            sku_col,
                            date_col,
                            quantity_col,
                            price_col,
                            name_col,
                            variant_col,
                        )

                    qbo_df = qbo_df.dropna(subset=['Sales Receipt No.', 'Product/Service'])

                    if qbo_df.empty:
                        pd_stream.warning("⚠️ Line items matching cash transactions don't have valid SKU codes.")
                    else:
                        pd_stream.success(f"🎉 Success! Found {len(qbo_df)} itemized cash lines.")
                        pd_stream.dataframe(qbo_df.head(10), height=420)
                        pd_stream.caption("Scroll inside the table preview to view all output columns.")

                        output_format = pd_stream.radio(
                            "Choose output file type for download",
                            ["XLSX", "CSV"],
                            index=0,
                            horizontal=True,
                        )

                        if output_format == "XLSX":
                            output_data = io.BytesIO()
                            with pd.ExcelWriter(output_data, engine='openpyxl') as writer:
                                qbo_df.to_excel(writer, index=False, sheet_name='Sheet1')
                            download_data = output_data.getvalue()
                            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            extension = "xlsx"
                            label = "📥 Download Clean QuickBooks XLSX"
                        else:
                            download_data = qbo_df.to_csv(index=False).encode('utf-8')
                            mime = "text/csv"
                            extension = "csv"
                            label = "📥 Download Clean QuickBooks CSV"

                        output_filename = f"qbo_cash_import.{extension}"
                        if "-" in itemized_file.name:
                            parts = itemized_file.name.replace(".xlsx", "").split("-")
                            if len(parts) >= 2:
                                output_filename = f"qbo_cash_import_{parts[-2]}_{parts[-1]}.{extension}"

                        pd_stream.download_button(
                            label=label,
                            data=download_data,
                            file_name=output_filename,
                            mime=mime,
                        )

    except Exception as e:
        pd_stream.error(f"❌ An unexpected processing error occurred: {e}")