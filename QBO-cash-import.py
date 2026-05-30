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

def load_excel_safely(uploaded_file, target_column):
    """
    Dynamically scans the Excel file to find which row actually contains 
    the header column, bypassing any variable number of metadata rows at the top.
    """
    # Read the first 30 rows without skipping to locate the headers
    df_scan = pd.read_excel(uploaded_file, header=None, nrows=30)
    
    header_row_index = None
    for idx, row in df_scan.iterrows():
        # Check if our target column is anywhere in this row's values
        row_str_values = row.astype(str).str.strip().tolist()
        if target_column in row_str_values:
            header_row_index = idx
            break
            
    if header_row_index is None:
        raise ValueError(f"Could not find the expected column '{target_column}' anywhere in the file. Please verify this is the correct report.")
        
    # Re-read the file starting exactly from the discovered header row
    df = pd.read_excel(uploaded_file, skiprows=header_row_index)
    df.columns = df.columns.str.strip()
    return df

if itemized_file and receipts_file:
    pd_stream.markdown("---")
    pd_stream.subheader("Step 2: Verify and Process")
    
    try:
        # Load files using the auto-detect header logic
        items_df = load_excel_safely(itemized_file, 'Receipt number')
        receipts_df = load_excel_safely(receipts_file, 'Receipt number')

        # Check for other required columns to fail gracefully with a helpful message
        if 'Payment method' not in receipts_df.columns:
            pd_stream.error("❌ The receipts file is missing the 'Payment method' column. Check if files were swapped.")
        elif 'SKU' not in items_df.columns:
            pd_stream.error("❌ The itemized sales file is missing the 'SKU' column. Check if files were swapped.")
        else:
            # Isolate cash transactions safely
            cash_receipts = receipts_df[receipts_df['Payment method'].astype(str).str.lower() == 'cash'].copy()
            
            if cash_receipts.empty:
                pd_stream.error("❌ No cash transactions found in the receipts file.")
            else:
                # Merge datasets
                merged_df = pd.merge(
                    items_df,
                    cash_receipts[['Receipt number', 'Payment method']],
                    on='Receipt number',
                    how='inner'
                )

                if merged_df.empty:
                    pd_stream.warning("⚠️ No matching lines found. Make sure the date ranges of both files overlap perfectly.")
                else:
                    # Clean up descriptions
                    merged_df['Variant'] = merged_df['Variant'].fillna('').astype(str)
                    merged_df['Line_Description'] = merged_df.apply(
                        lambda r: f"{r['Name']} ({r['Variant']})" if r['Variant'] != '' else r['Name'], axis=1
                    )

                    # Build QBO export DataFrame
                    qbo_df = pd.DataFrame()
                    qbo_df['Sales Receipt No.'] = merged_df['Receipt number']
                    qbo_df['Transaction Date'] = merged_df['Date']
                    qbo_df['Customer'] = "Cash Customer"
                    qbo_df['Product/Service'] = merged_df['SKU']
                    qbo_df['Description'] = merged_df['Line_Description']
                    qbo_df['Quantity'] = merged_df['Quantity']
                    qbo_df['Rate'] = merged_df['Price (USD)']
                    qbo_df['Deposit To'] = "Undeposited Funds"
                    qbo_df['Payment Method'] = "Cash"

                    qbo_df = qbo_df.dropna(subset=['Sales Receipt No.', 'Product/Service'])

                    pd_stream.success(f"🎉 Success! Found {len(qbo_df)} itemized cash lines.")
                    pd_stream.dataframe(qbo_df[['Sales Receipt No.', 'Product/Service', 'Quantity', 'Rate']].head(10))

                    # Prepare download
                    csv_buffer = io.StringIO()
                    qbo_df.to_csv(csv_buffer, index=False)
                    csv_data = csv_buffer.getvalue()

                    output_filename = "qbo_cash_import.csv"
                    if "-" in itemized_file.name:
                        parts = itemized_file.name.replace(".xlsx", "").split("-")
                        if len(parts) >= 2:
                            output_filename = f"qbo_cash_import_{parts[-2]}_{parts[-1]}.csv"

                    pd_stream.download_button(
                        label="📥 Download Clean QuickBooks CSV",
                        data=csv_data,
                        file_name=output_filename,
                        mime="text/csv"
                    )

    except Exception as e:
        pd_stream.error(f"❌ An unexpected processing error occurred: {e}")