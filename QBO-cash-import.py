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

def load_excel_with_flexible_headers(uploaded_file, target_column):
    """
    Reads the entire Excel sheet, locates the exact row containing the target_column,
    strips the metadata above it, and perfectly aligns the column headers.
    """
    # Read the file completely without headers or skipping anything
    df = pd.read_excel(uploaded_file, header=None)
    
    header_row_idx = None
    # Look at every row to see where our target column header lives
    for idx, row in df.iterrows():
        row_values = row.astype(str).str.strip().str.lower().tolist()
        if target_column.lower() in row_values:
            header_row_idx = idx
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

if itemized_file and receipts_file:
    pd_stream.markdown("---")
    pd_stream.subheader("Step 2: Verify and Process")
    
    try:
        # Load files completely dynamically regardless of empty rows or metadata length
        items_df = load_excel_with_flexible_headers(itemized_file, 'Receipt number')
        receipts_df = load_excel_with_flexible_headers(receipts_file, 'Receipt number')

        if 'Payment method' not in receipts_df.columns:
            pd_stream.error("❌ The receipts file is missing the 'Payment method' column. Check if files were swapped.")
        elif 'SKU' not in items_df.columns:
            pd_stream.error("❌ The itemized sales file is missing the 'SKU' column. Check if files were swapped.")
        else:
            # Fuzzy match to capture variations of 'Cash' safely
            cash_mask = receipts_df['Payment method'].astype(str).str.lower().str.contains('cash', na=False)
            cash_receipts = receipts_df[cash_mask].copy()
            
            if cash_receipts.empty:
                pd_stream.error("❌ No cash transactions found in the receipts file.")
                unique_payments = receipts_df['Payment method'].dropna().unique()
                pd_stream.info(f"The unique payment values discovered were: {list(unique_payments)}")
            else:
                # Merge the dynamically parsed datasets together on the cleaned receipt keys
                merged_df = pd.merge(
                    items_df,
                    cash_receipts[['Receipt number', 'Payment method']],
                    on='Receipt number',
                    how='inner'
                )

                if merged_df.empty:
                    pd_stream.warning("⚠️ Found cash sales in the payment log, but couldn't link them to inventory rows. Verify both reports cover the exact same date ranges.")
                else:
                    # Clean up variant descriptions
                    merged_df['Variant'] = merged_df['Variant'].fillna('').astype(str)
                    merged_df['Line_Description'] = merged_df.apply(
                        lambda r: f"{r['Name']} ({r['Variant']})" if r['Variant'] != '' else r['Name'], axis=1
                    )

                    # Build final formatted QBO output dataframe
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

                    if qbo_df.empty:
                        pd_stream.warning("⚠️ Line items matching cash transactions don't have valid SKU codes.")
                    else:
                        pd_stream.success(f"🎉 Success! Found {len(qbo_df)} itemized cash lines.")
                        pd_stream.dataframe(qbo_df[['Sales Receipt No.', 'Product/Service', 'Quantity', 'Rate']].head(10))

                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            qbo_df.to_excel(writer, index=False, sheet_name='Sheet1')
                        excel_data = excel_buffer.getvalue()

                        output_filename = "qbo_cash_import.xlsx"
                        if "-" in itemized_file.name:
                            parts = itemized_file.name.replace(".xlsx", "").split("-")
                            if len(parts) >= 2:
                                output_filename = f"qbo_cash_import_{parts[-2]}_{parts[-1]}.xlsx"

                        pd_stream.download_button(
                            label="📥 Download Clean QuickBooks XLSX",
                            data=excel_data,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

    except Exception as e:
        pd_stream.error(f"❌ An unexpected processing error occurred: {e}")