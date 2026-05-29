import importlib.util
import os
from unittest.mock import patch

import pandas as pd

MODULE_PATH = os.path.join(os.path.dirname(__file__), '..', 'QBO-cash-import.py')
SPEC = importlib.util.spec_from_file_location('qbo_cash_import', MODULE_PATH)
qbo_cash_import = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(qbo_cash_import)


def test_build_qbo_dataframe_returns_none_when_no_cash_payments():
    items_df = pd.DataFrame({
        'Receipt number': [1],
        'Date': ['2024-01-01'],
        'SKU': ['SKU1'],
        'Name': ['Item A'],
        'Variant': [None],
        'Quantity': [1],
        'Price (USD)': [10.00],
    })
    receipts_df = pd.DataFrame({
        'Receipt number': [1],
        'Payment method': ['Credit'],
    })

    result = qbo_cash_import.build_qbo_dataframe(items_df, receipts_df)
    assert result is None


def test_build_qbo_dataframe_returns_none_when_receipts_do_not_match():
    items_df = pd.DataFrame({
        'Receipt number': [1],
        'Date': ['2024-01-01'],
        'SKU': ['SKU1'],
        'Name': ['Item A'],
        'Variant': [None],
        'Quantity': [1],
        'Price (USD)': [10.00],
    })
    receipts_df = pd.DataFrame({
        'Receipt number': [2],
        'Payment method': ['Cash'],
    })

    result = qbo_cash_import.build_qbo_dataframe(items_df, receipts_df)
    assert result is None


def test_build_qbo_dataframe_creates_expected_qbo_rows():
    items_df = pd.DataFrame({
        'Receipt number': [1, 1],
        'Date': ['2024-01-01', '2024-01-01'],
        'SKU': ['SKU1', 'SKU2'],
        'Name': ['Item A', 'Item B'],
        'Variant': ['Red', None],
        'Quantity': [1, 2],
        'Price (USD)': [10.00, 5.00],
    })
    receipts_df = pd.DataFrame({
        'Receipt number': [1],
        'Payment method': ['Cash'],
    })

    result = qbo_cash_import.build_qbo_dataframe(items_df, receipts_df)

    assert result is not None
    assert list(result.columns) == [
        'Sales Receipt No.',
        'Transaction Date',
        'Customer',
        'Product/Service',
        'Description',
        'Quantity',
        'Rate',
        'Deposit To',
        'Payment Method',
    ]
    assert result.loc[0, 'Description'] == 'Item A (Red)'
    assert result.loc[1, 'Description'] == 'Item B'
    assert all(result['Customer'] == 'Cash Customer')
    assert all(result['Deposit To'] == 'Undeposited Funds')
    assert all(result['Payment Method'] == 'Cash')


@patch.object(qbo_cash_import, 'safe_read_excel')
def test_parse_reports_uses_skiprows_and_strips_columns(mock_safe_read_excel):
    raw_itemized = pd.DataFrame({'  Receipt number  ': [1], '  Name  ': ['Item A']})
    raw_receipts = pd.DataFrame({'  Receipt number  ': [1], '  Payment method  ': ['Cash']})
    mock_safe_read_excel.side_effect = [raw_itemized, raw_receipts]

    items_df, receipts_df = qbo_cash_import.parse_reports('itemized.xlsx', 'receipts.xlsx')

    assert list(items_df.columns) == ['Receipt number', 'Name']
    assert list(receipts_df.columns) == ['Receipt number', 'Payment method']
    assert mock_safe_read_excel.call_count == 2
    mock_safe_read_excel.assert_any_call('itemized.xlsx', skiprows=6)
    mock_safe_read_excel.assert_any_call('receipts.xlsx', skiprows=16)
