import importlib.util
import os
from unittest.mock import patch

import pandas as pd
import pytest

MODULE_PATH = os.path.join(os.path.dirname(__file__), '..', 'QBO-cash-import.py')
SPEC = importlib.util.spec_from_file_location('qbo_cash_import', MODULE_PATH)
qbo_cash_import = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(qbo_cash_import)


def test_load_excel_with_flexible_headers_finds_header_row_and_strips_metadata():
    raw_df = pd.DataFrame([
        ['Ignore this', 'Ignore this', 'Ignore this'],
        ['Receipt number', 'Date', 'SKU'],
        [1, '2024-01-01', 'SKU1'],
    ])

    with patch.object(qbo_cash_import.pd, 'read_excel', return_value=raw_df) as mock_read_excel:
        result = qbo_cash_import.load_excel_with_flexible_headers('fake.xlsx', 'Receipt number')

    assert list(result.columns) == ['Receipt number', 'Date', 'SKU']
    assert result.shape == (1, 3)
    assert result.iloc[0].to_dict() == {
        'Receipt number': 1,
        'Date': '2024-01-01',
        'SKU': 'SKU1',
    }
    mock_read_excel.assert_called_once_with('fake.xlsx', header=None)


def test_load_excel_with_flexible_headers_normalizes_whitespace_and_case():
    raw_df = pd.DataFrame([
        ['meta', 'meta', 'meta'],
        ['  receipt NUMBER  ', ' Date ', '  sKu  '],
        [2, '2024-01-02', 'SKU2'],
    ])

    with patch.object(qbo_cash_import.pd, 'read_excel', return_value=raw_df):
        result = qbo_cash_import.load_excel_with_flexible_headers('fake.xlsx', 'Receipt number')

    assert list(result.columns) == ['receipt NUMBER', 'Date', 'sKu']
    assert result.iloc[0]['receipt NUMBER'] == 2


def test_load_excel_with_flexible_headers_raises_when_target_column_not_found():
    raw_df = pd.DataFrame([
        ['meta', 'meta'],
        ['Date', 'SKU'],
        ['2024-01-01', 'SKU1'],
    ])

    with patch.object(qbo_cash_import.pd, 'read_excel', return_value=raw_df):
        with pytest.raises(ValueError, match="Could not find the expected column 'Receipt number'"):
            qbo_cash_import.load_excel_with_flexible_headers('fake.xlsx', 'Receipt number')


def test_load_excel_with_flexible_headers_accepts_receipt_number_aliases():
    raw_df = pd.DataFrame([
        ['meta', 'meta', 'meta'],
        ['Receipt #', 'Date', 'SKU'],
        [3, '2024-01-03', 'SKU3'],
    ])

    with patch.object(qbo_cash_import.pd, 'read_excel', return_value=raw_df):
        result = qbo_cash_import.load_excel_with_flexible_headers(
            'fake.xlsx',
            'Receipt number',
            aliases=['Receipt #', 'Receipt no.'],
        )

    assert list(result.columns) == ['Receipt #', 'Date', 'SKU']
    assert result.iloc[0]['Receipt #'] == 3


def test_standardize_dataframe_columns_lowercases_and_strips_headers():
    df = pd.DataFrame({
        '  Receipt number  ': [1],
        ' Payment Method ': ['Cash'],
        'SKU': ['SKU1'],
    })

    result = qbo_cash_import.standardize_dataframe_columns(df)

    assert list(result.columns) == ['receipt number', 'payment method', 'sku']


def test_find_column_matches_synonyms():
    df = pd.DataFrame(columns=['payment type', 'receipt #'])

    assert qbo_cash_import.find_column(df, ['payment method', 'payment type', 'payment']) == 'payment type'
    assert qbo_cash_import.find_column(df, ['receipt number', 'receipt #']) == 'receipt #'


def test_build_qbo_output_df_formats_date_and_includes_total():
    merged_df = pd.DataFrame({
        'receipt number': [1, 2],
        'date': ['2024-01-01 12:34:56', '2024-01-02 00:00:00'],
        'sku': ['SKU1', 'SKU2'],
        'name': ['Widget', 'Gadget'],
        'quantity': [2, 3],
        'rate': [10.00, 15.50],
    })

    qbo_df = qbo_cash_import.build_qbo_output_df(
        merged_df,
        items_receipt_col='receipt number',
        sku_col='sku',
        date_col='date',
        quantity_col='quantity',
        price_col='rate',
        name_col='name',
        variant_col=None,
    )

    assert list(qbo_df.columns)[:8] == [
        'Sales Receipt No.',
        'Transaction Date',
        'Customer',
        'Product/Service',
        'Description',
        'Quantity',
        'Rate',
        'Total',
    ]
    assert qbo_df['Transaction Date'].tolist() == ['2024-01-01', '2024-01-02']
    assert qbo_df['Total'].tolist() == [20.0, 46.5]
