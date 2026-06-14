import importlib.util
import os
from pathlib import Path
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


def test_load_excel_with_flexible_headers_finds_header_row_with_metadata_rows():
    raw_df = pd.DataFrame([
        ['meta', 'meta', 'meta'],
        ['more metadata', 'more metadata', 'more metadata'],
        ['', '', ''],
        ['Receipt number', 'Date', 'SKU'],
        [4, '2024-01-04', 'SKU4'],
    ])

    with patch.object(qbo_cash_import.pd, 'read_excel', return_value=raw_df):
        result = qbo_cash_import.load_excel_with_flexible_headers('fake.xlsx', 'Receipt number')

    assert list(result.columns) == ['Receipt number', 'Date', 'SKU']
    assert result.shape == (1, 3)
    assert result.iloc[0].to_dict() == {
        'Receipt number': 4,
        'Date': '2024-01-04',
        'SKU': 'SKU4',
    }


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

    assert list(qbo_df.columns) == [
        'Sales Receipt No.',
        'Sales Receipt Date',
        'Customer',
        'Product/Service',
        'Description',
        'SKU',
        'Qty',
        'Rate',
        'Total',
        'Deposit To',
        'Payment Method',
        'Ref No.',
    ]
    assert qbo_df['Sales Receipt No.'].tolist() == ['1000-1', '1001-2']
    assert qbo_df['Sales Receipt Date'].tolist() == ['01/01/2024', '01/02/2024']
    assert qbo_df['SKU'].tolist() == ['SKU1', 'SKU2']
    assert qbo_df['Product/Service'].tolist() == ['SKU1', 'SKU2']
    assert qbo_df['Description'].tolist() == ['Widget', 'Gadget']
    assert qbo_df['Ref No.'].tolist() == ['1', '2']
    assert qbo_df['Total'].tolist() == [20.0, 46.5]


def test_build_qbo_output_df_includes_variant_in_description():
    merged_df = pd.DataFrame({
        'receipt number': [1],
        'date': ['2024-01-01'],
        'sku': ['SKU1'],
        'name': ['Widget'],
        'variant': ['Large'],
        'quantity': [1],
        'rate': [10.00],
    })

    qbo_df = qbo_cash_import.build_qbo_output_df(
        merged_df,
        items_receipt_col='receipt number',
        sku_col='sku',
        date_col='date',
        quantity_col='quantity',
        price_col='rate',
        name_col='name',
        variant_col='variant',
    )

    assert qbo_df['Product/Service'].tolist() == ['SKU1']
    assert qbo_df['Description'].tolist() == ['Widget (Large)']
    assert qbo_df['Sales Receipt No.'].tolist() == ['1000-1']
    assert qbo_df['Ref No.'].tolist() == ['1']


def test_build_qbo_output_df_uses_custom_prefix_for_sales_receipt_no():
    merged_df = pd.DataFrame({
        'receipt number': [10, 20],
        'date': ['2024-01-01', '2024-01-02'],
        'sku': ['SKU1', 'SKU2'],
        'name': ['Widget', 'Gadget'],
        'quantity': [1, 1],
        'rate': [5.00, 7.00],
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
        receipt_prefix=5000,
    )

    assert qbo_df['Sales Receipt No.'].tolist() == ['5000-10', '5001-20']
    assert qbo_df['SKU'].tolist() == ['SKU1', 'SKU2']
    assert qbo_df['Product/Service'].tolist() == ['SKU1', 'SKU2']
    assert qbo_df['Description'].tolist() == ['Widget', 'Gadget']
    assert qbo_df['Ref No.'].tolist() == ['10', '20']


def test_build_reference_map_reads_mapping_columns():
    df = pd.DataFrame({
        'SKU': ['SKU1', 'SKU2'],
        'QBO SKU': ['QBO1', 'QBO2'],
        'Product/Service': ['Shirt', 'Soda'],
    })

    reference_map = qbo_cash_import.build_reference_map(qbo_cash_import.standardize_dataframe_columns(df))

    assert reference_map['sku1']['sku'] == 'QBO1'
    assert reference_map['sku1']['product_service'] == 'Shirt'
    assert reference_map['sku2']['sku'] == 'QBO2'
    assert reference_map['sku2']['product_service'] == 'Soda'


def test_build_reference_map_accepts_sku_only_mapping_rows():
    df = pd.DataFrame({
        'SKU': ['SKU1', 'SKU2'],
    })

    reference_map = qbo_cash_import.build_reference_map(qbo_cash_import.standardize_dataframe_columns(df))

    assert reference_map['sku1']['sku'] == 'SKU1'
    assert reference_map['sku1']['product_service'] == 'SKU1'
    assert reference_map['sku2']['sku'] == 'SKU2'
    assert reference_map['sku2']['product_service'] == 'SKU2'


def test_build_reference_map_accepts_category_and_sku_columns():
    df = pd.DataFrame({
        'Category': ['Apparel', 'Drink'],
        'SKU': ['SKU1', 'SKU2'],
        'QBO SKU': ['QBO1', 'QBO2'],
        'Product/Service': ['Shirt', 'Soda'],
    })

    reference_map = qbo_cash_import.build_reference_map(qbo_cash_import.standardize_dataframe_columns(df))

    assert reference_map['sku1']['sku'] == 'QBO1'
    assert reference_map['sku1']['product_service'] == 'Shirt'
    assert reference_map['sku2']['sku'] == 'QBO2'
    assert reference_map['sku2']['product_service'] == 'Soda'


def test_build_reference_map_reads_product_service_name_key():
    df = pd.DataFrame({
        'Product/Service Name': ['Accessories:Blanket blue white', 'Accessories:JCFC Beanie'],
        'SKU': ['blkt-blue-white', 'JCFC-beanie'],
    })

    reference_map = qbo_cash_import.build_reference_map(qbo_cash_import.standardize_dataframe_columns(df))

    assert reference_map['blkt-blue-white']['sku'] == 'blkt-blue-white'
    assert reference_map['blkt-blue-white']['product_service'] == 'Accessories:Blanket blue white'
    assert reference_map['jcfc-beanie']['sku'] == 'JCFC-beanie'
    assert reference_map['jcfc-beanie']['product_service'] == 'Accessories:JCFC Beanie'


def test_build_reference_map_loads_actual_xls_product_service_sample():
    sample_path = Path(__file__).resolve().parent / 'QBO-Product-Service-List-sample.xls'
    reference_df = qbo_cash_import.load_reference_file(sample_path)
    reference_map = qbo_cash_import.build_reference_map(reference_df)

    assert 'blkt-blue-white' in reference_map
    assert reference_map['blkt-blue-white']['sku'] == 'blkt-blue-white'
    assert reference_map['blkt-blue-white']['product_service'] == 'Accessories:Blanket blue white'
    assert 'jcfc-beanie' in reference_map
    assert reference_map['jcfc-beanie']['sku'] == 'JCFC-beanie'
    assert reference_map['jcfc-beanie']['product_service'] == 'Accessories:JCFC Beanie'


def test_apply_reference_mapping_creates_mapped_columns():
    merged_df = pd.DataFrame({
        'category': ['Apparel', 'Drink'],
        'sku': ['SKU1', 'SKU2'],
        'receipt number': [1, 2],
    })
    reference_map = {
        'sku1': {'sku': 'QBO1', 'product_service': 'Shirt'},
        'sku2': {'sku': 'QBO2', 'product_service': 'Soda'},
    }

    result = qbo_cash_import.apply_reference_mapping(
        merged_df,
        reference_map,
        sku_col='sku',
    )

    assert result['reference_lookup_key'].tolist() == ['sku1', 'sku2']
    assert result['mapped_sku'].tolist() == ['QBO1', 'QBO2']
    assert result['mapped_product_service'].tolist() == ['Shirt', 'Soda']


def test_apply_reference_mapping_falls_back_to_sku_only_when_no_category_available():
    merged_df = pd.DataFrame({
        'sku': ['SKU1', 'SKU2'],
        'receipt number': [1, 2],
    })
    reference_map = {
        'sku1': {'sku': 'QBO1', 'product_service': 'Widget'},
        'sku2': {'sku': 'QBO2', 'product_service': 'Gadget'},
    }

    result = qbo_cash_import.apply_reference_mapping(
        merged_df,
        reference_map,
        sku_col='sku',
    )

    assert result['reference_lookup_key'].tolist() == ['sku1', 'sku2']
    assert result['mapped_sku'].tolist() == ['QBO1', 'QBO2']
    assert result['mapped_product_service'].tolist() == ['Widget', 'Gadget']


def test_build_qbo_output_df_uses_mapped_values_when_available():
    merged_df = pd.DataFrame({
        'receipt number': [1],
        'date': ['2024-01-01'],
        'sku': ['SKU1'],
        'name': ['Widget'],
        'variant': ['Large'],
        'mapped_sku': ['QBO1'],
        'mapped_product_service': ['QBO Widget'],
        'quantity': [1],
        'rate': [10.00],
    })

    qbo_df = qbo_cash_import.build_qbo_output_df(
        merged_df,
        items_receipt_col='receipt number',
        sku_col='sku',
        date_col='date',
        quantity_col='quantity',
        price_col='rate',
        name_col='name',
        variant_col='variant',
        mapped_sku_col='mapped_sku',
        mapped_product_service_col='mapped_product_service',
    )

    assert qbo_df['SKU'].tolist() == ['QBO1']
    assert qbo_df['Product/Service'].tolist() == ['QBO Widget']
    assert qbo_df['Description'].tolist() == ['Widget (Large)']
