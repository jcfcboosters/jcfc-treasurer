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
