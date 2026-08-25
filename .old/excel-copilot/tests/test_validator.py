"""Unit tests for ActionValidator."""

import pytest
from excel_copilot.models.protocol import ActionProtocol, ActionIntent, ActionItem
from excel_copilot.models.semantic import WorkbookModel, WorksheetModel, TableModel, ColumnModel
from excel_copilot.validator.action_validator import ActionValidator


@pytest.fixture
def mock_workbook_model():
    cols = [
        ColumnModel(name="Revenue", index=2, data_type="numeric"),
        ColumnModel(name="Cost", index=3, data_type="numeric"),
        ColumnModel(name="Profit", index=4, data_type="formula"),
    ]
    tbl = TableModel(name="Sales", worksheet="Sheet1", range="A1:D10", is_native_table=True, columns=cols, row_count=9)
    ws = WorksheetModel(name="Sheet1", index=0, tables=[tbl])
    return WorkbookModel(filename="test.xlsx", worksheets=[ws])


def test_validator_success(mock_workbook_model):
    protocol = ActionProtocol(
        intent=ActionIntent.MODIFY_WORKBOOK,
        actions=[
            ActionItem(action="add_column", table="Sales", column="Margin %"),
            ActionItem(action="insert_formula", table="Sales", column="Margin %", formula="{Profit}/{Revenue}"),
        ],
    )
    val = ActionValidator.validate(protocol, mock_workbook_model)
    assert val.is_valid is True
    assert len(val.errors) == 0
    assert len(val.action_previews) == 2


def test_validator_missing_sheet(mock_workbook_model):
    protocol = ActionProtocol(
        intent=ActionIntent.MODIFY_WORKBOOK,
        actions=[
            ActionItem(action="rename_sheet", sheet="NonExistentSheet", new_name="NewSheet"),
        ],
    )
    val = ActionValidator.validate(protocol, mock_workbook_model)
    assert val.is_valid is False
    assert any("not found" in e for e in val.errors)
