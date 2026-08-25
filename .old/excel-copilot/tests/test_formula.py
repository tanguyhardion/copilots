"""Unit tests for FormulaTranslator."""

import pytest
from excel_copilot.utils.formula_translator import FormulaTranslator


def test_formula_table_syntax():
    formula = FormulaTranslator.translate(
        semantic_formula="{Profit}/{Revenue}",
        is_native_table=True,
    )
    assert formula == "=[@Profit]/[@Revenue]"


def test_formula_cell_syntax():
    col_map = {"profit": "D", "revenue": "C"}
    formula = FormulaTranslator.translate(
        semantic_formula="{Profit}/{Revenue}",
        col_name_to_letter=col_map,
        row_index=2,
        is_native_table=False,
    )
    assert formula == "=D2/C2"
