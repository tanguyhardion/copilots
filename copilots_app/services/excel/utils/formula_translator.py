"""Formula Translator: Resolves semantic variable formulas into Excel Table syntax or Cell references."""

import re
from typing import Dict, Optional
import openpyxl.utils


class FormulaTranslator:
    """Translates semantic formulas like `{Profit}/{Revenue}` into Excel formulas."""

    @classmethod
    def translate(
        cls,
        semantic_formula: str,
        col_name_to_letter: Optional[Dict[str, str]] = None,
        row_index: Optional[int] = None,
        is_native_table: bool = True,
    ) -> str:
        """Translate a semantic formula template.

        Args:
            semantic_formula: Formula with {Column} placeholders, e.g. "{Profit}/{Revenue}" or "SUM({Revenue})"
            col_name_to_letter: Map of column header name (lowercased) -> column letter (e.g. 'profit': 'D')
            row_index: 1-based target row number for cell-reference formulas (e.g. 2 for D2/C2)
            is_native_table: If True, uses Excel Table structured syntax `[@Profit]` when possible
        """
        if not semantic_formula:
            return ""

        formula = semantic_formula.strip()
        if not formula.startswith("="):
            formula = "=" + formula

        # Find all {Column} placeholders
        variables = re.findall(r"\{([a-zA-Z0-9_\s%]+)\}", formula)

        if not variables:
            return formula

        translated = formula
        col_map = {k.lower(): v for k, v in (col_name_to_letter or {}).items()}

        for var in set(variables):
            var_clean = var.strip()
            var_key = var_clean.lower()
            placeholder = f"{{{var}}}"

            if is_native_table and not row_index:
                # Table structured reference syntax: [@ColumnName]
                replacement = f"[@{var_clean}]"
            elif var_key in col_map and row_index:
                # Standard cell reference syntax: D2
                col_letter = col_map[var_key]
                replacement = f"{col_letter}{row_index}"
            else:
                # Default fallback structured reference
                replacement = f"[@{var_clean}]"

            translated = translated.replace(placeholder, replacement)

        return translated
