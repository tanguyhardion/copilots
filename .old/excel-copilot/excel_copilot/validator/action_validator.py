"""Action Validator: Verifies requested actions against the Workbook Semantic Model before execution."""

import re
from typing import List, Optional
from excel_copilot.models.semantic import WorkbookModel
from excel_copilot.models.protocol import ActionProtocol, ActionItem, ValidationResult, ActionIntent


class ActionValidator:
    """Validates action protocol objects against workbook semantic models."""

    SUPPORTED_ACTIONS = {
        # Workbook
        "rename_sheet", "duplicate_sheet", "delete_sheet", "create_sheet",
        # Search
        "find_sheet", "find_table", "find_column", "find_formula", "search_text",
        # Data
        "add_column", "remove_column", "rename_column", "insert_rows", "delete_rows", "replace_values", "write_dataframe",
        # Formula
        "insert_formula", "replace_formula", "fill_formula", "add_formula_column",
        # Formatting
        "autofit_columns", "apply_style", "freeze_panes", "conditional_format",
        # Charts
        "create_chart", "update_chart", "delete_chart",
    }

    @classmethod
    def validate(cls, protocol: ActionProtocol, model: Optional[WorkbookModel]) -> ValidationResult:
        """Validate protocol actions against workbook model."""
        errors: List[str] = []
        warnings: List[str] = []
        previews: List[str] = []

        if protocol.intent == ActionIntent.NO_ACTION or not protocol.actions:
            return ValidationResult(
                is_valid=True,
                errors=[],
                warnings=["No workbook actions found in input protocol block."],
                action_previews=["No actions to execute."],
            )

        if not model:
            return ValidationResult(
                is_valid=False,
                errors=["No workbook loaded. Please open an Excel workbook before executing actions."],
                warnings=[],
                action_previews=[],
            )

        for idx, item in enumerate(protocol.actions, start=1):
            action_name = item.action.lower().strip()

            # 1. Check supported action name
            if action_name not in cls.SUPPORTED_ACTIONS:
                errors.append(f"Action #{idx}: Unsupported action '{item.action}'.")
                continue

            # 2. Per-action validation logic
            cls._validate_action(idx, item, model, errors, warnings, previews)

        is_valid = len(errors) == 0

        return ValidationResult(
            is_valid=is_valid,
            errors=errors,
            warnings=warnings,
            action_previews=previews,
        )

    @classmethod
    def _validate_action(
        cls,
        idx: int,
        item: ActionItem,
        model: WorkbookModel,
        errors: List[str],
        warnings: List[str],
        previews: List[str],
    ) -> None:
        act = item.action.lower().strip()

        # Sheet check helper
        target_sheet = item.sheet
        target_table = item.table

        # Resolve table if provided
        table_obj = model.get_table(target_table) if target_table else None
        sheet_obj = model.get_sheet(target_sheet) if target_sheet else None

        if table_obj and not sheet_obj:
            sheet_obj = model.get_sheet(table_obj.worksheet)

        # Action specific rules
        if act == "create_sheet":
            sheet_name = item.new_name or item.sheet
            if not sheet_name:
                errors.append(f"Action #{idx} (create_sheet): Target sheet name is missing.")
            elif model.get_sheet(sheet_name):
                warnings.append(f"Action #{idx} (create_sheet): Sheet '{sheet_name}' already exists.")
            else:
                previews.append(f"Create new worksheet '{sheet_name}'")

        elif act == "rename_sheet":
            if not target_sheet:
                errors.append(f"Action #{idx} (rename_sheet): Missing target sheet name.")
            elif not sheet_obj:
                errors.append(f"Action #{idx} (rename_sheet): Sheet '{target_sheet}' not found in workbook.")
            elif not item.new_name:
                errors.append(f"Action #{idx} (rename_sheet): Missing new_name parameter.")
            else:
                previews.append(f"Rename worksheet '{target_sheet}' to '{item.new_name}'")

        elif act in ("duplicate_sheet", "delete_sheet"):
            if not target_sheet:
                errors.append(f"Action #{idx} ({act}): Missing target sheet name.")
            elif not sheet_obj:
                errors.append(f"Action #{idx} ({act}): Sheet '{target_sheet}' not found in workbook.")
            else:
                previews.append(f"{act.replace('_', ' ').capitalize()} '{target_sheet}'")

        elif act in ("add_column", "add_formula_column"):
            col_name = item.column or item.new_name
            if not col_name:
                errors.append(f"Action #{idx} ({act}): Missing column name.")
            elif not target_table and not target_sheet:
                errors.append(f"Action #{idx} ({act}): Must specify target table or sheet.")
            else:
                target_desc = f"table '{target_table}'" if target_table else f"sheet '{target_sheet}'"
                previews.append(f"Add column '{col_name}' to {target_desc}")

        elif act == "remove_column":
            col_name = item.column
            if not col_name:
                errors.append(f"Action #{idx} (remove_column): Missing column name.")
            elif target_table and table_obj:
                cols = [c.name.lower() for c in table_obj.columns]
                if col_name.lower() not in cols:
                    warnings.append(f"Action #{idx} (remove_column): Column '{col_name}' may not exist in table '{target_table}'.")
                previews.append(f"Remove column '{col_name}' from table '{target_table}'")
            else:
                previews.append(f"Remove column '{col_name}'")

        elif act == "rename_column":
            if not item.column or not item.new_name:
                errors.append(f"Action #{idx} (rename_column): Requires both 'column' and 'new_name'.")
            else:
                previews.append(f"Rename column '{item.column}' to '{item.new_name}'")

        elif act in ("insert_formula", "replace_formula", "fill_formula"):
            if not item.formula:
                errors.append(f"Action #{idx} ({act}): Formula string is missing.")
            else:
                # Check semantic variables inside formula {VarName}
                vars_found = re.findall(r"\{([a-zA-Z0-9_\s%]+)\}", item.formula)
                if vars_found and table_obj:
                    table_cols = [c.name.lower() for c in table_obj.columns]
                    for v in vars_found:
                        if v.lower() not in table_cols:
                            warnings.append(f"Action #{idx} ({act}): Semantic variable '{{{v}}}' not found in table '{target_table}' columns.")

                target_dest = f"column '{item.column}'" if item.column else f"table '{target_table}'"
                previews.append(f"Insert formula '{item.formula}' into {target_dest}")

        elif act in ("create_chart", "update_chart"):
            chart_name = item.chart or "Chart"
            previews.append(f"{act.replace('_', ' ').capitalize()} '{chart_name}'")

        elif act == "delete_chart":
            if not item.chart:
                errors.append(f"Action #{idx} (delete_chart): Missing chart name or title.")
            else:
                previews.append(f"Delete chart '{item.chart}'")

        elif act.startswith("find_") or act == "search_text":
            query = item.search_query or item.column or item.sheet or ""
            previews.append(f"Search workbook for '{query}' ({act})")

        else:
            previews.append(f"Execute {act} on target sheet/table ({target_sheet or target_table or 'Workbook'})")
