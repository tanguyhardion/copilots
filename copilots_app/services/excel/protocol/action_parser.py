"""Action Parser: Extracts and validates excel-action JSON blocks from LLM markdown responses."""

import re
import json
from typing import Tuple, Optional
from copilots_app.services.excel.models.protocol import ActionProtocol, ActionIntent, ActionItem


class ActionParser:
    """Parses LLM responses to extract structured ActionProtocol instances."""

    @classmethod
    def parse_response(cls, text: str) -> ActionProtocol:
        """Extract excel-action block from LLM output string."""
        if not text or not text.strip():
            return ActionProtocol(intent=ActionIntent.NO_ACTION, explanation="Empty input")

        # 1. Regex search for ```excel-action ... ``` or ```json ... ``` containing intent
        pattern = r"```(?:excel-action|json)?\s*(\{.*?\})\s*```"
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)

        raw_json_str = None
        if match:
            raw_json_str = match.group(1)
        else:
            # Fallback search for bare JSON block starting with {"intent":
            bare_match = re.search(r"(\{\s*\"intent\"\s*:.*?\})", text, re.DOTALL | re.IGNORECASE)
            if bare_match:
                raw_json_str = bare_match.group(1)

        if not raw_json_str:
            return ActionProtocol(
                intent=ActionIntent.NO_ACTION,
                actions=[],
                explanation=text.strip(),
            )

        # Clean JSON text
        sanitized_json = cls._sanitize_json_string(raw_json_str)

        try:
            data = json.loads(sanitized_json)
            intent_val = data.get("intent", "modify_workbook")
            actions_raw = data.get("actions", [])

            actions_list = []
            for act in actions_raw:
                if isinstance(act, dict):
                    action_name = act.get("action", "")
                    sheet = act.get("sheet") or act.get("worksheet")
                    table = act.get("table")
                    column = act.get("column")
                    new_name = act.get("new_name") or act.get("name")
                    formula = act.get("formula")
                    values = act.get("values")
                    search_query = act.get("search_query") or act.get("query")
                    chart = act.get("chart") or act.get("chart_title")
                    chart_type = act.get("chart_type") or act.get("type")
                    data_range = act.get("data_range") or act.get("range")

                    # Extra custom params
                    known_keys = {
                        "action", "sheet", "worksheet", "table", "column",
                        "new_name", "name", "formula", "values", "search_query",
                        "query", "chart", "chart_title", "chart_type", "type",
                        "data_range", "range",
                    }
                    extra_params = {k: v for k, v in act.items() if k not in known_keys}

                    actions_list.append(
                        ActionItem(
                            action=action_name,
                            sheet=sheet,
                            table=table,
                            column=column,
                            new_name=new_name,
                            formula=formula,
                            values=values,
                            search_query=search_query,
                            chart=chart,
                            chart_type=chart_type,
                            data_range=data_range,
                            params=extra_params,
                        )
                    )

            # Strip protocol block from explanation text
            explanation = text.replace(match.group(0) if match else raw_json_str, "").strip()

            return ActionProtocol(
                intent=ActionIntent(intent_val) if intent_val in ActionIntent.__members__.values() else ActionIntent.MODIFY_WORKBOOK,
                actions=actions_list,
                explanation=explanation,
            )

        except Exception as err:
            return ActionProtocol(
                intent=ActionIntent.NO_ACTION,
                actions=[],
                explanation=f"Error parsing action protocol block: {str(err)}\n\nOriginal text:\n{text}",
            )

    @classmethod
    def _sanitize_json_string(cls, raw: str) -> str:
        """Fix common JSON formatting errors like trailing commas or unescaped quotes."""
        # Remove trailing commas before closing braces/brackets
        s = re.sub(r",\s*([\}\]])", r"\1", raw)
        return s
