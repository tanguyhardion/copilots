"""Unit tests for ActionParser."""

import pytest
from excel_copilot.protocol.action_parser import ActionParser
from excel_copilot.models.protocol import ActionIntent


def test_parse_excel_action_block():
    text = """
I will add a Margin % column and update the dashboard chart.

```excel-action
{
  "intent": "modify_workbook",
  "actions": [
    {
      "action": "add_column",
      "table": "SalesData",
      "column": "Margin %"
    },
    {
      "action": "insert_formula",
      "table": "SalesData",
      "column": "Margin %",
      "formula": "{Profit}/{Revenue}"
    }
  ]
}
```
"""
    protocol = ActionParser.parse_response(text)
    assert protocol.intent == ActionIntent.MODIFY_WORKBOOK
    assert len(protocol.actions) == 2
    assert protocol.actions[0].action == "add_column"
    assert protocol.actions[0].table == "SalesData"
    assert protocol.actions[0].column == "Margin %"
    assert protocol.actions[1].formula == "{Profit}/{Revenue}"


def test_parse_no_action_conversation():
    text = "The total revenue in January was $4,500 based on your Sales table."
    protocol = ActionParser.parse_response(text)
    assert protocol.intent == ActionIntent.NO_ACTION
    assert len(protocol.actions) == 0
    assert protocol.explanation == text
