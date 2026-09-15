"""
Email Copilot package using pywin32 Outlook MAPI automation.
"""

from copilots_app.services.email_copilot.dsl_parser import (
    parse_email_dsl,
    EmailCommand,
)
from copilots_app.services.email_copilot.outlook_connector import (
    OutlookConnector,
    EmailItemSummary,
    EmailExecutionResult,
)

__all__ = [
    "parse_email_dsl",
    "EmailCommand",
    "OutlookConnector",
    "EmailItemSummary",
    "EmailExecutionResult",
]
