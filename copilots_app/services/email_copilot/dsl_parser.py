"""
Email DSL Parser: parses user/LLM generated email queries into structured commands.

Supported syntax examples:
  SEARCH query="quarterly roadmap" limit=10 max_chars=300
  SEARCH folder=inbox sender="alice@example.com" unread_only=true
  SEARCH subject="Project Apollo" after="2026-01-01" has_attachment=true
  SUMMARIZE_UNREAD folder=inbox limit=15 max_chars=250
"""

from __future__ import annotations

import re
import shlex
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any


@dataclass
class EmailCommand:
    action: str  # 'SEARCH' or 'SUMMARIZE_UNREAD'
    folder: str = "inbox"  # 'inbox', 'sent', 'junk', 'archive', or custom folder name
    query: Optional[str] = None  # Free-text search or keyword
    sender: Optional[str] = None
    subject: Optional[str] = None
    after: Optional[str] = None  # ISO date YYYY-MM-DD or relative
    before: Optional[str] = None  # ISO date YYYY-MM-DD or relative
    unread_only: bool = False
    has_attachment: Optional[bool] = None
    importance: Optional[str] = None  # 'high', 'normal', 'low'
    limit: int = 20
    max_chars: int = 300
    sort_by: str = "received_desc"  # 'received_desc', 'received_asc'
    options: Dict[str, Any] = field(default_factory=dict)


def _parse_bool(val: str) -> bool:
    return str(val).strip().lower() in ("true", "1", "yes", "y")


def _tokenize_line(line: str) -> List[str]:
    """Tokenize a line respecting double and single quotes."""
    try:
        return shlex.split(line, posix=False)
    except Exception:
        return line.split()


def parse_email_dsl(dsl_text: str, default_limit: int = 20, default_max_chars: int = 300) -> List[EmailCommand]:
    """
    Parse one or more lines of Email DSL into EmailCommand objects.
    Comments starting with '#' or '//' are ignored.
    """
    commands: List[EmailCommand] = []
    lines = dsl_text.strip().splitlines()

    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#") or line.startswith("//"):
            continue

        tokens = _tokenize_line(line)
        if not tokens:
            continue

        verb = tokens[0].upper()
        if verb not in ("SEARCH", "SUMMARIZE_UNREAD", "UNREAD"):
            # If the user typed a plain query without verb, default to SEARCH query="..."
            verb = "SEARCH"
            args_tokens = tokens
        else:
            args_tokens = tokens[1:]

        action = "SUMMARIZE_UNREAD" if verb in ("SUMMARIZE_UNREAD", "UNREAD") else "SEARCH"
        unread_only = True if action == "SUMMARIZE_UNREAD" else False
        folder = "inbox"
        query: Optional[str] = None
        sender: Optional[str] = None
        subject: Optional[str] = None
        after: Optional[str] = None
        before: Optional[str] = None
        has_attachment: Optional[bool] = None
        importance: Optional[str] = None
        limit = default_limit
        max_chars = default_max_chars
        sort_by = "received_desc"
        options: Dict[str, Any] = {}

        for tok in args_tokens:
            tok = tok.strip()
            if not tok:
                continue

            if "=" in tok:
                key, val = tok.split("=", 1)
                key = key.strip().lower()
                val = val.strip().strip('"').strip("'")

                if key == "folder":
                    folder = val.lower()
                elif key in ("query", "q", "search"):
                    query = val
                elif key in ("sender", "from"):
                    sender = val
                elif key in ("subject", "title"):
                    subject = val
                elif key in ("after", "since", "from_date", "start"):
                    after = val
                elif key in ("before", "until", "to_date", "end"):
                    before = val
                elif key in ("unread", "unread_only"):
                    unread_only = _parse_bool(val)
                elif key in ("attachment", "has_attachment", "attachments"):
                    has_attachment = _parse_bool(val)
                elif key in ("importance", "priority"):
                    importance = val.lower()
                elif key in ("limit", "max", "top", "count"):
                    try:
                        limit = max(1, min(100, int(val)))
                    except ValueError:
                        pass
                elif key in ("max_chars", "chars", "snippet_len", "snippet_length"):
                    try:
                        max_chars = max(50, min(2000, int(val)))
                    except ValueError:
                        pass
                elif key in ("sort", "sort_by", "order"):
                    sort_by = val.lower()
                else:
                    options[key] = val
            else:
                # Positional free-text term: if query not yet set, set as query
                clean_tok = tok.strip('"').strip("'")
                if query is None:
                    query = clean_tok
                else:
                    query = f"{query} {clean_tok}"

        cmd = EmailCommand(
            action=action,
            folder=folder,
            query=query,
            sender=sender,
            subject=subject,
            after=after,
            before=before,
            unread_only=unread_only,
            has_attachment=has_attachment,
            importance=importance,
            limit=limit,
            max_chars=max_chars,
            sort_by=sort_by,
            options=options,
        )
        commands.append(cmd)

    return commands
