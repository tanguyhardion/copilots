"""
Outlook Connector using pywin32 COM (MAPI Namespace) for desktop Outlook.
Provides safe querying, truncation, capping, and formatted Markdown context generation.
"""

from __future__ import annotations

import os
import re
import html
import datetime
from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional, Tuple

import pythoncom
import win32com.client

from copilots_app.services.email_copilot.dsl_parser import EmailCommand


# OlDefaultFolders enumeration in Outlook COM
OL_FOLDER_MAP = {
    "inbox": 6,          # olFolderInbox
    "sent": 5,           # olFolderSentMail
    "drafts": 16,        # olFolderDrafts
    "deleted": 3,        # olFolderDeletedItems
    "junk": 23,          # olFolderJunk
    "outbox": 4,         # olFolderOutbox
    "calendar": 9,       # olFolderCalendar
    "contacts": 10,      # olFolderContacts
    "tasks": 13,         # olFolderTasks
}


@dataclass
class EmailItemSummary:
    id: str
    entry_id: str
    subject: str
    sender_name: str
    sender_email: str
    received_time: str
    unread: bool
    has_attachments: bool
    attachment_names: List[str]
    importance: str  # 'high', 'normal', 'low'
    body_snippet: str
    to_recipients: List[str]
    cc_recipients: List[str]
    categories: List[str]

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class EmailExecutionResult:
    success: bool
    message: str
    total_matched: int
    returned_count: int
    items: List[EmailItemSummary]
    context_markdown: str
    error: Optional[str] = None


class OutlookConnector:
    """Interacts with local Microsoft Outlook via pywin32 COM."""

    def __init__(self):
        pass

    def _get_mapi_namespace(self):
        """Initialise COM apartment and retrieve MAPI namespace."""
        pythoncom.CoInitialize()
        try:
            # Try connecting to running instance first or launch Outlook
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook.GetNamespace("MAPI")

    def get_accounts_and_folders(self) -> Dict[str, Any]:
        """List available mail stores and standard folders."""
        try:
            mapi = self._get_mapi_namespace()
            accounts = []
            for folder in mapi.Folders:
                subfolders = []
                try:
                    for sub in folder.Folders:
                        subfolders.append(sub.Name)
                except Exception:
                    pass
                accounts.append({
                    "name": folder.Name,
                    "subfolders": subfolders,
                })
            return {"success": True, "accounts": accounts}
        except Exception as err:
            return {"success": False, "error": f"Failed to connect to Outlook: {err}"}

    def _resolve_folder(self, mapi, folder_spec: str):
        """Resolve folder name or default enumeration."""
        norm = folder_spec.strip().lower()
        if norm in OL_FOLDER_MAP:
            return mapi.GetDefaultFolder(OL_FOLDER_MAP[norm])
        
        # Search among top-level accounts and subfolders
        for store in mapi.Folders:
            if store.Name.lower() == norm:
                return store
            for sub in store.Folders:
                if sub.Name.lower() == norm:
                    return sub
                
        # Default fallback: Inbox
        return mapi.GetDefaultFolder(6)

    @staticmethod
    def _clean_body_snippet(raw_body: str, max_chars: int) -> str:
        """Strip tags, clean whitespaces, truncate to max_chars."""
        if not raw_body:
            return "(No body content)"
        # Remove HTML tags if present
        text = re.sub(r"<[^>]+>", " ", raw_body)
        # Unescape HTML entities
        text = html.unescape(text)
        # Normalize consecutive spaces and newlines
        text = re.sub(r"[\r\n]+", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        if len(text) > max_chars:
            return text[:max_chars].rstrip() + "…"
        return text

    def execute_command(self, cmd: EmailCommand) -> EmailExecutionResult:
        """Execute a parsed EmailCommand against Outlook COM."""
        try:
            mapi = self._get_mapi_namespace()
            folder = self._resolve_folder(mapi, cmd.folder)
            items = folder.Items
            
            # Sort newest first by ReceivedTime
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            # Build DASL / Jet filter when possible, or perform safe in-memory scanning
            # We combine Outlook Items filtering where safe and perform clean Python validation
            matched_items: List[EmailItemSummary] = []
            total_matches = 0

            # Outlook 0=olImportanceLow, 1=olImportanceNormal, 2=olImportanceHigh
            importance_map = {0: "low", 1: "normal", 2: "high"}

            # Iterate through items safely (1-indexed in COM or python iterator)
            count = 0
            # Guard against infinite loops or huge folders by setting a scan limit
            max_scan = 500

            for item in items:
                count += 1
                if count > max_scan and total_matches >= cmd.limit:
                    break

                # Ensure it's a MailItem or PostItem (skip ReportItem / MeetingItem errors if any)
                try:
                    # Message class check
                    msg_class = getattr(item, "MessageClass", "")
                    if msg_class and not msg_class.startswith("IPM.Note"):
                        continue
                except Exception:
                    pass

                try:
                    is_unread = bool(getattr(item, "UnRead", False))
                except Exception:
                    is_unread = False

                if cmd.unread_only and not is_unread:
                    continue

                # Subject
                try:
                    subj = str(getattr(item, "Subject", "") or "").strip()
                except Exception:
                    subj = ""

                # Sender Name & Email
                try:
                    sender_name = str(getattr(item, "SenderName", "") or "").strip()
                except Exception:
                    sender_name = ""

                try:
                    sender_email = str(getattr(item, "SenderEmailAddress", "") or "").strip()
                except Exception:
                    sender_email = ""

                # Received Time
                try:
                    recv_time = getattr(item, "ReceivedTime", None)
                    if recv_time:
                        # Convert COM pywintypes.Time or datetime
                        dt = datetime.datetime(
                            recv_time.year, recv_time.month, recv_time.day,
                            recv_time.hour, recv_time.minute, recv_time.second
                        )
                        recv_str = dt.strftime("%Y-%m-%d %H:%M")
                    else:
                        dt = None
                        recv_str = "Unknown"
                except Exception:
                    dt = None
                    recv_str = "Unknown"

                # Filter by after date
                if cmd.after and dt:
                    try:
                        after_dt = datetime.datetime.fromisoformat(cmd.after[:10])
                        if dt < after_dt:
                            continue
                    except Exception:
                        pass

                # Filter by before date
                if cmd.before and dt:
                    try:
                        before_dt = datetime.datetime.fromisoformat(cmd.before[:10])
                        if dt > before_dt:
                            continue
                    except Exception:
                        pass

                # Filter by sender
                if cmd.sender:
                    s_lower = cmd.sender.lower()
                    if s_lower not in sender_name.lower() and s_lower not in sender_email.lower():
                        continue

                # Filter by subject
                if cmd.subject:
                    if cmd.subject.lower() not in subj.lower():
                        continue

                # Attachments
                att_names = []
                has_att = False
                try:
                    attachments = getattr(item, "Attachments", None)
                    if attachments and attachments.Count > 0:
                        has_att = True
                        for i in range(1, min(attachments.Count + 1, 10)):
                            att_names.append(attachments.Item(i).FileName)
                except Exception:
                    pass

                if cmd.has_attachment is not None:
                    if cmd.has_attachment != has_att:
                        continue

                # Free-text Query search (matches subject, sender, or body snippet)
                try:
                    body = str(getattr(item, "Body", "") or "")
                except Exception:
                    body = ""

                if cmd.query:
                    q_lower = cmd.query.lower()
                    matches_q = (
                        q_lower in subj.lower() or
                        q_lower in sender_name.lower() or
                        q_lower in sender_email.lower() or
                        q_lower in body.lower()
                    )
                    if not matches_q:
                        continue

                # Importance
                try:
                    imp_val = getattr(item, "Importance", 1)
                    importance = importance_map.get(imp_val, "normal")
                except Exception:
                    importance = "normal"

                if cmd.importance and cmd.importance.lower() != importance:
                    continue

                # To / CC recipients
                to_list = []
                cc_list = []
                try:
                    to_recip = str(getattr(item, "To", "") or "")
                    if to_recip:
                        to_list = [r.strip() for r in to_recip.split(";") if r.strip()]
                    cc_recip = str(getattr(item, "CC", "") or "")
                    if cc_recip:
                        cc_list = [r.strip() for r in cc_recip.split(";") if r.strip()]
                except Exception:
                    pass

                # Categories
                categories = []
                try:
                    cats = str(getattr(item, "Categories", "") or "")
                    if cats:
                        categories = [c.strip() for c in cats.split(",") if c.strip()]
                except Exception:
                    pass

                entry_id = ""
                try:
                    entry_id = str(getattr(item, "EntryID", "") or "")
                except Exception:
                    pass

                snippet = self._clean_body_snippet(body, cmd.max_chars)

                total_matches += 1
                if len(matched_items) < cmd.limit:
                    summary = EmailItemSummary(
                        id=f"email-{total_matches}",
                        entry_id=entry_id,
                        subject=subj or "(No Subject)",
                        sender_name=sender_name or "Unknown Sender",
                        sender_email=sender_email,
                        received_time=recv_str,
                        unread=is_unread,
                        has_attachments=has_att,
                        attachment_names=att_names,
                        importance=importance,
                        body_snippet=snippet,
                        to_recipients=to_list,
                        cc_recipients=cc_list,
                        categories=categories,
                    )
                    matched_items.append(summary)

            # Generate formatted Markdown context for LLM
            context_md = self.format_markdown_context(
                cmd=cmd,
                items=matched_items,
                total_matches=total_matches,
            )

            msg = f"Found {total_matches} email(s). Showing top {len(matched_items)}."
            return EmailExecutionResult(
                success=True,
                message=msg,
                total_matched=total_matches,
                returned_count=len(matched_items),
                items=matched_items,
                context_markdown=context_md,
            )

        except Exception as err:
            return EmailExecutionResult(
                success=False,
                message=f"Outlook query failed: {err}",
                total_matched=0,
                returned_count=0,
                items=[],
                context_markdown="",
                error=str(err),
            )

    @staticmethod
    def format_markdown_context(cmd: EmailCommand, items: List[EmailItemSummary], total_matches: int) -> str:
        """Format email summaries into a clear, LLM-ready markdown prompt context."""
        lines = []
        action_title = "Unread Email Triage" if cmd.action == "SUMMARIZE_UNREAD" else f"Email Search Results ('{cmd.query or 'All'}')"
        lines.append(f"# {action_title}")
        lines.append(f"- **Folder**: {cmd.folder.title()}")
        lines.append(f"- **Total Matched**: {total_matches}")
        lines.append(f"- **Displayed**: Top {len(items)} (capped at limit={cmd.limit}, max_chars={cmd.max_chars})")
        if total_matches > len(items):
            lines.append(f"> [!NOTE]\n> Showing top {len(items)} of {total_matches} total matching emails. Remaining {total_matches - len(items)} omitted.\n")
        lines.append("")
        lines.append("## Emails")
        lines.append("")

        if not items:
            lines.append("*(No emails found matching the criteria)*")
            return "\n".join(lines)

        for idx, itm in enumerate(items, 1):
            unread_badge = "[UNREAD] " if itm.unread else ""
            lines.append(f"### {idx}. {unread_badge}{itm.subject}")
            sender_str = f"{itm.sender_name} <{itm.sender_email}>" if itm.sender_email else itm.sender_name
            lines.append(f"- **From**: {sender_str}")
            lines.append(f"- **Date**: {itm.received_time}")
            if itm.importance != "normal":
                lines.append(f"- **Importance**: {itm.importance.upper()}")
            if itm.has_attachments:
                att_str = ", ".join(itm.attachment_names) if itm.attachment_names else "Yes"
                lines.append(f"- **Attachments**: {att_str}")
            if itm.categories:
                lines.append(f"- **Categories**: {', '.join(itm.categories)}")
            lines.append(f"- **Snippet**: {itm.body_snippet}")
            lines.append("")

        return "\n".join(lines)
