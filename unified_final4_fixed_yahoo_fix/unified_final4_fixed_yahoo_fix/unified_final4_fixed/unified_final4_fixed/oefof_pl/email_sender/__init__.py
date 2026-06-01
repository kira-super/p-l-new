"""Outlook email sender for the OEFOF P&L pipeline."""
from .outlook import (
    EmailMessage,
    EmailSendResult,
    build_email_html_body,
    build_summary_html_table,
    normalize_recipients,
    send_via_outlook,
)

__all__ = [
    "EmailMessage", "EmailSendResult",
    "build_email_html_body", "build_summary_html_table",
    "normalize_recipients", "send_via_outlook",
]
