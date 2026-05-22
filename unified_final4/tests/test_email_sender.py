"""Tests for ``oefof_pl.email_sender.outlook``."""
from __future__ import annotations

import ast
from pathlib import Path

import pytest
from openpyxl import Workbook

from oefof_pl.email_sender import outlook as O
from oefof_pl.email_sender.outlook import (
    EmailMessage,
    EmailSendResult,
    build_email_html_body,
    build_summary_html_table,
    normalize_recipients,
    send_via_outlook,
)


# ─── normalize_recipients ────────────────────────────────────────────────────

@pytest.mark.parametrize("raw,expected", [
    ("", ""),
    (None, ""),
    ("a@x.com", "a@x.com"),
    ("a@x.com, b@y.com", "a@x.com;b@y.com"),
    ("a@x.com; b@y.com ;c@z.com", "a@x.com;b@y.com;c@z.com"),
    ("  a@x.com ", "a@x.com"),
    (";; ,, ", ""),
])
def test_normalize_recipients(raw, expected):
    assert normalize_recipients(raw) == expected


# ─── build_summary_html_table ────────────────────────────────────────────────

def _write_summary_workbook(path: Path, *, with_summary: bool = True,
                            with_data: bool = True) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Other"
    if with_summary:
        s = wb.create_sheet("Summary")
        # KPI cards (rows 4-11): label on r, value on r+1, in cols 1/3/5.
        s.cell(4, 1, "Total Positions"); s.cell(5, 1, 137)
        s.cell(4, 3, "MV End (EUR)"); s.cell(5, 3, 1500.0)
        s.cell(5, 3).number_format = '"€"#,##0'
        s.cell(4, 5, "Total P&L %"); s.cell(5, 5, 0.5)
        s.cell(5, 5).number_format = "0.00%"
        if with_data:
            # Section header row (col 1 only) for Instrument Breakdown.
            s.cell(13, 1, "Instrument Breakdown")
            headers = ["Type", "Positions", "Cost", "MV End", "P&L"]
            for c, h in enumerate(headers, start=1):
                s.cell(14, c, h)
            s.cell(15, 1, "ORD"); s.cell(15, 2, 100); s.cell(15, 3, 1000.0)
            s.cell(15, 4, 1500.0); s.cell(15, 5, -200.0)
            s.cell(16, 1, "TOTAL"); s.cell(16, 2, 100); s.cell(16, 3, 1000.0)
            s.cell(16, 4, 1500.0); s.cell(16, 5, -200.0)
    wb.save(path)


def test_build_summary_html_table_returns_none_when_file_missing(tmp_path: Path):
    assert build_summary_html_table(tmp_path / "missing.xlsx") is None


def test_build_summary_html_table_returns_none_without_summary_sheet(tmp_path: Path):
    p = tmp_path / "v.xlsx"
    _write_summary_workbook(p, with_summary=False)
    assert build_summary_html_table(p) is None


def test_build_summary_html_table_returns_none_when_no_data_rows(tmp_path: Path):
    p = tmp_path / "v.xlsx"
    _write_summary_workbook(p, with_data=False)
    # With no Instrument/By Analyst section, only the KPI grid is returned —
    # still a non-empty HTML fragment, never None.
    out = build_summary_html_table(p)
    assert out is not None
    assert "Instrument Breakdown" not in out


def test_build_summary_html_table_renders_table_with_total_row(tmp_path: Path):
    p = tmp_path / "v.xlsx"
    _write_summary_workbook(p)
    html_table = build_summary_html_table(p)
    assert html_table is not None
    assert html_table.startswith("<table")
    assert "Total Positions" in html_table  # KPI label
    assert "Instrument Breakdown" in html_table
    assert "ORD" in html_table
    assert "1,500" in html_table        # numeric formatting
    assert "50.00%" in html_table       # percent formatting (×100)
    assert "TOTAL" in html_table
    # Negative numbers get the loss color.
    assert "#b83a3a" in html_table


def test_build_summary_html_table_escapes_html(tmp_path: Path):
    p = tmp_path / "v.xlsx"
    wb = Workbook()
    s = wb.active
    s.title = "Summary"
    s.cell(13, 1, "Instrument Breakdown")
    for c, h in enumerate(["<script>", "B", "C", "D"], start=1):
        s.cell(14, c, h)
    s.cell(15, 1, "<img src=x>")
    s.cell(15, 2, 1.0); s.cell(15, 3, 2.0); s.cell(15, 4, 3.0)
    wb.save(p)
    html_table = build_summary_html_table(p)
    assert "&lt;script&gt;" in html_table
    assert "&lt;img" in html_table
    assert "<script>" not in html_table


# ─── build_email_html_body ───────────────────────────────────────────────────

def test_build_email_html_body_paragraphs_and_escape():
    body = "Line 1 <b>bold</b>\n\nLine 2"
    out = build_email_html_body(body, None)
    assert "<html>" in out and "</html>" in out
    assert out.count("<p ") == 2   # two paragraphs
    assert "&lt;b&gt;bold&lt;/b&gt;" in out


def test_build_email_html_body_includes_summary_block():
    out = build_email_html_body("hi", "<table>X</table>")
    assert "Summary" in out
    assert "<table>X</table>" in out


def test_build_email_html_body_renders_footer_lines():
    out = build_email_html_body(
        "hi", None,
        footer_lines=("Computed at: 2026-04-29 14:30", "Period: 2025-12-31 to 2026-04-29"),
    )
    assert "Computed at: 2026-04-29 14:30" in out
    assert "Period: 2025-12-31 to 2026-04-29" in out
    assert "border-top:" in out


def test_build_email_html_body_no_footer_when_empty():
    out = build_email_html_body("hi", None)
    assert "border-top:" not in out


# ─── send_via_outlook (with injected fake dispatcher) ────────────────────────

class _FakeAttachments:
    def __init__(self):
        self.added: list[str] = []

    def Add(self, path):
        self.added.append(path)


class _FakeMail:
    def __init__(self):
        self.To = ""
        self.CC = ""
        self.Subject = ""
        self.Body = ""
        self.HTMLBody = ""
        self.SendUsingAccount = None
        self.Attachments = _FakeAttachments()
        self.sent = False

    def Send(self):
        self.sent = True


class _FakeAccount:
    def __init__(self, smtp):
        self.SmtpAddress = smtp


class _FakeSession:
    def __init__(self, accounts):
        self.Accounts = accounts


class _FakeOutlook:
    def __init__(self, *, accounts=(), fail_create=False, fail_send=False):
        self.Session = _FakeSession([_FakeAccount(a) for a in accounts])
        self._fail_create = fail_create
        self._fail_send = fail_send
        self.last_mail: _FakeMail | None = None

    def CreateItem(self, kind):
        if self._fail_create:
            raise RuntimeError("create failed")
        assert kind == 0
        mail = _FakeMail()
        if self._fail_send:
            mail.Send = lambda: (_ for _ in ()).throw(RuntimeError("send failed"))
        self.last_mail = mail
        return mail


def _make_attachment(tmp_path: Path, name: str = "report.xlsx") -> Path:
    p = tmp_path / name
    p.write_bytes(b"x")
    return p


def test_send_via_outlook_no_recipients_returns_failure(tmp_path: Path):
    msg = EmailMessage(to="", subject="s", body_text="b",
                       attachments=(_make_attachment(tmp_path),))
    result = send_via_outlook(msg, dispatcher=lambda _: _FakeOutlook())
    assert result.sent is False
    assert result.error == "no recipients"


def test_send_via_outlook_no_existing_attachments_returns_failure(tmp_path: Path):
    msg = EmailMessage(to="a@x.com", subject="s", body_text="b",
                       attachments=(tmp_path / "missing.xlsx",))
    result = send_via_outlook(msg, dispatcher=lambda _: _FakeOutlook())
    assert result.sent is False
    assert "no attachments" in result.error


def test_send_via_outlook_dispatcher_import_error_returns_failure(tmp_path: Path):
    def bad_dispatch(_):
        raise ImportError("pywin32 missing")

    msg = EmailMessage(to="a@x.com", subject="s", body_text="b",
                       attachments=(_make_attachment(tmp_path),))
    result = send_via_outlook(msg, dispatcher=bad_dispatch)
    assert result.sent is False
    assert result.error == "pywin32 not installed"


def test_send_via_outlook_dispatcher_runtime_error_returns_failure(tmp_path: Path):
    def boom(_):
        raise RuntimeError("Outlook is sulking")

    msg = EmailMessage(to="a@x.com", subject="s", body_text="b",
                       attachments=(_make_attachment(tmp_path),))
    result = send_via_outlook(msg, dispatcher=boom)
    assert result.sent is False
    assert "dispatch failed" in result.error


def test_send_via_outlook_send_failure_returns_failure(tmp_path: Path):
    outlook = _FakeOutlook(fail_send=True)
    msg = EmailMessage(to="a@x.com", subject="s", body_text="b",
                       attachments=(_make_attachment(tmp_path),))
    result = send_via_outlook(msg, dispatcher=lambda _: outlook)
    assert result.sent is False
    assert "send failed" in result.error


def test_send_via_outlook_happy_path_with_html_and_cc(tmp_path: Path):
    outlook = _FakeOutlook()
    att1 = _make_attachment(tmp_path, "a.xlsx")
    att2 = _make_attachment(tmp_path, "b.xlsx")
    msg = EmailMessage(
        to="a@x.com", cc="b@x.com", subject="Subject Line",
        body_text="plain", body_html="<html>hi</html>",
        attachments=(att1, att2),
    )
    result = send_via_outlook(msg, dispatcher=lambda _: outlook)

    assert isinstance(result, EmailSendResult)
    assert result.sent is True
    assert result.attachment_count == 2
    assert outlook.last_mail.To == "a@x.com"
    assert outlook.last_mail.CC == "b@x.com"
    assert outlook.last_mail.Subject == "Subject Line"
    assert outlook.last_mail.Body == "plain"
    assert outlook.last_mail.HTMLBody == "<html>hi</html>"
    assert outlook.last_mail.Attachments.added == [str(att1), str(att2)]
    assert outlook.last_mail.sent is True


def test_send_via_outlook_finds_named_account(tmp_path: Path):
    outlook = _FakeOutlook(accounts=("self@x.com", "team@x.com"))
    msg = EmailMessage(
        to="dest@x.com", subject="s", body_text="b",
        attachments=(_make_attachment(tmp_path),),
        from_smtp="TEAM@X.COM",
    )
    result = send_via_outlook(msg, dispatcher=lambda _: outlook)
    assert result.sent is True
    assert result.account_used == "team@x.com"
    assert outlook.last_mail.SendUsingAccount.SmtpAddress == "team@x.com"


def test_send_via_outlook_silently_ignores_unknown_account(tmp_path: Path):
    outlook = _FakeOutlook(accounts=("self@x.com",))
    msg = EmailMessage(
        to="dest@x.com", subject="s", body_text="b",
        attachments=(_make_attachment(tmp_path),),
        from_smtp="missing@x.com",
    )
    result = send_via_outlook(msg, dispatcher=lambda _: outlook)
    assert result.sent is True
    assert result.account_used == ""
    assert outlook.last_mail.SendUsingAccount is None


# ─── AST: COM imports must be function-local ─────────────────────────────────

def test_outlook_module_does_not_import_win32com_at_top_level():
    src = Path(O.__file__).read_text(encoding="utf-8")
    tree = ast.parse(src)
    for node in tree.body:
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            mod = getattr(node, "module", None)
            names = [a.name for a in node.names] if isinstance(node, ast.Import) else []
            assert mod != "win32com" and "win32com" not in names
            assert mod != "pythoncom" and "pythoncom" not in names
