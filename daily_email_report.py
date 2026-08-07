"""Scheduled script (not part of the interactive Streamlit app): reads the
local Plex export folder (last written by daily_plex_export.py), runs the
same cleaning + Total PO $ + Report pipeline main.py's UI does, and emails
the combined result (one sheet per cleaned project, plus a Project Summary
sheet with its Profit/Loss chart) as a single Excel attachment to RECIPIENT.

Sends via Outlook desktop app COM automation (win32com), using whatever
account is signed into Outlook on this machine — no separate SMTP
credentials needed. Windows + Outlook desktop only.

Meant to run daily via Windows Task Scheduler, after daily_plex_export.py's
7:00 AM run (e.g. 9:00 AM, to leave time for Drive sync / a same-morning
PO Registry update). "Start in" set to this file's directory so
secrets.toml resolves.

Run manually: python daily_email_report.py
Preview without sending (opens a draft in Outlook instead): python daily_email_report.py --draft
"""

import sys
import tempfile
from pathlib import Path

import pandas as pd
import win32com.client

from plex_data import (
    add_totals_row,
    build_daily_report_excel,
    build_project_summary,
    clean_sheets,
    fetch_local_files,
    load_all_sheets,
    load_po_registry,
    normalize_pn,
)

RECIPIENT = "MBell@avna.com"


def send_daily_report_email(
    subject: str, body: str, attachment_bytes: bytes, attachment_filename: str, send: bool = True
) -> None:
    """Composes via the Outlook desktop app (COM automation), so it uses
    whichever account is already signed in there. Attachments.Add needs a
    file path, not raw bytes, so the attachment is written to a temp file
    named the way we want it to display. send=False opens it as a draft
    (mail.Display()) instead of actually sending — for previewing before the
    scheduled task starts sending for real."""
    tmp_dir = Path(tempfile.mkdtemp())
    tmp_path = tmp_dir / attachment_filename
    tmp_path.write_bytes(attachment_bytes)
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = RECIPIENT
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(str(tmp_path))
        if send:
            mail.Send()
        else:
            mail.Display()
    finally:
        tmp_path.unlink(missing_ok=True)
        tmp_dir.rmdir()


def main() -> None:
    files = fetch_local_files()
    if not files:
        print("No local export files found — nothing to email.")
        return

    raw_sheets = load_all_sheets(files)
    processed, missing_report = clean_sheets(raw_sheets)
    for name, missing in missing_report.items():
        print(f"  {name}: missing columns {missing} — skipped")

    if not processed:
        print("No sheets cleaned successfully — nothing to email.")
        return

    po_registry = load_po_registry()
    final_sheets = {}
    for name, df in processed.items():
        df = df.assign(**{"Total PO $": df["Okay PN"].map(normalize_pn).map(po_registry).fillna(0.0)})
        df["Budget Left"] = df["Total PO $"] - df["Total Cost"]
        final_sheets[name] = add_totals_row(df)

    project_summary = build_project_summary(final_sheets).sort_values(
        "Profit or Loss ($)", ascending=False, ignore_index=True
    )
    report_bytes = build_daily_report_excel(final_sheets, project_summary)

    draft_only = "--draft" in sys.argv
    today = pd.Timestamp.now().strftime("%Y-%m-%d")
    send_daily_report_email(
        subject=f"Daily Tooling Report - {today}",
        body="Attached is today's cleaned tooling data and profit/loss report.",
        attachment_bytes=report_bytes,
        attachment_filename=f"Daily_Tooling_Report_{today}.xlsx",
        send=not draft_only,
    )
    action = "Opened draft for" if draft_only else "Emailed daily report"
    print(f"{action} ({len(final_sheets)} sheet(s)) to {RECIPIENT}")


if __name__ == "__main__":
    main()
