"""
batch_report_builder.py -- Builds the "Batch & Failures Report", the SECOND
of the two Excel files the pipeline now produces (the first is the existing
per-invoice report_builder.py output -- unchanged, still one row per
invoice).

This file has TWO SHEETS, not two files:

  Success sheet  -- one row per NON-quarantined source email. Legacy format:
                    Total, Delta, Comment, Email Name, Source Email.
                      Total   = invoice count sent from that source email
                      Delta   = 0 for every success row (confirmed with Vamsi)
                      Comment = "Sent"
                      Email Name = every Batch ID that source email produced,
                                   comma-joined (a source email can split
                                   into several batches -- one per distinct
                                   LE it contained)
                      Source Email = the source-email folder name

  Failures sheet -- one row per failed invoice (or whole-PDF error) inside
                    a QUARANTINED source email: Source Email, Invoice #,
                    Supplier, Failure Reason.

Consumes the output of routing.process_batches()'s "complete" event
directly -- see flatten_success_rows() / flatten_failure_rows().
"""

from typing import Dict, List

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import routing

SUCCESS_COLUMNS = ["Total", "Delta", "Comment", "Email Name", "Source Email"]
FAILURE_COLUMNS = ["Source Email", "Invoice #", "Supplier", "Failure Reason"]

SUCCESS_COMMENT = "Sent"


# ---------------------------------------------------------------------------
# Flattening: routing.process_batches() output -> report rows
# ---------------------------------------------------------------------------
def flatten_success_rows(
    clean_folders: List[dict],
    email_to_batch_ids: Dict[str, List[str]],
) -> List[dict]:
    """One row per non-quarantined source-email folder."""
    rows = []
    for fr in clean_folders:
        total = routing.count_invoices(fr)
        batch_ids = email_to_batch_ids.get(fr["name"], [])
        rows.append({
            "Total": total,
            "Delta": 0,
            "Comment": SUCCESS_COMMENT,
            "Email Name": ",".join(batch_ids),
            "Source Email": fr["name"],
        })
    return rows


def flatten_failure_rows(failure_rows: List[dict]) -> List[dict]:
    """routing.build_failure_rows() already produces the right shape per
    folder; this just renames keys to the report's column headers and
    flattens the list-of-lists routing.py yields across folders."""
    return [
        {
            "Source Email": r.get("source_email"),
            "Invoice #": r.get("invoice_number"),
            "Supplier": r.get("supplier"),
            "Failure Reason": r.get("failure_reason"),
        }
        for r in failure_rows
    ]


# ---------------------------------------------------------------------------
# Excel building
# ---------------------------------------------------------------------------
def _write_sheet(ws, columns: List[str], rows: List[dict]) -> None:
    ws.append(columns)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"

    for row in rows:
        ws.append([row.get(col) for col in columns])

    # Autosize columns off the header + actual content widths.
    for i, col in enumerate(columns, start=1):
        letter = get_column_letter(i)
        max_len = len(str(col))
        for row in rows:
            val = row.get(col)
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[letter].width = min(max_len + 2, 60)


def build_batch_report(success_rows: List[dict], failure_rows: List[dict]) -> bytes:
    """Returns the two-sheet workbook as raw .xlsx bytes, ready to write to
    disk or hand to a Streamlit download button."""
    import io

    wb = Workbook()
    success_ws = wb.active
    success_ws.title = "Success"
    _write_sheet(success_ws, SUCCESS_COLUMNS, success_rows)

    failures_ws = wb.create_sheet("Failures")
    _write_sheet(failures_ws, FAILURE_COLUMNS, failure_rows)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save_batch_report(success_rows: List[dict], failure_rows: List[dict],
                      out_path: str) -> None:
    """Convenience wrapper for CLI use (mirrors report_builder.py's
    save_excel_report pattern)."""
    from pathlib import Path
    Path(out_path).write_bytes(build_batch_report(success_rows, failure_rows))