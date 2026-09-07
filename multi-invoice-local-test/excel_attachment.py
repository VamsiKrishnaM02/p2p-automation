"""
excel_attachment.py -- Attach a supporting Excel sheet to the last invoice.

Rule set:
  * ZERO Excel files in the folder -> do nothing, log that none was found.
  * MORE THAN ONE Excel file -> do nothing, log a warning (ambiguous --
    we don't guess which one is correct).
  * Exactly ONE Excel file -> scan every sheet for the invoice number:
      - Sheets with NO real content are skipped entirely (never a
        candidate), fixing the earlier bug of attaching blank sheets.
      - Among sheets that DO contain the invoice number, the LAST one
        by sheet/tab order is the one attached (tie-break rule).
      - If NO non-empty sheet contains the invoice number, nothing is
        attached (flagged with a reason) rather than guessing.
    The chosen sheet is rendered as a PDF table in LANDSCAPE orientation
    and appended to the last split invoice PDF.

MATCHING: the invoice number is compared against every digit-run found in
each cell (via regex), not the whole cell string, and leading zeros are
stripped on both sides before comparing. This handles:
  - a cell containing the number embedded in other text, e.g.
    "Invoice INV00009531 dated 01/05/2024" -- the invoice number is one
    of several digit-runs in that cell, alongside date numbers.
  - Excel silently dropping leading zeros when a value looks numeric
    (e.g. "00009531" stored/read back as 9531) -- both sides are
    normalized the same way before comparing.
  - the rare case where an invoice number has no digits at all, in which
    case matching falls back to a plain case-insensitive substring check.

Uses openpyxl (read the sheet's cell values) + reportlab (render as a PDF
table) -- both free, no license/watermark, no external dependency on
Excel/LibreOffice being installed. This renders a plain table of the
sheet's values, not a pixel-exact copy of its Excel formatting.

Wire into main.py: call `attach_excel_to_last_invoice(pdf_path, written,
invoice_number)` after the LAST split invoice's invoice_number is known
(i.e. after the extraction loop, using the last invoice's confirmed
invoice_number) -- pass the ORIGINAL source `pdf_path`, not a split
output, since that's where the Excel file lives.
"""

import logging
import re
from pathlib import Path
from typing import List, Optional, Set, Tuple

import openpyxl
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

log = logging.getLogger("excel_attachment")

_EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")

# Keep the rendered page readable rather than crushing a huge sheet down
# to unreadable text.
_MAX_ROWS = 200
_MAX_COLS = 20
_FONT_SIZE = 7

_DIGIT_RUN_RE = re.compile(r"\d+")
_NON_ALNUM_RE = re.compile(r"[^a-z0-9]")


# ---------------------------------------------------------------------------
# Step 1: find Excel file(s) next to the invoice PDF
# ---------------------------------------------------------------------------
def _find_excel_files(pdf_path: str) -> List[Path]:
    """All Excel workbooks in the same folder as `pdf_path` (skips Excel's
    ~$ lock files)."""
    folder = Path(pdf_path).parent
    return [
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in _EXCEL_EXTENSIONS
        and not p.name.startswith("~$")
    ]


# ---------------------------------------------------------------------------
# Step 2: find which sheet(s) contain the invoice number, skipping empties
# ---------------------------------------------------------------------------
def _digit_runs(text: str) -> Set[str]:
    """Every run of digits in `text`, each with leading zeros stripped, as
    a set. Used to compare an invoice number against a cell that may
    contain OTHER numbers too (dates, quantities, etc.) alongside it --
    e.g. "INV00009531 dated 01/05/2024" has digit-runs {"9531","1","5","2024"}."""
    return {run.lstrip("0") or "0" for run in _DIGIT_RUN_RE.findall(text or "")}


def _normalize_alnum(text: str) -> str:
    """Lowercase, strip everything except letters and digits."""
    return _NON_ALNUM_RE.sub("", (text or "").lower())


def _is_pure_digits(text: str) -> bool:
    """True if, after stripping separators/punctuation, `text` is made up
    ONLY of digits -- e.g. "00009531" or "0000-9531". These are the ONLY
    invoice-number shapes at real risk of Excel's numeric-cell leading-
    zero-drop, since Excel only auto-converts a cell to a number when its
    content looks purely numeric. Anything with a LETTER in it (e.g.
    "D073/26", "E1A2410078") is stored as text and Excel never reformats
    it, so digit-only comparison isn't needed -- and is actively wrong,
    see _sheet_scan()."""
    stripped = re.sub(r"[^a-z0-9]", "", (text or "").lower())
    return bool(stripped) and stripped.isdigit()


def _sheet_scan(ws, invoice_number: str) -> Tuple[bool, bool]:
    """One pass over every cell in the sheet. Returns (is_empty, matched).

    is_empty: True if every cell is None/blank -- handles Excel's 'phantom
    dimensions' problem where a sheet reports having rows/cols even after
    all its content was deleted.

    matched: depends on the SHAPE of the invoice number --
      - PURE DIGITS (e.g. "00009531"): compared against each cell's
        digit-runs individually (see _digit_runs), leading zeros stripped
        on both sides. This is the case genuinely at risk of Excel
        silently dropping leading zeros from a numeric-typed cell.
      - MIXED ALPHANUMERIC (has a letter, e.g. "D073/26", "E1A2410078"):
        compared as a normalized (punctuation-stripped, lowercased)
        substring instead. IMPORTANT: this does NOT decompose into
        separate digit-runs and glue them together -- doing that for
        "D073/26" would wrongly compare against "073"+"26" glued as
        "7326", which matches nothing real. Excel stores letter-containing
        values as text verbatim, so there's no leading-zero risk to guard
        against here in the first place.
    """
    inv = (invoice_number or "").strip()
    pure_digits = _is_pure_digits(inv)

    target_run = ""
    target_norm = ""
    if pure_digits:
        target_run = re.sub(r"[^0-9]", "", inv).lstrip("0") or "0"
    else:
        target_norm = _normalize_alnum(inv)

    has_content = False
    matched = False
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if v is None:
                continue
            s = str(v).strip()
            if not s:
                continue
            has_content = True
            if pure_digits:
                if target_run and target_run in _digit_runs(s):
                    matched = True
            else:
                cell_norm = _normalize_alnum(s)
                if target_norm and (target_norm in cell_norm or cell_norm in target_norm):
                    matched = True
    return (not has_content, matched)


def _find_matching_sheet(xlsx_path: str, invoice_number: str) -> Tuple[Optional[str], List[str]]:
    """Scan every sheet for the invoice number, skipping empty sheets
    entirely. Returns (chosen_sheet_name_or_None, all_matching_sheet_names).
    The chosen sheet is the LAST match by tab order, per the tie-break rule."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    matches: List[str] = []
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            is_empty, matched = _sheet_scan(ws, invoice_number)
            if is_empty:
                continue  # never a candidate, regardless of match
            if matched:
                matches.append(name)
    finally:
        wb.close()
    chosen = matches[-1] if matches else None
    return chosen, matches


# ---------------------------------------------------------------------------
# Step 3: render the chosen sheet's values as a standalone PDF table
# ---------------------------------------------------------------------------
def _cell_to_str(value) -> str:
    return "" if value is None else str(value)


def _sheet_to_pdf(xlsx_path: str, sheet_name: str, out_pdf_path: str) -> None:
    """Read the given sheet and render its cell values as a landscape PDF
    table."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb[sheet_name]

    rows_data = []
    for row in ws.iter_rows(values_only=True):
        rows_data.append([_cell_to_str(v) for v in row])
        if len(rows_data) >= _MAX_ROWS:
            break
    wb.close()

    total_cols = max((len(r) for r in rows_data), default=0)
    if total_cols > _MAX_COLS:
        rows_data = [r[:_MAX_COLS] for r in rows_data]

    doc = SimpleDocTemplate(
        out_pdf_path,
        pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm,
        topMargin=10 * mm, bottomMargin=10 * mm,
    )

    if not rows_data:
        # Defensive fallback only -- _find_matching_sheet() already
        # excludes empty sheets, so this shouldn't normally be reached.
        elements = [Table([["(empty sheet)"]])]
    else:
        table = Table(rows_data, repeatRows=1)
        table.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), _FONT_SIZE),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        elements = [table]

    doc.build(elements)


# ---------------------------------------------------------------------------
# Step 4: append that PDF to the end of the last invoice PDF
# ---------------------------------------------------------------------------
def _append_pdf(base_pdf_path: str, addition_pdf_path: str) -> None:
    writer = PdfWriter()
    for reader_path in (base_pdf_path, addition_pdf_path):
        reader = PdfReader(reader_path)
        for page in reader.pages:
            writer.add_page(page)
    with open(base_pdf_path, "wb") as fh:
        writer.write(fh)


# ---------------------------------------------------------------------------
# Orchestrator -- call this one from main.py
# ---------------------------------------------------------------------------
def attach_excel_to_last_invoice(source_pdf_path: str,
                                 written_invoice_paths: List[str],
                                 invoice_number: Optional[str]) -> dict:
    """
    Returns:
      {"attached": bool, "excel_path": str or None,
       "sheet_name": str or None, "candidate_sheets": List[str],
       "reason": str or None}
    """
    if not written_invoice_paths:
        return {"attached": False, "excel_path": None, "sheet_name": None,
                "candidate_sheets": [], "reason": "no split invoices to attach to"}

    excel_files = _find_excel_files(source_pdf_path)

    if not excel_files:
        log.info("No Excel file present in %s -- nothing to attach.",
                 Path(source_pdf_path).parent)
        return {"attached": False, "excel_path": None, "sheet_name": None,
                "candidate_sheets": [], "reason": "no Excel file in folder"}

    if len(excel_files) > 1:
        log.warning("Multiple Excel files found in %s -- skipping attachment "
                    "(ambiguous which one to use): %s",
                    Path(source_pdf_path).parent, [f.name for f in excel_files])
        return {"attached": False, "excel_path": None, "sheet_name": None,
                "candidate_sheets": [], "reason": "multiple Excel files found"}

    # Exactly one Excel file -- proceed.
    excel_path = excel_files[0]
    last_invoice_path = written_invoice_paths[-1]

    try:
        sheet_name, matches = _find_matching_sheet(str(excel_path), invoice_number)
    except Exception as e:
        log.warning("Failed to scan %s for invoice number %r: %s",
                    excel_path, invoice_number, e)
        return {"attached": False, "excel_path": str(excel_path), "sheet_name": None,
                "candidate_sheets": [], "reason": str(e)}

    if sheet_name is None:
        log.info("No non-empty sheet in %s matched invoice number %r -- "
                 "nothing to attach.", excel_path.name, invoice_number)
        return {"attached": False, "excel_path": str(excel_path), "sheet_name": None,
                "candidate_sheets": [], "reason": "no sheet matched the invoice number"}

    if len(matches) > 1:
        log.info("%d sheets in %s matched invoice number %r (%s) -- "
                 "using the last one: %r",
                 len(matches), excel_path.name, invoice_number, matches, sheet_name)

    tmp_sheet_pdf = str(Path(last_invoice_path).with_name(
        Path(last_invoice_path).stem + "__excel_sheet_tmp.pdf"))

    try:
        _sheet_to_pdf(str(excel_path), sheet_name, tmp_sheet_pdf)
        _append_pdf(last_invoice_path, tmp_sheet_pdf)
        log.info("Attached sheet %r from %s to %s",
                 sheet_name, excel_path.name, Path(last_invoice_path).name)
        return {"attached": True, "excel_path": str(excel_path),
                "sheet_name": sheet_name, "candidate_sheets": matches, "reason": None}
    except Exception as e:
        log.warning("Failed to attach Excel sheet from %s: %s", excel_path, e)
        return {"attached": False, "excel_path": str(excel_path), "sheet_name": None,
                "candidate_sheets": matches, "reason": str(e)}
    finally:
        Path(tmp_sheet_pdf).unlink(missing_ok=True)