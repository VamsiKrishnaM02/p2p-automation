"""
main.py -- Orchestrator for the multi-invoice pipeline.

Wires the three stages together:
    signals.py     -> per-page supplier_name + invoice_number
    (this module)  -> segmentation (pure Python) + PDF splitting
    extraction.py  -> per-invoice bill_to + po_numbers (reusing supplier/inv_no)

Segmentation model (from signals only -- supplier + invoice number):
  * KEY SUPPLIER = supplier on the first page that yields one.
  * A page STARTS a new invoice when its supplier fuzzy-matches the key
    supplier AND it has an invoice number that is non-null and different from
    the current invoice's number. (The first key-supplier page opens invoice 1.)
  * A page is a CONTINUATION when its number equals the current invoice's, OR
    its number is null but the supplier still matches the key supplier.
  * Anything else (different supplier / unrelated number) is a SUPPORTING page
    -- it rides along in the current invoice's bundle but is not a MAIN page
    (extraction never reads supporting pages).
  * A failed signal (None) attaches to the current invoice and flags it.
  * Pages before the first real invoice (e.g. a summary) are dropped.

Run:
    python main.py <path-to-pdf>
Requires GROQ_API_KEY.
"""

import json
import logging
import os
import re
import sys
from pathlib import Path
from typing import Optional, List

from pypdf import PdfReader, PdfWriter
from rapidfuzz import fuzz

# Local modules (must be importable from the same folder).
import signals
import extraction
import tags
import routing
import batch_report_builder
import email_sender
from matcher_le import SensitiveMatcher, LEMatcher,LE_THRESHOLD
from vietnam_renamer import (
    VietnamSupplierMatcher, GoogleTranslateTranslator, process_vietnam_invoice, QwenTranslator
)

translator = QwenTranslator()

# File was renamed msg_reader.py -> email_reader.py (added .eml support
# alongside .msg) -- aliased here so the 9 existing msg_reader.X
# references elsewhere in this file don't all need renaming individually.
import email_reader as msg_reader
vietnam_translator = QwenTranslator()

# Excel-sheet attachment is off. Set True to re-enable; the returned
# "excel_attachment" key exists either way, so nothing downstream breaks.
ENABLE_EXCEL_ATTACHMENT = False

# Kofax routing emails + the human-agent report are DRY-RUN by default --
# messages are built and logged, no SMTP connection is opened, nothing is
# actually sent. Flip to False only once email_sender.SENDER_EMAIL /
# ROUTING_CC / REPORT_AGENT_EMAIL are filled in with real addresses and
# you've reviewed a dry-run's log output.
EMAIL_DRY_RUN = False


SENSITIVE_REFERENCE_PATH = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\p2p-prefinal-version\Supporting Documents\sensitive_reference.json"
sensitive_matcher = SensitiveMatcher(SENSITIVE_REFERENCE_PATH)

LE_REFERENCE_PATH = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\p2p-prefinal-version\Supporting Documents\le_reference.json"
le_matcher = LEMatcher(LE_REFERENCE_PATH)

VIETNAM_REFERENCE_PATH = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\p2p-prefinal-version\Supporting Documents\vietnam_reference.json"
vietnam_matcher = VietnamSupplierMatcher(VIETNAM_REFERENCE_PATH)


# --- Logging ---------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("main")
logging.getLogger("pypdf").setLevel(logging.ERROR)  # silence /Info warnings

# --- Config ----------------------------------------------------------------
# Fuzzy supplier-match threshold (0-100).

SUPPLIER_MATCH_THRESHOLD = 85

# --- Kofax routing table (loaded from Excel) --------------------------------
import pandas as pd

ROUTING_XLSX = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\p2p-prefinal-version\Supporting Documents\Legal Entity Lists.xlsx"


def load_routing_table(xlsx_path: str, sheet: str = "kofax_table") -> dict:
    """Load the Kofax routing Excel and build a lookup dict.

    The LE list uses short source_site names (e.g. "Swindon", "PNG", "CN")
    while the routing table uses longer forms ("UK\\Swindon",
    "Malaysia/Penang/PNG", "China"). To bridge this, we split each routing
    Source Site on / and \\ and index every fragment, so "PNG" and "Swindon"
    still find their row.

    Returns: {normalized_fragment: {appo: email, tidecr: email, ...}}
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet)

    # Find columns by partial match
    col_site = [c for c in df.columns if "source" in str(c).lower() and "site" in str(c).lower()][0]
    col_appo = [c for c in df.columns if "appo" in str(c).lower()][0]
    col_tide = [c for c in df.columns if "tide" in str(c).lower()][0]
    # Grab optional columns if they exist (for future extension)
    col_nonpo = [c for c in df.columns if "non po" in str(c).lower() or "nonpo" in str(c).lower()]
    col_archive = [c for c in df.columns if "archive" in str(c).lower()]
    col_rits = [c for c in df.columns if "rits" in str(c).lower()]

    lookup = {}

    for _, row in df.iterrows():
        raw_site = str(row[col_site]).strip() if pd.notna(row[col_site]) else ""
        if not raw_site or raw_site.lower() in ("nan", "none", ""):
            continue

        entry = {
            "source_site_full": raw_site,
            "appo": str(row[col_appo]).strip() if pd.notna(row[col_appo]) else "",
            "tidecr": str(row[col_tide]).strip() if pd.notna(row[col_tide]) else "",
            "nonpo": str(row[col_nonpo[0]]).strip() if col_nonpo and pd.notna(row[col_nonpo[0]]) else "",
            "archive": str(row[col_archive[0]]).strip() if col_archive and pd.notna(row[col_archive[0]]) else "",
            "rits": str(row[col_rits[0]]).strip() if col_rits and pd.notna(row[col_rits[0]]) else "",
        }

        # Index by the full name AND every fragment after splitting on / and \
        # so "Malaysia/Penang/PNG" is findable by "PNG", "Penang", or "Malaysia"
        # and "UK\Swindon" is findable by "Swindon" or "UK".
        fragments = re.split(r'[/\\]', raw_site)
        keys = [raw_site.strip().lower()] + [f.strip().lower() for f in fragments if f.strip()]
        for k in keys:
            lookup[k] = entry

    return lookup


def route_invoice(source_site: str, po_numbers: list, routing: dict) -> dict:
    """Determine the Kofax email for an invoice.

    Logic:
      - PO numbers present -> APPO Invoice
      - PO numbers empty   -> TIDEeCR Invoice

    Returns:
      {
        "invoice_type": "APPO" | "TIDEeCR",
        "kofax_email": "...",
        "routing_site": "...",     # the full Source Site row that matched
        "routed": True/False,
      }
    """
    if not source_site:
        return {"invoice_type": None, "kofax_email": None,
                "routing_site": None, "routed": False}

    # Lookup: try exact lowercase, then strip trailing spaces / common suffixes
    key = source_site.strip().lower()
    entry = routing.get(key)

    # If not found, try common normalizations
    if not entry:
        # "Ireland " -> "ireland", "Swindon " -> "swindon"
        for rk in routing:
            if key in rk or rk in key:
                entry = routing[rk]
                break

    if not entry:
        return {"invoice_type": None, "kofax_email": None,
                "routing_site": None, "routed": False}

    # PO present -> APPO, PO empty -> TIDEeCR
    if po_numbers:
        inv_type = "APPO"
        email = entry["appo"]
    else:
        inv_type = "TIDEeCR"
        email = entry["tidecr"]

    return {
        "invoice_type": inv_type,
        "kofax_email": email,
        "routing_site": entry["source_site_full"],
        "routed": True,
    }


# Load the routing table once at startup.
routing_table = load_routing_table(ROUTING_XLSX)

# ---------------------------------------------------------------------------
# Fuzzy supplier matching
# ---------------------------------------------------------------------------
def _norm(name: Optional[str]) -> str:
    return "".join(ch for ch in (name or "").lower() if ch.isalnum())


def _supplier_matches(page_supplier: Optional[str], key_supplier: Optional[str]) -> bool:
    """partial_ratio match on normalized names (handles short-vs-full legal
    name). Guards very short strings with plain ratio to avoid spurious
    substring matches.

    IMPORTANT: for non-Latin names (Chinese, Japanese, etc.), _norm may strip
    everything and return an empty string (behavior of [^a-z0-9] on CJK varies
    by Python build). In that case, fall back to comparing the RAW lowercased
    names, so a Chinese supplier still matches itself and isn't misclassified."""
    a, b = _norm(page_supplier), _norm(key_supplier)
    # If normalization emptied either side (non-Latin script), compare the raw
    # names instead of failing the match.
    if not a or not b:
        ra = (page_supplier or "").strip().lower()
        rb = (key_supplier or "").strip().lower()
        if not ra or not rb:
            return False
        score = fuzz.partial_ratio(ra, rb)
        log.debug("      supplier match (raw) %r vs key %r -> %.0f",
                  page_supplier, key_supplier, score)
        return score >= SUPPLIER_MATCH_THRESHOLD
    score = fuzz.ratio(a, b) if min(len(a), len(b)) < 4 else fuzz.partial_ratio(a, b)
    log.debug("      supplier match %r vs key %r -> %.0f", page_supplier, key_supplier, score)
    return score >= SUPPLIER_MATCH_THRESHOLD


# ---------------------------------------------------------------------------
# Segmentation (pure Python, no model calls)
# ---------------------------------------------------------------------------
def find_key_supplier(sigs: List[Optional["signals.PageSignal"]]) -> Optional[str]:
    for s in sigs:
        if s is not None and s.supplier_name_on_page:
            return s.supplier_name_on_page
    return None


def classify_pages(sigs: List[Optional["signals.PageSignal"]],
                   key_supplier: Optional[str]) -> dict:
    """
    Stateful walk -> invoice bundles. Returns:
      {
        "invoices": [ {start, main_pages, support_pages, all_pages,
                       invoice_number, supplier_name, unknown_pages} ],
        "unknown_pages": [ints],   # signal failed
        "dropped_pages": [ints],   # before first invoice (e.g. summary)
      }
    All page indices 0-based.
    """
    invoices: List[dict] = []
    unknown_pages: List[int] = []
    dropped_pages: List[int] = []
    current: Optional[dict] = None
    current_number: Optional[str] = None
    pending_pages: List[int] = []   # key-supplier pages w/ no number, waiting
                                     # for the invoice that will claim them

    def open_invoice(idx: int, supplier: Optional[str], number: Optional[str]) -> dict:
        inv = {"start": idx, "main_pages": [], "support_pages": [],
               "all_pages": [], "invoice_number": number,
               "supplier_name": supplier, "unknown_pages": []}
        invoices.append(inv)
        return inv

    for idx, s in enumerate(sigs):
        # Failed signal -> attach to current invoice, flag; if none open, drop.
        if s is None:
            unknown_pages.append(idx)
            if current is not None:
                current["all_pages"].append(idx)
                current["support_pages"].append(idx)
                current["unknown_pages"].append(idx)
            else:
                dropped_pages.append(idx)
            continue

        supplier_ok = _supplier_matches(s.supplier_name_on_page, key_supplier)
        num = s.invoice_number
        norm_num = _norm(num) if num else None
        norm_cur = _norm(current_number) if current_number else None

        # SUPPORTING DOC: same supplier, no invoice number (summary/TOC/etc).
        # Never a MAIN page -- always rides with the nearest invoice: the
        # one already open, or (if none open yet) the next one that opens.
        if supplier_ok and num is None:
            if current is not None:
                current["support_pages"].append(idx)
                current["all_pages"].append(idx)
            else:
                pending_pages.append(idx)
            continue

        # NEW INVOICE START: matches key supplier, has a real number, and
        # that number differs from the currently open invoice's.
        is_start = supplier_ok and num is not None and norm_num != norm_cur
        if current is None and supplier_ok and num is not None:
            is_start = True   # first real-numbered key-supplier page opens invoice 1

        if is_start:
            current = open_invoice(idx, s.supplier_name_on_page or key_supplier, num)
            current_number = num
            # Flush any held supporting pages into THIS invoice, in order.
            if pending_pages:
                for p in pending_pages:
                    current["support_pages"].append(p)
                    current["all_pages"].append(p)
                pending_pages = []
            current["main_pages"].append(idx)
            current["all_pages"].append(idx)
            continue

        # Nothing open yet and this isn't a start -> drop (leading non-invoice).
        if current is None:
            dropped_pages.append(idx)
            continue

        # CONTINUATION: same invoice number as the currently open invoice.
        same_number = num is not None and norm_num == norm_cur
        if same_number:
            current["main_pages"].append(idx)
            current["all_pages"].append(idx)
            continue

        # SUPPORTING: different supplier / unrelated number.
        current["support_pages"].append(idx)
        current["all_pages"].append(idx)

    # Fallback: nothing detected -> treat the whole doc as one invoice.
    if not invoices:
        log.warning("No invoices detected; treating whole document as one invoice.")
        all_idx = list(range(len(sigs)))
        invoices = [{"start": 0, "main_pages": all_idx, "support_pages": [],
                     "all_pages": all_idx, "invoice_number": None,
                     "supplier_name": key_supplier, "unknown_pages": unknown_pages}]

    return {"invoices": invoices, "unknown_pages": unknown_pages,
            "dropped_pages": dropped_pages}


# ---------------------------------------------------------------------------
# Splitting + naming
# ---------------------------------------------------------------------------
def _sanitize(text: Optional[str], fallback: str) -> str:
    if not text or not text.strip():
        return fallback
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", text.strip()).strip("_.")[:60]
    return safe or fallback


def split_and_name(pdf_path: str, invoices: List[dict], output_dir: str) -> List[str]:
    """Write one PDF per invoice (all its pages, in order). Named
    <supplier>_<invoice_number>.pdf, sanitized, collision-safe."""
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(pdf_path)
    written: List[str] = []
    used = set()

    log.info("Splitting into %d file(s) under %s ...", len(invoices), out_dir)
    for i, inv in enumerate(invoices, start=1):
        writer = PdfWriter()
        for p in sorted(inv["all_pages"]):
            writer.add_page(reader.pages[p])

        supp = _sanitize(inv.get("supplier_name"), "unknown_supplier")
        num = _sanitize(inv.get("invoice_number"), f"noinv_{i:02d}")
        stem = f"{supp}_{num}"
        candidate, n = stem, 2
        while candidate in used or (out_dir / f"{candidate}.pdf").exists():
            candidate = f"{stem}_{n}"
            n += 1
        used.add(candidate)

        out_path = out_dir / f"{candidate}.pdf"
        with open(out_path, "wb") as fh:
            writer.write(fh)
        written.append(str(out_path))
        log.info("  wrote %s (pages %s)", out_path.name, [p + 1 for p in sorted(inv["all_pages"])])

    return written


# ---------------------------------------------------------------------------
# End-to-end
# ---------------------------------------------------------------------------
def process_invoice_pdf(pdf_path: str, output_dir: Optional[str] = None) -> dict:
    pdf_path = str(pdf_path)

    log.info("=" * 70)
    log.info("START: %s", Path(pdf_path).name)
    log.info("=" * 70)

    # 1) Rasterize + signals (from signals.py)
    images = signals.pdf_to_base64_images(pdf_path)
    sigs = signals.get_page_signals(images)

    # 2) Key supplier + segmentation (pure Python here)
    key_supplier = find_key_supplier(sigs)
    log.info("Key supplier: %r", key_supplier)

    det = classify_pages(sigs, key_supplier)
    invoices = det["invoices"]
    log.info("Detected %d invoice(s). Dropped pages: %s. Unknown pages: %s",
             len(invoices),
             [p + 1 for p in det["dropped_pages"]],
             [p + 1 for p in det["unknown_pages"]])
    for i, inv in enumerate(invoices, 1):
        log.info("  invoice %d: start=p%d supplier=%r inv_no=%r main=%s support=%s",
                 i, inv["start"] + 1, inv["supplier_name"], inv["invoice_number"],
                 [p + 1 for p in inv["main_pages"]], [p + 1 for p in inv["support_pages"]])

    # 3) Output dir named by key supplier
    if output_dir is None:
        output_dir = str(Path(pdf_path).parent /
                         f"{_sanitize(key_supplier, Path(pdf_path).stem)}__split")

    # 4) Split PDFs
    written = split_and_name(pdf_path, invoices, output_dir)

    # 5) Extract bill_to + PO per invoice (from extraction.py), reusing
    #    supplier + invoice_number from segmentation, then run the
    #    sensitive-supplier check and LE matching.
    log.info("Extracting bill_to + PO per invoice...")
    results = []
    for i, (inv, out_file) in enumerate(zip(invoices, written), start=1):
        log.info("Invoice %d/%d (%s):", i, len(invoices), Path(out_file).name)
        fields = extraction.extract_invoice(
            images=images,
            main_pages=inv["main_pages"],
            supplier_name=inv["supplier_name"],
            invoice_number=inv["invoice_number"],
        )

        # --- Sensitive-supplier check (Phase 2) ---
        sens = sensitive_matcher.check(fields["supplier_name"], translator=translator)
        log.info("  sensitive check: %s (score=%s, matched=%r)",
                 sens["is_sensitive"], sens["score"], sens["matched_name"])

        # --- Tag derivation: TELECOM / UTILITY / DD (feeds batch_id suffix) ---
        # IMPORTANT: source_folder must be `pdf_path` (the ORIGINAL,
        # pre-split PDF, living under 01_extracted/), NOT `out_file` (the
        # split-out PDF written to 03_processed/) -- email_metadata.json
        # only ever exists next to the original in 01_extracted/, the same
        # place excel_attachment.py already looks for its paired Excel.
        # load_email_metadata() accepts a file path directly and walks up
        # to 2 levels, so this works whether pdf_path sits flat in the
        # message folder or one level down in a per-invoice subfolder.
        # For loose-PDF-folder / single_pdf runs with no email_metadata.json
        # at all, this degrades gracefully to supplier/bill-to-only matching.
        invoice_tags = tags.derive_tags(
            supplier_name=fields["supplier_name"],
            bill_to_name=fields["bill_to_client_name"],
            source_folder=pdf_path,
        )
        if invoice_tags:
            log.info("  tags: %s", invoice_tags)

        # --- LE matching (Phase 2) ---
        le_result = le_matcher.match(fields["bill_to_client_name"], translator=translator)
        if le_result["matched"]:
            log.info("  LE match: %s -> %s (site=%s, country=%s, score=%s, method=%s)",
                     fields["bill_to_client_name"], le_result["le"],
                     le_result["source_site"], le_result["country"],
                     le_result["score"], le_result["method"])
        else:
            log.warning("  LE match: NO MATCH for %r (best=%s, score=%s)",
                        fields["bill_to_client_name"],
                        le_result["matched_le_name"], le_result["score"])

         # --- Kofax routing (Phase 2) ---
        routing_result = route_invoice(
            source_site=le_result["source_site"],
            po_numbers=fields["po_numbers"],
            routing=routing_table,
        )
        if routing_result["routed"]:
            log.info("  routing: %s -> %s (%s)",
                     routing_result["invoice_type"],
                     routing_result["kofax_email"],
                     routing_result["routing_site"])
        else:
            log.warning("  routing: FAILED — source_site %r not in routing table",
                        le_result["source_site"])

        # --- Vietnam supplier renaming (Phase 2, LE763 only) ---
        vietnam_result = None
        if le_result["le"] == "LE763":
            log.info(" Processing a vietnam supplier ")
            vietnam_result = process_vietnam_invoice(
                pdf_path=out_file,
                supplier_name=fields["supplier_name"],
                po_numbers=fields["po_numbers"],
                invoice_number=fields["invoice_number"],
                matcher=vietnam_matcher,
                translator=vietnam_translator,
            )
            log.info(vietnam_result)
            if vietnam_result["renamed"]:
                log.info("  vietnam rename: %s -> %s",
                         Path(out_file).name,
                         Path(vietnam_result["new_pdf_path"]).name)
                out_file = vietnam_result["new_pdf_path"]  # keep results in sync
            else:
                log.warning("  vietnam rename: SKIPPED (%s)", vietnam_result["reason"])

        review_reasons = list(fields.get("review_reasons", []))

        if inv["unknown_pages"]:
            pages = [p + 1 for p in inv["unknown_pages"]]
            review_reasons.append({
                "code": "signal_extraction_failed",
                "message": f"Could not read supplier/invoice number on "
                          f"page(s) {pages} -- model extraction failed for these pages.",
            })

        if not le_result["matched"]:
            review_reasons.append({
                "code": "le_no_match",
                "message": (
                    f"No confident Legal Entity match for bill-to "
                    f"{fields['bill_to_client_name']!r} -- closest match was "
                    f"{le_result['matched_le_name']!r} (score {le_result['score']}, "
                    f"threshold {LE_THRESHOLD})."
                ),
            })

        if not routing_result["routed"]:
            review_reasons.append({
                "code": "routing_failed",
                "message": (
                    f"No routing entry found for source site "
                    f"{le_result['source_site']!r} -- Kofax email address not resolved."
                ),
            })

        if le_result["le"] == "LE763" and not vietnam_result["renamed"]:
            review_reasons.append({
                "code": "vietnam_rename_failed",
                "message": f"LE763 invoice, but file rename failed: {vietnam_result['reason']}.",
            })

        review = bool(review_reasons)

        results.append({
            "invoice_index": i,
            "output_pdf": out_file,
            "pages": [p + 1 for p in inv["all_pages"]],
            "main_pages": [p + 1 for p in inv["main_pages"]],
            "support_pages": [p + 1 for p in inv["support_pages"]],
            "supplier_name": fields["supplier_name"],
            "invoice_number": fields["invoice_number"],
            "bill_to_client_name": fields["bill_to_client_name"],
            "po_numbers": fields["po_numbers"],
            "is_sensitive": sens["is_sensitive"],
            "tags": invoice_tags,
            "sensitive_match": {
                "matched_name": sens["matched_name"],
                "score": sens["score"],
                "method": sens["method"],
                "translated": sens.get("translated", False),          
                "translated_query": sens.get("translated_query"),
            },
            "le": le_result["le"],
            "source_site": le_result["source_site"],
            "le_country": le_result["country"],
            "le_match": {
                "matched_le_name": le_result["matched_le_name"],
                "score": le_result["score"],
                "method": le_result["method"],
                "translated": le_result.get("translated", False),     
                "translated_query": le_result.get("translated_query"),
            },
            "invoice_type": routing_result["invoice_type"],
            "kofax_email": routing_result["kofax_email"],
            "vietnam_supplier_id": vietnam_result["supplier_match"]["supplier_id"] if vietnam_result else None,
            "vietnam_supplier_name": vietnam_result["supplier_match"]["translated_query"] if vietnam_result else None,
            "vietnam_renamed": vietnam_result["renamed"] if vietnam_result else None,
            "needs_manual_review": review,
            "review_reasons": review_reasons,
        })

    # 6) Attach a supporting Excel sheet to the last invoice's split PDF.
    #    DISABLED -- flip ENABLE_EXCEL_ATTACHMENT at the top of this file
    #    to turn it back on. The stub below keeps "excel_attachment" in the
    #    returned dict either way, so report_builder.py and app.py never
    #    hit a missing key.
    if ENABLE_EXCEL_ATTACHMENT:
        from excel_attachment import attach_excel_to_last_invoice
        last_invoice_number = results[-1]["invoice_number"] if results else None
        excel_result = attach_excel_to_last_invoice(pdf_path, written, last_invoice_number)
        if excel_result["attached"]:
            log.info("Excel attachment: sheet %r from %s appended to %s",
                     excel_result["sheet_name"],
                     Path(excel_result["excel_path"]).name,
                     Path(written[-1]).name)
        else:
            log.info("Excel attachment: %s", excel_result["reason"])
    else:
        excel_result = {"attached": False, "excel_path": None,
                        "sheet_name": None, "candidate_sheets": [],
                        "reason": "disabled"}

    # NOTE: the following were previously indented one level too deep (inside
    # this for-loop), which caused process_invoice_pdf() to return after
    # only the FIRST invoice. Fixed here -- both now run after the loop.
    log.info("DONE: %d invoice(s) from %d page(s). Output in %s",
             len(invoices), len(images), output_dir)
    log.info("=" * 70)

    return {
        "source_file": pdf_path,
        "total_pages": len(images),
        "key_supplier": key_supplier,
        "invoice_count": len(invoices),
        "output_dir": output_dir,
        "dropped_pages": [p + 1 for p in det["dropped_pages"]],
        "unknown_pages": [p + 1 for p in det["unknown_pages"]],
        "excel_attachment": excel_result,
        "invoices": results,
    }


def process_invoice_folders(scan_folders: List[dict], target_path: str):
    """
    Batch-process every PDF across multiple source subfolders, splitting
    output into a mirrored structure under `target_path`. This owns the
    batch loop AND the final consolidated Excel report -- a UI (or a CLI
    script) just drives it, it doesn't reimplement the looping or call
    report_builder itself.

    This is a GENERATOR so a caller (e.g. a Streamlit progress bar) can
    show live progress without process_invoice_folders having to know
    anything about how progress is displayed. It yields three kinds of
    events:

      {"type": "progress", "current": int, "total": int,
       "folder_name": str, "file_name": str}
        -- one PDF is about to start.

      {"type": "folder_done", "folder_result": dict}
        -- one subfolder's PDFs have all finished. `folder_result` has
        the same shape app.py's UI already expects: {"name",
        "source_path", "target_path", "pdf_count", "excel_count",
        "pdf_results": [...]}.

      {"type": "complete", "results": List[dict], "report_bytes": bytes}
        -- everything is done. `results` is the full list of
        folder_result dicts; `report_bytes` is a ready-to-save .xlsx
        (one row per invoice, see report_builder.py) or None if no
        invoices were found at all.

    `scan_folders`: list of dicts, each at least {"name": str,
    "path": Path, "pdfs": List[Path]} -- matches the folder entries
    produced by a UI's own folder scan (e.g. app.py's
    scan_source_folder()). An "excels" key is used for the reported
    excel_count if present, but isn't otherwise needed here -- Excel
    attachment happens inside process_invoice_pdf() itself.
    """
    from report_builder import flatten_folder_results, build_excel_report

    target_root = Path(target_path)
    total_pdfs = sum(len(f["pdfs"]) for f in scan_folders)
    pdf_count = 0
    all_results: List[dict] = []

    for folder_info in scan_folders:
        folder_name = folder_info["name"]
        folder_out = target_root / folder_name
        folder_out.mkdir(parents=True, exist_ok=True)

        folder_result = {
            "name": folder_name,
            "source_path": str(folder_info["path"]),
            "target_path": str(folder_out),
            "pdf_count": len(folder_info["pdfs"]),
            "excel_count": len(folder_info.get("excels", [])),
            "pdf_results": [],
        }

        for pdf_file in folder_info["pdfs"]:
            pdf_count += 1
            yield {"type": "progress", "current": pdf_count, "total": total_pdfs,
                   "folder_name": folder_name, "file_name": pdf_file.name}

            try:
                result = process_invoice_pdf(str(pdf_file), output_dir=str(folder_out))
                folder_result["pdf_results"].append(result)
            except Exception as e:
                folder_result["pdf_results"].append({
                    "source_file": str(pdf_file),
                    "error": str(e),
                    "invoices": [],
                })

        all_results.append(folder_result)
        yield {"type": "folder_done", "folder_result": folder_result}

    report_rows = flatten_folder_results(all_results)
    report_bytes = build_excel_report(report_rows) if report_rows else None
    yield {"type": "complete", "results": all_results, "report_bytes": report_bytes}


# ---------------------------------------------------------------------------
# Folder scanning
# ---------------------------------------------------------------------------
_PDF_EXT = {".pdf"}
_EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}


def scan_pdf_folder(source_path: str) -> List[dict]:
    """Scan a folder of subfolders, each holding invoice PDFs.

    This is the "loose PDFs in a folder" testing workflow -- msg_reader is
    NOT involved. Mirrors what app.py's own scan_source_folder() does, so
    the same layout works from the CLI.

    Returns the scan_folders list process_invoice_folders() expects.
    """
    root = Path(source_path)
    if not root.is_dir():
        raise NotADirectoryError(f"{source_path} is not a directory")

    subdirs = sorted([d for d in root.iterdir() if d.is_dir()])
    if not subdirs:                     # no subfolders -> treat root itself as one
        subdirs = [root]

    folders = []
    for d in subdirs:
        pdfs = sorted([f for f in d.iterdir()
                      if f.is_file() and f.suffix.lower() in _PDF_EXT])
        if not pdfs:
            continue
        excels = sorted([f for f in d.iterdir()
                        if f.is_file() and f.suffix.lower() in _EXCEL_EXT
                        and not f.name.startswith("~$")])
        folders.append({"name": d.name, "path": d, "pdfs": pdfs, "excels": excels})
    return folders


def scan_extracted_folder(extracted_root: str) -> List[dict]:
    """Scan msg_reader's 01_extracted/ output into scan_folders entries.

    ONE ENTRY PER MESSAGE -- every invoice from a single email stays
    together in one output folder, which keeps the email-level traceability
    that the batching/routing step needs.

    Message folders come in two shapes, and a naive one-level scan misses
    the second entirely:
        <msg>/invoice.pdf                (flat -- message had <=1 PDF)
        <msg>/INV001/invoice.pdf         (grouped -- message had 2+ PDFs)
    rglob covers both. msg_reader collapses deeper zip paths into a single
    folder name (a/b -> a_b), so today this is always at most two levels --
    but rglob means this keeps working if that ever changes.

    NOTE: collecting a message's PDFs into one entry does NOT undo
    msg_reader's grouping. excel_attachment.py locates its Excel from the
    PDF's OWN parent directory on disk, not from this entry, so each PDF
    still pairs with the single Excel sitting beside it.
    """
    root = Path(extracted_root)
    if not root.is_dir():
        raise NotADirectoryError(f"{extracted_root} is not a directory")

    folders = []
    for msg_dir in sorted(d for d in root.iterdir() if d.is_dir()):
        pdfs = sorted(p for p in msg_dir.rglob("*")
                     if p.is_file() and p.suffix.lower() in _PDF_EXT)
        if not pdfs:
            continue                    # blocked or attachment-less message
        excels = sorted(p for p in msg_dir.rglob("*")
                       if p.is_file() and p.suffix.lower() in _EXCEL_EXT
                       and not p.name.startswith("~$"))
        folders.append({"name": msg_dir.name, "path": msg_dir,
                       "pdfs": pdfs, "excels": excels})
    return folders


# ---------------------------------------------------------------------------
# End-to-end runs
# ---------------------------------------------------------------------------
def _run_extracted(run: dict, extracted_root: str):
    """Shared tail: scan an extracted folder, process it, save the report
    into the run folder. Used by both .msg runs and latest-run reruns."""
    folders = scan_extracted_folder(extracted_root)
    if not folders:
        yield {"type": "error",
               "message": f"No PDFs found under {extracted_root}"}
        return

    for event in process_invoice_folders(folders, run["processed"]):
        if event["type"] == "complete" and event.get("report_bytes"):
            Path(run["report_path"]).write_bytes(event["report_bytes"])
            event = dict(event, report_path=run["report_path"])

            # --- Batching + second report (routing.py / batch_report_builder.py) ---
            batch_complete = None
            for batch_event in routing.process_batches(event["results"], run):
                if batch_event["type"] == "complete":
                    batch_complete = batch_event
                else:
                    yield batch_event  # "quarantine" / "batch_done" progress

            if batch_complete is not None:
                success_rows = batch_report_builder.flatten_success_rows(
                    batch_complete["clean_folders"],
                    batch_complete["email_to_batch_ids"],
                )
                failure_rows = batch_report_builder.flatten_failure_rows(
                    batch_complete["failure_rows"]
                )
                batch_report_bytes = batch_report_builder.build_batch_report(
                    success_rows, failure_rows
                )
                batch_report_name = msg_reader.report_filename(
                    "batch_report", Path(run["date_dir"]).name, Path(run["run_root"]).name
                )
                batch_report_path = str(Path(run["run_root"]) / batch_report_name)
                Path(batch_report_path).write_bytes(batch_report_bytes)
                event = dict(event, batch_report_path=batch_report_path,
                            batches=batch_complete["batches"])

                # --- Send emails: Kofax routing, then the human-agent report ---
                smtp_password = "Genpact@147258ab" #os.environ.get("SMTP_PASSWORD")
                if smtp_password is None:
                    log.warning("SMTP_PASSWORD not set in the environment -- "
                              "skipping ALL email sending for this run "
                              "(routing emails and the human-agent report).")
                else:
                    annotated_batches = email_sender.send_routing_emails(
                        batch_complete["batches"],
                        sender=email_sender.SENDER_EMAIL,
                        password=smtp_password,
                        cc=email_sender.ROUTING_CC,
                        login_account=email_sender.LOGIN_ACCOUNT,
                        dry_run=EMAIL_DRY_RUN,
                    )
                    yield {"type": "routing_emails_done", "batches": annotated_batches}

                    run_stats = email_sender.compute_run_stats(
                        annotated_batches, batch_complete["quarantined"]
                    )
                    email_sender.send_report_email(
                        run_stats, batch_complete["quarantined"],
                        invoice_report_path=run["report_path"],
                        batch_report_path=batch_report_path,
                        run_root=run["run_root"],
                        agent_email=email_sender.REPORT_AGENT_EMAIL,
                        sender=email_sender.SENDER_EMAIL,
                        password=smtp_password,
                        cc=email_sender.ROUTING_CC,
                        login_account=email_sender.LOGIN_ACCOUNT,
                        run_label=run["run_name"],
                        dry_run=EMAIL_DRY_RUN,
                    )
                    event = dict(event, email_stats=run_stats)

        yield event


def run_from_msg_source(msg_folder: Optional[str] = None,
                       results_root: Optional[str] = None,
                       label: Optional[str] = None,
                       move_blocked_msg: bool = False):
    """Full .msg pipeline: extract attachments -> process invoices ->
    write the report, all inside one timestamped run folder.

    Generator, so a UI can show progress. Yields everything
    process_invoice_folders() yields, plus:

      {"type": "run_ready", "run": {...}}
        -- FIRST, carrying every run path so the UI can display where
        output is going before any slow work starts.

      {"type": "msg_progress", "current", "total", "file_name"}
        -- one .msg parsed (extraction happens before PDF processing).

      {"type": "msg_done", "parsed", "blocked", "failed", "results"}
        -- extraction finished, PDF processing about to begin.

      {"type": "error", "message"}
        -- nothing usable found; the run stops here.
    """
    msg_folder = msg_folder or msg_reader.DEFAULT_MSG_SOURCE
    run = msg_reader.make_run_folders(results_root, label=label)
    yield {"type": "run_ready", "run": run}

    source = Path(msg_folder)
    if not source.is_dir():
        yield {"type": "error", "message": f"{msg_folder} is not a directory"}
        return

    msg_files = sorted(p for p in source.iterdir()
                      if p.is_file() and p.suffix.lower() in msg_reader.SUPPORTED_EMAIL_TYPES)
    if not msg_files:
        yield {"type": "error", "message": f"No .msg/.eml files in {msg_folder}"}
        return

    # Parsed one at a time rather than via scan_email_folder() so the UI gets
    # per-file progress instead of one long silent pause.
    msg_results = []
    for i, msg_file in enumerate(msg_files, start=1):
        yield {"type": "msg_progress", "current": i, "total": len(msg_files),
               "file_name": msg_file.name}
        try:
            msg_results.append(msg_reader.parse_email(
                str(msg_file), run["extracted"],
                blocked_root=run["blocked"],
                move_blocked_msg=move_blocked_msg))
        except Exception as e:
            log.warning("Failed to parse %s: %s", msg_file.name, e)
            msg_results.append({"msg_path": str(msg_file), "error": str(e),
                               "blocked": False, "block_reasons": []})

    blocked = sum(1 for r in msg_results if r.get("blocked"))
    failed = sum(1 for r in msg_results if r.get("error"))
    yield {"type": "msg_done",
           "parsed": len(msg_results) - blocked - failed,
           "blocked": blocked, "failed": failed, "results": msg_results}

    yield from _run_extracted(run, run["extracted"])


def run_from_latest_extracted(results_root: Optional[str] = None,
                             date: Optional[str] = None):
    """Re-process the most recent run's already-extracted attachments,
    without re-reading the .msg files. Same events as
    run_from_msg_source(), minus the msg_* ones."""
    run = msg_reader.find_latest_run(results_root, date=date)
    if run is None:
        day = date or "today"
        yield {"type": "error",
               "message": f"No run folder found for {day} under "
                          f"{results_root or msg_reader.DEFAULT_RESULTS_ROOT}"}
        return

    yield {"type": "run_ready", "run": run}
    Path(run["processed"]).mkdir(parents=True, exist_ok=True)
    yield from _run_extracted(run, run["extracted"])


if __name__ == "__main__":
    # ------------------------------------------------------------------
    # MODE picks what this script does. "single_pdf" is the default and
    # behaves exactly as before -- the loose-PDF testing workflow is
    # untouched.
    #
    #   "single_pdf"  one PDF file            -> process_invoice_pdf()
    #   "pdf_folder"  folder of subfolders    -> scan_pdf_folder()
    #                 holding PDFs                + process_invoice_folders()
    #                 (no msg_reader involved)
    #   "msg_source"  folder of .msg files    -> run_from_msg_source()
    #                 (extract + process, new run folder)
    #   "latest_run"  reuse today's latest    -> run_from_latest_extracted()
    #                 extracted output
    # ------------------------------------------------------------------
    MODE = "single_pdf"

    SINGLE_PDF = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\sample-invoices\LE100 PO.pdf"
    PDF_FOLDER = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\sample-invoices"
    MSG_FOLDER = msg_reader.DEFAULT_MSG_SOURCE
    RESULTS_ROOT = msg_reader.DEFAULT_RESULTS_ROOT

    def _drive(events):
        """Consume a run generator, printing progress as it goes."""
        final = None
        for event in events:
            kind = event["type"]
            if kind == "run_ready":
                run = event["run"]
                print(f"\nRun folder : {run['run_name']}")
                print(f"  extracted: {run['extracted']}")
                print(f"  blocked  : {run['blocked']}")
                print(f"  processed: {run['processed']}\n")
            elif kind == "msg_progress":
                print(f"  [msg {event['current']}/{event['total']}] {event['file_name']}")
            elif kind == "msg_done":
                print(f"\n  extracted: {event['parsed']} parsed, "
                      f"{event['blocked']} blocked, {event['failed']} failed\n")
            elif kind == "progress":
                print(f"  [pdf {event['current']}/{event['total']}] "
                      f"{event['folder_name']}/{event['file_name']}")
            elif kind == "folder_done":
                fr = event["folder_result"]
                n = sum(len(pr.get("invoices", [])) for pr in fr["pdf_results"])
                print(f"    {fr['name']}: {n} invoice(s)")
            elif kind == "quarantine":
                print(f"    QUARANTINED: {event['folder_name']} "
                      f"(original -> {event.get('copied_to')})")
            elif kind == "batch_done":
                b = event["batch"]
                print(f"    BATCH {b['batch_id']} -> {b['kofax_email']} "
                      f"({b['invoice_count']} invoice(s))")
            elif kind == "routing_emails_done":
                sent = sum(1 for b in event["batches"] if b.get("email_sent"))
                failed = len(event["batches"]) - sent
                print(f"    Routing emails: {sent} sent, {failed} failed")
            elif kind == "error":
                print(f"\nERROR: {event['message']}")
            elif kind == "complete":
                final = event
        if final:
            total = sum(len(pr.get("invoices", []))
                       for fr in final["results"] for pr in fr["pdf_results"])
            print(f"\nDONE: {total} invoice(s) across "
                  f"{len(final['results'])} folder(s)")
            if final.get("report_path"):
                print(f"Report: {final['report_path']}")
            if final.get("batch_report_path"):
                print(f"Batch report: {final['batch_report_path']}")
            if final.get("email_stats"):
                s = final["email_stats"]
                print(f"Emails: {s['total_successful']} successful, "
                      f"{s['total_failed']} failed/blocked "
                      f"(of {s['total_invoices']} total)")
        return final

    if MODE == "single_pdf":
        pdf_path = sys.argv[1] if len(sys.argv) > 1 else SINGLE_PDF
        result = process_invoice_pdf(pdf_path)
        print(json.dumps(result, indent=2, ensure_ascii=False))

        # Excel Report
        # from report_builder import flatten_single_pdf_result, save_excel_report
        # report_rows = flatten_single_pdf_result(result)
        # if report_rows:
        #     report_path = str(Path(result["output_dir"]) / "invoice_report.xlsx")
        #     save_excel_report(report_rows, report_path)
        #     print(f"\nExcel report written to: {report_path}")
        # else:
        #     print("\nNo invoices extracted -- no report written.")

    elif MODE == "pdf_folder":
        source = sys.argv[1] if len(sys.argv) > 1 else PDF_FOLDER
        folders = scan_pdf_folder(source)
        print(f"Found {len(folders)} folder(s) with PDFs under {source}")
        out_dir = str(Path(source) / "_processed")
        final = _drive(process_invoice_folders(folders, out_dir))
        if final and final.get("report_bytes"):
            report_path = Path(out_dir) / "invoice_report.xlsx"
            report_path.write_bytes(final["report_bytes"])
            print(f"Report: {report_path}")

    elif MODE == "msg_source":
        source = sys.argv[1] if len(sys.argv) > 1 else MSG_FOLDER
        _drive(run_from_msg_source(source, RESULTS_ROOT))

    elif MODE == "latest_run":
        _drive(run_from_latest_extracted(RESULTS_ROOT))

    else:
        print(f"Unknown MODE {MODE!r} -- expected one of: "
              f"single_pdf, pdf_folder, msg_source, latest_run")