"""
routing.py -- Batch ID generation, all-or-nothing quarantine, and physical
staging for outbound Kofax batches.

Runs AFTER a run's invoices are already processed (i.e. after
process_invoice_folders()'s "complete" event, or the equivalent for a
.msg-source run). It consumes each invoice's already-computed
`kofax_email`, `le`, `needs_manual_review`, and `review_reasons` -- it does
NOT redo Kofax routing or LE matching.

===============================================================================
BATCH ID FORMAT: <LE>.<ConfigCode><DayLetter><Counter>
===============================================================================
e.g. LE100.GH01  (LE=LE100, ConfigCode=G, DayLetter=H (Thursday), Counter=01)
     LE420.LW11-TELCO-DD  (a batch containing at least one TELECOM-tagged
                           invoice AND at least one DD-tagged invoice)

  ConfigCode  -- fixed constant below (CONFIG_CODE). Per Vamsi's explicit
                 call, NOT read from an external config file/sheet for now.
                 If the real value ever needs to vary, this is the one
                 place to change it.

  DayLetter   -- Mon=M Tue=T Wed=W Thu=H Fri=F Sat=S Sun=U.

  Counter     -- 2-digit, zero-padded (grows past 2 digits naturally if it
                 ever exceeds 99 in a day -- f"{n:02d}" is a MINIMUM width,
                 not a truncation). Shared GLOBALLY across every LE and
                 every batch within one RUN. The first batch of the run
                 gets 01; each SUBSEQUENT batch's counter = the previous
                 batch's counter + the previous batch's invoice count --
                 so the counter tracks each invoice's cumulative position
                 in the day's outbound stream, not a simple batch-sequence
                 number.

  Tag suffix  -- optional, appended in a FIXED order: -TELCO, then -DD,
                 then -SENSITIVE. A batch can contain several invoices,
                 each with its own independent tags (from tags.py's
                 derive_tags(), stored per-invoice as invoice["tags"]) and
                 its own independent invoice["is_sensitive"] flag from the
                 sensitive-supplier matcher. The batch-level rule (Vamsi's
                 explicit call): a batch gets -TELCO if ANY of its invoices
                 carries the TELECOM tag, -DD if ANY carries DD, -SENSITIVE
                 if ANY is_sensitive -- these are independent checks and
                 ALL applicable suffixes are appended together (a batch can
                 be "-TELCO-DD-SENSITIVE" all at once). UTILITY is
                 deliberately NOT part of the Batch ID suffix -- only
                 TELECOM, DD, and SENSITIVE are.
                 Requires main.py to populate invoice["tags"] via
                 tags.derive_tags() -- see main.py's wiring.

  ACCEPTED TRADEOFF -- per Vamsi's explicit call: the counter resets to 01
  at the START OF EVERY RUN rather than persisting across multiple runs on
  the same calendar day (which the written spec, taken literally, calls
  for). A second run the same day restarts at 01 rather than continuing
  where the first run left off. If this becomes a real problem, the fix is
  to swap BatchCounter's in-memory state for a small JSON file at
  <results_root>/<date>/batch_counter.json, read-incremented-written per
  batch instead of held in a Python object -- process_batches() below is
  the only place that would need to change.

===============================================================================
GROUPING
===============================================================================
Invoices are grouped by (kofax_email, le) -- but ONLY within a single
source-email folder. The same (kofax_email, le) combination appearing in a
LATER, different source email in the same run gets its OWN new Batch ID; it
is never merged with an earlier batch. This is what makes "one source email
can produce several comma-joined Batch IDs" (see batch_report_builder.py)
meaningful -- a batch is 1:1 with (source_email, kofax_email, le).

===============================================================================
QUARANTINE (all-or-nothing per source email)
===============================================================================
If ANY invoice from a source-email folder has needs_manual_review == True
(or a whole-PDF processing error occurred for that folder), the ENTIRE
folder is excluded from batching:
  - none of its invoices are staged into 04_batches/
  - the original .msg/.eml is COPIED (not moved) into the run's blocked
    folder (run["blocked"], i.e. 02_blocked/ -- reusing the same folder
    email_reader.py's own password/no-PDF blocking already uses)
  - the already-split PDFs are left exactly where they are in 03_processed/
    for forensic tracking -- nothing is deleted or moved there
  - every invoice in that folder becomes a row in the Failures sheet

===============================================================================
PHYSICAL STAGING
===============================================================================
04_batches/<Kofax_Email>/<Batch_ID>/<pdf files>
under the run root, alongside the existing 01_extracted / 02_blocked /
03_processed folders.
"""

import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import email_reader

log = logging.getLogger("routing")

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
# Fixed per Vamsi's decision -- NOT read from an external config source.
# If this ever needs to become configurable, this is the one line to change.
CONFIG_CODE = "G"

RUN_BATCHES_DIR = "04_batches"

_DAY_LETTERS = {0: "M", 1: "T", 2: "W", 3: "H", 4: "F", 5: "S", 6: "U"}  # Mon=0


def day_letter(dt: Optional[datetime] = None) -> str:
    return _DAY_LETTERS[(dt or datetime.now()).weekday()]


def _sanitize_folder_name(text: str) -> str:
    """Make an email address (or anything else) safe as a Windows/POSIX
    folder name."""
    return re.sub(r'[<>:"/\\|?*]', "_", (text or "unknown").strip())


# ---------------------------------------------------------------------------
# Batch ID counter
# ---------------------------------------------------------------------------
class BatchCounter:
    """Global-within-a-run counter. See module docstring for the exact
    increment rule and the per-run-reset tradeoff."""

    def __init__(self, start: int = 1):
        self.next_counter = start

    def next_id(self, le: str, letter: str, suffix: str = "") -> str:
        return f"{le}.{CONFIG_CODE}{letter}{self.next_counter:02d}{suffix}"

    def advance(self, invoice_count: int) -> None:
        """Call once a batch's invoice list is finalized. The NEXT batch's
        counter = this counter + THIS batch's invoice count (not +1)."""
        self.next_counter += invoice_count


# ---------------------------------------------------------------------------
# Quarantine pass
# ---------------------------------------------------------------------------
def _folder_needs_quarantine(folder_result: dict) -> bool:
    for pr in folder_result["pdf_results"]:
        if pr.get("error"):
            return True
        for inv in pr.get("invoices", []):
            if inv.get("needs_manual_review"):
                return True
    return False


def count_invoices(folder_result: dict) -> int:
    """Total invoice count across every PDF in one source-email folder.
    Shared by batch_report_builder.py's Success-sheet 'Total' column and
    by email_sender.py's run-level stats, so there's one place that knows
    how to count invoices in a folder_result."""
    return sum(len(pr.get("invoices", [])) for pr in folder_result["pdf_results"])


def find_quarantined_folders(
    folder_results: List[dict],
) -> Tuple[List[dict], List[dict]]:
    """Split a run's folder_results into (clean, quarantined)."""
    clean, quarantined = [], []
    for fr in folder_results:
        (quarantined if _folder_needs_quarantine(fr) else clean).append(fr)
    return clean, quarantined


def quarantine_source(folder_result: dict, blocked_root: str) -> dict:
    """Copy the original .msg/.eml for one quarantined source-email folder
    into `blocked_root`. Does NOT touch the already-split PDFs in
    03_processed/ -- those are left in place for forensic tracking.

    Returns {"folder_name", "msg_path", "copied_to"} -- copied_to is None
    if no original email file could be located (e.g. a loose-PDF-folder
    run with no email_metadata.json at all). Callers in process_batches()
    add an "invoice_count" key onto this dict afterward.
    """
    meta = email_reader.load_email_metadata(folder_result["source_path"])
    msg_path = meta.get("msg_path") if meta else None

    copied_to = None
    if msg_path and Path(msg_path).is_file():
        dest_dir = Path(blocked_root)
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / Path(msg_path).name
        n = 2
        while dest.exists():
            dest = dest_dir / f"{Path(msg_path).stem}_{n}{Path(msg_path).suffix}"
            n += 1
        shutil.copy2(msg_path, dest)
        copied_to = str(dest)
    else:
        log.warning(
            "Quarantine: no original email file found for folder %r "
            "(msg_path=%r) -- only the Failures sheet will reflect this.",
            folder_result["name"], msg_path,
        )

    return {
        "folder_name": folder_result["name"],
        "msg_path": msg_path,
        "copied_to": copied_to,
    }


def build_failure_rows(folder_result: dict) -> List[dict]:
    """One row per failed invoice (or per whole-PDF processing error)
    inside a quarantined source-email folder. Keys match what
    batch_report_builder.py's Failures sheet expects."""
    rows: List[dict] = []
    source_name = folder_result["name"]

    for pr in folder_result["pdf_results"]:
        if pr.get("error"):
            rows.append({
                "source_email": source_name,
                "invoice_number": None,
                "supplier": None,
                "failure_reason": f"pdf_processing_error: {pr['error']}",
            })
            continue

        for inv in pr.get("invoices", []):
            if not inv.get("needs_manual_review"):
                continue
            reasons = inv.get("review_reasons") or []
            if not reasons:
                rows.append({
                    "source_email": source_name,
                    "invoice_number": inv.get("invoice_number"),
                    "supplier": inv.get("supplier_name"),
                    "failure_reason": "needs_manual_review (no reason recorded)",
                })
                continue
            for r in reasons:
                rows.append({
                    "source_email": source_name,
                    "invoice_number": inv.get("invoice_number"),
                    "supplier": inv.get("supplier_name"),
                    "failure_reason": r.get("code") or r.get("message") or str(r),
                })
    return rows


def _batch_tag_suffix(invoices: List[dict]) -> str:
    """Collapse the tags of every invoice in ONE batch into a single Batch
    ID suffix. See module docstring's "Tag suffix" section for the exact
    rule (any-invoice-triggers, all applicable suffixes appended, fixed
    order TELCO/DD/SENSITIVE, UTILITY excluded)."""
    has_telco = any("TELECOM" in (inv.get("tags") or []) for inv in invoices)
    has_dd = any("DD" in (inv.get("tags") or []) for inv in invoices)
    has_sensitive = any(inv.get("is_sensitive") for inv in invoices)

    suffix = ""
    if has_telco:
        suffix += "-TELCO"
    if has_dd:
        suffix += "-DD"
    if has_sensitive:
        suffix += "-SENSITIVE"
    return suffix


# ---------------------------------------------------------------------------
# Grouping + batch construction
# ---------------------------------------------------------------------------
def group_invoices_by_destination(
    folder_result: dict,
) -> Dict[Tuple[Optional[str], Optional[str]], List[dict]]:
    """Group every invoice in ONE (already-clean) source-email folder by
    (kofax_email, le). Never merged across folders -- see module docstring."""
    groups: Dict[Tuple[Optional[str], Optional[str]], List[dict]] = {}
    for pr in folder_result["pdf_results"]:
        for inv in pr.get("invoices", []):
            key = (inv.get("kofax_email"), inv.get("le"))
            groups.setdefault(key, []).append(inv)
    return groups


def build_batches(
    clean_folders: List[dict],
    counter: BatchCounter,
    letter: str,
    batches_root: str,
) -> Tuple[List[dict], Dict[str, List[str]]]:
    """Walk every clean source-email folder (in order) and, for each
    (kofax_email, le) group inside it:
      - assign a fresh Batch ID
      - physically COPY each invoice's split PDF into
        04_batches/<kofax_email>/<batch_id>/
      - stamp `batch_id` onto each invoice dict in place

    Returns (batches, email_to_batch_ids) --
      batches: [{"batch_id", "kofax_email", "le", "source_email",
                 "invoice_count", "staged_dir"}, ...]
      email_to_batch_ids: {source_email_folder_name: [batch_id, ...]}
        for the Success sheet's comma-joined "Email Name" column.
    """
    batches: List[dict] = []
    email_to_batch_ids: Dict[str, List[str]] = {}

    for fr in clean_folders:
        groups = group_invoices_by_destination(fr)
        # Deterministic order within a folder so re-runs are reproducible.
        for key in sorted(groups.keys(), key=lambda k: (k[0] or "", k[1] or "")):
            kofax_email, le = key
            invs = groups[key]

            if not kofax_email or not le:
                log.warning(
                    "Skipping batch for folder %r: incomplete routing "
                    "(kofax_email=%r, le=%r) on a folder that wasn't "
                    "flagged for review -- check the review_reasons logic.",
                    fr["name"], kofax_email, le,
                )
                continue

            suffix = _batch_tag_suffix(invs)
            batch_id = counter.next_id(le, letter, suffix)
            staged_dir = Path(batches_root) / _sanitize_folder_name(kofax_email) / batch_id
            staged_dir.mkdir(parents=True, exist_ok=True)

            for inv in invs:
                src = inv.get("output_pdf")
                if src and Path(src).is_file():
                    shutil.copy2(src, staged_dir / Path(src).name)
                else:
                    log.warning("Batch %s: missing output PDF for invoice %r",
                               batch_id, inv.get("invoice_number"))
                inv["batch_id"] = batch_id

            counter.advance(len(invs))

            batches.append({
                "batch_id": batch_id,
                "kofax_email": kofax_email,
                "le": le,
                "source_email": fr["name"],
                "invoice_count": len(invs),
                "staged_dir": str(staged_dir),
            })
            email_to_batch_ids.setdefault(fr["name"], []).append(batch_id)

    return batches, email_to_batch_ids


# ---------------------------------------------------------------------------
# Top-level entry point
# ---------------------------------------------------------------------------
def process_batches(all_results: List[dict], run: dict):
    """Generator. Call once, after a run's invoice processing is fully
    complete, with the full list of folder_result dicts.

    Yields:
      {"type": "quarantine", "folder_name", "copied_to"}
        -- one source-email folder was quarantined.
      {"type": "batch_done", "batch": {...}}
        -- one batch finished staging.
      {"type": "complete", "batches", "quarantined", "failure_rows",
       "email_to_batch_ids", "clean_folders"}
        -- everything is done; feed this straight into
        batch_report_builder.flatten_success_rows() /
        flatten_failure_rows().
    """
    batches_root = Path(run["run_root"]) / RUN_BATCHES_DIR
    letter = day_letter()
    counter = BatchCounter()

    clean, quarantined = find_quarantined_folders(all_results)

    failure_rows: List[dict] = []
    quarantine_log: List[dict] = []
    for fr in quarantined:
        q = quarantine_source(fr, run["blocked"])
        q["invoice_count"] = count_invoices(fr)
        quarantine_log.append(q)
        failure_rows.extend(build_failure_rows(fr))
        yield {"type": "quarantine", "folder_name": fr["name"],
               "copied_to": q["copied_to"]}

    batches, email_to_batch_ids = build_batches(clean, counter, letter, str(batches_root))
    for b in batches:
        yield {"type": "batch_done", "batch": b}

    yield {
        "type": "complete",
        "batches": batches,
        "quarantined": quarantine_log,
        "failure_rows": failure_rows,
        "email_to_batch_ids": email_to_batch_ids,
        "clean_folders": clean,
    }