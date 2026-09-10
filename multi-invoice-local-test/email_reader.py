"""
msg_reader.py -- Turn a folder of email (.msg and .eml) files into per-message working
folders that the existing invoice pipeline can consume.

For each email file, creates:

    <output_root>/<msg_folder_name>/
        email_metadata.json     sender, subject, date, attachment inventory
        email_body.txt          full body text (kept separate -- bodies can be
                                long and make the JSON unpleasant to read)
        <attachments...>

ATTACHMENT LAYOUT -- two cases:

  1. SIMPLE (at most one PDF): attachments are written FLAT into the message
     folder, alongside the two metadata files. No subfolders. This is the
     common case and keeps things shallow.

  2. MULTIPLE PDFs: attachments are grouped into one subfolder per PDF,
     named after that PDF's stem, so each PDF sits with its own supporting
     Excel/XML:
         <msg_folder>/INV001/INV001.pdf
         <msg_folder>/INV001/INV001.xlsx
     A non-PDF attachment is placed with a PDF when their filename stems
     match (normalized, or by prefix -- so "INV001_supporting.xlsx" pairs
     with "INV001.pdf"). Anything that can't be paired is left flat in the
     message folder, outside every per-PDF folder.

WHY GROUP AT ALL: excel_attachment.py finds the Excel to attach by scanning
the PDF's OWN folder, and deliberately refuses to guess when it sees more
than one Excel there. So several PDFs and several Excels sitting flat in one
folder means nothing gets attached at all. Grouping gives each PDF its own
unambiguous folder.

ZIP ATTACHMENTS: .zip attachments are extracted automatically and their
contents treated as ordinary attachments. Nested zips are followed up to
MAX_ZIP_DEPTH. Subfolders INSIDE the archive are preserved and become the
grouping directly -- a zip laid out as INV001/invoice.pdf +
INV001/data.xlsx already says which files belong together, so that beats
guessing from filenames. Every path component is sanitized and '..' /
absolute-root parts are dropped, which blocks zip-slip traversal. Nested
paths are collapsed to a single folder name (a/b -> a_b) so the layout
never goes deeper than <msg_folder>/<group>/<file>.

BLOCKED MESSAGES: if a message carries a password-protected zip, a
password-protected PDF, or NO PDF AT ALL, the whole message is BLOCKED -- it is written to
<output_root>/_blocked/<msg_folder_name>/ instead of the normal location,
-- nothing is extracted or written for it. The ONLY thing that lands in
<output_root>/_blocked/ is the original email file itself, so that folder
stays a flat pile of messages ready to be dragged into an Outlook folder
for manual handling. Pass move_blocked_msg=True to move the email out of
the source folder rather than copy it. The returned dict carries
"blocked": True, "block_reasons": [...] and "blocked_msg_copy".
  Note on PDFs: pypdf reports is_encrypted=True for PDFs that are merely
  permission-restricted (no printing/copying) yet open fine with an empty
  password. Those are NOT blocked -- only PDFs that genuinely fail to
  decrypt with an empty password are.

Metadata files (.json/.txt) are ignored by the rest of the pipeline --
scan_source_folder() and excel_attachment.py both filter by extension
(.pdf / .xlsx|.xlsm|.xls), so neither picks them up.

RUN FOLDERS: make_run_folders() creates a dated, timestamped folder per
run (<results_root>/<YYYY-MM-DD>/run_<HH-MM-SS>/) holding 01_extracted/,
02_blocked/ and 03_processed/, so a second run on the same day never
overwrites the first. DEFAULT_MSG_SOURCE / DEFAULT_RESULTS_ROOT live here
too, since both the pipeline and the UI need the same values.

Requires: pip install extract-msg pypdf
"""

import json
import logging
import re
import shutil
import zipfile
from datetime import datetime
from email import policy
from email.parser import BytesParser
import email.utils
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import extract_msg
from pypdf import PdfReader

log = logging.getLogger("msg_reader")

SUPPORTED_EMAIL_TYPES = {".msg", ".eml"}

_PDF_EXT = {".pdf"}
_EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}
_XML_EXT = {".xml"}
_ZIP_EXT = {".zip"}

BLOCKED_DIR = "_blocked"

# How many levels of zip-inside-zip to follow before giving up.
MAX_ZIP_DEPTH = 3

# Shortest normalized filename stem allowed to take part in PREFIX pairing
# (exact matches are always allowed). Guards against a one- or two-char PDF
# stem prefix-matching every unrelated attachment that happens to start
# with the same letters.
_MIN_STEM_FOR_PREFIX = 4


# ---------------------------------------------------------------------------
# Defaults + per-run output folders
# ---------------------------------------------------------------------------
# Shared by the pipeline (main.py) and the UI (app.py) so both show and use
# the same paths -- keeping one copy avoids the two drifting apart.

# Folder the email files are picked up from.
DEFAULT_MSG_SOURCE = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\msg-test"

# Root under which every run's dated output folder is created.
DEFAULT_RESULTS_ROOT = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\msg-result"

# Subfolder names inside a single run. Named RUN_* to keep them clearly
# apart from BLOCKED_DIR above, which is something different: the fallback
# blocked folder used when parse_email is called WITHOUT an explicit
# blocked_root.
RUN_EXTRACTED_DIR = "01_extracted"
RUN_BLOCKED_DIR = "02_blocked"
RUN_PROCESSED_DIR = "03_processed"
RUN_REPORT_NAME = "invoice_report.xlsx"  # legacy default; actual
# filenames now go through report_filename() below, which adds the
# run's own date+time so a report file pulled out of its run folder
# (e.g. a Streamlit download, or emailed onward) still identifies
# which run it came from by its filename alone.

def report_filename(prefix: str, date_dir_name: str, run_root_name: str) -> str:
    """'<prefix>_<YYYY-MM-DD>_<HH-MM-SS[_label]>.xlsx', built from the
    run folder's OWN name components -- NEVER datetime.now() -- so
    make_run_folders(), find_latest_run(), and list_runs() all produce
    the IDENTICAL filename for the same run. Re-inspecting an old run
    later must reproduce that run's original timestamp, not today's.
    """
    time_part = (run_root_name[len("run_"):]
                if run_root_name.startswith("run_") else run_root_name)
    return f"{prefix}_{date_dir_name}_{time_part}.xlsx"


def find_latest_run(results_root: Optional[str] = None,
                   date: Optional[str] = None) -> Optional[dict]:
    """Find the most recent run folder and return the same path dict that
    make_run_folders() produces (so callers can use either interchangeably)."""
    root = Path(results_root or DEFAULT_RESULTS_ROOT)
    day = date or datetime.now().strftime("%Y-%m-%d")
    date_dir = root / day

    if not date_dir.is_dir():
        return None

    runs = sorted((d for d in date_dir.iterdir()
                  if d.is_dir() and d.name.startswith("run_")),
                 key=lambda d: d.name)
    if not runs:
        return None

    run_root = runs[-1]
    return {
        "results_root": str(root),
        "date_dir": str(date_dir),
        "run_root": str(run_root),
        "extracted": str(run_root / RUN_EXTRACTED_DIR),
        "blocked": str(run_root / RUN_BLOCKED_DIR),
        "processed": str(run_root / RUN_PROCESSED_DIR),
        "report_path": str(run_root / report_filename("invoice_report", date_dir.name, run_root.name)),
        "run_name": f"{date_dir.name}/{run_root.name}",
    }


def list_runs(results_root: Optional[str] = None,
             date: Optional[str] = None) -> List[dict]:
    """Every run folder for a day, oldest first."""
    root = Path(results_root or DEFAULT_RESULTS_ROOT)
    day = date or datetime.now().strftime("%Y-%m-%d")
    date_dir = root / day

    if not date_dir.is_dir():
        return []

    out = []
    for d in sorted((d for d in date_dir.iterdir()
                    if d.is_dir() and d.name.startswith("run_")),
                   key=lambda d: d.name):
        out.append({
            "results_root": str(root),
            "date_dir": str(date_dir),
            "run_root": str(d),
            "extracted": str(d / RUN_EXTRACTED_DIR),
            "blocked": str(d / RUN_BLOCKED_DIR),
            "processed": str(d / RUN_PROCESSED_DIR),
            "report_path": str(d / report_filename("invoice_report", date_dir.name, d.name)),
            "run_name": f"{date_dir.name}/{d.name}",
        })
    return out


def make_run_folders(results_root: Optional[str] = None,
                    label: Optional[str] = None,
                    create: bool = True) -> dict:
    """Build (and by default create) this run's output folders."""
    root = Path(results_root or DEFAULT_RESULTS_ROOT)

    now = datetime.now()
    date_dir = root / now.strftime("%Y-%m-%d")

    run_name = f"run_{now.strftime('%H-%M-%S')}"
    if label:
        safe_label = "".join(c if (c.isalnum() or c in "-_") else "_"
                            for c in str(label).strip())
        if safe_label:
            run_name = f"{run_name}_{safe_label}"

    run_root = date_dir / run_name
    base, n = run_root, 2
    while run_root.exists():
        run_root = Path(f"{base}_{n}")
        n += 1

    paths = {
        "results_root": str(root),
        "date_dir": str(date_dir),
        "run_root": str(run_root),
        "extracted": str(run_root / RUN_EXTRACTED_DIR),
        "blocked": str(run_root / RUN_BLOCKED_DIR),
        "processed": str(run_root / RUN_PROCESSED_DIR),
        "report_path": str(run_root / report_filename("invoice_report", date_dir.name, run_root.name)),
        "run_name": f"{date_dir.name}/{run_root.name}",
    }

    if create:
        for key in ("extracted", "blocked", "processed"):
            Path(paths[key]).mkdir(parents=True, exist_ok=True)

    return paths

# Filenames Outlook adds that are never real invoice attachments.
_SKIP_ATTACHMENT_NAMES = {"image001.png", "image002.png", "image003.png",
                          "oledata.mso", "winmail.dat"}


# ---------------------------------------------------------------------------
# EML parsing
# ---------------------------------------------------------------------------
def load_eml(path: Path) -> dict:
    """Extract standard metadata and attachments from an .eml file."""
    with open(path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)

    subject = msg.get('subject', '')
    sender = msg.get('from', '')

    # Normalize the date to the SAME ISO format the .msg path already uses
    # (msg_obj.date.isoformat()). Left as the raw RFC-2822 header string
    # ("Wed, 09 Sep 2026 10:30:00 +0000"), received_date would be in two
    # different formats depending on which file type produced it -- fine
    # for a human reading the JSON, but a trap for anything that later
    # tries to parse or sort on this field.
    date_str = ""
    raw_date = msg.get('date', '')
    if raw_date:
        try:
            date_str = email.utils.parsedate_to_datetime(raw_date).isoformat()
        except (TypeError, ValueError):
            date_str = str(raw_date)   # keep SOMETHING rather than drop it

    body = ""
    body_part = msg.get_body(preferencelist=('plain',))
    if body_part:
        try:
            body = body_part.get_content()
        except Exception:
            body = str(body_part.get_payload(decode=True))
    elif not msg.is_multipart():
        try:
            body = msg.get_content()
        except Exception:
            body = str(msg.get_payload(decode=True))

    attachments = []
    for part in msg.iter_attachments():
        # iter_attachments() returns BOTH real attachments AND inline
        # parts (a signature logo, a tracking pixel) -- the .msg path
        # filters these by filename (_SKIP_ATTACHMENT_NAMES), but that
        # list assumes Outlook's naming convention, which non-Outlook
        # senders (Gmail, other mail clients) won't necessarily follow.
        # Content-Disposition is the reliable signal instead: skip
        # anything that isn't explicitly "attachment".
        if part.get_content_disposition() != "attachment":
            continue
        filename = part.get_filename()
        if filename:
            data = part.get_payload(decode=True)
            if data is not None:
                attachments.append({
                    "filename": filename,
                    "data": data
                })

    return {
        "subject": subject,
        "sender": sender,
        "date": date_str,
        "body": body,
        "attachments": attachments
    }


# ---------------------------------------------------------------------------
# Filename helpers
# ---------------------------------------------------------------------------
def _sanitize(text: Optional[str], fallback: str, max_len: int = 60) -> str:
    if not text or not str(text).strip():
        return fallback
    safe = re.sub(r"[^A-Za-z0-9._\- ]+", "_", str(text).strip())
    safe = re.sub(r"[\s_]+", "_", safe).strip("._-")
    return safe[:max_len] or fallback


def _sanitize_filename(text: Optional[str], fallback: str,
                      max_len: int = 60) -> str:
    if not text or not str(text).strip():
        return fallback

    raw = str(text).strip()
    suffix = Path(raw).suffix
    if not re.fullmatch(r"\.[A-Za-z0-9]{1,10}", suffix):
        suffix = ""
    stem = raw[:len(raw) - len(suffix)] if suffix else raw

    safe_stem = re.sub(r"[^A-Za-z0-9._\- ]+", "_", stem)
    safe_stem = re.sub(r"[\s_]+", "_", safe_stem).strip("._-")
    safe_suffix = re.sub(r"[^A-Za-z0-9.]+", "", suffix)

    if not safe_stem:
        safe_stem = Path(fallback).stem or "file"

    keep = max(1, max_len - len(safe_suffix))
    return f"{safe_stem[:keep]}{safe_suffix}"


def _norm_stem(stem: str) -> str:
    return re.sub(r"[^a-z0-9]", "", stem.lower())


def _unique_path(folder: Path, name: str) -> Path:
    candidate = folder / name
    if not candidate.exists():
        return candidate
    stem, suffix = Path(name).stem, Path(name).suffix
    n = 2
    while (folder / f"{stem}_{n}{suffix}").exists():
        n += 1
    return folder / f"{stem}_{n}{suffix}"


def _attachment_filename(att, index: int) -> str:
    for attr in ("longFilename", "shortFilename", "displayName"):
        value = getattr(att, attr, None)
        if value:
            return str(value).strip()
    return f"attachment_{index}"


# ---------------------------------------------------------------------------
# Sender extraction
# ---------------------------------------------------------------------------
def _parse_sender(raw_sender: Optional[str]) -> Dict[str, str]:
    if not raw_sender:
        return {"sender_name": "", "sender_email": ""}

    raw = str(raw_sender).strip()

    angled = re.search(r"<([^>]+)>", raw)
    if angled:
        email = angled.group(1).strip()
        name = raw[:angled.start()].strip().strip('"').strip()
        return {"sender_name": name, "sender_email": email}

    bare = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", raw)
    if bare:
        email = bare.group(0)
        name = raw.replace(email, "").strip().strip('"').strip("<>").strip()
        return {"sender_name": name, "sender_email": email}

    return {"sender_name": raw.strip('"'), "sender_email": ""}


# ---------------------------------------------------------------------------
# Zip extraction + password/lock detection
# ---------------------------------------------------------------------------
def _zip_block_reason(zip_path: Path) -> Optional[str]:
    try:
        with zipfile.ZipFile(zip_path) as zf:
            if any(zi.flag_bits & 0x1 for zi in zf.infolist()):
                return f"password-protected zip: {zip_path.name}"
        return None
    except zipfile.BadZipFile:
        return f"unreadable or corrupt zip: {zip_path.name}"
    except Exception as e:
        return f"could not open zip {zip_path.name}: {e}"


def _pdf_block_reason(pdf_path: Path) -> Optional[str]:
    try:
        reader = PdfReader(str(pdf_path))
        if reader.is_encrypted:
            try:
                if not reader.decrypt(""):
                    return f"password-protected PDF: {pdf_path.name}"
            except Exception:
                return f"password-protected PDF: {pdf_path.name}"
        return None
    except Exception as e:
        log.warning("  could not inspect PDF %s (%s) -- not blocking, "
                   "letting it through", pdf_path.name, e)
        return None


def _safe_relative_parts(member_name: str) -> Optional[Tuple[List[str], str, bool]]:
    raw = Path(member_name.replace("\\", "/"))
    unsafe = any(p in ("..", "/", "\\") for p in raw.parts)
    parts = [p for p in raw.parts if p not in ("..", "/", "\\", ".")]
    if not parts:
        return None
    filename = parts[-1]
    subdirs = [_sanitize(p, "folder") for p in parts[:-1]]
    return subdirs, filename, unsafe


def _extract_zip(zip_path: Path, dest_dir: Path,
                depth: int = 0) -> Tuple[List[Path], List[str]]:
    extracted: List[Path] = []
    reasons: List[str] = []

    if depth >= MAX_ZIP_DEPTH:
        return extracted, [f"zip nested deeper than {MAX_ZIP_DEPTH} levels: "
                          f"{zip_path.name}"]

    reason = _zip_block_reason(zip_path)
    if reason:
        return extracted, [reason]

    dest_dir.mkdir(parents=True, exist_ok=True)
    unsafe_members = []
    try:
        with zipfile.ZipFile(zip_path) as zf:
            for zi in zf.infolist():
                if zi.is_dir():
                    continue
                parsed = _safe_relative_parts(zi.filename)
                if parsed is None:
                    continue
                subdirs, filename, was_unsafe = parsed
                if was_unsafe:
                    unsafe_members.append(zi.filename)
                if filename.lower() in _SKIP_ATTACHMENT_NAMES:
                    continue
                target_dir = dest_dir.joinpath(*subdirs) if subdirs else dest_dir
                target_dir.mkdir(parents=True, exist_ok=True)
                target = _unique_path(target_dir,
                                     _sanitize_filename(filename, "extracted"))
                with zf.open(zi) as src, open(target, "wb") as out:
                    shutil.copyfileobj(src, out)
                extracted.append(target)
    except Exception as e:
        return extracted, [f"failed extracting {zip_path.name}: {e}"]

    if unsafe_members:
        reasons.append(
            f"unsafe path(s) in zip {zip_path.name} "
            f"(contained, not extracted outside): {unsafe_members[:3]}")

    for nested in [p for p in list(extracted) if p.suffix.lower() in _ZIP_EXT]:
        sub_paths, sub_reasons = _extract_zip(nested, nested.parent, depth + 1)
        reasons.extend(sub_reasons)
        extracted.remove(nested)
        nested.unlink(missing_ok=True)
        extracted.extend(sub_paths)

    return extracted, reasons


def _expand_and_check(saved: List[Path],
                     staging: Path) -> Tuple[List[Path], List[str]]:
    reasons: List[str] = []
    files: List[Path] = []

    for path in saved:
        if path.suffix.lower() in _ZIP_EXT:
            extracted, zip_reasons = _extract_zip(path, staging)
            reasons.extend(zip_reasons)
            files.extend(extracted)
            path.unlink(missing_ok=True)
        else:
            files.append(path)

    for path in files:
        if path.suffix.lower() in _PDF_EXT:
            pdf_reason = _pdf_block_reason(path)
            if pdf_reason:
                reasons.append(pdf_reason)

    return files, reasons


# ---------------------------------------------------------------------------
# Attachment grouping
# ---------------------------------------------------------------------------
def _group_attachments(saved: List[Path],
                      staging: Optional[Path] = None) -> Dict[str, List[Path]]:
    groups: Dict[str, List[Path]] = {}
    top_level: List[Path] = []

    for path in saved:
        rel_parent = Path(".")
        if staging is not None:
            try:
                rel_parent = path.parent.relative_to(staging)
            except ValueError:
                rel_parent = Path(".")
        if str(rel_parent) in (".", ""):
            top_level.append(path)
        else:
            group_name = _sanitize("_".join(rel_parent.parts), "group")
            groups.setdefault(group_name, []).append(path)

    pdfs = [p for p in top_level if p.suffix.lower() in _PDF_EXT]
    others = [p for p in top_level if p.suffix.lower() not in _PDF_EXT]

    if len(pdfs) <= 1:
        if top_level:
            groups.setdefault("", []).extend(top_level)
        return groups

    pdf_norms = []
    for pdf in pdfs:
        folder = _sanitize(pdf.stem, "invoice")
        base, n = folder, 2
        while folder in groups:
            folder = f"{base}_{n}"
            n += 1
        groups[folder] = [pdf]
        pdf_norms.append((_norm_stem(pdf.stem), folder))

    for other in others:
        onorm = _norm_stem(other.stem)
        target = None
        for pnorm, folder in pdf_norms:
            if pnorm and onorm == pnorm:
                target = folder
                break
        if target is None:
            for pnorm, folder in pdf_norms:
                if not pnorm or not onorm:
                    continue
                if min(len(pnorm), len(onorm)) < _MIN_STEM_FOR_PREFIX:
                    continue
                if onorm.startswith(pnorm) or pnorm.startswith(onorm):
                    target = folder
                    break
        groups.setdefault(target or "", []).append(other)

    return groups


# ---------------------------------------------------------------------------
# Single Email Parsing (.msg / .eml)
# ---------------------------------------------------------------------------
def parse_email(email_path: str, output_root: str,
             blocked_root: Optional[str] = None,
             move_blocked_msg: bool = False) -> dict:
    """Parse one email file (.msg or .eml) into its own folder.

    Normally that folder is <output_root>/<folder_name>/. If the message
    is BLOCKED (password-protected zip or PDF, no PDF at all, or an unsafe
    zip path) NOTHING is extracted -- only the original email is placed in
    <blocked_root>/ (default <output_root>/_blocked/), flat, so that folder
    is just a pile of messages to drag into an Outlook folder.

    `move_blocked_msg=False` (the default) COPIES the email, leaving the
    source folder untouched. Pass True to MOVE it out of the source folder
    instead, which keeps re-runs clean but is destructive.
    """
    email_path = Path(email_path)
    output_root = Path(output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    blocked_root = Path(blocked_root) if blocked_root else output_root / BLOCKED_DIR

    ext = email_path.suffix.lower()
    if ext not in SUPPORTED_EMAIL_TYPES:
        raise ValueError(f"Unsupported email format: {ext}")

    staging = None
    msg_obj = None

    try:
        # --- Normalization Layer ---
        if ext == ".msg":
            msg_obj = extract_msg.Message(str(email_path))
            subject = (msg_obj.subject or "").strip()
            body = msg_obj.body or ""
            sender = _parse_sender(msg_obj.sender)

            received = msg_obj.date
            if isinstance(received, datetime):
                received_str = received.isoformat()
            else:
                received_str = str(received) if received else ""

            normalized_attachments = []
            for i, att in enumerate(msg_obj.attachments, start=1):
                normalized_attachments.append({
                    "filename": _attachment_filename(att, i),
                    "data": att.data,
                    "is_binary": isinstance(att.data, bytes)
                })

        elif ext == ".eml":
            eml_data = load_eml(email_path)
            subject = (eml_data.get("subject") or "").strip()
            body = eml_data.get("body") or ""
            sender = _parse_sender(eml_data.get("sender"))
            received_str = str(eml_data.get("date") or "")

            normalized_attachments = []
            for i, att in enumerate(eml_data.get("attachments", []), start=1):
                normalized_attachments.append({
                    "filename": att.get("filename") or f"attachment_{i}",
                    "data": att.get("data"),
                    "is_binary": isinstance(att.get("data"), bytes)
                })

        folder_name = _sanitize(subject, _sanitize(email_path.stem, "message"))

        # --- Staging & Attachment Processing ---
        staging = output_root / f"__staging_{folder_name}"
        base, n = staging, 2
        while staging.exists():
            staging = Path(f"{base}_{n}")
            n += 1
        staging.mkdir(parents=True)

        saved: List[Path] = []
        for i, att_dict in enumerate(normalized_attachments, start=1):
            filename = att_dict["filename"]
            if filename.lower() in _SKIP_ATTACHMENT_NAMES:
                continue
            
            if not att_dict["is_binary"]:
                log.info("  skipping non-binary attachment %r in %s",
                        filename, email_path.name)
                continue
                
            dest = _unique_path(staging,
                               _sanitize_filename(filename, f"attachment_{i}"))
            dest.write_bytes(att_dict["data"])
            saved.append(dest)

        # Expand zips, then check every PDF for a password.
        saved, block_reasons = _expand_and_check(saved, staging)

        # A message with no PDF at all has nothing for the pipeline to
        # process -- quarantine it for manual handling rather than
        # producing an empty folder that silently does nothing.
        if not any(p.suffix.lower() in _PDF_EXT for p in saved):
            block_reasons.append("no PDF attachment in this message")

        blocked = bool(block_reasons)

        # --- BLOCKED: keep ONLY the email file --------------------------
        if blocked:
            blocked_root.mkdir(parents=True, exist_ok=True)
            shutil.rmtree(staging, ignore_errors=True)
            staging = None

            msg_dest = _unique_path(blocked_root, email_path.name)
            try:
                if move_blocked_msg:
                    shutil.move(str(email_path), str(msg_dest))
                else:
                    shutil.copy2(str(email_path), str(msg_dest))
                msg_dest_str = str(msg_dest)
            except Exception as e:
                log.warning("  could not %s original email for %s: %s",
                           "move" if move_blocked_msg else "copy",
                           email_path.name, e)
                msg_dest_str = None

            log.warning("BLOCKED %s -> %s (%s)", email_path.name,
                       blocked_root, "; ".join(block_reasons))

            return {
                "msg_path": str(email_path),
                "folder_name": None,
                "output_folder": None,
                "sender_name": sender["sender_name"],
                "sender_email": sender["sender_email"],
                "subject": subject,
                "body": body,
                "received_date": received_str,
                "attachments": {"pdfs": [], "excels": [], "xmls": [], "other": []},
                "groups": {},
                "blocked": True,
                "block_reasons": block_reasons,
                "blocked_msg_copy": msg_dest_str,
                "error": None,
            }

        # --- NOT BLOCKED: normal folder + attachments ------------------
        output_root.mkdir(parents=True, exist_ok=True)
        base, n = folder_name, 2
        while (output_root / folder_name).exists():
            folder_name = f"{base}_{n}"
            n += 1
        out_folder = output_root / folder_name
        out_folder.mkdir(parents=True)

        groups = _group_attachments(saved, staging)
        final_groups: Dict[str, List[str]] = {}
        for folder_name_part, paths in groups.items():
            target_dir = out_folder / folder_name_part if folder_name_part else out_folder
            target_dir.mkdir(parents=True, exist_ok=True)
            moved = []
            for p in paths:
                dest = _unique_path(target_dir, p.name)
                shutil.move(str(p), str(dest))
                moved.append(str(dest))
            final_groups[folder_name_part] = moved

        shutil.rmtree(staging, ignore_errors=True)
        staging = None

        # Flat inventory by type, across every group.
        all_paths = [Path(p) for paths in final_groups.values() for p in paths]
        attachments = {
            "pdfs":   sorted(str(p) for p in all_paths if p.suffix.lower() in _PDF_EXT),
            "excels": sorted(str(p) for p in all_paths if p.suffix.lower() in _EXCEL_EXT),
            "xmls":   sorted(str(p) for p in all_paths if p.suffix.lower() in _XML_EXT),
            "other":  sorted(str(p) for p in all_paths
                            if p.suffix.lower() not in _PDF_EXT | _EXCEL_EXT | _XML_EXT),
        }

        metadata = {
            "msg_path": str(email_path),
            "folder_name": folder_name,
            "output_folder": str(out_folder),
            "sender_name": sender["sender_name"],
            "sender_email": sender["sender_email"],
            "subject": subject,
            "received_date": received_str,
            "attachments": attachments,
            "groups": final_groups,
            "blocked": False,
            "block_reasons": [],
            "blocked_msg_copy": None,
        }

        (out_folder / "email_metadata.json").write_text(
            json.dumps(metadata, indent=2, ensure_ascii=False), encoding="utf-8")
        (out_folder / "email_body.txt").write_text(body, encoding="utf-8")

        log.info("Parsed %s -> %s (%d pdf, %d excel, %d xml)",
                email_path.name, folder_name, len(attachments["pdfs"]),
                len(attachments["excels"]), len(attachments["xmls"]))

        result = dict(metadata)
        result["body"] = body
        result["error"] = None
        return result

    finally:
        if msg_obj is not None:
            msg_obj.close()
        if staging is not None:
            shutil.rmtree(staging, ignore_errors=True)


# ---------------------------------------------------------------------------
# Folder of Emails
# ---------------------------------------------------------------------------
def scan_email_folder(source_folder: str, output_root: str,
                   blocked_root: Optional[str] = None,
                   move_blocked_msg: bool = False) -> List[dict]:
    """Parse every supported email file in `source_folder` into `output_root`."""
    source_folder = Path(source_folder)
    if not source_folder.is_dir():
        raise NotADirectoryError(f"{source_folder} is not a directory")

    email_files = sorted(p for p in source_folder.iterdir()
                      if p.is_file() and p.suffix.lower() in SUPPORTED_EMAIL_TYPES)

    log.info("Found %d email file(s) in %s", len(email_files), source_folder)

    results = []
    for email_file in email_files:
        try:
            results.append(parse_email(str(email_file), str(output_root),
                                    blocked_root=blocked_root,
                                    move_blocked_msg=move_blocked_msg))
        except Exception as e:
            log.warning("Failed to parse %s: %s", email_file.name, e)
            results.append({
                "msg_path": str(email_file),
                "folder_name": None,
                "output_folder": None,
                "sender_name": "", "sender_email": "",
                "subject": "", "body": "", "received_date": "",
                "attachments": {"pdfs": [], "excels": [], "xmls": [], "other": []},
                "groups": {},
                "blocked": False,
                "block_reasons": [],
                "blocked_msg_copy": None,
                "error": str(e),
            })

    blocked = sum(1 for r in results if r.get("blocked"))
    failed = sum(1 for r in results if r.get("error"))
    log.info("Parsed %d email file(s): %d ok, %d blocked, %d failed",
            len(results), len(results) - blocked - failed, blocked, failed)
    return results


# ---------------------------------------------------------------------------
# Read back metadata for a folder the pipeline is processing
# ---------------------------------------------------------------------------
def load_email_metadata(folder: str) -> Optional[dict]:
    start = Path(folder)
    if start.is_file():
        start = start.parent

    for candidate in (start, start.parent):
        meta_file = candidate / "email_metadata.json"
        if meta_file.is_file():
            data = json.loads(meta_file.read_text(encoding="utf-8"))
            body_file = candidate / "email_body.txt"
            data["body"] = body_file.read_text(encoding="utf-8") if body_file.is_file() else ""
            return data
    return None


# ---------------------------------------------------------------------------
# Direct run
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%H:%M:%S",
    )

    MSG_FOLDER = DEFAULT_MSG_SOURCE
    RESULTS_ROOT = DEFAULT_RESULTS_ROOT

    run = make_run_folders(RESULTS_ROOT)
    results = scan_email_folder(MSG_FOLDER, run["extracted"],
                             blocked_root=run["blocked"])

    print()
    print("=" * 72)
    print(f"Parsed {len(results)} email file(s) from {MSG_FOLDER}")
    print(f"Run folder: {run['run_name']}")
    print("=" * 72)

    for r in results:
        print()
        if r.get("error"):
            print(f"  FAILED  {Path(r['msg_path']).name}: {r['error']}")
            continue

        att = r["attachments"]
        marker = "BLOCKED " if r.get("blocked") else ""
        print(f"  {marker}{Path(r['msg_path']).name}")
        print(f"    subject : {r['subject']!r}")
        print(f"    sender  : {r['sender_name']!r} <{r['sender_email']}>")
        print(f"    received: {r['received_date']}")

        if r.get("blocked"):
            for reason in r["block_reasons"]:
                print(f"    !! {reason}")
            if r.get("blocked_msg_copy"):
                print(f"    email   -> {r['blocked_msg_copy']}")
            print(f"    -> NOT processed; email quarantined for manual routing")
            continue

        print(f"    body    : {len(r['body'])} chars")
        print(f"    output  : {r['output_folder']}")
        print(f"    files   : {len(att['pdfs'])} pdf, {len(att['excels'])} excel, "
              f"{len(att['xmls'])} xml, {len(att['other'])} other")

        for folder_key, paths in r["groups"].items():
            where = folder_key or "(flat in message folder)"
            print(f"      {where}: {[Path(p).name for p in paths]}")

        text = f"{r['subject']} {r['body']}".lower()
        if "direct debit" in text:
            print(f"      -> 'direct debit' found (DD tag would apply)")

    blocked = sum(1 for r in results if r.get("blocked"))
    failed = sum(1 for r in results if r.get("error"))
    ok = len(results) - blocked - failed
    print()
    print(f"Done: {ok} parsed, {blocked} blocked, {failed} failed.")
    print(f"  run root  -> {run['run_root']}")
    print(f"  extracted -> {run['extracted']}")
    if blocked:
        print(f"  blocked   -> {run['blocked']}")