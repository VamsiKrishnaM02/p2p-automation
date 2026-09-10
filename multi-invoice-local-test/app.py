"""
app.py -- AI-Powered Invoice Routing System (Streamlit UI)

Reads a SOURCE folder containing subfolders (each subfolder holds one or
more invoice PDFs + optional Excel attachments), processes every PDF
through the pipeline (signals → segmentation → extraction → LE matching
→ routing), saves split invoices into a mirrored directory structure
under a TARGET folder, and displays full results in a two-panel layout.

Run:
    streamlit run app.py
"""

import base64
import os
import shutil
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
import os
os.environ["TRANSFORMERS_VERBOSITY"] = "error"

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="AI Invoice Router",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Cached pipeline import (heavy: loads matchers + reference JSONs once)
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner="Loading pipeline and reference data...")
def load_pipeline():
    import main as pipeline
    return pipeline


# ---------------------------------------------------------------------------
# Folder scanning (instant, no AI calls)
# ---------------------------------------------------------------------------
_PDF_EXT = {".pdf"}
_EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}


def scan_source_folder(source_path: str) -> dict:
    """Walk immediate subdirectories. Return inventory per subfolder."""
    root = Path(source_path)
    if not root.is_dir():
        return {"error": f"'{source_path}' is not a valid directory."}

    folders = []
    total_pdfs, total_excels = 0, 0

    subdirs = sorted([d for d in root.iterdir() if d.is_dir()])
    # If no subdirs, treat root itself as the only "subfolder"
    if not subdirs:
        subdirs = [root]

    for d in subdirs:
        pdfs = sorted([f for f in d.iterdir()
                       if f.is_file() and f.suffix.lower() in _PDF_EXT])
        excels = sorted([f for f in d.iterdir()
                         if f.is_file() and f.suffix.lower() in _EXCEL_EXT
                         and not f.name.startswith("~$")])
        other = [f for f in d.iterdir()
                 if f.is_file() and f.suffix.lower() not in _PDF_EXT | _EXCEL_EXT
                 and not f.name.startswith("~$")]

        folders.append({
            "path": d,
            "name": d.name if d != root else root.name,
            "pdfs": pdfs,
            "excels": excels,
            "other": other,
        })
        total_pdfs += len(pdfs)
        total_excels += len(excels)

    return {
        "root": root,
        "folders": folders,
        "total_pdfs": total_pdfs,
        "total_excels": total_excels,
        "total_folders": len(folders),
    }


# ---------------------------------------------------------------------------
# Display helpers
# ---------------------------------------------------------------------------
def page_range_str(pages):
    if not pages:
        return "—"
    pages = sorted(pages)
    ranges, start, prev = [], pages[0], pages[0]
    for p in pages[1:]:
        if p == prev + 1:
            prev = p
            continue
        ranges.append(f"{start}–{prev}" if start != prev else f"{start}")
        start = prev = p
    ranges.append(f"{start}–{prev}" if start != prev else f"{start}")
    return ", ".join(ranges)


def _dot(ok: bool) -> str:
    """Green dot for a good state, red dot for a flagged state."""
    return "🟢" if ok else "🔴"


def _status(ok: bool, good_label: str, bad_label: str) -> str:
    """Dot plus the word describing the ACTUAL state, so a red dot is never
    ambiguous about what it's flagging."""
    return f"{'🟢' if ok else '🔴'} {good_label if ok else bad_label}"


def render_invoice_detail(inv: dict):
    """Middle panel: compact invoice summary with a collapsed expander
    for match-details (sensitive-supplier / LE / routing specifics)."""
    st.markdown(f"**{inv.get('supplier_name') or 'Unknown supplier'}**"
               + (f"  ·  #{inv['invoice_number']}" if inv.get("invoice_number") else ""))

    st.write(f"Bill-to: {inv.get('bill_to_client_name') or '—'}")
    po = ", ".join(inv.get("po_numbers", [])) or "None found"
    st.write(f"PO: {po}")

    # Compact one-line status row
    sensitive_ok = not inv.get("is_sensitive")
    le_ok = bool(inv.get("le"))
    routing_ok = bool(inv.get("kofax_email"))
    review_ok = not inv.get("needs_manual_review")

    le_label = f"LE: {inv.get('le')}"
    st.markdown(
        f"{_status(sensitive_ok, 'Not sensitive', 'SENSITIVE')} &nbsp;&nbsp; "
        f"{_status(le_ok, le_label, 'LE: not matched')} &nbsp;&nbsp; "
        f"{_status(routing_ok, 'Routed', 'Not routed')} &nbsp;&nbsp; "
        f"{_status(review_ok, 'No review needed', 'REVIEW NEEDED')}",
        unsafe_allow_html=True,
    )

    st.caption(f"Pages — all: {page_range_str(inv.get('pages', []))} · "
              f"main: {page_range_str(inv.get('main_pages', []))} · "
              f"support: {page_range_str(inv.get('support_pages', []))}")

    if inv.get("needs_manual_review"):
        for r in inv.get("review_reasons", []):
            msg = r.get("message", str(r)) if isinstance(r, dict) else str(r)
            st.caption(f"⚠️ {msg}")

    # Match details -- tucked away, opened on demand
    with st.expander("🔍 Match details"):
        sm = inv.get("sensitive_match", {})
        st.write(f"**Sensitive-supplier match:** {sm.get('matched_name') or '—'} "
                f"(score {sm.get('score', '—')}, method {sm.get('method', '—')})")
        if sm.get("translated"):                                                    # ADD
            st.caption(f"↳ matched via translation: {sm.get('translated_query')}")   # ADD

        lm = inv.get("le_match", {})
        st.write(f"**LE match:** {lm.get('matched_le_name') or '—'} "
                f"(score {lm.get('score', '—')}, method {lm.get('method', '—')})")
        if lm.get("translated"):                                                    # ADD
            st.caption(f"↳ matched via translation: {lm.get('translated_query')}")   # ADD
        st.write(f"**LE country:** {inv.get('le_country') or '—'} · "
                f"**Source site:** {inv.get('source_site') or '—'}")

        st.write(f"**Routing:** {inv.get('invoice_type') or '—'} → "
                f"{inv.get('kofax_email') or 'not resolved'}")

        if inv.get("le") == "LE763":
            renamed = inv.get("vietnam_renamed")
            st.write(f"**Vietnam rename:** {_dot(bool(renamed))} "
                    f"{'renamed' if renamed else 'skipped'} "
                    f"(supplier ID: {inv.get('vietnam_supplier_id') or '—'})")

    out = inv.get("output_pdf", "")
    if out:
        st.caption(f"Saved to: `{out}`")


def render_pdf_preview(inv: dict):
    """Right panel: embedded PDF preview for the selected invoice."""
    out = inv.get("output_pdf", "")
    if not out or not Path(out).is_file():
        st.info("No file to preview.")
        return
    pdf_bytes = Path(out).read_bytes()
    b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_iframe = (
        f'<iframe src="data:application/pdf;base64,{b64_pdf}" '
        f'width="100%" height="560" '
        f'style="border: 1px solid #ccc; border-radius: 8px;">'
        f'</iframe>'
    )
    st.markdown(pdf_iframe, unsafe_allow_html=True)


def render_folder_block(folder_result: dict):
    """Renders one subfolder's block: header + a 3-column layout
    (invoice list | details | PDF preview). The list column can be
    swapped for a popover button (compact mode), and the preview column
    can be hidden entirely -- both controlled by the global toggles set
    once near the top of the page."""
    folder_name = folder_result["name"]
    folder_invoice_count = sum(
        len(pr.get("invoices", []))
        for pr in folder_result["pdf_results"]
    )

    st.markdown("---")
    st.markdown(f"""
    <div style="
        background: #2d5f8a;
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.9rem;
    ">
        📂 {folder_name} &nbsp;·&nbsp;
        {folder_result['pdf_count']} PDF(s) &nbsp;·&nbsp;
        {folder_result['excel_count']} Excel &nbsp;·&nbsp;
        {folder_invoice_count} invoice(s)
    </div>
    """, unsafe_allow_html=True)

    invoice_options = []
    invoice_data = []
    for pr in folder_result["pdf_results"]:
        if pr.get("error"):
            st.error(f"Failed to process {Path(pr['source_file']).name}: "
                    f"{pr['error']}")
            continue
        src_name = Path(pr["source_file"]).name
        for inv in pr.get("invoices", []):
            label = (f"{src_name} → Inv {inv['invoice_index']}: "
                    f"{inv.get('supplier_name', 'Unknown')}")
            if inv.get("invoice_number"):
                label += f" #{inv['invoice_number']}"
            if inv.get("needs_manual_review"):
                label = "⚠️ " + label
            invoice_options.append(label)
            invoice_data.append(inv)

    if not invoice_options:
        st.info("No invoices detected in this folder.")
        return

    compact_list = st.session_state.get("compact_list", False)
    show_preview = st.session_state.get("show_preview", True)

    def _list_widget():
        return st.radio(
            "Select an invoice",
            range(len(invoice_options)),
            format_func=lambda i: invoice_options[i],
            key=f"radio_{folder_name}",
            label_visibility="collapsed",
        )

    if compact_list:
        # List collapses into a popover button -- frees width for the
        # other two columns.
        with st.popover(f"📋 {len(invoice_options)} invoice(s) ▾"):
            selected_idx = _list_widget()
        if show_preview:
            detail_col, preview_col = st.columns([1, 1])
        else:
            detail_col, preview_col = st.columns([1, 0.001])
    else:
        if show_preview:
            list_col, detail_col, preview_col = st.columns([1, 1.3, 1.3])
        else:
            list_col, detail_col = st.columns([1, 2])
            preview_col = None
        with list_col:
            selected_idx = _list_widget()

    if selected_idx is None or selected_idx >= len(invoice_data):
        return
    selected_inv = invoice_data[selected_idx]

    with detail_col:
        render_invoice_detail(selected_inv)

    if show_preview and preview_col is not None:
        with preview_col:
            render_pdf_preview(selected_inv)


# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------

# --- Header ---
st.markdown("""
<div style="
    background: linear-gradient(135deg, #1e3a5f 0%, #2d5f8a 100%);
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
">
    <h1 style="color: white; margin: 0; font-size: 1.8rem;">
        🧾 AI-Powered Invoice Routing System
    </h1>
    <p style="color: #b0c4de; margin: 0.5rem 0 0 0; font-size: 0.95rem;">
        Multi-invoice splitting · 50+ languages · Scanned & digital PDFs ·
        Automated LE matching & Kofax routing
    </p>
</div>
""", unsafe_allow_html=True)


# --- Layout controls (apply to every folder block below) ---
ctrl1, ctrl2, _ = st.columns([1, 1, 3])
with ctrl1:
    st.toggle("📋 Compact invoice list", key="compact_list",
             help="Collapse the invoice list into a dropdown button "
                  "instead of a fixed column, to give details/preview more room.")
with ctrl2:
    st.toggle("👁 Show PDF preview", value=True, key="show_preview",
             help="Turn off to hide the preview column and widen the details panel.")


# --- Source type ------------------------------------------------------------
source_mode = st.radio(
    "Source",
    options=["pdf_folder", "email_source", "latest_run"],
    format_func=lambda m: {
        "pdf_folder": "📂 Loose PDF folders",
        "email_source": "📧 Email files (.msg / .eml)",
        "latest_run": "🔁 Reuse latest run",
    }[m],
    horizontal=True,
    key="source_mode",
)

pipeline = load_pipeline()

source_path = target_path = None   # used by the pdf_folder branch below

if source_mode == "pdf_folder":
    col_src, col_tgt = st.columns(2)
    with col_src:
        source_path = st.text_input(
            "📂 Source folder",
            placeholder="C:\\path\\to\\invoices",
            help="Parent folder containing subfolders. Each subfolder holds "
                 "invoice PDFs and optional Excel attachments.",
        )
    with col_tgt:
        target_path = st.text_input(
            "📁 Target folder",
            placeholder="C:\\path\\to\\output",
            help="Where split invoices will be saved, mirroring the source "
                 "directory structure.",
        )

elif source_mode == "email_source":
    col_src, col_tgt = st.columns(2)
    with col_src:
        email_source_path = st.text_input(
            "📧 Email source folder (.msg / .eml)",
            value=pipeline.msg_reader.DEFAULT_MSG_SOURCE,
            help="Folder of .msg and/or .eml files. Both types in the same "
                 "folder are picked up automatically.",
        )
    with col_tgt:
        results_root_path = st.text_input(
            "📁 Results root",
            value=pipeline.msg_reader.DEFAULT_RESULTS_ROOT,
            help="Each run creates its own dated, timestamped folder here.",
        )

else:  # latest_run
    results_root_path = st.text_input(
        "📁 Results root",
        value=pipeline.msg_reader.DEFAULT_RESULTS_ROOT,
        help="Reprocesses the most recent run found under this folder for "
             "today, without re-reading the original emails.",
    )
    latest = pipeline.msg_reader.find_latest_run(results_root_path)
    if latest:
        st.info(f"Will reuse: **{latest['run_name']}**")
    else:
        st.warning("No run found for today under this results root yet.")


# --- Input paths (loose PDF folders) -----------------------------------------


# --- Shared event driver -----------------------------------------------------
def _drive_events(events, progress_bar, status_text, live_results_area):
    """Consume any of the pipeline's run generators and update the UI as
    events arrive. Shared across all three source modes so the progress/
    blocked/error handling is written once, not three times.
    """
    for event in events:
        kind = event["type"]

        if kind == "run_ready":
            run = event["run"]
            st.session_state["run_info"] = run
            with live_results_area:
                st.info(f"📁 Run: **{run['run_name']}**")
                st.caption(f"Extracted: `{run['extracted']}`  ·  "
                          f"Processed: `{run['processed']}`")

        elif kind == "msg_progress":
            status_text.text(f"Reading emails: {event['file_name']} "
                             f"({event['current']}/{event['total']})")
            progress_bar.progress(event["current"] / max(event["total"], 1))

        elif kind == "msg_done":
            st.session_state["msg_done_info"] = event
            if event["blocked"] or event["failed"]:
                with live_results_area:
                    with st.expander(
                        f"⚠️ {event['blocked']} blocked, {event['failed']} failed "
                        f"during email extraction", expanded=True):
                        for r in event["results"]:
                            if r.get("blocked"):
                                name = Path(r["msg_path"]).name
                                st.warning(f"**{name}** — "
                                         f"{'; '.join(r['block_reasons'])}")
                            elif r.get("error"):
                                name = Path(r["msg_path"]).name
                                st.error(f"**{name}** — {r['error']}")

        elif kind == "progress":
            status_text.text(f"Processing {event['folder_name']}/{event['file_name']} "
                             f"({event['current']}/{event['total']})")
            progress_bar.progress(event["current"] / max(event["total"], 1))

        elif kind == "folder_done":
            folder_result = event["folder_result"]
            st.session_state["results"].append(folder_result)
            with live_results_area:
                render_folder_block(folder_result)

        elif kind == "complete":
            st.session_state["report_bytes"] = event.get("report_bytes")
            if event.get("report_path"):
                st.session_state["report_path"] = event["report_path"]

        elif kind == "error":
            st.error(f"❌ {event['message']}")
            st.session_state["run_error"] = event["message"]


# --- Scan + Summary: PDF FOLDER mode ----------------------------------------
if source_mode == "pdf_folder" and source_path and target_path:
    scan = scan_source_folder(source_path)

    if "error" in scan:
        st.error(scan["error"])
        st.stop()

    st.markdown("---")
    st.subheader("Summary")
    m1, m2, m3 = st.columns(3)
    m1.metric("Subfolders", scan["total_folders"])
    m2.metric("PDF files", scan["total_pdfs"])
    m3.metric("Excel files", scan["total_excels"])

    with st.expander("File inventory (click to expand)", expanded=False):
        for f in scan["folders"]:
            st.markdown(f"**{f['name']}/** — "
                       f"{len(f['pdfs'])} PDF(s), "
                       f"{len(f['excels'])} Excel, "
                       f"{len(f['other'])} other")
            for pdf in f["pdfs"]:
                st.caption(f"  📄 {pdf.name}")
            for xl in f["excels"]:
                st.caption(f"  📊 {xl.name}")

    if scan["total_pdfs"] == 0:
        st.warning("No PDF files found in any subfolder.")
        st.stop()

    if st.button("🚀 Process All Invoices", type="primary",
                 use_container_width=True):
        st.session_state["results"] = []
        st.session_state["report_bytes"] = None
        progress_bar = st.progress(0)
        status_text = st.empty()
        live_results_area = st.container()

        _drive_events(
            pipeline.process_invoice_folders(scan["folders"], target_path),
            progress_bar, status_text, live_results_area,
        )

        progress_bar.empty()
        status_text.empty()
        st.rerun()


# --- Scan + Summary: EMAIL SOURCE mode --------------------------------------
elif source_mode == "email_source" and email_source_path and results_root_path:
    st.markdown("---")
    if not Path(email_source_path).is_dir():
        st.error(f"'{email_source_path}' is not a valid directory.")
        st.stop()

    email_files = [p for p in Path(email_source_path).iterdir()
                  if p.is_file() and p.suffix.lower() in {".msg", ".eml"}]
    st.subheader("Summary")
    m1, m2 = st.columns(2)
    m1.metric(".msg files", sum(1 for p in email_files if p.suffix.lower() == ".msg"))
    m2.metric(".eml files", sum(1 for p in email_files if p.suffix.lower() == ".eml"))

    if not email_files:
        st.warning("No .msg or .eml files found in this folder.")
        st.stop()

    if st.button("📧 Extract Emails && Process Invoices", type="primary",
                 use_container_width=True):
        st.session_state["results"] = []
        st.session_state["report_bytes"] = None
        progress_bar = st.progress(0)
        status_text = st.empty()
        live_results_area = st.container()

        _drive_events(
            pipeline.run_from_msg_source(email_source_path, results_root_path),
            progress_bar, status_text, live_results_area,
        )

        progress_bar.empty()
        status_text.empty()
        st.rerun()


# --- LATEST RUN mode ---------------------------------------------------------
elif source_mode == "latest_run" and results_root_path:
    st.markdown("---")
    if st.button("🔁 Reprocess Latest Run", type="primary",
                 use_container_width=True,
                 disabled=not pipeline.msg_reader.find_latest_run(results_root_path)):
        st.session_state["results"] = []
        st.session_state["report_bytes"] = None
        progress_bar = st.progress(0)
        status_text = st.empty()
        live_results_area = st.container()

        _drive_events(
            pipeline.run_from_latest_extracted(results_root_path),
            progress_bar, status_text, live_results_area,
        )

        progress_bar.empty()
        status_text.empty()
        st.rerun()


# --- Results display ---
if "results" in st.session_state:
    results = st.session_state["results"]

    st.markdown("---")

    # Post-processing summary
    total_invoices = sum(
        len(pr.get("invoices", []))
        for fr in results for pr in fr["pdf_results"]
    )
    needs_review = sum(
        1 for fr in results for pr in fr["pdf_results"]
        for inv in pr.get("invoices", []) if inv.get("needs_manual_review")
    )
    errors = sum(
        1 for fr in results for pr in fr["pdf_results"]
        if pr.get("error")
    )

    st.subheader("Processing Complete")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Folders processed", len(results))
    r2.metric("Invoices detected", total_invoices)
    r3.metric("Needs review", needs_review)
    r4.metric("Errors", errors)

    if total_invoices > 0:
        # process_invoice_folders() already builds this once, at the end
        # of the batch -- use it directly instead of recomputing. Falls
        # back to rebuilding only if session state predates this (e.g.
        # results carried over from an older run).
        report_bytes = st.session_state.get("report_bytes")
        if not report_bytes:
            from report_builder import flatten_folder_results, build_excel_report
            report_bytes = build_excel_report(flatten_folder_results(results))
        st.download_button(
            "📊 Download Excel report",
            data=report_bytes,
            file_name=f"invoice_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # --- Per-folder blocks (same renderer used for live + post-rerun) ---
    for folder_result in results:
        render_folder_block(folder_result)

    # --- Clear button ---
    st.markdown("---")
    if st.button("🗑️ Clear results and start over"):
        del st.session_state["results"]
        st.rerun()

else:
    if source_mode == "pdf_folder" and not (source_path and target_path):
        st.info("Enter source and target folder paths above to begin.")
    elif source_mode == "email_source" and not (email_source_path and results_root_path):
        st.info("Enter the email source and results root paths above to begin.")
    elif source_mode == "latest_run" and not results_root_path:
        st.info("Enter a results root above to begin.")