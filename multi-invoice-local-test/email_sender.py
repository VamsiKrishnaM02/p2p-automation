"""
email_sender.py -- The last stage of the pipeline: sends the outbound
Kofax routing emails, then one summary report email to a human agent.

Two entry points, both built on send_mail.send_mail():

  send_routing_emails(batches, ...)
    One email PER BATCH (from routing.process_batches()'s "complete"
    event): attachment-only, subject = Batch ID, body empty, receiver =
    that batch's resolved Kofax mailbox, cc = a fixed distribution list.
    Annotates each batch dict in place with "email_sent"/"email_error" so
    compute_run_stats() (and the caller) can tell which batches actually
    reached Kofax versus which were processed fine but failed to SEND.

  send_report_email(...)
    ONE email, sent AFTER every routing email has been attempted, to a
    human agent: small HTML stats (assigned / successful / failed, with a
    breakdown of WHY something failed), the two Excel reports as normal
    attachments, and one zip ("needs_attention.zip") containing every
    quarantined .msg/.eml copy AND the staged PDFs of any batch that
    failed to SEND -- see _build_report_zip() for exactly how those two
    are kept in separate subfolders.

STATS DEFINITION (per Vamsi's explicit call):
  An invoice counts as SUCCESSFUL only if it was both (a) processed
  cleanly (not quarantined) AND (b) its batch's routing email actually
  reached Kofax. A batch whose SMTP send failed moves its invoices into
  the FAILED bucket too -- a processed-but-undelivered invoice needs a
  human's attention just as much as a processing failure does.

DRY RUN: both entry points accept dry_run=True, which threads straight
through to send_mail.send_mail() -- messages are built and logged, no
SMTP connection is opened. Use this to test against real batch/run data
before flipping to real sends.
"""

import logging
import zipfile
from pathlib import Path
from typing import List, Optional

from send_mail import send_mail

log = logging.getLogger("email_sender")

# ---------------------------------------------------------------------------
# Config -- fixed constants, same pattern as routing.CONFIG_CODE. Fill in
# real values before this goes anywhere near production.
# ---------------------------------------------------------------------------
SENDER_EMAIL = "vamsix.modala@intel.com"     
ROUTING_CC = "alekhyax.vemuri@intel.com"  
REPORT_AGENT_EMAIL = "vamsix.modala@intel.com"   


# ---------------------------------------------------------------------------
# Routing emails -- one per batch
# ---------------------------------------------------------------------------
def send_routing_emails(batches: List[dict], sender: str, password: str,
                        cc: Optional[str] = None,
                        dry_run: bool = False) -> List[dict]:
    """Send one attachment-only email per batch. Mutates and returns
    `batches` -- each dict gains "email_sent" (bool) and "email_error"
    (str or None)."""
    for batch in batches:
        staged_dir = Path(batch["staged_dir"])
        pdfs = sorted(str(p) for p in staged_dir.glob("*.pdf"))

        if not pdfs:
            log.warning("Batch %s has no PDFs staged in %s -- sending anyway "
                       "per spec, but this usually means a copy failed earlier.",
                       batch["batch_id"], staged_dir)

        result = send_mail(
            sender=sender,
            receiver=batch["kofax_email"],
            password=password,
            subject=batch["batch_id"],
            body="",
            attachment_list=pdfs,
            cc=cc,
            dry_run=dry_run,
        )
        batch["email_sent"] = result["sent"]
        batch["email_error"] = result["error"]

        if not result["sent"]:
            log.error("Batch %s FAILED to send to %s: %s",
                     batch["batch_id"], batch["kofax_email"], result["error"])

    return batches


# ---------------------------------------------------------------------------
# Run-level stats
# ---------------------------------------------------------------------------
def compute_run_stats(batches: List[dict], quarantine_log: List[dict]) -> dict:
    """
    `batches` must already be annotated by send_routing_emails() (i.e. call
    this AFTER send_routing_emails(), not before).
    `quarantine_log` is process_batches()'s "complete"->"quarantined" list
    -- each entry already carries "invoice_count" (see routing.py).
    """
    total_successful = 0
    total_send_failed_invoices = 0
    failed_batches = []

    for b in batches:
        if b.get("email_sent"):
            total_successful += b["invoice_count"]
        else:
            total_send_failed_invoices += b["invoice_count"]
            failed_batches.append({
                "batch_id": b["batch_id"],
                "kofax_email": b["kofax_email"],
                "invoice_count": b["invoice_count"],
                "error": b.get("email_error"),
                "staged_dir": b["staged_dir"],
            })

    total_quarantined_invoices = sum(q.get("invoice_count", 0) for q in quarantine_log)
    total_failed = total_send_failed_invoices + total_quarantined_invoices

    return {
        "total_invoices": total_successful + total_failed,
        "total_successful": total_successful,
        "total_failed": total_failed,
        "total_quarantined_invoices": total_quarantined_invoices,
        "total_send_failed_invoices": total_send_failed_invoices,
        "total_batches": len(batches),
        "total_quarantined_emails": len(quarantine_log),
        "failed_batches": failed_batches,
    }


# ---------------------------------------------------------------------------
# Report zip: quarantined source emails + any batch that failed to SEND
# ---------------------------------------------------------------------------
def _build_report_zip(quarantine_log: List[dict], failed_batches: List[dict],
                      run_root: str) -> Optional[str]:
    """Zip everything the human agent needs to act on by hand:

      quarantined_source_emails/<original.msg>   -- one copy per quarantined
        source email (processing failures) -- routing.py already copied
        these into 02_blocked/; this just bundles them for the email.

      failed_to_send/<batch_id>/<pdf files>       -- the ALREADY-SPLIT PDFs
        for any batch whose Kofax send failed. These are staged correctly
        in 04_batches/ already; a send failure is a per-BATCH problem, not
        a per-source-email one (one source email can produce several
        batches, and only one might fail to send) -- so this bundles the
        specific batch's files rather than touching the original .msg,
        which could wrongly implicate a sibling batch that sent fine.

    Returns None only if there's truly nothing to include (every invoice
    was processed cleanly AND every batch sent successfully)."""
    copied_paths = [q["copied_to"] for q in quarantine_log if q.get("copied_to")]
    if not copied_paths and not failed_batches:
        return None

    zip_path = Path(run_root) / "needs_attention.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in copied_paths:
            if Path(p).is_file():
                zf.write(p, arcname=f"quarantined_source_emails/{Path(p).name}")
            else:
                log.warning("Report zip: expected quarantined file missing, "
                          "skipping: %s", p)

        for fb in failed_batches:
            staged_dir = Path(fb["staged_dir"])
            if not staged_dir.is_dir():
                log.warning("Report zip: staged dir missing for failed batch "
                          "%s, skipping: %s", fb["batch_id"], staged_dir)
                continue
            for pdf in sorted(staged_dir.glob("*.pdf")):
                zf.write(pdf, arcname=f"failed_to_send/{fb['batch_id']}/{pdf.name}")

    return str(zip_path)


# ---------------------------------------------------------------------------
# Human-agent report email
# ---------------------------------------------------------------------------
def _build_stats_html(stats: dict, run_label: str = "") -> str:
    rows = [
        ("Total invoices assigned", stats["total_invoices"]),
        ("Successful (sent to Kofax)", stats["total_successful"]),
        ("Failed / blocked (total)", stats["total_failed"]),
        ("&nbsp;&nbsp;- Quarantined during processing", stats["total_quarantined_invoices"]),
        ("&nbsp;&nbsp;- Processed OK but failed to SEND", stats["total_send_failed_invoices"]),
        ("Batches sent", stats["total_batches"]),
        ("Source emails quarantined", stats["total_quarantined_emails"]),
    ]
    table_rows = "".join(
        f"<tr><td style='padding:4px 12px 4px 0'>{label}</td>"
        f"<td style='padding:4px 0;text-align:right'><b>{value}</b></td></tr>"
        for label, value in rows
    )

    failed_batch_rows = ""
    if stats["failed_batches"]:
        items = "".join(
            f"<li>{b['batch_id']} → {b['kofax_email']} "
            f"({b['invoice_count']} invoice(s)) -- {b['error']}</li>"
            for b in stats["failed_batches"]
        )
        failed_batch_rows = (
            "<p><b>Batches that failed to send (retry needed):</b></p>"
            f"<ul>{items}</ul>"
        )

    title = f"Invoice Pipeline Report{f' — {run_label}' if run_label else ''}"
    return (
        f"<h2>{title}</h2>"
        f"<table>{table_rows}</table>"
        f"{failed_batch_rows}"
        "<p>See attached: the per-invoice report, the batch &amp; failures "
        "report, and (if anything needs manual handling) a zip containing "
        "the original quarantined emails and/or the PDFs for any batch "
        "that failed to send.</p>"
    )


def send_report_email(stats: dict, quarantine_log: List[dict],
                      invoice_report_path: str, batch_report_path: str,
                      run_root: str, agent_email: str, sender: str,
                      password: str, cc: Optional[str] = None,
                      run_label: str = "", dry_run: bool = False) -> dict:
    """Send the single end-of-run summary email to the human agent."""
    zip_path = _build_report_zip(quarantine_log, stats["failed_batches"], run_root)

    attachments = [invoice_report_path, batch_report_path]
    if zip_path:
        attachments.append(zip_path)

    subject = f"Invoice Pipeline Report{f' - {run_label}' if run_label else ''}"
    body = _build_stats_html(stats, run_label)

    result = send_mail(
        sender=sender,
        receiver=agent_email,
        password=password,
        subject=subject,
        body=body,
        attachment_list=attachments,
        cc=cc,
        dry_run=dry_run,
    )
    if not result["sent"]:
        log.error("Report email to %s FAILED: %s", agent_email, result["error"])
    return result