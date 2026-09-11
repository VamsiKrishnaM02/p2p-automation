"""
send_mail.py -- Core SMTP send primitive, used by email_sender.py for both
the per-batch Kofax routing emails and the human-agent report email.

Changes from the original version (kept functionally identical for the
part that already matched spec -- filename sanitization and fallback
naming were already correct and are untouched):

  - Returns a structured {"sent", "error", "dry_run"} dict instead of
    printing and swallowing every exception. email_sender.py needs this to
    know which batches actually reached Kofax (a batch that failed to SEND
    still needs to show up as a failure in the human report, even though
    it was processed correctly).
  - Uses `logging` instead of `print`, consistent with the rest of the
    pipeline (main.py, routing.py, ...).
  - Adds `dry_run`: builds the message and logs what WOULD be sent
    (receiver, cc, subject, attachment filenames) without opening an SMTP
    connection at all. Use this to test the Kofax-routing flow against
    real batch data before it's allowed to actually send.
  - REMOVED the hardcoded fallback password from __main__. There is no
    string fallback for a credential, ever -- SMTP_PASSWORD must be set in
    the environment or the script refuses to run.
"""

import logging
import os
import re
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional

log = logging.getLogger("send_mail")

SMTP_SERVER = "smtpauth.intel.com"
SMTP_PORT = 587


def send_mail(sender: str, receiver: str, password: str, subject: str = "",
             body: str = "", attachment_list: Optional[List[str]] = None,
             cc: Optional[str] = None, dry_run: bool = False,
             login_account: Optional[str] = None) -> dict:
    """
    Sends an email using the internal SMTP server.

    login_account: the account that actually AUTHENTICATES to the SMTP
    server (server.login()). Defaults to `sender` when not given, so
    every existing single-account caller is completely unchanged.

    Pass login_account explicitly to send AS a different address than the
    one that logs in -- e.g. login_account="new_lead@intel.com" while
    sender="gspo@intel.com". This mirrors the legacy system's proven
    behavior for this exact relay: the authenticated account is used
    ONLY for server.login(); both the "From" header AND the SMTP
    envelope-from (server.sendmail()'s first argument) use `sender`, not
    login_account. That does mean the envelope-from won't match the
    authenticated login -- if smtpauth.intel.com's permissions ever
    change to require them to match, this is the first place to look.

    Returns {"sent": bool, "error": str or None, "dry_run": bool}.
    """
    attached_names: List[str] = []
    login_account = login_account or sender
    try:
        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = receiver

        # Build recipient list
        rcpts = [receiver]
        if cc:
            msg["Cc"] = cc
            rcpts.extend([c.strip() for c in cc.split(",") if c.strip()])

        # Attach HTML body
        if body:
            msg.attach(MIMEText(body, "html"))

        # Process attachments
        if attachment_list:
            for filepath in attachment_list:
                if not os.path.isfile(filepath):
                    log.warning("Attachment not found, skipping: %s", filepath)
                    continue

                file_name = os.path.basename(filepath)
                with open(filepath, "rb") as file:
                    attachment = MIMEBase("application", "octet-stream")
                    attachment.set_payload(file.read())
                encoders.encode_base64(attachment)

                # Sanitize the filename -- letters, digits, periods only.
                basename = re.sub(r"[^a-zA-Z0-9.]", "_", file_name)
                if not basename.split(".")[0]:
                    basename = (subject[:6] if subject else "attached") + ".pdf"

                attachment.add_header("Content-Disposition",
                                     f'attachment; filename="{basename}"')
                msg.attach(attachment)
                attached_names.append(basename)

        if dry_run:
            log.info("[DRY RUN] to=%s cc=%s subject=%r attachments=%s "
                     "(From=%s, login_account=%s)",
                     receiver, cc, subject, attached_names, sender, login_account)
            return {"sent": True, "error": None, "dry_run": True}

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(login_account, password)
            server.sendmail(sender, rcpts, msg.as_string())

        log.info("Email sent to %s (From=%s, login_account=%s, cc=%s, "
                "subject=%r, %d attachment(s))",
                receiver, sender, login_account, cc, subject, len(attached_names))
        return {"sent": True, "error": None, "dry_run": False}

    except Exception as e:
        log.error("Failed to send mail to %s (subject=%r): %s", receiver, subject, e)
        return {"sent": False, "error": str(e), "dry_run": dry_run}


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s | %(levelname)-7s | %(message)s")

    my_email = "vamsix.modala@intel.com"
    my_password = os.environ.get("SMTP_PASSWORD")
    if not my_password:
        raise SystemExit(
            "SMTP_PASSWORD is not set in the environment -- no fallback "
            "password is stored in code. Set it and try again."
        )

    result = send_mail(
        sender=my_email,
        receiver="vamsix.modala@intel.com",
        password=my_password,
        subject="Invoice Pipeline Status",
        body="<h1>Run Complete</h1><p>The processing has finished successfully.</p>",
        dry_run=True,  # flip to False once you've confirmed this looks right
    )
    print(result)