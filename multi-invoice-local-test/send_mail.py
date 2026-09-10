# import os
# import re
# import smtplib
# from email import encoders
# from email.mime.base import MIMEBase
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText

# def send_mail(sender: str, receiver: str, password: str, subject: str = "", body: str = "", attachment_list: list = None, cc: str = None):
#     """
#     Sends an email using the internal SMTP server.
#     """
#     try:
#         msg = MIMEMultipart()
#         msg["Subject"] = subject
#         msg["From"] = sender
#         msg["To"] = receiver
        
#         # Build recipient list
#         rcpts = [receiver]
#         if cc:
#             msg["Cc"] = cc
#             rcpts.extend([c.strip() for c in cc.split(",") if c.strip()])
            
#         # Attach HTML body
#         if body:
#             msg.attach(MIMEText(body, 'html'))

#         # Process attachments
#         if attachment_list:
#             for filepath in attachment_list:
#                 if os.path.isfile(filepath):
#                     file_name = os.path.basename(filepath)
                    
#                     with open(filepath, 'rb') as file:
#                         attachment = MIMEBase('application', 'octet-stream')
#                         attachment.set_payload(file.read())
                        
#                     encoders.encode_base64(attachment)
                    
#                     # Sanitize the filename
#                     basename = re.sub(r'[^a-zA-Z0-9.]', '_', file_name)
#                     if not basename.split(".")[0]:
#                         basename = (subject[:6] if subject else "attached") + ".pdf"
                        
#                     attachment.add_header('Content-Disposition', f'attachment; filename="{basename}"')
#                     msg.attach(attachment)

#         # Connect to server and send
#         smtp_server = "smtpauth.intel.com"
#         smtp_port = 587
        
#         with smtplib.SMTP(smtp_server, smtp_port) as server:
#             server.starttls()  # Secure the connection
#             server.login(sender, password)
#             server.sendmail(sender, rcpts, msg.as_string())
            
#         print(f"Email successfully sent to {receiver}!")
        
#     except Exception as e:
#         print(f"Exception occurred while sending mail: {e}")

# if __name__ == "__main__":
#     # Example usage:
#     # Use environment variables for passwords instead of hardcoding them in plain text.
#     my_email = "vamsix.modala@intel.com"
#     my_password = os.environ.get("SMTP_PASSWORD") or "Banana@1026" 
    
#     send_mail(
#         sender=my_email,
#         receiver="vamsix.modala@intel.com",
#         password=my_password,
#         subject="Invoice Pipeline Status",
#         body="<h1>Run Complete</h1><p>The processing has finished successfully.</p>",
#         attachment_list=[r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\msg-result\2026-09-10\run_13-16-08\02_blocked\FW_ _HRC_ _PI_ Deloitte - Hard Copy of Invoices ILH LE 755_ 750 and 870_ Soft Copy of Invoice ILH LE 778 - August 2026.msg"]
#     )


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
             cc: Optional[str] = None, dry_run: bool = False) -> dict:
    """
    Sends an email using the internal SMTP server.

    Returns {"sent": bool, "error": str or None, "dry_run": bool}.
    Never raises -- any failure (missing attachment aside, which is just
    skipped with a warning) is caught and reported in the return dict.
    """
    attached_names: List[str] = []
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
            log.info("[DRY RUN] to=%s cc=%s subject=%r attachments=%s",
                     receiver, cc, subject, attached_names)
            return {"sent": True, "error": None, "dry_run": True}

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, rcpts, msg.as_string())

        log.info("Email sent to %s (cc=%s, subject=%r, %d attachment(s))",
                 receiver, cc, subject, len(attached_names))
        return {"sent": True, "error": None, "dry_run": False}

    except Exception as e:
        log.error("Failed to send mail to %s (subject=%r): %s", receiver, subject, e)
        return {"sent": False, "error": str(e), "dry_run": dry_run}


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s | %(levelname)-7s | %(message)s")

    my_email = "vamsix.modala@intel.com"
    my_password = "Banana@1026"

    # if not my_password:
    #     raise SystemExit(
    #         "SMTP_PASSWORD is not set in the environment -- no fallback "
    #         "password is stored in code. Set it and try again."
    #     )

    result = send_mail(
        sender=my_email,
        receiver="vamsix.modala@intel.com",
        password=my_password,
        subject="Invoice Pipeline Status",
        body="<h1>Run Complete</h1><p>The processing has finished successfully.</p>",
        dry_run=True,  # flip to False once you've confirmed this looks right
    )
    print(result)