"""
email_watcher.py
Watches a Gmail mailbox for CommonSku SO confirmation emails delivered via a
Google Group, downloads PDF attachments, and processes them through the Monday
integration.

Runs continuously, polling every 60 seconds.

Inbound flow:
    Sales order email --> so@avidapparel.ca group --> bret@ mailbox
    A Gmail filter on bret@ applies label `so-inbox`, skips Inbox, marks read.
    Old dedicated mailbox is set to forward to so@ for the transition period.

Watcher behaviour:
    Polls for messages with label `so-inbox` (pending).
    After processing, removes `so-inbox` and adds `so-processed`.
    Labels are auto-created on first run if missing.

Usage:
    python3 email_watcher.py           # run continuously
    python3 email_watcher.py --once    # process current pending emails once and exit

Setup:
    1. Place your client_secret_*.json file in the same directory as this script
    2. Run once interactively to complete OAuth flow (opens browser) as bret@
    3. After authentication, token.json is saved — subsequent runs are silent
    4. Run with systemd or screen for continuous operation (see README)
"""

import os
import sys
import base64
import time
import re
import ssl
import socket
import http.client
import tempfile
import logging
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

# Google API imports
import httplib2
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Local imports
sys.path.insert(0, str(Path(__file__).parent))
from alerts import logger, notify_slack, alert_error

SCOPES            = ["https://www.googleapis.com/auth/gmail.modify"]
CREDENTIALS_FILE  = next(Path(__file__).parent.glob("client_secret_*.json"), None)
TOKEN_FILE        = Path(__file__).parent / "token.json"
POLL_INTERVAL     = int(os.environ.get("EMAIL_POLL_INTERVAL", 60))  # seconds
LABEL_PENDING     = os.environ.get("EMAIL_LABEL_PENDING", "so-inbox")
LABEL_PROCESSED   = os.environ.get("EMAIL_LABEL_PROCESSED", "so-processed")

_label_id_cache = {}

# Transport-level failures: the connection died, as opposed to the API rejecting
# the request. googleapiclient sits on httplib2, which keeps one persistent TLS
# connection per host and does NOT detect a stale socket. Between 60s polls the
# connection goes idle, Google's frontend (and the droplet's NAT conntrack) drops
# it, and the next request writes into a dead socket -> SSLEOFError
# "EOF occurred in violation of protocol". The cure is to rebuild the service.
TRANSPORT_ERRORS = (
    ssl.SSLError,              # includes ssl.SSLEOFError
    socket.timeout,
    socket.error,              # OSError: ConnectionReset/Aborted/BrokenPipe
    http.client.HTTPException,
    httplib2.HttpLib2Error,
)

# googleapiclient will retry transport failures itself, but only if asked.
API_RETRIES = 3


# ---------------------------------------------------------------------------
# Gmail authentication
# ---------------------------------------------------------------------------

def get_gmail_service():
    """
    Authenticate with Gmail API and return a service object.
    On first run, opens a browser for OAuth consent.
    On subsequent runs, uses saved token.json silently.
    """
    if not CREDENTIALS_FILE:
        raise RuntimeError(
            "No client_secret_*.json file found in the parser directory. "
            "Download it from Google Cloud Console and place it here."
        )

    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            # Loopback flow: opens a browser and listens on a local port for the redirect.
            # Requires a browser on this machine — for headless servers, run this script
            # once on a workstation, then copy the resulting token.json over.
            creds = flow.run_local_server(port=0, prompt="consent")

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds, cache_discovery=False)


# ---------------------------------------------------------------------------
# Email processing
# ---------------------------------------------------------------------------

def get_label_id(service, name):
    """Return Gmail label id for `name`, creating the label if absent."""
    if name in _label_id_cache:
        return _label_id_cache[name]

    labels = service.users().labels().list(userId="me").execute(num_retries=API_RETRIES).get("labels", [])
    for lbl in labels:
        if lbl["name"].lower() == name.lower():
            _label_id_cache[name] = lbl["id"]
            return lbl["id"]

    created = service.users().labels().create(
        userId="me",
        body={
            "name": name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        },
    ).execute(num_retries=API_RETRIES)
    logger.info(f"Created Gmail label: {name}")
    _label_id_cache[name] = created["id"]
    return created["id"]


def get_pending_messages(service):
    """Fetch messages tagged with the pending label (delivered via group, not yet processed)."""
    pending_id = get_label_id(service, LABEL_PENDING)
    result = service.users().messages().list(
        userId="me",
        labelIds=[pending_id],
        maxResults=50,
    ).execute(num_retries=API_RETRIES)
    return result.get("messages", [])


def get_message_detail(service, message_id):
    """Fetch full message detail including attachments."""
    return service.users().messages().get(
        userId="me",
        id=message_id,
        format="full"
    ).execute(num_retries=API_RETRIES)


def get_pdf_attachments(service, message):
    """
    Extract all PDF attachments from a message.
    Returns list of (filename, bytes) tuples.
    """
    pdfs = []
    parts = message.get("payload", {}).get("parts", [])

    def walk_parts(parts):
        for part in parts:
            # Recurse into nested parts
            if part.get("parts"):
                walk_parts(part["parts"])
                continue

            filename = part.get("filename", "")
            mime     = part.get("mimeType", "")

            if not filename.lower().endswith(".pdf") and mime != "application/pdf":
                continue

            body = part.get("body", {})
            attachment_id = body.get("attachmentId")

            if attachment_id:
                # Large attachment — fetch separately
                attachment = service.users().messages().attachments().get(
                    userId="me",
                    messageId=message["id"],
                    id=attachment_id
                ).execute(num_retries=API_RETRIES)
                data = base64.urlsafe_b64decode(attachment["data"])
            elif body.get("data"):
                # Small attachment — inline
                data = base64.urlsafe_b64decode(body["data"])
            else:
                continue

            pdfs.append((filename or "attachment.pdf", data))

    walk_parts(parts)
    return pdfs


def mark_processed(service, message_id):
    """Move a message from the pending label to the processed label so it isn't re-processed."""
    pending_id   = get_label_id(service, LABEL_PENDING)
    processed_id = get_label_id(service, LABEL_PROCESSED)
    service.users().messages().modify(
        userId="me",
        id=message_id,
        body={
            "removeLabelIds": [pending_id, "UNREAD"],
            "addLabelIds":    [processed_id],
        },
    ).execute(num_retries=API_RETRIES)


def get_sender(message):
    """Extract sender email from message headers."""
    headers = message.get("payload", {}).get("headers", [])
    for h in headers:
        if h["name"].lower() == "from":
            return h["value"]
    return "unknown sender"


def get_subject(message):
    """Extract subject from message headers."""
    headers = message.get("payload", {}).get("headers", [])
    for h in headers:
        if h["name"].lower() == "subject":
            return h["value"]
    return "(no subject)"


# ---------------------------------------------------------------------------
# Main processing loop
# ---------------------------------------------------------------------------

def process_message(service, message_id):
    """
    Process a single email message:
    - Download PDF attachments
    - Run each through the Monday integration pipeline
    - Move from pending label to processed label when done (even if processing failed)
    """
    from monday_api import process_pdf

    message = get_message_detail(service, message_id)
    sender  = get_sender(message)
    subject = get_subject(message)

    logger.info(f"Processing email from {sender}: {subject}")

    pdfs = get_pdf_attachments(service, message)

    if not pdfs:
        logger.warning(f"No PDF attachments found in email: {subject}")
        notify_slack(
            f":warning: *Email received but no PDF found*\n"
            f"From: {sender}\n"
            f"Subject: {subject}\n"
            f"Email has been marked as processed.",
            level="warning"
        )
        mark_processed(service, message_id)
        return

    any_failed = False

    for filename, pdf_bytes in pdfs:
        logger.info(f"  Processing attachment: {filename}")

        # Write PDF to a temp file using the original filename
        safe_filename = re.sub(r'[^\w\s\-.]', '_', filename)  # sanitise for filesystem
        tmp_path = Path(tempfile.gettempdir()) / safe_filename
        tmp_path.write_bytes(pdf_bytes)

        try:
            process_pdf(str(tmp_path))
        except Exception as e:
            any_failed = True
            logger.error(f"  Failed to process {filename}: {e}")
            notify_slack(
                f":red_circle: *Failed to process email attachment*\n"
                f"File: {filename}\n"
                f"From: {sender}\n"
                f"Subject: {subject}\n"
                f"Error: `{str(e)[:300]}`",
                level="error"
            )
        finally:
            # Always clean up temp file
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

    # Move to processed regardless of success/failure
    # Prevents infinite retry loops on persistently broken PDFs
    mark_processed(service, message_id)

    if any_failed:
        logger.warning(f"Email processed with errors: {subject}")
    else:
        logger.info(f"Email fully processed: {subject}")


def run_once(service):
    """Check pending label once and process all messages found."""
    messages = get_pending_messages(service)

    if not messages:
        logger.debug("No pending messages.")
        return 0

    logger.info(f"Found {len(messages)} pending message(s)")
    processed = 0

    for msg in messages:
        try:
            process_message(service, msg["id"])
            processed += 1
        except TRANSPORT_ERRORS:
            # Connection died mid-message. Bubble up so run_continuously can
            # rebuild the Gmail service; the message keeps its pending label
            # and gets retried on the next poll.
            raise
        except Exception as e:
            logger.error(f"Unexpected error processing message {msg['id']}: {e}")
            # Don't move to processed — will retry next poll
            # But do notify so the team knows
            notify_slack(
                f":red_circle: *Unexpected error in email watcher*\n"
                f"Message ID: {msg['id']}\n"
                f"Error: `{str(e)[:300]}`\n"
                f"Email left in pending label for retry.",
                level="error"
            )

    return processed


def run_continuously(service):
    """Poll inbox every POLL_INTERVAL seconds indefinitely."""
    logger.info(f"Email watcher started. Polling every {POLL_INTERVAL}s.")
    notify_slack(
        f":white_check_mark: *Email watcher started*\n"
        f"Polling every {POLL_INTERVAL} seconds.",
        level="success"
    )

    consecutive_transport_errors = 0

    while True:
        try:
            run_once(service)
            consecutive_transport_errors = 0
        except HttpError as e:
            logger.error(f"Gmail API error: {e}")
            notify_slack(f":red_circle: *Gmail API error in watcher*\n`{str(e)[:300]}`", level="error")
        except TRANSPORT_ERRORS as e:
            consecutive_transport_errors += 1
            logger.warning(
                f"Gmail connection dropped ({type(e).__name__}: {e}). "
                f"Rebuilding service (consecutive failures: {consecutive_transport_errors})."
            )
            _label_id_cache.clear()
            try:
                service = get_gmail_service()
                logger.info("Gmail service rebuilt.")
            except Exception as rebuild_err:
                logger.error(f"Failed to rebuild Gmail service: {rebuild_err}")

            # A single dropped keep-alive is routine and self-healing — only page
            # the team once it looks like a real outage.
            if consecutive_transport_errors == 3:
                notify_slack(
                    f":red_circle: *Email watcher losing its Gmail connection*\n"
                    f"3 consecutive transport failures. Latest: `{type(e).__name__}: {str(e)[:200]}`\n"
                    f"Still retrying every {POLL_INTERVAL}s.",
                    level="error",
                )
        except Exception as e:
            logger.error(f"Watcher loop error: {e}")
            notify_slack(f":red_circle: *Unexpected watcher error*\n`{str(e)[:300]}`", level="error")

        time.sleep(POLL_INTERVAL)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    try:
        service = get_gmail_service()
    except Exception as e:
        logger.error(f"Failed to authenticate with Gmail: {e}")
        sys.exit(1)

    if "--once" in sys.argv:
        count = run_once(service)
        logger.info(f"Done. Processed {count} message(s).")
    else:
        run_continuously(service)