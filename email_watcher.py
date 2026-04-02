"""
email_watcher.py
Watches a dedicated Gmail inbox for CommonSku SO confirmation emails,
downloads PDF attachments, and processes them through the Monday integration.

Runs continuously, polling every 60 seconds.

First run: opens a browser for OAuth authentication and saves a token.
Subsequent runs: uses the saved token silently.

Usage:
    python3 email_watcher.py           # run continuously
    python3 email_watcher.py --once    # process current unread emails once and exit

Setup:
    1. Place your client_secret_*.json file in the same directory as this script
    2. Run once interactively to complete OAuth flow (opens browser)
    3. After authentication, token.json is saved — subsequent runs are silent
    4. Run with systemd or screen for continuous operation (see README)
"""

import os
import sys
import base64
import time
import tempfile
import logging
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

# Google API imports
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
            # run_console: prints a URL to open on any browser, then asks for the auth code
            # Works on headless servers with no browser
            # For headless servers: print auth URL, user pastes code back
            flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
            auth_url, _ = flow.authorization_url(prompt="consent")
            print(f"\nPlease open this URL in your browser:\n\n{auth_url}\n")
            code = input("Enter the authorization code from the browser: ").strip()
            flow.fetch_token(code=code)
            creds = flow.credentials

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


# ---------------------------------------------------------------------------
# Email processing
# ---------------------------------------------------------------------------

def get_unread_messages(service):
    """Fetch all unread messages in the inbox."""
    result = service.users().messages().list(
        userId="me",
        labelIds=["INBOX", "UNREAD"],
        maxResults=50
    ).execute()
    return result.get("messages", [])


def get_message_detail(service, message_id):
    """Fetch full message detail including attachments."""
    return service.users().messages().get(
        userId="me",
        id=message_id,
        format="full"
    ).execute()


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
                ).execute()
                data = base64.urlsafe_b64decode(attachment["data"])
            elif body.get("data"):
                # Small attachment — inline
                data = base64.urlsafe_b64decode(body["data"])
            else:
                continue

            pdfs.append((filename or "attachment.pdf", data))

    walk_parts(parts)
    return pdfs


def mark_as_read(service, message_id):
    """Mark a message as read so it isn't processed again."""
    service.users().messages().modify(
        userId="me",
        id=message_id,
        body={"removeLabelIds": ["UNREAD"]}
    ).execute()


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
    - Mark email as read when done (even if processing failed)
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
            f"Email has been marked as read.",
            level="warning"
        )
        mark_as_read(service, message_id)
        return

    any_failed = False

    for filename, pdf_bytes in pdfs:
        logger.info(f"  Processing attachment: {filename}")

        # Write PDF to a temp file for the pipeline
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        try:
            process_pdf(tmp_path)
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

    # Mark as read regardless of success/failure
    # Prevents infinite retry loops on persistently broken PDFs
    mark_as_read(service, message_id)

    if any_failed:
        logger.warning(f"Email processed with errors: {subject}")
    else:
        logger.info(f"Email fully processed: {subject}")


def run_once(service):
    """Check inbox once and process all unread messages."""
    messages = get_unread_messages(service)

    if not messages:
        logger.debug("No unread messages.")
        return 0

    logger.info(f"Found {len(messages)} unread message(s)")
    processed = 0

    for msg in messages:
        try:
            process_message(service, msg["id"])
            processed += 1
        except Exception as e:
            logger.error(f"Unexpected error processing message {msg['id']}: {e}")
            # Don't mark as read — will retry next poll
            # But do notify so the team knows
            notify_slack(
                f":red_circle: *Unexpected error in email watcher*\n"
                f"Message ID: {msg['id']}\n"
                f"Error: `{str(e)[:300]}`\n"
                f"Email left unread for retry.",
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

    while True:
        try:
            run_once(service)
        except HttpError as e:
            logger.error(f"Gmail API error: {e}")
            notify_slack(f":red_circle: *Gmail API error in watcher*\n`{str(e)[:300]}`", level="error")
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