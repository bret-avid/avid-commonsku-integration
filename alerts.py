"""
alerts.py
Handles logging, Monday update comments, and Slack notifications
for the CommonSku → Monday integration.

Reads from .env:
    SLACK_WEBHOOK_URL   — Slack incoming webhook URL
    LOG_FILE            — path to log file (default: ~/avid_integration.log)
"""

import os
import json
import logging
import traceback
import requests
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

SLACK_WEBHOOK_URL = os.environ.get("SLACK_WEBHOOK_URL")
LOG_FILE          = os.environ.get("LOG_FILE", str(Path.home() / "avid_integration.log"))
MONDAY_API_URL    = "https://api.monday.com/v2"
API_TOKEN         = os.environ.get("MONDAY_API_TOKEN")


# ---------------------------------------------------------------------------
# Logging setup — writes to file and stdout simultaneously
# ---------------------------------------------------------------------------

def setup_logging():
    logger = logging.getLogger("avid")
    if logger.handlers:
        return logger  # already set up

    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")

    # File handler — always on
    fh = logging.FileHandler(LOG_FILE)
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    return logger


logger = setup_logging()


# ---------------------------------------------------------------------------
# Slack
# ---------------------------------------------------------------------------

def notify_slack(message, level="error"):
    """Post a message to Slack. Silently skips if webhook not configured."""
    if not SLACK_WEBHOOK_URL or SLACK_WEBHOOK_URL == "your_slack_webhook_url_here":
        return

    emoji = {"error": ":red_circle:", "warning": ":warning:", "success": ":white_check_mark:"}.get(level, ":information_source:")

    payload = {
        "text": f"{emoji} *Avid Integration*\n{message}",
        "unfurl_links": False,
    }

    try:
        resp = requests.post(SLACK_WEBHOOK_URL, json=payload, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        logger.warning(f"Slack notification failed: {e}")


# ---------------------------------------------------------------------------
# Monday update comment
# ---------------------------------------------------------------------------

def post_monday_update(item_id, message):
    """Post an update comment on a Monday item."""
    if not API_TOKEN or not item_id:
        return

    mutation = """
    mutation ($item_id: ID!, $body: String!) {
        create_update(item_id: $item_id, body: $body) {
            id
        }
    }
    """
    try:
        resp = requests.post(
            MONDAY_API_URL,
            headers={
                "Authorization": API_TOKEN,
                "Content-Type": "application/json",
                "API-Version": "2024-01",
            },
            json={"query": mutation, "variables": {"item_id": item_id, "body": message}},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        if "errors" in data:
            logger.warning(f"Monday update post failed: {data['errors']}")
    except Exception as e:
        logger.warning(f"Monday update post failed: {e}")


# ---------------------------------------------------------------------------
# High-level alert helpers
# ---------------------------------------------------------------------------

def log_success(so_number, client, pdf_path, item_ids):
    """Log a successful run."""
    items_str = ", ".join(str(i) for i in item_ids)
    msg = f"SO {so_number} ({client}) — {len(item_ids)} item(s) created [{items_str}] from {Path(pdf_path).name}"
    logger.info(msg)


def alert_field_warnings(item_id, so_number, missing_fields):
    """
    Post a Monday update if any expected fields couldn't be populated.
    Only flags fields that are genuinely missing (not correctly-null fields).
    """
    # Fields that are legitimately blank on some orders — don't alert on these
    expected_nulls = {
        "PRODUCTION NOTE", "CLIENT PO", "GARMENT STYLE",
        "NECK TAG TYPE", "NECK TAG DETAILS", "CLIP LABEL DETAILS",
        "ACCOUNT REP", "GARMENT DESCRIPTION", "SUPPLIER",
        "GARMENT ORIGIN", "PRINTER", "DECORATION PO",
        "SUPPLIER PO#", "PRIORITY", "RUSH",
    }

    unexpected = [f for f in missing_fields if f not in expected_nulls]
    if not unexpected:
        return

    msg = (
        f"⚠️ Auto-import for SO {so_number} completed but could not populate: "
        + ", ".join(unexpected)
        + ". Please review and fill in manually."
    )
    post_monday_update(item_id, msg)
    logger.warning(f"SO {so_number} — field warnings posted to item {item_id}: {unexpected}")


def alert_error(pdf_path, so_number, client, error, item_id=None):
    """
    Handle a processing error:
    - Log it with full traceback
    - Post to Slack
    - Post to Monday item if one was created
    """
    pdf_name = Path(pdf_path).name
    short_err = str(error)[:300]
    tb = traceback.format_exc()

    # Log full detail
    logger.error(f"FAILED: {pdf_name} (SO {so_number} / {client})\n{tb}")

    # Slack — concise
    slack_msg = (
        f"*Failed to process {pdf_name}*\n"
        f"SO: {so_number} | Client: {client}\n"
        f"Error: `{short_err}`"
    )
    notify_slack(slack_msg, level="error")

    # Monday — if item was partially created
    if item_id:
        monday_msg = (
            f"⚠️ Auto-import error for SO {so_number}.\n"
            f"Some fields may be missing. Error: {short_err}\n"
            f"Please review this item manually."
        )
        post_monday_update(item_id, monday_msg)


def alert_parse_warning(pdf_path, so_number, warning):
    """Log a non-fatal parsing warning."""
    logger.warning(f"Parse warning — {Path(pdf_path).name} (SO {so_number}): {warning}")