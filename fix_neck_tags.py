"""
fix_neck_tags.py
Targeted one-off: updates NECK TAG TYPE and NECK TAG DETAILS on existing
Monday items that were imported with blank neck labels.

Usage:
    python3 fix_neck_tags.py --dry-run   # preview only
    python3 fix_neck_tags.py             # live update
"""

import os
import sys
import json
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

sys.path.insert(0, str(Path(__file__).parent))
from parse_commonsku_so import parse, extract_text
from transform import _neck_tag_type, _neck_tag_details

MONDAY_API_URL = "https://api.monday.com/v2"
API_TOKEN      = os.environ.get("MONDAY_API_TOKEN")
BOARD_ID       = "3607906471"  # PRODUCTION TRACKER (env has sandbox)

NECK_TAG_TYPE_COL    = "color64"
NECK_TAG_DETAILS_COL = "long_text2"
AVID_SO_COL          = "text15"

PDFS = [
    "/tmp/attachments5/Sales Order #74137.pdf",
    "/tmp/attachments5/Sales Order #74138.pdf",
    "/tmp/attachments5/Sales Order #74139.pdf",
    "/tmp/attachments5/Sales Order #74140.pdf",
    "/tmp/attachments5/Sales Order #74141.pdf",
    "/tmp/attachments5/Sales Order #74142.pdf",
    "/tmp/attachments5/Sales Order #74143.pdf",
    "/tmp/attachments5/Sales Order #74144.pdf",
    "/tmp/attachments5/Sales Order #74148.pdf",
]


def _headers():
    return {
        "Authorization": API_TOKEN,
        "Content-Type":  "application/json",
        "API-Version":   "2024-01",
    }


def _gql(query, variables=None):
    payload = {"query": query}
    if variables:
        payload["variables"] = variables
    resp = requests.post(MONDAY_API_URL, headers=_headers(), json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    if "errors" in data:
        raise RuntimeError(f"Monday error: {data['errors']}")
    return data["data"]


def find_items(so_number):
    q = """
    query ($board_id: ID!, $col_id: String!, $col_val: String!) {
        items_page_by_column_values(
            board_id: $board_id,
            columns: [{column_id: $col_id, column_values: [$col_val]}]
            limit: 50
        ) { items { id name created_at } }
    }
    """
    data = _gql(q, {"board_id": BOARD_ID, "col_id": AVID_SO_COL, "col_val": str(so_number)})
    items = data["items_page_by_column_values"]["items"]
    return sorted(items, key=lambda x: x["created_at"])


def update_neck_tags(item_id, neck_type, neck_details, dry_run=False):
    col_values = {}
    if neck_type:
        col_values[NECK_TAG_TYPE_COL] = {"label": neck_type}
    if neck_details:
        col_values[NECK_TAG_DETAILS_COL] = neck_details

    if not col_values:
        print(f"  [{item_id}] nothing to set — skipping")
        return

    if dry_run:
        print(f"  [DRY RUN] item {item_id}: NECK TAG TYPE={neck_type!r}  DETAILS={neck_details!r}")
        return

    mutation = """
    mutation ($item_id: ID!, $board_id: ID!, $column_values: JSON!) {
        change_multiple_column_values(
            item_id: $item_id,
            board_id: $board_id,
            column_values: $column_values
        ) { id }
    }
    """
    _gql(mutation, {
        "item_id":       item_id,
        "board_id":      BOARD_ID,
        "column_values": json.dumps(col_values),
    })
    print(f"  Updated item {item_id}: NECK TAG TYPE={neck_type!r}  DETAILS={neck_details!r}")


def main(dry_run=False):
    for pdf_path in PDFS:
        if not Path(pdf_path).exists():
            print(f"MISSING: {pdf_path} — skipping")
            continue

        order = parse(pdf_path)
        full_text = extract_text(pdf_path)
        so = order.get("so_number", "?")
        print(f"\nSO {so} ({Path(pdf_path).name})")

        existing = find_items(so)
        if not existing:
            print(f"  No Monday items found for SO {so} — skipping")
            continue

        for i, product in enumerate(order["products"]):
            dec_locs = product.get("decoration_locations", [])
            neck_type    = _neck_tag_type(dec_locs, full_text)
            neck_details = _neck_tag_details(full_text)

            print(f"  Line {i+1}: type={neck_type!r}  details={neck_details!r}")

            if i < len(existing):
                update_neck_tags(existing[i]["id"], neck_type, neck_details, dry_run=dry_run)
            else:
                print(f"  Line {i+1}: no matching Monday item (only {len(existing)} found)")


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv
    if dry_run:
        print("[DRY RUN — no changes]\n")
    main(dry_run=dry_run)
