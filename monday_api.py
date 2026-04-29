"""
monday_api.py
Creates Monday.com items from transformed CommonSku SO data.

Requires a .env file in the same directory containing:
    MONDAY_API_TOKEN=your_token_here
    MONDAY_BOARD_ID=your_board_id_here
    MONDAY_GROUP_ID=your_group_id_here

Usage:
    python3 monday_api.py SALES_ORDER-70443.pdf --dry-run
    python3 monday_api.py SALES_ORDER-70443.pdf
"""

import os
import sys
import json
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

import sys as _sys
_sys.path.insert(0, str(Path(__file__).parent))
from alerts import (
    logger, log_success, alert_error,
    alert_field_warnings, notify_slack, post_monday_update
)

MONDAY_API_URL = "https://api.monday.com/v2"
API_TOKEN      = os.environ.get("MONDAY_API_TOKEN")
BOARD_ID       = os.environ.get("MONDAY_BOARD_ID")
GROUP_ID       = os.environ.get("MONDAY_GROUP_ID")


COLUMN_MAP = {
    "AVID SO #":              "text15",
    "CLIENT PO":              "text3",
    "GARMENT STYLE":          "text32",
    "GARMENT COLOUR":         "text1",
    "TTL SO QUANTITY":        "numeric7",
    "TERMS $":                "color",
    "Customer Expected Date": "date_mm07a2w6",
    "PRODUCTION NOTE":        "long_text6",
    "LOCATIONS":              "dropdown44",
    "DECORATION TYPE":        "dropdown9",
    "NECK TAG TYPE":          "color64",
    "NECK TAG DETAILS":       "long_text2",
    "CLIP LABEL DETAILS":     "long_text23",
    "CLIP LABEL NEEDED?":     "color63",
    "HANG TAG":               "label5",
    "POLY BAG":               "label8",
    "BARCODE NEEDED":         "dup__of_poly_bag",
    "CURRENCY":               None,
    "Troll Co Order?":        "color_mm092wkj",
    "REPEAT ORDER?":          "repeat_order_",
    "PO VALUE":               "numeric8",
    "Troll Co Style #":       "text_mm03p9v6",
    "ARTWORK":                "dup__of_urgent___notes",
}

# Maps parsed terms strings to Monday's exact status labels
TERMS_MAP = {
    "NET 90":                "NET 90",
    "NET 60":                "NET 60",
    "NET 45":                "NET 45",
    "NET 30":                "NET 30",
    "NET 20":                "NET20",
    "NET 15":                "NET 15",
    "NET 7":                 "NET7",
    "COD":                   "C.O.D",
    "C.O.D":                 "C.O.D",
    "50% DEPOSIT / BALANCE": "50% DEPOSIT 50% ON SHIPMENT",
    "50% DEPOSIT":           "50% DEPOSIT 50% ON SHIPMENT",
}


def _headers():
    if not API_TOKEN:
        raise RuntimeError("MONDAY_API_TOKEN not set in .env")
    return {
        "Authorization": API_TOKEN,
        "Content-Type":  "application/json",
        "API-Version":   "2024-01",
    }


def _run_query(query, variables=None):
    payload = {"query": query}
    if variables:
        payload["variables"] = variables
    resp = requests.post(MONDAY_API_URL, headers=_headers(), json=payload, timeout=30)
    if not resp.ok:
        raise RuntimeError(f"Monday API {resp.status_code}: {resp.text[:500]}")
    resp.raise_for_status()
    data = resp.json()
    if "errors" in data:
        raise RuntimeError(f"Monday API error: {data['errors']}")
    return data["data"]


def fetch_board_columns():
    if not BOARD_ID:
        raise RuntimeError("MONDAY_BOARD_ID not set.")
    query = """
    query ($board_id: ID!) {
        boards(ids: [$board_id]) {
            columns { id title type }
        }
    }
    """
    data = _run_query(query, {"board_id": BOARD_ID})
    return data["boards"][0]["columns"]


def fetch_board_groups():
    if not BOARD_ID:
        raise RuntimeError("MONDAY_BOARD_ID not set.")
    query = """
    query ($board_id: ID!) {
        boards(ids: [$board_id]) {
            groups { id title }
        }
    }
    """
    data = _run_query(query, {"board_id": BOARD_ID})
    return data["boards"][0]["groups"]


def _split_top_level(s):
    """Split on commas outside parentheses — handles DECORATION TYPE specialty ink notation."""
    parts, depth, current = [], 0, []
    for char in s:
        if char == '(':
            depth += 1
            current.append(char)
        elif char == ')':
            depth -= 1
            current.append(char)
        elif char == ',' and depth == 0:
            parts.append(''.join(current).strip())
            current = []
        else:
            current.append(char)
    if current:
        parts.append(''.join(current).strip())
    return [p for p in parts if p]


def _build_column_values(monday_dict):
    """
    Convert a to_monday() dict into a Python dict of Monday column values.
    Each value is a plain Python object — caller does ONE json.dumps on the whole dict.

    Monday value formats by type:
      text/long_text  ->  "plain string"
      numbers         ->  255
      date            ->  {"date": "YYYY-MM-DD"}
      status          ->  {"label": "Option Name"}
      dropdown        ->  {"labels": ["Option 1", "Option 2"]}
    """
    column_values = {}

    for field_name, column_id in COLUMN_MAP.items():
        if column_id is None:
            continue
        value = monday_dict.get(field_name)
        if value is None:
            continue

        if field_name == "Customer Expected Date":
            column_values[column_id] = {"date": str(value)}

        elif field_name in ("TTL SO QUANTITY", "PO VALUE"):
            # Monday numeric columns reject comma-formatted strings (e.g. "13,860.00")
            if isinstance(value, str):
                value = float(value.replace(",", ""))
            column_values[column_id] = value  # plain int/float

        elif field_name in ("DECORATION TYPE", "LOCATIONS"):
            column_values[column_id] = {"labels": _split_top_level(str(value))}

        elif field_name == "TERMS $":
            label = TERMS_MAP.get(str(value).upper(), str(value))
            column_values[column_id] = {"label": label}

        elif field_name in (
            "CLIP LABEL NEEDED?", "HANG TAG", "POLY BAG", "BARCODE NEEDED",
            "NECK TAG TYPE", "Troll Co Order?", "REPEAT ORDER?",
        ):
            column_values[column_id] = {"label": str(value)}

        else:
            column_values[column_id] = str(value)  # plain text

    return column_values


def find_existing_items(so_number):
    """
    Search the board for items where AVID SO # matches so_number.
    Returns a list of items sorted by created_at ascending (oldest first = line 1).
    """
    if not BOARD_ID:
        raise RuntimeError("MONDAY_BOARD_ID not set.")

    query = """
    query ($board_id: ID!, $col_id: String!, $col_val: String!) {
        items_page_by_column_values(
            board_id: $board_id,
            columns: [{column_id: $col_id, column_values: [$col_val]}]
            limit: 50
        ) {
            items {
                id
                name
                created_at
            }
        }
    }
    """
    data = _run_query(query, {
        "board_id": BOARD_ID,
        "col_id":   COLUMN_MAP["AVID SO #"],
        "col_val":  str(so_number),
    })
    items = data["items_page_by_column_values"]["items"]
    # Sort by creation time so positional matching works
    return sorted(items, key=lambda x: x["created_at"])


def update_item(item_id, monday_dict, dry_run=False):
    """Update an existing Monday item with new column values."""
    column_values = _build_column_values(monday_dict)
    item_name = monday_dict.get("Name") or "Unnamed Item"

    if dry_run:
        print(f"\n[DRY RUN] Would UPDATE existing Monday item {item_id}:")
        print(f"  Name:   {item_name}")
        print(f"  Columns ({len(column_values)} mapped):")
        for col_id, val in column_values.items():
            human = next((k for k, v in COLUMN_MAP.items() if v == col_id), col_id)
            print(f"    {human:<30} ({col_id}): {val}")
        return item_id

    mutation = """
    mutation ($item_id: ID!, $board_id: ID!, $column_values: JSON!) {
        change_multiple_column_values(
            item_id: $item_id,
            board_id: $board_id,
            column_values: $column_values
        ) {
            id
        }
    }
    """
    variables = {
        "item_id":       item_id,
        "board_id":      BOARD_ID,
        "column_values": json.dumps(column_values),
    }
    data = _run_query(mutation, variables)
    print(f"  Updated Monday item {item_id}: {item_name}")
    return item_id


def upsert_item(monday_dict, line_index, existing_items, dry_run=False):
    """
    Create or update a Monday item for a single product line.

    Matches by position within the SO:
      - line_index 0 = first item created for this SO
      - line_index 1 = second item created, etc.

    If an existing item exists at this position: update it.
    If not: create a new item.
    """
    if not BOARD_ID:
        raise RuntimeError("MONDAY_BOARD_ID not set.")
    if not GROUP_ID and line_index >= len(existing_items):
        raise RuntimeError("MONDAY_GROUP_ID not set. Run --list-groups to find it.")

    item_name = monday_dict.get("Name") or "Unnamed Item"
    column_values = _build_column_values(monday_dict)

    if line_index < len(existing_items):
        # UPDATE existing item
        existing = existing_items[line_index]
        item_id = existing["id"]

        if dry_run:
            print(f"\n[DRY RUN] Would UPDATE existing item {item_id} (line {line_index + 1}):")
            print(f"  Name:   {item_name}")
            print(f"  Columns ({len(column_values)} mapped):")
            for col_id, val in column_values.items():
                human = next((k for k, v in COLUMN_MAP.items() if v == col_id), col_id)
                print(f"    {human:<30} ({col_id}): {val}")
            unmapped = [k for k, v in COLUMN_MAP.items()
                        if v is None and monday_dict.get(k) is not None]
            if unmapped:
                print(f"\n  [!] These fields have values but no column ID yet:")
                for u in unmapped:
                    print(f"      {u}: {monday_dict.get(u)}")
            return item_id

        mutation = """
        mutation ($item_id: ID!, $board_id: ID!, $column_values: JSON!) {
            change_multiple_column_values(
                item_id: $item_id,
                board_id: $board_id,
                column_values: $column_values
            ) {
                id
            }
        }
        """
        _run_query(mutation, {
            "item_id":       item_id,
            "board_id":      BOARD_ID,
            "column_values": json.dumps(column_values),
        })
        logger.info(f"  Updated item {item_id}: {item_name}")
        return item_id

    else:
        # CREATE new item
        if dry_run:
            print(f"\n[DRY RUN] Would CREATE new Monday item (line {line_index + 1}):")
            print(f"  Board:  {BOARD_ID}")
            print(f"  Group:  {GROUP_ID}")
            print(f"  Name:   {item_name}")
            print(f"  Columns ({len(column_values)} mapped):")
            for col_id, val in column_values.items():
                human = next((k for k, v in COLUMN_MAP.items() if v == col_id), col_id)
                print(f"    {human:<30} ({col_id}): {val}")
            unmapped = [k for k, v in COLUMN_MAP.items()
                        if v is None and monday_dict.get(k) is not None]
            if unmapped:
                print(f"\n  [!] These fields have values but no column ID yet:")
                for u in unmapped:
                    print(f"      {u}: {monday_dict.get(u)}")
            return None

        mutation = """
        mutation ($board_id: ID!, $group_id: String!, $item_name: String!, $column_values: JSON!) {
            create_item(
                board_id: $board_id,
                group_id: $group_id,
                item_name: $item_name,
                column_values: $column_values
            ) {
                id
            }
        }
        """
        data = _run_query(mutation, {
            "board_id":      BOARD_ID,
            "group_id":      GROUP_ID,
            "item_name":     item_name,
            "column_values": json.dumps(column_values),
        })
        item_id = data["create_item"]["id"]
        logger.info(f"  Created item {item_id}: {item_name}")
        return item_id


def process_pdf(pdf_path, dry_run=False):
    sys.path.insert(0, str(Path(__file__).parent))
    from parse_commonsku_so import parse, extract_text
    from transform import to_monday

    logger.info(f"Processing: {Path(pdf_path).name}")

    try:
        order = parse(pdf_path)
        full_text = extract_text(pdf_path)
    except Exception as e:
        alert_error(pdf_path, "unknown", "unknown", e)
        raise

    so_number = order.get("so_number", "unknown")
    client    = order.get("client", "unknown")

    logger.info(f"SO {so_number} - {client} - {len(order['products'])} product line(s)")

    # Find any existing Monday items for this SO number
    # Search runs in both live and dry-run mode so output accurately shows create vs update
    existing_items = []
    try:
        existing_items = find_existing_items(so_number)
        if existing_items:
            action = "update (dry run)" if dry_run else "update"
            logger.info(f"  Found {len(existing_items)} existing item(s) for SO {so_number} — will {action}")
    except Exception as e:
        logger.warning(f"  Could not search for existing items: {e} — will create new")

    upserted_ids = []
    created_count = 0
    updated_count = 0

    for i, product in enumerate(order["products"]):
        monday_dict = to_monday(order, product, full_text)
        style = product.get("style_code") or "no style code"
        color = product.get("color") or "no color"
        action = "Update" if i < len(existing_items) else "Create"
        logger.info(f"  Line {i+1}: [{action}] {style} / {color} / {product.get('total_units')} units")

        item_id = None
        try:
            item_id = upsert_item(monday_dict, i, existing_items, dry_run=dry_run)
        except Exception as e:
            alert_error(pdf_path, so_number, client, e, item_id=None)
            notify_slack(
                f"*SO {so_number} Line {i+1} failed* ({Path(pdf_path).name})\n"
                f"Client: {client} | Style: {style} / {color}\n"
                f"Error: `{str(e)[:200]}`",
                level="error"
            )
            raise

        if item_id:
            upserted_ids.append(item_id)
            if i < len(existing_items):
                updated_count += 1
            else:
                created_count += 1
            missing = [k for k, v in monday_dict.items() if v is None]
            alert_field_warnings(item_id, so_number, missing)

    if not dry_run and upserted_ids:
        log_success(so_number, client, pdf_path, upserted_ids)
        parts = []
        if created_count:
            parts.append(f"{created_count} created")
        if updated_count:
            parts.append(f"{updated_count} updated")

        # Build Monday item links
        base_url = f"https://avidapparel.monday.com/boards/{BOARD_ID}/pulses"
        item_links = " | ".join(f"<{base_url}/{item_id}|View item>" for item_id in upserted_ids)

        notify_slack(
            f"*SO {so_number} processed* — {client}\n"
            f"{', '.join(parts)} from {Path(pdf_path).name}\n"
            f"{item_links}",
            level="success"
        )

    return upserted_ids


if __name__ == "__main__":
    args = sys.argv[1:]

    if "--list-columns" in args:
        print(f"\nColumns for board {BOARD_ID}:\n")
        cols = fetch_board_columns()
        print(f"  {'ID':<20} {'TYPE':<20} TITLE")
        print("  " + "-" * 70)
        for c in cols:
            print(f"  {c['id']:<20} {c['type']:<20} {c['title']}")
        sys.exit(0)

    if "--list-groups" in args:
        print(f"\nGroups for board {BOARD_ID}:\n")
        for g in fetch_board_groups():
            print(f"  {g['id']:<30} {g['title']}")
        sys.exit(0)

    if not args or args[0].startswith("--"):
        print(__doc__)
        sys.exit(1)

    pdf_path = args[0]
    dry_run = "--dry-run" in args

    if dry_run:
        print("[DRY RUN MODE - no items will be created]")

    process_pdf(pdf_path, dry_run=dry_run)