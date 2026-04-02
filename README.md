# CommonSku → Monday.com Integration

Automatically parses CommonSku Sales Order PDFs and creates items in the Monday.com production tracker board.

## What it does

When a CommonSku SO PDF is processed, the script:
- Parses all order and product fields from the PDF
- Transforms them into Monday.com-ready values
- Creates one Monday item per product line
- Posts a Slack notification on success or failure
- Logs all activity to a local log file
- Posts a warning comment on any Monday item where fields couldn't be auto-populated

## Files

| File | Purpose |
|------|---------|
| `parse_commonsku_so.py` | Extracts fields from the CommonSku PDF |
| `transform.py` | Maps parsed fields to Monday column values |
| `monday_api.py` | Creates Monday items via the API |
| `alerts.py` | Logging, Slack notifications, Monday update comments |
| `.env.template` | Template for credentials — copy to `.env` and fill in |
| `requirements.txt` | Python dependencies |

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure credentials
```bash
cp .env.template .env
nano .env
```

Fill in:
```
MONDAY_API_TOKEN=your_token_here
MONDAY_BOARD_ID=your_board_id_here
MONDAY_GROUP_ID=topics
SLACK_WEBHOOK_URL=your_webhook_url_here
LOG_FILE=/home/bretgs/avid_integration.log
```

### 3. Find your Monday column IDs (first time only)
```bash
python3 monday_api.py --list-columns
python3 monday_api.py --list-groups
```

## Usage

```bash
# Dry run — shows what would be created, no API calls
python3 monday_api.py SALES_ORDER-70443.pdf --dry-run

# Live run — creates Monday items
python3 monday_api.py SALES_ORDER-70443.pdf

# Process multiple PDFs
for f in SALES_ORDER-*.pdf; do
    python3 monday_api.py "$f"
done
```

## Monitoring

Check the log file:
```bash
tail -50 ~/avid_integration.log
```

## Field mapping

| Monday Field | Source | Notes |
|---|---|---|
| AVID SO # | SO number from PDF | Exact |
| Name | Client name | Trailing punctuation stripped |
| CLIENT PO | Customer PO | Blank if not supplied |
| GARMENT STYLE | Style code | AV prefix stripped (AVN6210 → 6210), brand prefixes kept (TB0063 → TB0063) |
| GARMENT COLOUR | Color | Lowercased |
| TTL SO QUANTITY | Total units | Exact |
| TERMS $ | Payment terms | Mapped to Monday status labels |
| Customer Expected Date | In hands date | Reformatted to YYYY-MM-DD |
| PRODUCTION NOTE | All-caps instruction block | Blank if not present in PDF |
| LOCATIONS | Decoration locations | Filtered to valid Monday dropdown options |
| DECORATION TYPE | Imprint types + specialty ink detection | Mapped to Monday dropdown labels |
| NECK TAG TYPE | Decoration locations + repeat detection | PRINT - NEW or PRINT - REPEAT |
| NECK TAG DETAILS | Neck tag design name | Blank if no printed neck tag |
| CLIP LABEL NEEDED? | Decoration locations + services | YES/NO |
| CLIP LABEL DETAILS | Clip/woven label design name | Blank if no clip label |
| HANG TAG | Services + decoration locations | YES/NO |
| POLY BAG | Services | YES/NO |
| BARCODE NEEDED | Services + decoration locations | YES/NO |

**Always filled manually after creation:**
ACCOUNT REP, SUPPLIER, GARMENT ORIGIN, PRINTER, GARMENT DESCRIPTION, DECORATION PO, SUPPLIER PO#, PRIORITY, RUSH

## Adding new decoration types or locations

- **New imprint type**: add to `IMPRINT_TYPE_MAP` in `transform.py`
- **New terms value**: add to `TERMS_MAP` in `monday_api.py`
- **New decoration location**: add to `VALID_LOCATIONS` in `transform.py`
- **New location spelling variant**: add to `LOCATION_NORMALISE` in `transform.py`