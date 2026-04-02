"""
transform.py
Takes a parsed CommonSku SO dict (from parse_commonsku_so.py) and returns
a Monday.com-ready dict with field names and values matching the board.

Usage:
    from parse_commonsku_so import parse
    from transform import to_monday

    order = parse("SALES_ORDER-70443.pdf")
    for product in order["products"]:
        monday_item = to_monday(order, product)
"""

import re
from datetime import datetime


# ---------------------------------------------------------------------------
# Lookup tables
# ---------------------------------------------------------------------------

IMPRINT_TYPE_MAP = {
    "screenprinting":     "SCREEN PRINT",
    "embroidery":         "EMBROIDERY",
    "dtf":                "DTF",
    "direct to garment":  "DTF",
    "direct to film":     "DTF",
    "heat transfer":      "SCREEN PRINT",   # closest Monday equivalent
    "patch":              "PATCHES",
    "chenille patch":     "CHENILLE PATCH",
    "dtf patch":          "DTF PATCH",
    "tackle twill":       "Tackle Twill",
    "applique":           "APPLIQUE",
    "flocking":           "FLOCKING",
    "pad printing":       "PAD PRINTING",
    "custom fabric label": None,            # not a decoration type — drives label flags
}

SPECIALTY_INK_KEYWORDS = [
    "glow in the dark",
    "puff",
    "hd ink",
    "gloss",
    "foil",
    "metallic",
    "discharge",
    "neon",
    "water base",
    "suede",
]

# Valid LOCATIONS dropdown options from Monday board (exact match required)
# Anything not in this set is a finishing/label field handled separately
VALID_LOCATIONS = {
    "FRONT CHEST", "LEFT CHEST", "RIGHT CHEST",
    "RIGHT SLEEVE", "LEFT SLEEVE",
    "FULL BACK", "BACK PRINT", "BACK CENTER", "BACK POCKET",
    "LEFT THIGH", "RIGHT THIGH", "LEFT CALF", "RIGHT CALF",
    "FRONT CENTER", "BOTTOM LEFT",
    "HAT - RIGHT SIDE", "HAT - LEFT SIDE", "HAT - BACK CENTER", "HAT - FRONT",
    "WRIST", "NECK LABEL",
    "ALL OVER",
    "RIGHT HIP", "LEFT HIP", "RIGHT LEG", "LEFT LEG",
    "Left Cuff", "Right Cuff", "Hat Interior label",
    "PRINTED NECK TAG", "CUSTOM CLIP LABEL",
    "BLANK",
    # Hat-specific
    "FRONT CENTRE", "BACK CENTRE", "RIGHT SIDE", "LEFT SIDE",
    "CLIP LABEL", "FRONT PATCH",
    "LEFT SIDE - EMBROIDERY", "RIGHT SIDE - EMBROIDERY",
    "BACK - EMBROIDERY", "FRONT - EMBROIDERY",
}

# These valid locations are handled by dedicated Monday fields, not LOCATIONS column
LOCATIONS_SEPARATE = {
    "printed neck tag",
    "custom clip label",
    "custom woven label",
    "woven label",
    "woven neck label",
    "mockup",
    "clip label",
    "interior taping",
    "satin care tear away label",
    "woven brand clip label",
    "tech pack",
}


# ---------------------------------------------------------------------------
# Field transformers
# ---------------------------------------------------------------------------

def _client_name(client):
    """Strip trailing punctuation from client name."""
    return client.rstrip(".,") if client else None


def _garment_style(style_code):
    """
    Strip Avid supplier prefix (AV, AVN, AVL, AVB etc.) → numeric only.
    Preserve brand-specific prefixes (TB, TC, HB etc.) as-is.
    Examples: AVN6210 → 6210, AVL1605 → 1605, TB0063 → TB0063, TC2965 → TC2965
    """
    if not style_code:
        return None
    # Only strip if it starts with AV followed by optional single letter then digits
    stripped = re.sub(r"^AV[A-Z]?(?=\d)", "", style_code)
    return stripped


def _in_hands_date(date_str):
    """Convert 'Apr 06, 2026' → '2026-04-06'."""
    if not date_str:
        return None
    for fmt in ("%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return date_str  # return as-is if format unrecognised


# Normalise PDF location names to Monday dropdown values
LOCATION_NORMALISE = {
    "FRONT CENTRE":             "FRONT CENTER",
    "BACK CENTRE":              "BACK CENTER",
    "FRONT CENTER":             "FRONT CENTER",
    "BACK CENTER":              "BACK CENTER",
    # Hat locations
    "HAT - FRONT":              "HAT - FRONT",
    "HAT - BACK CENTER":        "HAT - BACK CENTER",
    "HAT - BACK CENTRE":        "HAT - BACK CENTER",
    "HAT - LEFT SIDE":          "HAT - LEFT SIDE",
    "HAT - RIGHT SIDE":         "HAT - RIGHT SIDE",
    "HAT INTERIOR LABEL":       "Hat Interior label",
    "HAT INTERIOR LABEL":       "Hat Interior label",
    # Hat embroidery descriptor style (e.g. "Front Centre", "Back Centre")
    "LEFT SIDE - EMBROIDERY":   "HAT - LEFT SIDE",
    "RIGHT SIDE - EMBROIDERY":  "HAT - RIGHT SIDE",
    "BACK - EMBROIDERY":        "HAT - BACK CENTER",
    "FRONT PATCH":              "HAT - FRONT",
}


def _locations(decoration_locations):
    """
    Build the LOCATIONS string from decoration locations.
    Normalises spelling variants and maps hat-specific descriptor locations
    to Monday dropdown values. Excludes label/finishing locations.
    """
    locs = []
    for loc in decoration_locations:
        loc_upper = loc.upper().strip()
        # Skip anything handled by a dedicated field
        if any(excl in loc.lower() for excl in LOCATIONS_SEPARATE):
            continue
        # Normalise to Monday's exact label if a mapping exists
        normalised = LOCATION_NORMALISE.get(loc_upper, loc_upper)
        # Only include if it's a known valid dropdown option (case-insensitive check)
        if normalised.upper() in {v.upper() for v in VALID_LOCATIONS}:
            locs.append(normalised)
    return ", ".join(locs) if locs else None


def _decoration_type(imprint_types, full_text):
    """
    Build DECORATION TYPE value(s) matching Monday's exact dropdown labels.
    Returns a comma-separated string of valid Monday decoration type labels.
    Specialty ink is detected and appended as Monday's closest matching label.
    """
    base_types = []
    for t in imprint_types:
        mapped = IMPRINT_TYPE_MAP.get(t.lower().strip())
        if mapped and mapped not in base_types:
            base_types.append(mapped)

    if not base_types:
        return None

    # Detect specialty inks and pick the closest Monday label
    text_lower = full_text.lower()
    has_glow = "glow in the dark" in text_lower
    has_neon = "neon" in text_lower
    has_puff = any(kw in text_lower for kw in ["puff", "hd ink", "gloss", "foil"])

    if has_puff or (has_glow and not has_neon):
        base_types.append("**SPECIALTY INK (PUFF, HD, GLOSS, GLOW IN THE DARK")
    elif has_glow or has_neon:
        base_types.append("**SPECIALTY INK (GLOW IN THE DARK, NEON)")

    return ", ".join(base_types)


def _neck_tag_type(decoration_locations, full_text):
    """
    Returns 'PRINT - REPEAT', 'PRINT - NEW', or None.
    PRINT - NEW:    "Printed Neck Tag" in decoration locations, first time.
    PRINT - REPEAT: above + design name contains "RN" or artwork notes say "Repeat of PO".
    """
    has_printed_neck = any(
        "printed neck tag" in loc.lower() for loc in decoration_locations
    )
    if not has_printed_neck:
        return None

    is_repeat = bool(
        re.search(r"\bRN\b", full_text) or
        re.search(r"Repeat of PO", full_text, re.IGNORECASE)
    )
    return "PRINT - REPEAT" if is_repeat else "PRINT - NEW"


def _neck_tag_details(full_text):
    """
    Extract the neck tag design name (e.g. 'TC6040N RN - 2025').
    Looks for a DESIGN NAME associated with a Printed Neck Tag location.
    """
    # Find design blocks where DESIGN LOCATION = Printed Neck Tag
    blocks = re.split(r"(?=DESIGN NAME)", full_text)
    for block in blocks:
        if "Printed Neck Tag" in block or "printed neck tag" in block.lower():
            m = re.search(r"DESIGN NAME\s+(.+)", block)
            if m:
                name = m.group(1).strip()
                # Skip mockup entries
                if "mockup" not in name.lower():
                    return name
    return None


def _clip_label_details(full_text):
    """
    Extract the clip/woven label design name if present.
    """
    blocks = re.split(r"(?=DESIGN NAME)", full_text)
    for block in blocks:
        loc_m = re.search(r"DESIGN LOCATION\s+(.+)", block)
        if not loc_m:
            continue
        loc = loc_m.group(1).lower()
        if "clip label" in loc or "woven label" in loc or "woven neck" in loc:
            name_m = re.search(r"DESIGN NAME\s+(.+)", block)
            if name_m:
                name = name_m.group(1).strip()
                if "mockup" not in name.lower():
                    return name
    return None


def _yes_no_flags(decoration_locations, services, full_text):
    """
    Derive YES/NO flags for finishing fields.
    Checks decoration locations, service names, and full text for keywords.
    """
    loc_text = " ".join(decoration_locations).lower()
    svc_text = " ".join(s["service"] for s in services).lower()
    combined = loc_text + " " + svc_text

    return {
        "CLIP LABEL NEEDED?": "YES" if any(
            kw in combined for kw in ["clip label", "woven label", "woven neck"]
        ) else "NO",
        "HANG TAG": "YES" if ("hang tag" in combined or "hangtag" in combined) else "NO",
        "POLY BAG": "YES" if any(
            kw in combined for kw in ["poly bag", "bagging", "polybag"]
        ) else "NO",
        "BARCODE NEEDED": "YES" if "barcode" in combined else "NO",
    }


# ---------------------------------------------------------------------------
# Main transform
# ---------------------------------------------------------------------------

def to_monday(order, product, full_text=""):
    """
    Takes an order-level dict and a single product dict from the parser,
    returns a flat dict ready for the Monday.com API.

    Pass full_text (raw extracted PDF text) for specialty ink and neck tag detection.
    If not passed, those fields will degrade gracefully.
    """
    flags = _yes_no_flags(
        product.get("decoration_locations", []),
        order.get("services", []),
        full_text
    )

    monday = {
        # Order-level fields
        "AVID SO #":              order.get("so_number"),
        "Name":                   _client_name(order.get("client")),
        "CLIENT PO":              order.get("customer_po"),
        "TERMS $":                order.get("terms", "").upper() if order.get("terms") else None,
        "Customer Expected Date": _in_hands_date(order.get("in_hands_date")),
        "PRODUCTION NOTE":        order.get("production_notes"),
        "CURRENCY":               order.get("currency"),

        # Product-level fields
        "GARMENT STYLE":          _garment_style(product.get("style_code")),
        "GARMENT COLOUR":         product.get("color", "").lower() if product.get("color") else None,
        "TTL SO QUANTITY":        product.get("total_units"),

        # Decoration
        "LOCATIONS":              _locations(product.get("decoration_locations", [])),
        "DECORATION TYPE":        _decoration_type(
                                      product.get("imprint_types", []), full_text
                                  ),
        "NECK TAG TYPE":          _neck_tag_type(
                                      product.get("decoration_locations", []), full_text
                                  ),
        "NECK TAG DETAILS":       _neck_tag_details(full_text),
        "CLIP LABEL DETAILS":     _clip_label_details(full_text),

        # YES/NO finishing flags
        **flags,

        # Always blank on creation — filled manually or operationally
        "ACCOUNT REP":            None,
        "GARMENT DESCRIPTION":    None,
        "SUPPLIER":               None,
        "GARMENT ORIGIN":         None,
        "PRINTER":                None,
        "DECORATION PO":          None,
        "SUPPLIER PO#":           None,
        "PRIORITY":               None,
        "RUSH":                   None,
    }

    return monday
# patch — run after file loads