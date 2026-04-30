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
    "screenprinting":                    "SCREEN PRINT",
    "embroidery":                        "EMBROIDERY",
    "dtf":                               "DTF",
    "direct to film":                    "DTF",
    "dtg":                               "DTF",   # Monday uses DTF for DTG
    "direct to garment":                 "DTF",
    "digital direct to garment":         "DTF",
    "digital direct to garment (dtg)":   "DTF",
    "heat transfer":                     "SCREEN PRINT",
    "patch":                             "PATCHES",
    "chenille patch":                    "CHENILLE PATCH",
    "dtf patch":                         "DTF PATCH",
    "tackle twill":                      "Tackle Twill",
    "applique":                          "APPLIQUE",
    "flocking":                          "FLOCKING",
    "pad printing":                      "PAD PRINTING",
    "custom fabric label":               None,   # drives label flags, not decoration type
    "other":                             None,   # mockups etc
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

# ---------------------------------------------------------------------------
# Client-level overrides
# Fields listed here will be forced to YES for matching clients,
# regardless of what the PDF says. Add clients as needed.
# Keys are lowercase substrings to match against the client name.
# ---------------------------------------------------------------------------

CLIENT_OVERRIDES = {
    # Format: "client name substring (lowercase)": {"FIELD": "VALUE", ...}
    # Example:
    # "troll co": {"BARCODE NEEDED": "YES", "POLY BAG": "YES"},
}


# ---------------------------------------------------------------------------
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
    "hang tag",
    "hangtag",
    "hang tag reference",
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
    # Centre/Center spelling variants
    "FRONT CENTRE":             "FRONT CENTER",
    "BACK CENTRE":              "BACK CENTER",
    "FRONT CENTER":             "FRONT CENTER",
    "BACK CENTER":              "BACK CENTER",
    # "Print" suffix variants — strip the qualifier, keep the location
    "FRONT CHEST PRINT":        "FRONT CHEST",
    "LEFT CHEST PRINT":         "LEFT CHEST",
    "BACK PRINT":               "FULL BACK",
    "FULL BACK PRINT":          "FULL BACK",
    "FRONT PRINT":              "FRONT CHEST",
    "BACK CENTER PRINT":        "FULL BACK",
    "FRONT CENTRE PRINT":       "FRONT CENTER",
    "BACK CENTRE PRINT":        "FULL BACK",
    # Hat locations
    "HAT - FRONT":              "HAT - FRONT",
    "HAT - BACK CENTER":        "HAT - BACK CENTER",
    "HAT - BACK CENTRE":        "HAT - BACK CENTER",
    "HAT - LEFT SIDE":          "HAT - LEFT SIDE",
    "HAT - RIGHT SIDE":         "HAT - RIGHT SIDE",
    "HAT INTERIOR LABEL":       "Hat Interior label",
    # Hat embroidery descriptor style
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
    Returns NECK TAG TYPE value based on what's in the decoration locations.

    Printed neck tag:
      PRINT - NEW    — Printed Neck Tag location, no repeat markers
      PRINT - REPEAT — Printed Neck Tag location + RN in design name or "Repeat of PO"

    Woven neck label:
      WOVEN - NEW    — woven/neck label location, no repeat markers
      WOVEN - REPEAT — woven/neck label location + repeat markers

    Woven takes priority if both are present (unusual but possible).
    """
    is_repeat = bool(
        re.search(r"\bRN\b", full_text) or
        re.search(r"Repeat of PO", full_text, re.IGNORECASE)
    )

    has_woven_neck = any(
        any(kw in loc.lower() for kw in ["woven neck", "woven label", "neck label", "woven clip"])
        and "hem" not in loc.lower()  # exclude woven hem labels — those are clip labels
        for loc in decoration_locations
    )
    if has_woven_neck:
        return "WOVEN - REPEAT" if is_repeat else "WOVEN - NEW"

    has_printed_neck = any("printed neck tag" in loc.lower() for loc in decoration_locations)
    if has_printed_neck:
        return "PRINT - REPEAT" if is_repeat else "PRINT - NEW"

    return None


def _neck_tag_details(full_text):
    """
    Extract the neck tag design name for both printed and woven neck tags.
    Checks for Printed Neck Tag first, then woven neck label locations.
    """
    blocks = re.split(r"(?=DESIGN NAME)", full_text)
    for block in blocks:
        loc_m = re.search(r"DESIGN LOCATION\s+(.+)", block)
        if not loc_m:
            continue
        loc = loc_m.group(1).lower()
        is_neck = (
            "printed neck tag" in loc or
            ("woven" in loc and "neck" in loc and "hem" not in loc) or
            ("neck label" in loc and "hem" not in loc)
        )
        if is_neck:
            name_m = re.search(r"DESIGN NAME\s+(.+)", block)
            if name_m:
                name = name_m.group(1).strip()
                if "mockup" not in name.lower():
                    return name
    return None


def _clip_label_details(full_text):
    """
    Extract the clip/woven label design name.
    Only matches hem/side clip labels, NOT woven neck labels
    (those go into NECK TAG DETAILS instead).
    """
    blocks = re.split(r"(?=DESIGN NAME)", full_text)
    for block in blocks:
        loc_m = re.search(r"DESIGN LOCATION\s+(.+)", block)
        if not loc_m:
            continue
        loc = loc_m.group(1).lower()
        is_clip = (
            "clip label" in loc or
            ("woven label" in loc and "neck" not in loc) or
            ("woven" in loc and "hem" in loc) or
            ("woven" in loc and "side" in loc)
        )
        if is_clip:
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

    # Woven neck labels are NOT clip labels — split the detection
    loc_str = " ".join(decoration_locations).lower()
    svc_str = " ".join(s["service"] for s in services).lower()

    # Clip label: woven hem/side labels, explicit clip labels
    is_clip = any(
        ("clip label" in t.lower()) or
        ("woven label" in t.lower() and "neck" not in t.lower()) or
        ("woven" in t.lower() and "hem" in t.lower())
        for t in decoration_locations + [s["service"] for s in services]
        if isinstance(t, str)
    )

    # Hang tag: explicit hang tag location/service OR neck label in services
    is_hang_tag = (
        "hang tag" in loc_str or "hangtag" in loc_str or
        "hang tag reference" in loc_str or
        "hang tag" in svc_str or "hangtag" in svc_str or
        "neck label" in svc_str
    )

    return {
        "CLIP LABEL NEEDED?": "YES" if is_clip else "NO",
        "HANG TAG":           "YES" if is_hang_tag else "NO",
        "POLY BAG": "YES" if any(
            kw in loc_str + " " + svc_str for kw in ["poly bag", "bagging", "polybag"]
        ) else "NO",
        "BARCODE NEEDED": "YES" if "barcode" in loc_str + " " + svc_str else "NO",
    }


# ---------------------------------------------------------------------------
# Main transform
# ---------------------------------------------------------------------------

def _is_troll_co(client_name):
    """Returns True if the client is Troll Co. (case-insensitive, handles variants)."""
    if not client_name:
        return False
    normalised = client_name.lower().replace(" ", "").replace(".", "")
    return "trollco" in normalised


def _apply_client_overrides(flags, client_name):
    """
    Apply client-level overrides to YES/NO flags.
    Overrides are additive — they can only force YES, never force NO.
    """
    if not client_name:
        return flags
    client_lower = client_name.lower()
    for substring, overrides in CLIENT_OVERRIDES.items():
        if substring in client_lower:
            for field, value in overrides.items():
                flags[field] = value
    return flags


def _best_product_title(full_text):
    """
    Find the most complete product title line.
    Prefers /// lines but falls back to // lines if the /// line is truncated
    (i.e. doesn't contain a TC number when one exists elsewhere in the text).
    """
    import re as _re
    triple = _re.search(r'[^\n]*///[^\n]+', full_text)
    # Find // lines, explicitly skipping any that are actually /// lines
    double = None
    for line in full_text.split('\n'):
        if '//' in line and '///' not in line and _re.search(r'\bTC\d{3,}\b', line):
            double = line.strip()
            break

    # If /// line contains a TC number, use it
    if triple and _re.search(r'\bTC\d{3,}\b', triple.group(0)):
        return triple.group(0).strip()
    # Fall back to // line if it has TC number (/// line may be truncated)
    if double and _re.search(r'\bTC\d{3,}\b', double):
        return double
    # Default to /// line
    return triple.group(0).strip() if triple else ''


def _get_tc_number(full_text):
    """
    Extract the TC style number (e.g. TC2320) from the best product title line,
    falling back to DESIGN NAME lines if not found in the title.
    """
    import re as _re
    title = _best_product_title(full_text)
    tc = _re.search(r'\bTC(\d{3,})\b', title)
    if tc:
        return f"TC{tc.group(1)}"
    # Final fallback: first TC#### anywhere in the text
    for line in full_text.split('\n'):
        tc = _re.search(r'\bTC(\d{3,})\b', line)
        if tc:
            return f"TC{tc.group(1)}"
    return None


def _get_artwork_name(full_text):
    """
    Extract the artwork/product name from the best product title line.
    Title structure: CLIENT_PO///STYLE DESC - COLOUR - TC#### - ARTWORK NAME - decoration
    Artwork name is the segment immediately after TC#### and before the decoration suffix.
    """
    import re as _re
    title = _best_product_title(full_text)
    if not title:
        return None

    tc_match = _re.search(r'\bTC\d{3,}\b', title)
    if not tc_match:
        return None

    after_tc = title[tc_match.end():]
    after_tc = _re.sub(r'^\s*[-\s]+', '', after_tc)

    decoration_kw = (r'Left Chest|Right Chest|Front Chest|Back Print|Full Back|'
                     r'Front Center|Left Sleeve|Right Sleeve|Hat|Replen|Replenishment')
    loc_match = _re.search(r'\s*-\s*(?:' + decoration_kw + r')', after_tc, _re.IGNORECASE)
    if loc_match:
        return after_tc[:loc_match.start()].strip() or None
    dash_match = _re.search(r'\s*-\s*', after_tc)
    return after_tc[:dash_match.start()].strip() if dash_match else after_tc.strip() or None


def _primary_design_name(full_text):
    """
    Extract the first non-neck-tag, non-mockup DESIGN NAME from artwork detail blocks.
    Used for the ARTWORK field on non-Troll Co orders.
    """
    skip_locs = LOCATIONS_SEPARATE | {"mockup for production", "mockup"}
    blocks = re.split(r"(?=DESIGN NAME)", full_text)
    for block in blocks:
        name_m = re.search(r"DESIGN NAME\s+(.+)", block)
        if not name_m:
            continue
        name = name_m.group(1).strip()
        if "mockup" in name.lower():
            continue
        loc_m = re.search(r"DESIGN LOCATION\s+(.+)", block)
        if loc_m:
            loc = loc_m.group(1).strip().lower()
            if any(excl in loc for excl in skip_locs):
                continue
        return name
    return None


def _is_repeat_order(full_text):
    """
    Returns True if any artwork description contains the word 'repeat'.
    Catches: 'Repeat of PO', 'Repeat PO', 'RN' (repeat neck), 'replenishment', etc.
    """
    import re as _re
    return bool(_re.search(r'\brepeat\b', full_text, _re.IGNORECASE))


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
    flags = _apply_client_overrides(flags, order.get("client"))

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

        # Troll Co order flag
        "Troll Co Order?":        "YES" if _is_troll_co(order.get("client", "")) else "NO",

        # Troll Co specific fields — only populated for Troll Co orders
        "Troll Co Style #":       _get_tc_number(full_text) if _is_troll_co(order.get("client", "")) else None,
        "ARTWORK":                _get_artwork_name(full_text) if _is_troll_co(order.get("client", "")) else _primary_design_name(full_text),

        # Repeat order detection
        "REPEAT ORDER?":          "REPEAT ORDER" if _is_repeat_order(full_text) else "NEW ORDER",

        # PO value from subtotal
        "PO VALUE":               order.get("subtotal"),

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