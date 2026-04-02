"""
parse_commonsku_so.py
Parses a CommonSku Sales Order PDF and outputs structured JSON.

Usage:
    python3 parse_commonsku_so.py SALES_ORDER-70443.pdf --pretty
    python3 parse_commonsku_so.py SALES_ORDER-70443.pdf --csv
    python3 parse_commonsku_so.py SALES_ORDER-70443.pdf > order.json
"""

import sys, re, json, csv, io
import pdfplumber


def _is_description_line(line):
    """True if line is garment spec copy rather than a product name."""
    # Weight/material specs are always descriptions
    if re.search(r"\b(oz\.|gsm|cotton|polyester|made in|neckline|fit\.|structure\.|finish\.|rib height|ounces|grams|ringspun)\b",
                 line, re.IGNORECASE):
        return True
    # Long mixed-case lines are descriptions UNLESS they look like product titles
    # (start with a style code pattern or contain //)
    if len(line) > 80 and sum(1 for c in line if c.islower()) > len(line) * 0.4:
        if re.match(r"^[A-Z0-9]", line) and ("//" in line or re.match(r"^\d+\s+\w", line)):
            return False  # looks like a product title
        return True
    return False

_SKIP_HEADER = re.compile(
    r"^(TERMS|SHIPPING|BILLING|PROJECT|CUSTOMER|CURRENCY|Page \d|"
    r"North York|Canada|St Regis|Fulfillment|The Avid|SALES ORDER|AVID APPAREL|Click to enlarge|"
    r"Archway|Guelph|Mississauga|Ontario|Ashley|James|Grant|Grace|Halifax|Sackville|Hobsons|"
    r"Unit B|\d{3,5} |Doors \d)",
    re.IGNORECASE
)


def extract_text(pdf_path):
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"
    return full_text


def parse_order_header(text):
    order = {}

    m = re.search(r"SALES ORDER for (.+)", text)
    order["client"] = m.group(1).strip() if m else None

    m = re.search(r"(\d{5})\s+(\d{5})\s+(\w+ \d+, \d{4})", text)
    if m:
        order["project_number"] = m.group(1)
        order["so_number"] = m.group(2)
        order["in_hands_date"] = m.group(3)
    else:
        order["project_number"] = order["so_number"] = order["in_hands_date"] = None

    m = re.search(r"\b(Net \d+|50% Deposit[^$\n\d]*|Due on Receipt|Prepaid|Credit Card)\b",
                  text, re.IGNORECASE)
    raw_terms = m.group(1).strip() if m else None
    order["terms"] = re.sub(r"\s+(USD|CAD)\s*$", "", raw_terms).strip() if raw_terms else None

    # Customer PO: everything between terms token and USD/CAD on that line, if it contains digits
    m = re.search(r"(?:Net \d+|50%[^$\n\d]+|Due[^$\n\d]+)\s+(.+?)\s+(?:USD|CAD)\s*$",
                  text, re.MULTILINE)
    if m:
        candidate = m.group(1).strip()
        # Strip known filler words that can bleed from terms (e.g. "Days" from "Net 7 Days")
        candidate = re.sub(r"^(?:Days|Due|Before|Shipment|Receipt)\s+", "", candidate, flags=re.IGNORECASE).strip()
        order["customer_po"] = candidate if re.search(r"\d", candidate) else None
    else:
        order["customer_po"] = None

    m = re.search(r"\b(USD|CAD)\b", text)
    order["currency"] = m.group(1) if m else None

    # Production notes: instruction block before "Questions about this sales order?"
    # Try all-caps first, fall back to any mixed-case sentence that isn't a garment description
    m = re.search(r"\n([A-Z][A-Z0-9 ,.\-!]+[!.])\s*\n(?:Questions|SUBTOTAL)", text)
    if not m:
        m = re.search(r"\n([A-Z][^\n]{10,}[!.])\s*\n(?:Questions|SUBTOTAL)", text)
        if m and re.search(
            r"\b(oz\.|gsm|cotton|polyester|ring.spun|seamed|tape|shrinkage)\b",
            m.group(1), re.IGNORECASE
        ):
            m = None
    order["production_notes"] = m.group(1).strip() if m else None

    for pattern, key in [
        (r"SUBTOTAL\s+\$([\d,]+\.\d{2})", "subtotal"),
        (r"TOTAL\s+(?:USD|CAD)\s+\$([\d,]+\.\d{2})", "total"),
        (r"LESS DEPOSIT:\s+\(\$([\d,]+\.\d{2})\)", "deposit"),
        (r"FINAL BALANCE:\s+\$([\d,]+\.\d{2})", "balance"),
    ]:
        m = re.search(pattern, text)
        order[key] = m.group(1) if m else None

    m = re.search(r"Questions about this sales order\?\s+([\w]+ [\w]+)\s+([\w.]+@[\w.]+)", text)
    if m:
        order["account_rep_name"] = m.group(1)
        order["account_rep_email"] = m.group(2)
    else:
        order["account_rep_name"] = order["account_rep_email"] = None

    m = re.search(r"last approved by (.+?) on (\w+ \d+, \d{4})", text)
    if m:
        order["approved_by"] = m.group(1)
        order["approval_date"] = m.group(2)
    else:
        order["approved_by"] = order["approval_date"] = None

    return order


def _extract_product_name(candidate_lines):
    """
    Given the candidate lines before an ITEM block, extract style_code and product_name.
    Handles:
      - Pure style-code headers:  AVL1605 Heavy Pigment Tee...
      - PO-prefixed titles:       CAD PO #518989-3//AVN6210 Ultimate Unisex...  (may wrap)
      - Freeform names:           Custom Edge T-Shirt
    """
    if not candidate_lines:
        return None, None

    # Look for a line containing an embedded style code (even with PO prefix)
    for i, line in enumerate(reversed(candidate_lines)):
        style_m = re.search(r"\b([A-Z]{0,4}\d{3,5})\b", line)
        if style_m and re.search(r"[A-Z]", style_m.group(1)):  # must have at least one letter
            style_code = style_m.group(1)
            after = line[style_m.end():].strip().lstrip('-').strip()
            # Check if the very next candidate line is a title continuation (not description)
            idx = len(candidate_lines) - 1 - i
            if idx + 1 < len(candidate_lines):
                nxt = candidate_lines[idx + 1]
                if not _is_description_line(nxt) and not _SKIP_HEADER.match(nxt):
                    after = after + " " + nxt
            return style_code, after.strip()

    # No style code found — use last line as freeform name
    return None, candidate_lines[-1]


def parse_products(text):
    products = []

    for item_match in re.finditer(r"ITEM QTY PRICE AMOUNT", text):
        block_start = item_match.start()

        before = text[max(0, block_start - 700):block_start]
        candidate_lines = [
            l.strip() for l in before.split("\n")
            if l.strip()
            and not _SKIP_HEADER.match(l.strip())
            and not _is_description_line(l.strip())
        ]

        style_code, product_name = _extract_product_name(candidate_lines)

        product = {
            "style_code": style_code,
            "product_name": product_name,
        }

        # Bound block end
        next_item = re.search(r"ITEM QTY PRICE AMOUNT", text[block_start + 1:])
        services_pos = text.find("SERVICE QTY PRICE AMOUNT", block_start)
        tc_pos = text.find("TERMS AND CONDITIONS", block_start)

        end_candidates = [len(text)]
        if next_item:
            end_candidates.append(block_start + 1 + next_item.start())
        if services_pos > block_start:
            end_candidates.append(services_pos)
        if tc_pos > block_start:
            end_candidates.append(tc_pos)

        block = text[block_start:min(end_candidates)]

        color_m = re.search(r"Size: \S+ - Color: (.+?)\s+\d", block)
        product["color"] = color_m.group(1).strip() if color_m else None

        size_rows = re.findall(
            r"Size: (\S+) - Color: \S.*?\s+(\d+)\s+\$([\d.]+)\s+\$([\d.]+)", block
        )
        product["sizes"] = [
            {"size": r[0], "qty": int(r[1]), "unit_price": float(r[2]), "amount": float(r[3])}
            for r in size_rows
        ]

        m = re.search(r"TOTAL UNITS\s+(\d+)", block)
        product["total_units"] = int(m.group(1)) if m else None

        m = re.search(r"\nTOTAL\s+\$([\d,]+\.\d{2})", block)
        product["total"] = m.group(1) if m else None

        locations = re.findall(r"DESIGN LOCATION\s+(.+)", block)
        product["decoration_locations"] = list(dict.fromkeys(
            [l for l in locations if "Mockup" not in l]
        ))

        imprint_types = re.findall(r"IMPRINT TYPE\s+(.+)", block)
        product["imprint_types"] = list(dict.fromkeys(
            [t for t in imprint_types if t != "Other"]
        ))

        products.append(product)

    return products


def parse_services(text):
    services = []

    m = re.search(
        r"SERVICE QTY PRICE AMOUNT\n(.*?)(?=\nPage \d+ of|\nSUBTOTAL|\nQuestions|\nMUST BE)",
        text, re.DOTALL
    )
    if not m:
        return services

    pending_name = None
    for line in m.group(1).strip().split("\n"):
        qp = re.search(r"(\d+)\s+\$([\d.]+)\s+\$([\d.]+)", line)
        if qp:
            inline_name = line[:qp.start()].strip()
            if pending_name and (not inline_name or len(inline_name) > 40):
                name = pending_name
            elif inline_name:
                name = inline_name
            elif pending_name:
                name = pending_name
            else:
                name = "Unknown Service"
            services.append({
                "service": name,
                "qty": int(qp.group(1)),
                "unit_price": float(qp.group(2)),
                "amount": float(qp.group(3))
            })
            pending_name = None
        else:
            clean = line.strip()
            if clean and not re.match(r"^(overview|Page \d|Troll Co\.|Barcoding|Barcode)", clean, re.IGNORECASE):
                if pending_name is None:
                    pending_name = clean

    return services


def parse(pdf_path):
    text = extract_text(pdf_path)
    order = parse_order_header(text)
    order["products"] = parse_products(text)
    order["services"] = parse_services(text)
    return order


def to_csv(order):
    output = io.StringIO()
    fieldnames = [
        "so_number", "project_number", "client", "customer_po", "in_hands_date",
        "terms", "currency", "account_rep_name", "account_rep_email",
        "approved_by", "approval_date", "production_notes",
        "subtotal", "total", "deposit", "balance",
        "style_code", "product_name", "color", "total_units", "product_total",
        "decoration_locations", "imprint_types", "sizes_detail"
    ]
    writer = csv.DictWriter(output, fieldnames=fieldnames)
    writer.writeheader()
    for product in order["products"]:
        sizes_detail = "; ".join(
            f"{s['size']}x{s['qty']}@${s['unit_price']}" for s in product["sizes"]
        )
        writer.writerow({
            "so_number": order.get("so_number"),
            "project_number": order.get("project_number"),
            "client": order.get("client"),
            "customer_po": order.get("customer_po"),
            "in_hands_date": order.get("in_hands_date"),
            "terms": order.get("terms"),
            "currency": order.get("currency"),
            "account_rep_name": order.get("account_rep_name"),
            "account_rep_email": order.get("account_rep_email"),
            "approved_by": order.get("approved_by"),
            "approval_date": order.get("approval_date"),
            "production_notes": order.get("production_notes"),
            "subtotal": order.get("subtotal"),
            "total": order.get("total"),
            "deposit": order.get("deposit"),
            "balance": order.get("balance"),
            "style_code": product.get("style_code"),
            "product_name": product.get("product_name"),
            "color": product.get("color"),
            "total_units": product.get("total_units"),
            "product_total": product.get("total"),
            "decoration_locations": ", ".join(product.get("decoration_locations", [])),
            "imprint_types": ", ".join(product.get("imprint_types", [])),
            "sizes_detail": sizes_detail,
        })
    return output.getvalue()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 parse_commonsku_so.py <pdf_path> [--pretty] [--csv]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    flags = sys.argv[2:]
    order = parse(pdf_path)
    if "--csv" in flags:
        print(to_csv(order))
    elif "--pretty" in flags:
        print(json.dumps(order, indent=2))
    else:
        print(json.dumps(order))