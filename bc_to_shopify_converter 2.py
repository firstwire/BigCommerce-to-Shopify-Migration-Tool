"""
BC -> Shopify CSV Converter  v13.0
===================================
Converts BigCommerce export CSV files into Shopify-ready import CSVs.

  Products  -> shopify_products.csv   (Shopify Admin -> Products -> Import)
  Customers -> shopify_customers.csv  (Shopify Admin -> Customers -> Import)
  Orders    -> shopify_orders.csv     (Matrixify app -> Import)

REQUIREMENTS:
  pip install pandas openpyxl

ORDERS NOTE:
  Shopify's native CSV importer does NOT support order import.
  Orders must be imported using the free Matrixify app:
    Shopify Admin -> Apps -> Matrixify -> Import -> shopify_orders.csv

USAGE:
  python bc_to_shopify_converter.py
"""

import os
import re
import csv
import sys
import pandas as pd
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR  = os.path.join(SCRIPT_DIR, "shopify_output")

# ─────────────────────────────────────────────────────────────
# Official Shopify product CSV headers
# Source: https://help.shopify.com/en/manual/products/import-export/using-csv
# ─────────────────────────────────────────────────────────────
PRODUCT_HEADERS = [
    "Handle", "Title", "Body (HTML)", "Vendor", "Type", "Tags", "Published",
    "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value",
    "Option3 Name", "Option3 Value",
    "Variant SKU", "Variant Grams", "Variant Inventory Tracker",
    "Variant Inventory Qty", "Variant Inventory Policy",
    "Variant Fulfillment Service", "Variant Price", "Variant Compare At Price",
    "Variant Requires Shipping", "Variant Taxable", "Variant Barcode",
    "Image Src", "Image Alt Text", "SEO Title", "SEO Description", "Status",
]  # 29 columns

# ─────────────────────────────────────────────────────────────
# Official Shopify customer CSV headers
# Source: https://help.shopify.com/en/manual/customers/import-export-customers
# ─────────────────────────────────────────────────────────────
CUSTOMER_HEADERS = [
    "First Name", "Last Name", "Email", "Accepts Email Marketing",
    "Default Address Company", "Default Address Address1",
    "Default Address Address2", "Default Address City",
    "Default Address Province Code", "Default Address Country Code",
    "Default Address Zip", "Default Address Phone", "Phone",
    "Accepts SMS Marketing", "Note", "Tax Exempt", "Tags",
]  # 17 columns

# ─────────────────────────────────────────────────────────────
# Matrixify order CSV headers (official Matrixify format)
# Source: https://matrixify.app/documentation/orders/
# Standard Shopify does NOT support order CSV import — use Matrixify app.
# Each order has multiple row types:
#   Row 1:  Line: Type = "Line Item"    (product line)
#   Row 2:  Line: Type = "Shipping Line" (shipping cost)
#   Row 3:  Line: Type = "Transaction"   (payment record)
# ─────────────────────────────────────────────────────────────
ORDER_HEADERS = [
    "Name",               # Order number e.g. #1001
    "Email",              # Customer email
    "Phone",              # Customer phone
    "Currency",           # ISO currency code e.g. USD
    "Payment: Status",    # paid / pending / refunded / voided
    "Processed At",       # Order date  YYYY-MM-DD HH:MM:SS +0000
    "Send Receipt",       # FALSE (don't email customer on import)
    "Inventory Behaviour",# bypass (don't adjust stock on import)
    "Note",               # Order notes
    "Tags",               # Order tags
    # ── Billing address ───────────────────────────────────────
    "Billing: First Name",
    "Billing: Last Name",
    "Billing: Company",
    "Billing: Phone",
    "Billing: Address 1",
    "Billing: Address 2",
    "Billing: City",
    "Billing: Province",       # Full name e.g. "New York"
    "Billing: Province Code",  # 2-letter e.g. NY
    "Billing: Country",        # Full name e.g. "United States"
    "Billing: Country Code",   # 2-letter ISO e.g. US
    "Billing: Zip",
    # ── Shipping address ──────────────────────────────────────
    "Shipping: First Name",
    "Shipping: Last Name",
    "Shipping: Company",
    "Shipping: Phone",
    "Shipping: Address 1",
    "Shipping: Address 2",
    "Shipping: City",
    "Shipping: Province",
    "Shipping: Province Code",
    "Shipping: Country",
    "Shipping: Country Code",
    "Shipping: Zip",
    # ── Line item columns (filled per row type) ───────────────
    "Line: Type",         # Line Item / Shipping Line / Transaction
    "Line: Title",        # Product name  OR  shipping method name
    "Line: Quantity",     # Units ordered (blank for non-line-item rows)
    "Line: Price",        # Unit price    (blank for non-line-item rows)
    "Line: SKU",          # Product SKU   (blank for non-line-item rows)
    "Line: Grams",        # Weight grams  (blank for non-line-item rows)
    "Line: Requires Shipping",  # TRUE / FALSE
    "Line: Taxable",      # TRUE / FALSE
    "Line: Discount",     # Discount amount (negative)
    # ── Tax ───────────────────────────────────────────────────
    "Tax 1: Title",       # Tax name e.g. "Tax"
    "Tax 1: Price",       # Tax amount
    "Tax 1: Rate",        # Tax rate  e.g. 0.08 = 8%
    # ── Transaction (payment record) ──────────────────────────
    "Transaction: Amount",   # Amount paid
    "Transaction: Currency", # Currency
    "Transaction: Kind",     # "sale" for normal payment
    "Transaction: Status",   # "success"
    "Transaction: Gateway",  # Payment gateway name
]  # 46 columns


# ═══════════════════════════════════════════════════════════════
# UTILITIES
# ═══════════════════════════════════════════════════════════════

def log(msg, level="INFO"):
    tag = {"INFO": "i ", "OK": "OK", "WARN": "! ", "ERROR": "X ", "STEP": ">>"}
    print("  [{}]  {}".format(tag.get(level, "  "), msg))


def ensure_output_dir():
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        return True
    except Exception as e:
        log("Cannot create output folder: {}".format(e), "ERROR")
        return False


def write_csv(rows, headers, filename):
    if not ensure_output_dir():
        return
    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, quoting=csv.QUOTE_ALL).writerows([headers] + rows)
    log("Saved -> {}  ({:,} bytes)".format(path, os.path.getsize(path)), "OK")


def read_file(path):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".csv":
            try:    return pd.read_csv(path, encoding="utf-8")
            except: return pd.read_csv(path, encoding="latin-1")
        elif ext in (".xlsx", ".xlsm"):
            return pd.read_excel(path, engine="openpyxl")
        elif ext == ".xls":
            return pd.read_excel(path, engine="xlrd")
        else:
            log("Unsupported file type: {}".format(ext), "ERROR")
            return None
    except Exception as e:
        log("Cannot read {}: {}".format(os.path.basename(path), e), "ERROR")
        return None


def sv(val, default=""):
    """Safe string — returns default for NaN / None / blank."""
    try:
        if val is None: return default
        if isinstance(val, float) and pd.isna(val): return default
        s = str(val).strip()
        return s if s.lower() not in ("nan", "none", "") else default
    except Exception:
        return default


def sfloat(val, default=""):
    """Safe 2-decimal float string. Empty for zero / missing."""
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)): return default
        f = float(val)
        return "{:.2f}".format(f) if f > 0 else default
    except Exception:
        return default


def sint(val, default=0):
    """Safe integer."""
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)): return default
        return int(float(val))
    except Exception:
        return default


def find_col(df, candidates):
    """Return first matching column name from a list of candidates."""
    for c in candidates:
        if c in df.columns: return c
    return None


def gv(row, col, default=""):
    """Get row value safely — handles None column."""
    if col is None: return default
    return sv(row.get(col, default), default)


def gn(row, col, default=""):
    """Get numeric row value safely."""
    if col is None: return default
    return sfloat(row.get(col, None), default)


# ═══════════════════════════════════════════════════════════════
# PRODUCT HELPERS
# ═══════════════════════════════════════════════════════════════

def make_handle(url_raw, title):
    url = sv(url_raw)
    if url and url not in ("/", ""):
        handle = url.strip("/").strip()
        handle = re.sub(r"\.html?$", "", handle, flags=re.IGNORECASE)
        return handle
    cleaned = re.sub(r"\[.*?\]", "", title)
    return re.sub(r"[^a-z0-9]+", "-", cleaned.lower()).strip("-")


def parse_images(raw):
    if not raw or sv(raw) == "": return []
    results = []
    for block in str(raw).split("|"):
        m = re.search(r"Product Image URL:\s*(\S+)", block)
        if not m: continue
        url = m.group(1).strip()
        url = re.sub(r"/%%[^%]+%%/", "/", url)
        url = re.sub(r"%%[^%]+%%",   "",  url)
        if url.startswith("http://"):
            url = "https://" + url[7:]
        if url.startswith("https://"):
            results.append(url)
    return results


def parse_categories(raw):
    if not raw or sv(raw) == "": return ""
    names = re.findall(r"Category Name:\s*([^,|]+)", str(raw))
    return ", ".join(list(dict.fromkeys([n.strip() for n in names if n.strip()])))


def parse_custom_fields(raw):
    if not raw or sv(raw) == "": return {}
    result = {}
    parts = re.split(r';(?=(?:[^"]*"[^"]*")*[^"]*$)', str(raw))
    for part in parts:
        part = part.strip().strip('"')
        if "=" not in part: continue
        key, _, val = part.partition("=")
        key = key.strip().strip('"')
        val = val.strip().strip('"').strip()
        if key and val and val.strip():
            result[key] = val
    return result


def pick_price_and_compare(row):
    try:
        calc   = float(sv(row.get("Calculated Price"), "0") or "0")
        retail = float(sv(row.get("Retail Price"),     "0") or "0")
    except Exception:
        calc = retail = 0.0
    price   = calc if calc > 0 else 0.0
    compare = retail if retail > 0 and retail > price else 0.0
    return (
        "{:.2f}".format(price)   if price   > 0 else "",
        "{:.2f}".format(compare) if compare > 0 else "",
    )


def lbs_to_grams(val):
    try:   return int(round(float(val or 0) * 453.592))
    except: return 0


def _blank_product_row(handle):
    row = [""] * len(PRODUCT_HEADERS)
    row[0] = handle
    return row


# ═══════════════════════════════════════════════════════════════
# PRODUCT CONVERTER
# ═══════════════════════════════════════════════════════════════

def convert_products(input_path):
    log("Reading: {}".format(os.path.basename(input_path)), "STEP")
    df = read_file(input_path)
    if df is None: return

    log("Loaded {} products | {} columns".format(len(df), len(df.columns)))
    print()

    all_rows  = []
    converted = 0
    skipped   = 0

    for _, row in df.iterrows():
        title = sv(row.get("Name", ""))
        if not title: skipped += 1; continue

        handle    = make_handle(sv(row.get("Product URL", "")), title)
        body      = sv(row.get("Description", ""))
        vendor    = sv(row.get("Brand", ""))
        tags      = parse_categories(sv(row.get("Category Details", "")))
        cf        = parse_custom_fields(row.get("Product Custom Fields", ""))
        prod_type = sv(cf.get("type", ""))
        barcode   = sv(cf.get("mfg-part-number", ""))
        seo_title = sv(row.get("Page Title", ""))
        seo_desc  = sv(row.get("META Description", ""))

        vis       = sv(row.get("Product Visible", "Y"))
        published = "TRUE"   if vis.upper() in ("Y","YES","1","TRUE") else "FALSE"
        status    = "active" if published == "TRUE" else "draft"

        price, compare = pick_price_and_compare(row)
        weight_g  = lbs_to_grams(row.get("Weight"))
        qty       = sint(row.get("Stock Level", 0))

        inventoried = sv(row.get("Product Inventoried", "Y"))
        inv_tracker = "shopify" if inventoried.upper() in ("Y","YES","1") else ""

        allow_purch = sv(row.get("Allow Purchases", "Y"))
        inv_policy  = "deny" if allow_purch.upper() in ("Y","YES","1") else "continue"

        images    = parse_images(sv(row.get("Images", "")))
        first_img = images[0] if images else ""

        # Main product row — 29 values
        all_rows.append([
            handle, title, body, vendor, prod_type, tags, published,
            "Title", "Default Title", "", "", "", "",
            sv(row.get("Code", "")), weight_g, inv_tracker, qty, inv_policy,
            "manual", price, compare, "TRUE", "TRUE", barcode,
            first_img, title if first_img else "",
            seo_title, seo_desc, status,
        ])

        # Extra image rows — same Handle, all other cols blank
        for img_url in images[1:]:
            img_row = _blank_product_row(handle)
            img_row[PRODUCT_HEADERS.index("Image Src")]      = img_url
            img_row[PRODUCT_HEADERS.index("Image Alt Text")] = title
            all_rows.append(img_row)

        converted += 1

    bad = [(i+2, len(r)) for i, r in enumerate(all_rows) if len(r) != len(PRODUCT_HEADERS)]
    if bad:
        log("BUG: {} rows have wrong column count!".format(len(bad)), "ERROR")
        return

    write_csv(all_rows, PRODUCT_HEADERS, "shopify_products.csv")
    print()
    log("Products converted : {}".format(converted), "OK")
    log("Products skipped   : {}".format(skipped), "OK" if skipped == 0 else "WARN")
    log("Total rows (inc. extra image rows): {}".format(len(all_rows)), "OK")
    print()
    log("Import -> Shopify Admin -> Products -> Import -> shopify_products.csv", "INFO")
    print()


# ═══════════════════════════════════════════════════════════════
# CUSTOMER CONVERTER
# ═══════════════════════════════════════════════════════════════

def convert_customers(input_path):
    log("Reading: {}".format(os.path.basename(input_path)), "STEP")
    df = read_file(input_path)
    if df is None: return

    log("Loaded {} rows | {} columns".format(len(df), len(df.columns)))

    df.columns = [
        str(c).strip().lower().replace(" ","_").replace("-","_").replace("/","_")
        for c in df.columns
    ]

    c_first   = find_col(df, ["first_name","firstname","fname","given_name","customer_first_name"])
    c_last    = find_col(df, ["last_name","lastname","lname","surname","customer_last_name"])
    c_email   = find_col(df, ["email","email_address","customer_email","e_mail"])
    c_phone   = find_col(df, ["phone","phone_number","mobile","telephone","cell"])
    c_company = find_col(df, ["company","company_name","business_name","organisation"])
    c_addr1   = find_col(df, ["address1","street_address","address","street","street_1",
                               "addr1","billing_address1"])
    c_addr2   = find_col(df, ["address2","street_2","addr2","address_line_2","suite","apt"])
    c_city    = find_col(df, ["city","town","billing_city","suburb"])
    c_prov    = find_col(df, ["state","province","state_province","billing_state",
                               "region","province_code","state_code"])
    c_zip     = find_col(df, ["zip","postal_code","postcode","billing_zip","pincode"])
    c_country = find_col(df, ["country_code","country_iso2","billing_country_iso2",
                               "country","billing_country"])
    c_accepts = find_col(df, ["accepts_marketing","newsletter","subscribed",
                               "email_opt_in","accepts_email_marketing"])
    c_tags    = find_col(df, ["tags","customer_tags","customer_group","group"])
    c_note    = find_col(df, ["note","notes","customer_note","comments","remarks"])

    COUNTRY_NAMES = {
        "united states":"US","usa":"US","united kingdom":"GB","uk":"GB",
        "canada":"CA","australia":"AU","india":"IN","germany":"DE",
        "france":"FR","italy":"IT","spain":"ES","netherlands":"NL",
        "new zealand":"NZ","singapore":"SG","ireland":"IE",
    }

    all_rows = []
    skipped  = 0

    for _, row in df.iterrows():
        email = gv(row, c_email)
        if not email or "@" not in email: skipped += 1; continue

        prov = gv(row, c_prov, "").upper().strip()
        if len(prov) > 4: prov = ""

        cc = gv(row, c_country, "US").strip()
        cc = COUNTRY_NAMES.get(cc.lower(), cc[:2].upper()) if len(cc) > 2 else cc.upper()

        accepts = "yes" if gv(row, c_accepts, "0").lower() in ("1","true","yes","y") else "no"

        phone = gv(row, c_phone, "")
        if phone:
            d = re.sub(r"[^\d+]", "", phone)
            if d and not d.startswith("+"):
                d = "+1" + d if len(d) == 10 else "+" + d
            phone = d

        all_rows.append([
            gv(row, c_first), gv(row, c_last), email, accepts,
            gv(row, c_company), gv(row, c_addr1), gv(row, c_addr2),
            gv(row, c_city), prov, cc,
            gv(row, c_zip), phone, phone,
            "no", gv(row, c_note), "no", gv(row, c_tags),
        ])

    write_csv(all_rows, CUSTOMER_HEADERS, "shopify_customers.csv")
    print()
    log("Customers converted: {}".format(len(all_rows)), "OK")
    log("Customers skipped  : {} (no valid email)".format(skipped),
        "OK" if skipped == 0 else "WARN")
    print()
    log("Import -> Shopify Admin -> Customers -> Import -> shopify_customers.csv", "INFO")
    print()


# ═══════════════════════════════════════════════════════════════
# ORDER CONVERTER  —  Matrixify format
# ═══════════════════════════════════════════════════════════════

# BC payment status -> Shopify / Matrixify payment status
PAYMENT_STATUS_MAP = {
    "completed":          "paid",
    "shipped":            "paid",
    "partially_shipped":  "paid",
    "paid":               "paid",
    "awaiting_fulfillment": "paid",
    "awaiting_shipment":  "paid",
    "pending":            "pending",
    "awaiting_payment":   "pending",
    "manual_verification_required": "pending",
    "incomplete":         "pending",
    "refunded":           "refunded",
    "partially_refunded": "partially_refunded",
    "cancelled":          "voided",
    "canceled":           "voided",
    "declined":           "voided",
    "voided":             "voided",
}


def format_order_date(raw):
    """Parse any date string and return Matrixify format: YYYY-MM-DD HH:MM:SS +0000"""
    try:
        return pd.to_datetime(raw, dayfirst=False).strftime("%Y-%m-%d %H:%M:%S +0000")
    except Exception:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S +0000")


def clean_phone(phone):
    """Return E.164-style phone number or empty string."""
    if not phone: return ""
    d = re.sub(r"[^\d+]", "", str(phone))
    if d and not d.startswith("+"):
        d = "+1" + d if len(d) == 10 else "+" + d
    return d


def _order_base(order_name, email, phone, currency, pay_status,
                processed_at, note, tags,
                b_fname, b_lname, b_company, b_phone,
                b_addr1, b_addr2, b_city, b_province, b_prov_code,
                b_country, b_country_code, b_zip,
                s_fname, s_lname, s_company, s_phone,
                s_addr1, s_addr2, s_city, s_province, s_prov_code,
                s_country, s_country_code, s_zip):
    """Return the 46-value base portion of an order row (all line cols blank)."""
    return [
        order_name, email, phone, currency, pay_status, processed_at,
        "FALSE",   # Send Receipt — don't email customer on import
        "bypass",  # Inventory Behaviour — don't deduct stock on import
        note, tags,
        # Billing
        b_fname, b_lname, b_company, b_phone,
        b_addr1, b_addr2, b_city, b_province, b_prov_code,
        b_country, b_country_code, b_zip,
        # Shipping
        s_fname, s_lname, s_company, s_phone,
        s_addr1, s_addr2, s_city, s_province, s_prov_code,
        s_country, s_country_code, s_zip,
        # Line cols — all blank for base; filled per row type below
        "", "", "", "", "", "", "", "", "",
        # Tax — all blank for base
        "", "", "",
        # Transaction — all blank for base
        "", "", "", "", "",
    ]   # 46 values


def convert_orders(input_path):
    log("Reading: {}".format(os.path.basename(input_path)), "STEP")
    df = read_file(input_path)
    if df is None: return

    log("Loaded {} rows | {} columns".format(len(df), len(df.columns)))

    # Normalize column names
    df.columns = [
        str(c).strip().lower()
               .replace(" ","_").replace("-","_")
               .replace("(","").replace(")","").replace("/","_")
        for c in df.columns
    ]

    # ── Detect columns ────────────────────────────────────────
    c_order_id   = find_col(df, ["order_id","id","order_number","bc_order_id","orderId"])
    c_email      = find_col(df, ["billing_email","email","customer_email",
                                 "billing_address_email","customer_e_mail"])
    c_phone      = find_col(df, ["billing_phone","phone","customer_phone"])
    c_date       = find_col(df, ["date_created","created_at","order_date","date",
                                 "placed_at","order_placed"])
    c_status     = find_col(df, ["status","order_status","payment_status",
                                 "financial_status","payment_state"])
    c_currency   = find_col(df, ["currency_code","currency","order_currency"])
    c_note       = find_col(df, ["order_notes","notes","note","customer_note",
                                 "staff_notes","comments"])

    # Billing address
    c_b_fname    = find_col(df, ["billing_first_name","first_name","customer_first_name"])
    c_b_lname    = find_col(df, ["billing_last_name","last_name","customer_last_name"])
    c_b_company  = find_col(df, ["billing_company","company","company_name"])
    c_b_phone    = find_col(df, ["billing_phone","phone"])
    c_b_addr1    = find_col(df, ["billing_address_line_1","billing_address1",
                                 "billing_street_1","billing_address"])
    c_b_addr2    = find_col(df, ["billing_address_line_2","billing_address2",
                                 "billing_street_2"])
    c_b_city     = find_col(df, ["billing_city","city"])
    c_b_state    = find_col(df, ["billing_state","billing_province","state","province"])
    c_b_zip      = find_col(df, ["billing_zip","billing_postal_code","zip","postal_code"])
    c_b_country  = find_col(df, ["billing_country","billing_country_iso2","country"])

    # Shipping address
    c_s_fname    = find_col(df, ["shipping_first_name","ship_first_name"])
    c_s_lname    = find_col(df, ["shipping_last_name","ship_last_name"])
    c_s_company  = find_col(df, ["shipping_company","ship_company"])
    c_s_phone    = find_col(df, ["shipping_phone","ship_phone"])
    c_s_addr1    = find_col(df, ["shipping_address_line_1","shipping_address1",
                                 "shipping_street_1","ship_address"])
    c_s_addr2    = find_col(df, ["shipping_address_line_2","shipping_address2",
                                 "shipping_street_2"])
    c_s_city     = find_col(df, ["shipping_city","ship_city"])
    c_s_state    = find_col(df, ["shipping_state","shipping_province","ship_state"])
    c_s_zip      = find_col(df, ["shipping_zip","shipping_postal_code","ship_zip"])
    c_s_country  = find_col(df, ["shipping_country","shipping_country_iso2","ship_country"])

    # Line item
    c_product    = find_col(df, ["product_name","item_name","name","line_item_name",
                                 "product"])
    c_sku        = find_col(df, ["product_sku","sku","item_sku","variant_sku"])
    c_qty        = find_col(df, ["quantity","qty","item_quantity","ordered_qty"])
    c_price      = find_col(df, ["base_price","product_price","price","unit_price",
                                 "item_price","line_price"])
    c_discount   = find_col(df, ["discount_amount","discount","coupon_discount",
                                 "total_discount","line_discount"])

    # Totals
    c_subtotal   = find_col(df, ["subtotal_inc_tax","subtotal","subtotal_ex_tax",
                                 "items_total","subtotal_price"])
    c_shipping   = find_col(df, ["shipping_cost_inc_tax","shipping_cost",
                                 "shipping_total","total_shipping","freight"])
    c_tax        = find_col(df, ["total_tax","tax_amount","tax","gst","vat","tax_total"])
    c_total      = find_col(df, ["total_inc_tax","total","grand_total",
                                 "order_total","total_price"])
    c_ship_method= find_col(df, ["shipping_method","shipping_zone","ship_method",
                                 "shipping_description"])

    all_rows = []
    skipped  = 0

    for _, row in df.iterrows():
        order_id = gv(row, c_order_id)
        if not order_id: skipped += 1; continue

        order_name   = "#{}".format(order_id)
        email        = gv(row, c_email)
        phone        = clean_phone(gv(row, c_phone))
        currency     = gv(row, c_currency, "USD")
        processed_at = format_order_date(gv(row, c_date))
        note         = gv(row, c_note)

        raw_status   = gv(row, c_status, "paid").lower().replace(" ", "_")
        pay_status   = PAYMENT_STATUS_MAP.get(raw_status, "paid")

        # Billing address
        b_fname     = gv(row, c_b_fname)
        b_lname     = gv(row, c_b_lname)
        b_company   = gv(row, c_b_company)
        b_phone     = clean_phone(gv(row, c_b_phone))
        b_addr1     = gv(row, c_b_addr1)
        b_addr2     = gv(row, c_b_addr2)
        b_city      = gv(row, c_b_city)
        b_state_raw = gv(row, c_b_state)
        b_prov_code = b_state_raw.upper()[:2] if len(b_state_raw) <= 4 else ""
        b_province  = b_state_raw if len(b_state_raw) > 2 else ""
        b_country_raw= gv(row, c_b_country, "US")
        b_cc        = b_country_raw.upper()[:2] if len(b_country_raw) <= 3 else "US"
        b_country   = b_country_raw if len(b_country_raw) > 3 else ""
        b_zip       = gv(row, c_b_zip)

        # Shipping address — fall back to billing if not present
        s_fname     = gv(row, c_s_fname) or b_fname
        s_lname     = gv(row, c_s_lname) or b_lname
        s_company   = gv(row, c_s_company) or b_company
        s_phone     = clean_phone(gv(row, c_s_phone)) or b_phone
        s_addr1     = gv(row, c_s_addr1) or b_addr1
        s_addr2     = gv(row, c_s_addr2) or b_addr2
        s_city      = gv(row, c_s_city)  or b_city
        s_state_raw = gv(row, c_s_state) or b_state_raw
        s_prov_code = s_state_raw.upper()[:2] if len(s_state_raw) <= 4 else b_prov_code
        s_province  = s_state_raw if len(s_state_raw) > 2 else b_province
        s_country_raw = gv(row, c_s_country, b_country_raw)
        s_cc        = s_country_raw.upper()[:2] if len(s_country_raw) <= 3 else b_cc
        s_country   = s_country_raw if len(s_country_raw) > 3 else b_country
        s_zip       = gv(row, c_s_zip) or b_zip

        # Values
        product_name  = gv(row, c_product, "Item")
        sku           = gv(row, c_sku)
        qty           = sv(sint(row.get(c_qty) if c_qty else None) or "1")
        unit_price    = sfloat(row.get(c_price) if c_price else None)
        discount_raw  = sfloat(row.get(c_discount) if c_discount else None)
        discount      = "-{}".format(discount_raw) if discount_raw else ""
        shipping_cost = sfloat(row.get(c_shipping) if c_shipping else None)
        ship_method   = gv(row, c_ship_method, "Shipping")
        tax_total     = sfloat(row.get(c_tax) if c_tax else None)
        order_total   = sfloat(row.get(c_total) if c_total else None)

        # ── ROW 1: Line Item (product) ────────────────────────
        line_row = _order_base(
            order_name, email, phone, currency, pay_status, processed_at,
            note, "",
            b_fname, b_lname, b_company, b_phone,
            b_addr1, b_addr2, b_city, b_province, b_prov_code,
            b_country, b_cc, b_zip,
            s_fname, s_lname, s_company, s_phone,
            s_addr1, s_addr2, s_city, s_province, s_prov_code,
            s_country, s_cc, s_zip,
        )
        # Fill in line item columns (indices 34-42)
        line_row[ORDER_HEADERS.index("Line: Type")]              = "Line Item"
        line_row[ORDER_HEADERS.index("Line: Title")]             = product_name
        line_row[ORDER_HEADERS.index("Line: Quantity")]          = qty
        line_row[ORDER_HEADERS.index("Line: Price")]             = unit_price
        line_row[ORDER_HEADERS.index("Line: SKU")]               = sku
        line_row[ORDER_HEADERS.index("Line: Requires Shipping")] = "TRUE"
        line_row[ORDER_HEADERS.index("Line: Taxable")]           = "TRUE"
        line_row[ORDER_HEADERS.index("Line: Discount")]          = discount
        # Fill in tax columns if available
        if tax_total:
            line_row[ORDER_HEADERS.index("Tax 1: Title")] = "Tax"
            line_row[ORDER_HEADERS.index("Tax 1: Price")] = tax_total
        all_rows.append(line_row)

        # ── ROW 2: Shipping Line ──────────────────────────────
        if shipping_cost:
            ship_row = _order_base(
                order_name, "", "", currency, pay_status, processed_at,
                "", "",
                b_fname, b_lname, b_company, b_phone,
                b_addr1, b_addr2, b_city, b_province, b_prov_code,
                b_country, b_cc, b_zip,
                s_fname, s_lname, s_company, s_phone,
                s_addr1, s_addr2, s_city, s_province, s_prov_code,
                s_country, s_cc, s_zip,
            )
            ship_row[ORDER_HEADERS.index("Line: Type")]  = "Shipping Line"
            ship_row[ORDER_HEADERS.index("Line: Title")] = ship_method
            ship_row[ORDER_HEADERS.index("Line: Price")] = shipping_cost
            all_rows.append(ship_row)

        # ── ROW 3: Transaction (payment record) ───────────────
        # Required to show "Paid by customer" amount in Shopify admin
        if order_total and pay_status == "paid":
            txn_row = _order_base(
                order_name, "", "", currency, pay_status, processed_at,
                "", "",
                b_fname, b_lname, b_company, b_phone,
                b_addr1, b_addr2, b_city, b_province, b_prov_code,
                b_country, b_cc, b_zip,
                s_fname, s_lname, s_company, s_phone,
                s_addr1, s_addr2, s_city, s_province, s_prov_code,
                s_country, s_cc, s_zip,
            )
            txn_row[ORDER_HEADERS.index("Line: Type")]              = "Transaction"
            txn_row[ORDER_HEADERS.index("Transaction: Amount")]     = order_total
            txn_row[ORDER_HEADERS.index("Transaction: Currency")]   = currency
            txn_row[ORDER_HEADERS.index("Transaction: Kind")]       = "sale"
            txn_row[ORDER_HEADERS.index("Transaction: Status")]     = "success"
            txn_row[ORDER_HEADERS.index("Transaction: Gateway")]    = "Custom Gateway"
            all_rows.append(txn_row)

    # Validate column count
    bad = [(i+2, len(r)) for i, r in enumerate(all_rows) if len(r) != len(ORDER_HEADERS)]
    if bad:
        log("BUG: {} rows with wrong column count!".format(len(bad)), "ERROR")
        for ri, rc in bad[:5]:
            log("  Row {}: {} cols (expected {})".format(ri, rc, len(ORDER_HEADERS)), "ERROR")
        return

    write_csv(all_rows, ORDER_HEADERS, "shopify_orders.csv")
    print()
    log("Orders converted   : {}".format(len(df) - skipped), "OK")
    log("Orders skipped     : {} (no order ID)".format(skipped),
        "OK" if skipped == 0 else "WARN")
    log("Total CSV rows     : {} (line + shipping + transaction rows)".format(
        len(all_rows)), "OK")
    print()
    log("IMPORTANT: Orders cannot be imported via Shopify's native importer.", "WARN")
    log("Use the free Matrixify app:", "INFO")
    log("  1. Shopify Admin -> Apps -> Matrixify", "INFO")
    log("  2. Click Import -> Add file -> shopify_orders.csv", "INFO")
    log("  3. Sheet name auto-detects as 'Orders'", "INFO")
    log("  4. Review and click Import", "INFO")
    print()


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════

def main():
    print()
    print("  +---------------------------------------------------------+")
    print("  |  BC -> Shopify CSV Converter  v13.0                     |")
    print("  |  Products + Customers + Orders  |  No API required      |")
    print("  +---------------------------------------------------------+")
    print()
    print("  Output folder: {}".format(OUTPUT_DIR))
    print()

    # ── Products ──────────────────────────────────────────────
    print("  =============================================================")
    print("  PRODUCTS")
    print("  =============================================================")
    prod_path = input("  BC products CSV path (Enter to skip):\n  > ").strip().strip('"\'')
    if prod_path and os.path.exists(prod_path):
        print()
        convert_products(prod_path)
    elif prod_path:
        log("File not found: {}".format(prod_path), "ERROR")

    # ── Customers ─────────────────────────────────────────────
    print("  =============================================================")
    print("  CUSTOMERS")
    print("  =============================================================")
    cust_path = input("  BC customers CSV path (Enter to skip):\n  > ").strip().strip('"\'')
    if cust_path and os.path.exists(cust_path):
        print()
        convert_customers(cust_path)
    elif cust_path:
        log("File not found: {}".format(cust_path), "ERROR")

    # ── Orders ────────────────────────────────────────────────
    print("  =============================================================")
    print("  ORDERS")
    print("  =============================================================")
    print("  Note: Output will be in Matrixify format.")
    print("  Import via: Shopify Admin -> Apps -> Matrixify -> Import")
    print()
    ord_path = input("  BC orders CSV path (Enter to skip):\n  > ").strip().strip('"\'')
    if ord_path and os.path.exists(ord_path):
        print()
        convert_orders(ord_path)
    elif ord_path:
        log("File not found: {}".format(ord_path), "ERROR")

    # ── Summary ───────────────────────────────────────────────
    print("  =============================================================")
    print("  OUTPUT FILES  ->  {}".format(OUTPUT_DIR))
    print("  =============================================================")
    if os.path.exists(OUTPUT_DIR):
        for fname in sorted(os.listdir(OUTPUT_DIR)):
            fpath = os.path.join(OUTPUT_DIR, fname)
            if os.path.isfile(fpath):
                print("    {:>12,} bytes  {}".format(os.path.getsize(fpath), fname))
    print()
    print("  HOW TO IMPORT:")
    print("  Products  -> Shopify Admin -> Products -> Import -> shopify_products.csv")
    print("  Customers -> Shopify Admin -> Customers -> Import -> shopify_customers.csv")
    print("  Orders    -> Matrixify app -> Import -> shopify_orders.csv")
    print()

if __name__ == "__main__":
    main()
