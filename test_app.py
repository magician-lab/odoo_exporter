import os
import sys
import json
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import app as app_mod
from app import app, build_rows, catalogue_products, catalogue_html

PASS = 0
FAIL = 0


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  [PASS] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name} {detail}")


print("=" * 70)
print("TEST 1: Schema-aware build_rows (live Odoo data)")
print("=" * 70)
rows = build_rows()
check("build_rows returns data", len(rows) > 100, f"got {len(rows)}")

variants = [r for r in rows if r["type"] == "VARIANT"]
templates = [r for r in rows if r["type"] == "TEMPLATE"]
print(f"  templates={len(templates)} variants={len(variants)}")

# Product types surfaced from product.template.type
types = {r["product_type"] for r in rows}
print(f"  product types: {types}")
check("Goods + Service types present", {"Goods", "Service"} <= types, f"got {types}")

services = [r for r in rows if r["product_type"] == "Service"]
check("services have no stock (N/A)", len(services) > 0 and all(r.get("stock") is None for r in services))
check("services not flagged Out Of Stock", all("Out Of Stock" not in (r.get("defects") or []) for r in services))
check("services not flagged Missing Buying Price", all("Missing Buying Price" not in (r.get("defects") or []) for r in services))
check("services not flagged Missing Image", all("Missing Image" not in (r.get("defects") or []) for r in services))

# Known-good sell prices read from stored lst_price (verified in DB)
by_id = {r["id"]: r for r in variants}
for vid, expected_sell in [(3755, 850.0), (3756, 850.0), (3822, 16500.0), (3819, 2000.0), (1327, 7000.0)]:
    v = by_id.get(vid)
    if v is None:
        check(f"variant {vid} present", False, "missing from rows")
        continue
    got = v.get("sell_raw")
    check(f"variant {vid} sell = {expected_sell}", got == expected_sell, f"got {got} (src: {v.get('price_source')})")

with_sell = sum(1 for r in variants if r.get("sell_raw"))
with_buy = sum(1 for r in variants if r.get("buy_raw"))
print(f"  variants with sell price: {with_sell}/{len(variants)}  buy: {with_buy}/{len(variants)}")
check("sell price coverage > 85%", with_sell / len(variants) > 0.85)
check("buy price coverage > 75%", with_buy / len(variants) > 0.75)

# template buy fallback to family majority / template standard_price
tmpl_by_id = {r["id"]: r for r in templates}
if 88 in tmpl_by_id:
    t = tmpl_by_id[88]
    print(f"  template 88 (PETG) buy={t.get('buy_raw')} src={t.get('price_source')}")
    check("PETG template has a buy price", t.get("buy_raw") and t.get("buy_raw") > 0, f"got {t.get('buy_raw')}")

# junk 1.0 placeholder prices must NOT be treated as real sell prices.
# A variant with no own price inherits a sibling price ONLY when that price
# is the family MAJORITY (>50%) - e.g. endmill family {850 x3, 3200} -> 850.
junk = by_id.get(3759)  # 6mm endmill: lst_price = 1.0 placeholder, family majority = 850
if junk is not None:
    check("variant 3759 inherits family majority sell (850)", junk.get("sell_raw") == 850.0, f"got {junk.get('sell_raw')}")
    check("variant 3759 price source is family lst_price", "family lst_price" in (junk.get("price_source") or ""))

# A truly mixed family (no >50% value) is NEVER inherited - variants are flagged.
# MIG Wires family {2500, 4000, 13500} all unique -> 4 unpriced variants stay missing.
mig_unpriced = [r for r in variants if r.get("template_id") == 3246 and r.get("sell_raw") is None]
check("MIG Wires mixed family: no invented prices", len(mig_unpriced) >= 3, f"got {len(mig_unpriced)}")
check("MIG Wires unpriced variants flagged", all("Missing Selling Price" in (r.get("defects") or []) for r in mig_unpriced))

# margin sanity: allow genuine negative (sell below cost) but never > 100
bad_margin = [r for r in rows if r.get("margin") is not None and r["margin"] > 100]
check("no absurd margins (>100%)", len(bad_margin) == 0, f"{len(bad_margin)} rows: {[(r['name'], r['margin']) for r in bad_margin[:5]]}")

# No product without a product_type label
check("every row has a product_type", all(r.get("product_type") for r in rows))

print("\n" + "-" * 70)
print("TEST 1b: Sales analytics fields on rows")
print("-" * 70)
check("rows carry units_30d", all("units_30d" in r for r in rows))
check("rows carry revenue_30d", all("revenue_30d" in r for r in rows))
check("rows carry abc_class", all("abc_class" in r for r in rows))
check("rows carry stock_value", all("stock_value" in r for r in rows))
check("rows carry create_date", all("create_date" in r for r in rows))
check("abc classes valid", all(r["abc_class"] in ("A", "B", "C", "No Sales") for r in rows), f"got {sorted({r['abc_class'] for r in rows})}")
check("no negative revenue", all(r["revenue_30d"] >= 0 for r in rows))
check("no negative units", all(r["units_30d"] >= 0 for r in rows))
# POS is the dominant channel: 619 pos orders / 1001 lines exist in this DB
sold = [r for r in rows if r["revenue_total"] > 0]
print(f"  rows with sales: {len(sold)} / {len(rows)}")
check("some rows have sales data", len(sold) > 10, f"got {len(sold)}")
top10 = sorted(sold, key=lambda r: -r["revenue_30d"])[:10]
check("top sellers have ABC A or B", all(r["abc_class"] in ("A", "B") for r in top10),
      f"got {[(r['name'], r['abc_class']) for r in top10[:5]]}")

print("\n" + "=" * 70)
print("TEST 2: Flask endpoints (test client)")
print("=" * 70)
client = app.test_client()
t0 = time.time()

r = client.get("/")
check("GET / returns 200", r.status_code == 200)
check("dashboard page mentions Product Intelligence", b"Odoo Product Intelligence" in r.data or b"Product Intelligence" in r.data)

r = client.get("/api/products")
data = r.get_json()
check("GET /api/products 200 + JSON array", r.status_code == 200 and isinstance(data, list), f"status {r.status_code}")
check("api/products has rows", len(data) > 100, f"got {len(data)}")

r = client.get("/api/products?issue=Missing+Selling+Price")
data = r.get_json()
check("issue filter works", all("Missing Selling Price" in (x.get("defects") or []) for x in data), f"got {len(data)} rows")

r = client.get("/api/products?search=endmill")
data = r.get_json()
check("search filter works", len(data) >= 5, f"got {len(data)} rows")
check("search finds 4 flute carbide endmill w/ real price", any(x.get("sell_raw") == 850 for x in data if x.get("type") == "VARIANT"))

r = client.get("/api/products?status=GOOD")
data = r.get_json()
check("status=GOOD filter works", all(x.get("status") == "GOOD" for x in data))

r = client.get("/api/products?ptype=Service")
data = r.get_json()
check("product type filter (Service) works", len(data) > 0 and all(x.get("product_type") == "Service" for x in data), f"got {len(data)} rows")

r = client.get("/api/products?abc=A")
data = r.get_json()
check("ABC=A filter works", len(data) > 0 and all(x.get("abc_class") == "A" for x in data), f"got {len(data)} rows")

r = client.get("/api/products?has_sales=1")
data = r.get_json()
check("has_sales filter works", len(data) > 0 and all((x.get("revenue_total") or 0) > 0 for x in data), f"got {len(data)} rows")

r = client.get("/api/products?min_stock=100")
data = r.get_json()
check("min_stock filter works", all((x.get("stock") or 0) >= 100 for x in data), f"got {len(data)} rows")

r = client.get("/api/product_image?id=88")
img = r.get_json()
check("product_image returns base64 image", r.status_code == 200 and img.get("image", "").startswith("data:image"), f"got {img.get('image', '')[:40]}")

r = client.get("/api/analytics")
a = r.get_json()
check("GET /api/analytics 200", r.status_code == 200)
check("analytics has kpis", "kpis" in a and a["kpis"]["products"] > 100, f"got {a.get('kpis', {}).get('products')}")
check("analytics kpis include stock_value + revenue_30d", "stock_value" in a["kpis"] and "revenue_30d" in a["kpis"])
check("analytics has by_category", isinstance(a.get("by_category"), list) and len(a["by_category"]) > 0, f"got {len(a.get('by_category', []))}")
check("analytics has defects", isinstance(a.get("defects"), list) and len(a["defects"]) > 0)
check("analytics has sales_trend (6 months)", len(a.get("sales_trend", [])) == 6, f"got {len(a.get('sales_trend', []))}")
check("analytics has top_sellers", isinstance(a.get("top_sellers"), list) and len(a["top_sellers"]) > 0)
check("analytics has abc breakdown", isinstance(a["kpis"].get("abc"), dict) and sum(a["kpis"]["abc"].values()) > 0, f"got {a['kpis'].get('abc')}")
check("analytics has price_dist", isinstance(a.get("price_dist"), list) and len(a["price_dist"]) == 8, f"got {len(a.get('price_dist', []))}")
check("analytics has locations", isinstance(a.get("locations"), list) and len(a["locations"]) >= 2, f"got {a.get('locations')}")

r = client.get("/api/options")
opts = r.get_json()
check("GET /api/options 200", r.status_code == 200 and "categories" in opts)
check("options has categories", len(opts["categories"]) > 0)
check("options has product_types", "Goods" in opts.get("product_types", []) and "Service" in opts.get("product_types", []), f"got {opts.get('product_types')}")

r = client.get("/export_report")
check("GET /export_report 200 xlsx", r.status_code == 200 and r.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"mimetype {r.mimetype}")

r = client.get("/export_report?abc=A")
check("export with abc filter 200 xlsx", r.status_code == 200, f"status {r.status_code}")

r = client.get("/catalogue")
check("GET /catalogue 200", r.status_code == 200)

r = client.get("/api/catalogue_products")
cp = r.get_json()
check("catalogue_products 200", r.status_code == 200 and "products" in cp)
check("catalogue products present", cp["total_products"] > 100, f"got {cp.get('total_products')}")
c_variants = [p for p in cp["products"] if p["type"] == "VARIANT"]
c_v3755 = next((p for p in c_variants if p.get("variant_id") == 3755), None)
check("catalogue variant 3755 sell=850", c_v3755 is not None and c_v3755["sell"] == 850, f"got {c_v3755}")
# internal expense-code services (no sell price) excluded from catalogue
check("catalogue excludes price-less services", all(p.get("product_type") != "Service" or p.get("sell") for p in cp["products"]))

r = client.get("/catalogue/review")
check("GET /catalogue/review 200", r.status_code == 200)
check("review page has no stray 'use' syntax error", b"\nuse\n" not in r.data)

r = client.post("/catalogue/preview", json={
    "title": "TEST CATALOGUE", "website": "test.com", "contacts": "123",
    "companyDescription": "Test company", "logo": "", "qr": "",
    "products": [{
        "name": "Test Product", "sell": 850, "discount": 10,
        "description": "A great product", "useSystemDesc": True,
        "showDescription": True, "showPrice": True, "showDiscount": True,
        "image_1920": "", "category": "3D Printing", "imageSize": "medium",
    }],
})
check("POST /catalogue/preview 200", r.status_code == 200)
check("preview contains product name", b"Test Product" in r.data)

print("\n" + "=" * 70)
print("TEST 3: API timing")
print("=" * 70)
print(f"  /api/products full run: {time.time()-t0:.1f}s")

print("\n" + "=" * 70)
print(f"RESULT: {PASS} passed, {FAIL} failed")
print("=" * 70)
sys.exit(1 if FAIL else 0)
