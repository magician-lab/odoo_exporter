# agent.md — Odoo Exporter Project Guide

Read `architecture.md` first — it documents the live Odoo system this app talks to.

## What this project is

A local Flask app ("Odoo Product Intelligence") that reads the Survitec Odoo 19.2 SaaS
instance over XML-RPC and produces:

1. **Product intelligence dashboard** — every template + variant with sell/buy prices, stock,
   locations, defects (Missing Selling Price, Out Of Stock, Low Stock, Missing Description,
   Missing Image, Missing Buying Price, Low Margin, Unpublished, Variant Price Inconsistency).
2. **Excel report export** (`/export_report`) — styled xlsx with Filters + Products sheets.
3. **PDF catalogue builder** (`/catalogue`, `/catalogue/review`) — pick products, group by
   category, add company profile/logo/QR, discounts, custom descriptions → wkhtmltopdf PDF.

## Layout

```
odoo_exporter/
├── app.py                 # main backend (Flask + XML-RPC client) - port 5000
├── sales_reports.py       # standalone sales service (Flask) - port 5050
├── test_app.py            # live-DB test suite (no mocks; requires working .env)
├── diag.py                # diagnostic counters/defect summary
├── .env                   # credentials (NEVER commit/share)
├── templates/
│   ├── dashboard.html     # main KPI/filter/table UI
│   ├── catalogue.html     # catalogue builder UI
│   ├── review.html        # catalogue review/preview UI
│   └── sales/
│       └── report.html    # sales service UI (own template folder)
├── exports/               # generated xlsx + pdf (timestamped)
├── temp_catalogue_images/ # converted product images for pdfkit
└── server_out.log / server_err.log  # app runtime logs
```

## Commands

| Task | Command |
|---|---|
| Run app | `python app.py` → http://localhost:5000 (debug, host 0.0.0.0) |
| Run sales service | `python sales_reports.py` → http://localhost:5050 (debug, host 0.0.0.0) |
| Run tests | `python test_app.py` (exits 1 on any failure; hits the LIVE Odoo DB) |
| Diagnostics | `python diag.py` (defect counts, missing-sell samples) |
| Required | Python 3.14, flask, pandas, pillow, python-dotenv, pdfkit, openpyxl, wkhtmltopdf at `C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe` |

## How data flows

```
Odoo SaaS (xmlrpc) ──> fetch_master_data()  ──> build_rows() ──> API JSON / xlsx / catalogue
       ▲                    │                     │
   safe_call()       cached 180s (TTL)      templates + variants rows
   (3 retries)       clear via ?refresh=1
```

- `fetch_master_data()` pulls `product.template` (405), `product.product` (756),
  `stock.location` (internal usage only), `stock.quant` (1,760) — cached for 180s.
- `fetch_sales_data()` pulls `pos.order.line` (1,001) + `sale.order.line` (249) with
  their orders; only `paid`/`done` (POS) and `sale` (SO) states count. POS revenue uses
  `price_subtotal` — **`pos.order.line` has NO `price_total` in this Odoo build** (verified
  via fields_get; `get_valid_model_fields` would silently drop it — always verify).
  Produces per-variant sales + a 6-month revenue trend.
- `build_rows(include_images=False)` produces one row per template + one per variant with a
  unified shape (see row keys in app.py). Prices resolved via `get_sale_price_from_record` /
  `get_cost_from_record`, defects via `defects()`. Sales fields: `units_30d`, `revenue_30d`,
  `units_total`, `revenue_total`, `avg_sale_price`, `last_sale_date`, `days_of_stock`,
  `stock_value` (stock × cost), `abc_class` (A/B/C by 30d revenue cumulative share, else
  "No Sales").
- `apply_filters()` honors query params: `issue`, `search`, `location`, `category`, `type`,
  `ptype`, `status`, `min_margin`, `abc`, `has_sales`, `min_stock`, `max_stock`.

## Hard-won domain rules (do NOT break)

1. **Selling price** = `product.product.lst_price` (stored). Template `list_price` is computed
   (first variant). `price_extra` is already folded in — never add it. `1.0` = junk placeholder.
2. **Buying price** = `product.product.standard_price` (stored). `supplierinfo` is USD — never
   use as KES cost.
3. **Family majority inheritance**: a variant with no price may inherit the family's value ONLY
   when one value holds >50% of the family. Mixed families (e.g. MIG Wires {2500,4000,13500})
   are never invented — variants stay flagged "Missing Selling Price".
4. **Services** (`type=service`, 13 of them): no stock, no buying price, no image/description
   defect checks, no margin. They are internal expense codes, not inventory.
5. **Stock**: `qty_available` authoritative; goods only; `stock_na=True` for services.
6. **Description**: only `description_ecommerce` matters in this DB (others empty).
7. Margin = `(sell - buy) / sell * 100`, only for goods, never >100.
8. Catalogue excludes services without a real sell price.

## HTTP endpoints

| Route | Method | Purpose |
|---|---|---|
| `/` | GET | dashboard (KPIs, 6 charts, filters, sortable/paginated table, product drawer) |
| `/api/products` | GET | rows JSON; `?refresh=1` clears cache; filters above; `?images=1` embeds base64 images |
| `/api/product_image` | GET | single template image as data-URI (`?id=<template_id>`) for the drawer |
| `/api/analytics` | GET | KPIs, by_category, defect distribution, price buckets, top sellers, ABC counts, 6-month sales trend, location totals, fetch health |
| `/api/options` | GET | categories, locations, product_types, generated_at |
| `/export_report` | GET | xlsx download (applies same filters) |
| `/catalogue` | GET | builder UI |
| `/api/catalogue_products` | GET | products with base64 images + defects |
| `/catalogue/review` | GET | review UI |
| `/catalogue/preview` | POST | HTML preview (data-URI images, no files) |
| `/catalogue/download` | POST | PDF via pdfkit/wkhtmltopdf (images saved to temp_catalogue_images) |

## Sales service (`sales_reports.py`, port 5050)

Standalone Flask process — own template folder (`templates/sales/`), own cache (300s TTL),
re-uses the same XML-RPC auth + `safe_call` / `get_valid_model_fields` patterns.

| Route | Method | Purpose |
|---|---|---|
| `/` | GET | monthly sales dashboard (KPIs, trend, channel split, product table) |
| `/api/sales` | GET | JSON; `?month=YYYY-MM` filters (default = latest month, `all` = all time); `?refresh=1` clears cache |
| `/api/export` | GET | CSV of the product table for the selected month |

- Fetches `pos.order.line` (revenue = `price_subtotal`, states `paid`/`done`) + `sale.order.line`
  (`price_total`, state `sale`) — same rules as app.py.
- Groups sales per product template (via `product.product.product_tmpl_id`); each product carries:
  units/revenue/orders/avg price/last sale for the selected month, plus full-history monthly
  breakdown, POS vs Sales Order split, and the variants that actually sold.
- KPIs + monthly trend + channel split for the selected month.

## Editing conventions

- `app.py` is single-file; keep it that way unless the change is large — the app is tested as a whole.
- Keep the Odoo schema-verification pattern: `get_valid_model_fields(model, preferred)` before any `search_read`.
- Use `safe_call` (retries) for every Odoo call.
- `image_src()` must handle WebP/JPEG/PNG magic prefixes before base64 embed.
- After any change run `python test_app.py` — it asserts the real data invariants above
  (variant 3755 sell=850, 3759 inherits 850 from family majority, MIG wires stay missing,
  sales/ABC/stock-value fields present, analytics endpoint shape, etc.).
- Never commit `.env`, `exports/`, or `temp_catalogue_images/` contents.
- No comments unless they encode Odoo-schema facts like the existing header block.