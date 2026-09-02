# Survitec Equipments Ltd — Odoo System Architecture

> Discovered live via the XML-RPC API (`saas~19.2+e`, Odoo SaaS 19.2) on 2026-08-18.
> Database: `survitec-equipments-ltd` · Host: `https://www.survitec.co.ke` · Admin UID 2.

## 1. Platform

| Item | Value |
|---|---|
| Server version | `saas~19.2+e` (Odoo SaaS 19.2, protocol v1) |
| Host | `https://www.survitec.co.ke` |
| DB | `survitec-equipments-ltd` |
| Login user | `survitecequipmentsltd@gmail.com` (uid 2, full admin) |
| Auth | XML-RPC via API key: `/xmlrpc/2/common` (authenticate) + `/xmlrpc/2/object` (execute_kw) |
| API keys on user 2 | "Bulk SMS", "product fetching key" |
| Models in registry | 828 (`ir.model`), 17,949 `ir.model.fields` |
| SaaS wrapper | Custom `/home/odoo/src/custom/trial/saas_trial` intercepts `ir.module.module` and some RPC — treat module-list queries as unreliable |
| Attachments | 7,038 (`ir.attachment`) |

## 2. Companies & localization

- **2 companies**: Survitec Equipments Ltd (id 1, the active one) and Dahawi Credit Ltd (id 2).
- **Country**: Kenya (id 114). **Currencies**: KES (id 93, rate 1.0 — functional) and USD (id 1, rate 0.007710 ≈ trading rate).
- **Kenya localization installed**: `l10n_ke_edi_oscu.code` (292 OSCU codes), `l10n.ke.hr.payroll.*` (SHIF/NSSF reporting wizards, tax deduction card), `l10n_ke.tax.deduction.card`, M-PESA transaction model (`transaction.lipa.na.mpesa`).
- **Taxes (15)**: sale — VAT 16%, VAT 0%, 0% EXPORT, 0% EXEMPT, 2% WH, 2% CTL (+legacy dupes "VAT"/"vat"/"16"); purchase — 16%, 8%, 0%, 0% EXEMPT, 16% IMPORT, 2% WH.

## 3. Installed application areas (inferred from models/registry)

Sales (incl. eCommerce), Inventory, Point of Sale, Invoicing/Accounting, Purchasing, Website, CRM, Helpdesk, Recruitment + HR Payroll (KE), Expenses, Email Marketing, Blog, Discuss/chat, Knowledge, Spreadsheets, Studio (custom forms/approvals), WhatsApp integration, SMS, eSign/portal.

## 4. Product catalog (core data)

| Model | Count |
|---|---|
| `product.template` | 405 — 392 Goods (consu), 13 Service, 0 Combo; 357 storable |
| `product.product` (variants) | 756 |
| `product.category` | 5: Goods (1), Services (3), Expenses (2), Food (5), Deliveries (6) |
| `product.attribute` / `product.attribute.value` | 72 / 430 — Color, Wire Diameter, Size, Weight, Material Type, Spool Weight, Shielding Type, Capacity, Nozzle Size, EndMill sizes, MIG Wire Type, ER-32 Collet Size, Tungsten Rod Sizes, Ceramic cups, Lathe Tool sizes, etc. |
| `product.supplierinfo` | 13 — ALL in USD from Shenzhen Kingroon 3D Tech Co., Ltd. **Never usable as KES cost.** |
| `product.pricelist` | 3 — "Default" (KES) x2, "USD" |
| `uom.uom` | 15 — Units, mm, m, km, m², ml, L, g, kg, Ton, KWH, Pack of 6, Minutes, Hours, Days |
| `product.tag` / `product.public.category` | present (website shop categories) |

### Field semantics (verified against live DB — critical)

- **Selling price**: `product.product.lst_price` — **STORED, variant-level, authoritative**. `product.template.list_price` is stored but computed from the first variant. `price_extra` is **already folded into lst_price** — never add it. A value of `1.0` is Odoo's placeholder → treat as missing.
- **Buying cost**: `product.product.standard_price` — **STORED, variant-level**. `product.template.standard_price` is computed (store=False, first variant).
- **Stock**: `qty_available` (computed, on-hand across internal locations). Template qty == sum of variants. Services hold no quants (qty always 0) → stock rules do NOT apply to them.
- **Descriptions**: goods carry sales copy in `description_ecommerce` (~81% filled); `website_description`, `description_sale`, `public_description`, `description` are effectively EMPTY in this DB.
- **Images**: `product.template.image_1920` (stored), `image_512` (stored, readonly). Products carry WebP (`UklGR…` magic) or JPEG (`/9j/…`) base64 data.
- **Published**: `product.template.website_published` (computed, store=False).
- **Type** selection on template: `consu`=Goods, `service`=Service, `combo`=Combo.

## 5. Inventory

- **2 warehouses** (both under company 1):
  - `WH` — Survitec Equipments Ltd (main)
  - `CBP` — Survitec Equipments Ltd Ruiru Bypass Warehouse
- **13 locations**: WH/Stock (internal, 574 quants), CBP/Stock (internal, 251), Customers (239), Vendors (38), Inventory adjustment (658), plus view/transit/production/supplier.
- **1,760 stock quants** (`stock.quant`). `stock.lot` = 0 → **no serial/lot tracking**.
- **651 pickings** (518 done, 127 assigned, 4 confirmed, 2 cancel):
  - 616 **PoS Orders** (POS is the dominant outbound flow),
  - 33 Delivery Orders, 1 Receipt, 1 Internal Transfer.
- Picking types: Receipts, Internal Transfers, Delivery Orders, PoS Orders (duplicated per warehouse config).

## 6. Business flows (actual usage)

| Flow | Records | Notes |
|---|---|---|
| Sales Orders | 181 (138 draft, 32 sale, 11 sent) | All KES; total ≈ KES 3.27M. Recent orders come from the website ("Public user" partner id 4) |
| POS | 619 orders (607 done, 9 paid, 3 cancel) | Total ≈ KES 5.53M — the primary retail channel; session-based, cash journals |
| Invoices/Accounting | 452 moves (394 entry, 43 out_invoice, 13 out_refund, 2 in_invoice; 430 posted) | 1,339 move lines |
| Payments | 213 (`account.payment`) | |
| Purchase Orders | 3 (2 draft, 1 purchase) | P00002/P00001 = KES 18,800? No — USD 18,800 to Kingroon (import stock) |
| Partners | 115 (15 companies, 100 individuals; 38 customers, 24 suppliers) | |
| Users | 16 active backend users | |

**Journals (14)**: Sales (INV), Purchases (BILL), Bank (BNK1), Misc Ops (MISC), Cash (CSH1), Petty Cash (CSH2), **M-PESA Payment (CSH3)**, Rhodah (CSH4), Cash Basis Taxes (CABA), Exchange Difference (EXCH), Point of Sale (POSS), Inventory Valuation (STJ), M-pesa purchase (BILL1), Salaries (SLR0).

## 7. Website / eCommerce

- 1 website "Survitec Equipments Ltd", 14 pages (Home, About Us, Services, Contact Us, Privacy, Return Policy, Terms, Cookie Policy, job pages…).
- Menus: Home, Shop, Blog, Jobs + category menus → **3D Printing, Woodworking, Metal Works & Machining, General Tools** (shop categories).
- Payment providers **enabled**: "Pay via Mpesa" (custom) and "Cash on Delivery". DPO Pay / Flutterwave / Lipa na Mpesa present but **disabled**.
- Business: equipment supplies for 3D printing (filaments: PLA/PETG/TPU/ABS/ASA/Nylon), welding (MIG/TIG rods, tungsten, ceramics), machining (endmills, carbide tools, collets, lathe tooling), woodworking, general tools — mostly sold via POS and the web shop.

## 8. Other modules present (light usage)

- CRM (`crm.lead`), Helpdesk (`helpdesk.ticket`), Recruitment (`hr.applicant`, `hr.job`), HR leaves/payslips (KE payroll), Expenses, Email marketing (`mailing.mailing`), Blog, WhatsApp, Knowledge, Studio custom models.

## 9. API notes for integrators

- Auth: `common.authenticate(DB, USER, API_KEY, {})` → uid 2; then `execute_kw` on `/xmlrpc/2/object`.
- Wrap calls in retry logic; transient failures are common (see `safe_call` in app.py).
- `fields_get` per model is reliable; guard every requested field with `get_valid_model_fields` because field sets differ from stock Odoo (e.g. `uom.category` and `mail.channel` DON'T exist here; `factor`/`category_id` on `uom.uom` is not fetchable via XML-RPC).
- `read_group` is blocked by the SaaS trial controller for some models — use `search_read`/`search_count` + Counter instead.
- `ir.module.module` search is intercepted by `saas_trial` — don't rely on it.
- All prices in KES except `product.supplierinfo` (USD) and POs to Kingroon (USD).