import app

print("== PETG (template 88) ==")
t = app.search_read("product.template", ["id", "name", "list_price", "standard_price"], [["id", "=", 88]])
print("template:", t[0] if t else None)
vs = app.search_read("product.product", ["id", "name", "lst_price", "list_price", "standard_price"], [["product_tmpl_id", "=", 88]])
for v in vs:
    print("   ", v["id"], v["name"][:45], "lst=", v["lst_price"], "std=", v["standard_price"])

print("\n== PETG template variant list (from build_rows) ==")
rows = app.build_rows()
by_id = {r["id"]: r for r in rows}
r88 = by_id.get(88)
print("template 88 row:", r88["buy_raw"], "|", r88["price_source"], "| defects:", r88["defects"])

print("\n== COVERAGE BREAKDOWN ==")
variants = [r for r in rows if r["type"] == "VARIANT"]
no_sell = [r for r in variants if not r["sell_raw"]]
no_buy = [r for r in variants if not r["buy_raw"]]
print("variants no sell:", len(no_sell), "no buy:", len(no_buy))
print("sample no-sell families:")
from collections import Counter
c = Counter((r.get("variant_of") or "") for r in no_sell)
for name, n in c.most_common(12):
    print(f"   {n}x  {name}")
print("sample no-buy families:")
c2 = Counter((r.get("variant_of") or "") for r in no_buy)
for name, n in c2.most_common(12):
    print(f"   {n}x  {name}")