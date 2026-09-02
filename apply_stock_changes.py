"""Apply expected-stock values to live Odoo stock.quants (resumable).
Backs up current quants first, applies via inventory_quantity +
action_apply_inventory (Odoo 17+ flow), then verifies in bulk.
Negatives are clamped to 0 and reported. Not-in-system rows are skipped.
Checkpoints progress to _apply_stock/checkpoint.json so it can be re-run
to resume after a timeout.
"""
import sys
import os
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import stock_reconcile as s

OUT = os.path.join(s.BASE, "_apply_stock")
CHECK = os.path.join(OUT, "checkpoint.json")


def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def build_changes(report):
    by_pid = report["by_pid"]
    cur_nbi, cur_rui = report["cur_nbi"], report["cur_rui"]
    sold = report["sold"]
    nbi_id, rui_id = report["nbi_id"], report["rui_id"]

    changes = []

    def consider(pid, loc, cur, exp, name):
        if abs(exp - cur) < 0.001:
            return
        tgt = max(exp, 0.0)
        if abs(tgt - cur) < 0.001:
            return
        changes.append({
            "pid": pid, "loc": loc, "name": name,
            "cur": cur, "exp": exp, "tgt": tgt, "neg": exp < 0,
        })

    for pid, b in by_pid.items():
        v = report["variants_by_id"].get(pid)
        name = v["display"] if v else f"pid {pid}"
        counted = fnum(b["count"]); cons = fnum(b["consignment"])
        fw = fnum(b["from_wh"]); wh = fnum(b["warehouse"])
        pi = fnum(b["pi"])
        cur_n = fnum(cur_nbi.get(pid)); cur_r = fnum(cur_rui.get(pid))
        sld = fnum(sold.get(pid))
        exp_n = (counted if counted else cur_n) + cons + fw - sld
        exp_r = (wh if wh else cur_r) - fw + pi
        consider(pid, nbi_id, cur_n, exp_n, name)
        consider(pid, rui_id, cur_r, exp_r, name)

    return changes


def fetch_quants(pid, loc):
    q = s.search_read("stock.quant",
                      ["id", "product_id", "location_id", "quantity"],
                      [("product_id", "=", pid), ("location_id", "=", loc)])
    return q


def save(ckpt):
    with open(CHECK, "w") as f:
        json.dump(ckpt, f, indent=1)


def main():
    os.makedirs(OUT, exist_ok=True)
    ckpt = {}
    if os.path.exists(CHECK):
        ckpt = json.load(open(CHECK))
        print(f"resuming: {len(ckpt.get('applied', []))} already applied")

    system = s.fetch_system()
    sales = s.fetch_sales()
    variants = s.build_variant_objects(system["variants"], system["templates"])
    variants_by_id = {v["id"]: v for v in variants}
    overrides = s.build_overrides(system["variants"])
    report = s.build_report(system, sales, overrides)
    report["variants_by_id"] = variants_by_id

    changes = build_changes(report)
    print(f"CHANGES: {len(changes)}")

    applied, failed = ckpt.get("applied", []), ckpt.get("failed", [])
    done_keys = {(a["pid"], a["loc"]) for a in applied} | {(f_["pid"], f_["loc"]) for f_ in failed}

    # backup current quants of affected products (also used to find existing quants)
    backup = {}     # (pid, loc) -> quant id
    backup_qty = {}  # (pid, loc) -> current quantity
    pids = sorted({c["pid"] for c in changes})
    for i in range(0, len(pids), 50):
        chunk = pids[i:i + 50]
        qs = s.search_read("stock.quant",
                           ["id", "product_id", "location_id", "quantity"],
                           [("product_id", "in", chunk),
                            ("location_id", "in", [report["nbi_id"], report["rui_id"]])])
        for q in qs:
            key = (q["product_id"][0], q["location_id"][0])
            backup[key] = q["id"]
            backup_qty[key] = backup_qty.get(key, 0.0) + fnum(q["quantity"])
    with open(os.path.join(OUT, "backup_quants.json"), "w") as f:
        json.dump([{"id": v, "pid": k[0], "loc": k[1], "qty": backup_qty[k]}
                   for k, v in backup.items()], f, indent=1)
    print(f"backup: {len(backup)} quants saved to {OUT}\\backup_quants.json")

    # skip anything already at target (e.g. applied by an interrupted run)
    already = [c for c in changes
               if (c["pid"], c["loc"]) not in done_keys
               and abs(backup_qty.get((c["pid"], c["loc"]), 0.0) - c["tgt"]) < 0.001]
    for c in already:
        applied.append({**c, "quant_id": backup.get((c["pid"], c["loc"])), "preapplied": True})
    print(f"already at target (skipped): {len(already)}")
    done_keys |= {(a["pid"], a["loc"]) for a in already}

    # apply
    todo = [c for c in changes if (c["pid"], c["loc"]) not in done_keys]
    print(f"to apply now: {len(todo)}")
    for i, c in enumerate(todo):
        pid, loc, tgt = c["pid"], c["loc"], c["tgt"]
        try:
            qid = backup.get((pid, loc))
            if qid is not None:
                s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD,
                                     "stock.quant", "write",
                                     [[qid], {"inventory_quantity": tgt}])
                try:
                    s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD,
                                         "stock.quant", "action_apply_inventory",
                                         [[qid]])
                except Exception:
                    pass  # None result cannot be marshaled; applied server-side
            else:
                qid = s._models.execute_kw(
                    s.ODOO_DB, s._uid, s.ODOO_PASSWORD, "stock.quant",
                    "create", [{"product_id": pid, "location_id": loc,
                                "inventory_quantity": tgt}])
                try:
                    s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD,
                                         "stock.quant", "action_apply_inventory",
                                         [[qid]])
                except Exception:
                    pass
            applied.append({**c, "quant_id": qid, "tried": True})
            if (i + 1) % 25 == 0 or i == len(todo) - 1:
                print(f"  ... {i+1}/{len(todo)} applied", flush=True)
        except Exception as e:
            failed.append({**c, "error": str(e)[:300]})
            print(f"FAIL {c['name']} (pid {pid} loc {loc}): {e}", flush=True)
        if (i + 1) % 10 == 0:
            save({"applied": applied, "failed": failed})

    save({"applied": applied, "failed": failed})
    print(f"\nAPPLIED: {len(applied)}  FAILED: {len(failed)}", flush=True)
    for f_ in failed:
        print(f"  FAILED: {f_['name']} loc {f_['loc']}: {f_['error']}", flush=True)

    # bulk verify everything applied so far
    if applied:
        pids = sorted({a["pid"] for a in applied})
        now = {}
        for i in range(0, len(pids), 50):
            chunk = pids[i:i + 50]
            qs = s.search_read("stock.quant",
                               ["id", "product_id", "location_id", "quantity"],
                               [("product_id", "in", chunk),
                                ("location_id", "in", [report["nbi_id"], report["rui_id"]])])
            for q in qs:
                now[(q["product_id"][0], q["location_id"][0])] = \
                    now.get((q["product_id"][0], q["location_id"][0]), 0.0) + fnum(q["quantity"])
        bad = []
        for a in applied:
            cur = now.get((a["pid"], a["loc"]), 0.0)
            if abs(cur - a["tgt"]) > 0.001:
                bad.append({**a, "now": cur})
        print(f"\nVERIFY: {len(applied) - len(bad)} ok, {len(bad)} mismatched", flush=True)
        for b_ in bad:
            print(f"  MISMATCH {b_['name']} loc {b_['loc']}: want {b_['tgt']}, now {b_['now']}", flush=True)
        with open(os.path.join(OUT, "mismatch.json"), "w") as f:
            json.dump(bad, f, indent=1)


if __name__ == "__main__":
    main()