"""Fix the 19 consumable products: clear negative move balances with intake
receipts (Vendors->NBI), switch templates to is_storable, then set target
stock. Targets come from _apply_stock/checkpoint.json failed list.
"""
import sys
import os
import json
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import stock_reconcile as s

NBI = 5
VENDORS = 1


def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def done_balance(pid):
    mv = s.search_read('stock.move', ['id', 'location_id', 'location_dest_id', 'quantity', 'state'],
                       [('product_id', '=', pid)])
    bal = 0.0
    for m in mv:
        if m['state'] != 'done':
            continue
        if m['location_id'][0] == NBI:
            bal -= fnum(m['quantity'])
        if m['location_dest_id'][0] == NBI:
            bal += fnum(m['quantity'])
    return bal


def make_receipt(pid, qty, label):
    pt = s.search_read('stock.picking.type', ['id', 'code'], [('code', '=', 'incoming')])
    p = s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD, 'stock.picking', 'create', [{
        'picking_type_id': pt[0]['id'],
        'location_id': VENDORS,
        'location_dest_id': NBI,
        'name': f'{label} {pid} {int(time.time())}',
        'move_ids': [(0, 0, {
            'product_id': pid,
            'location_id': VENDORS,
            'location_dest_id': NBI,
            'product_uom_qty': qty,
            'quantity': qty,
        })],
    }])
    s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD, 'stock.picking', 'action_confirm', [[p]])
    s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD, 'stock.picking', 'button_validate', [[p]])
    return p


def apply_quant(pid, loc, tgt):
    qs = s.search_read('stock.quant', ['id', 'quantity'], [('product_id', '=', pid), ('location_id', '=', loc)])
    if qs:
        qid = qs[0]['id']
        s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD, 'stock.quant', 'write',
                             [[qid], {'inventory_quantity': tgt}])
        try:
            s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD, 'stock.quant',
                                 'action_apply_inventory', [[qid]])
        except Exception:
            pass
        return qid
    qid = s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD, 'stock.quant', 'create',
                               [{'product_id': pid, 'location_id': loc, 'inventory_quantity': tgt}])
    try:
        s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD, 'stock.quant',
                             'action_apply_inventory', [[qid]])
    except Exception:
        pass
    return qid


def main():
    ck = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     '_apply_stock', 'checkpoint.json')))
    failed = ck['failed']
    print('failed entries:', len(failed))

    # 1. intake receipts for variants with negative NBI balance (qty = deficit + NBI target)
    nbi_targets = {}
    for f in failed:
        if f['loc'] == NBI and f['tgt'] > 0:
            nbi_targets[f['pid']] = f['tgt']
    print('NBI targets:', nbi_targets)
    for pid, tgt in sorted(nbi_targets.items()):
        deficit = -done_balance(pid)
        if deficit < -0.001:
            print(f'pid {pid} still negative ({deficit}); not receiving yet')
        qty = max(deficit, 0.0) + tgt
        if qty > 0.001:
            p = make_receipt(pid, qty, 'Stock reconciliation intake')
            print(f'  receipt pid {pid} qty {qty} -> picking {p}')

    # 2. switch the 7 blocked templates to storable
    for tid in [24, 25, 41, 48, 3209, 3224, 3288]:
        try:
            s._models.execute_kw(s.ODOO_DB, s._uid, s.ODOO_PASSWORD,
                                 'product.template', 'write', [[tid], {'is_storable': True}])
            print(f'template {tid} -> storable OK')
        except Exception as e:
            print(f'template {tid} -> FAIL {str(e)[-150:]}')

    # 3. apply all failed targets (RUI + any NBI not covered by receipts)
    for f in failed:
        pid, loc, tgt = f['pid'], f['loc'], f['tgt']
        try:
            apply_quant(pid, loc, tgt)
            print(f'  set pid {pid} loc {loc} -> {tgt}')
        except Exception as e:
            print(f'  FAIL pid {pid} loc {loc}: {str(e)[-150:]}')

    # 4. verify
    print('\nVERIFY:')
    ok = True
    for f in failed:
        pid, loc, tgt = f['pid'], f['loc'], f['tgt']
        qs = s.search_read('stock.quant', ['quantity'], [('product_id', '=', pid), ('location_id', '=', loc)])
        now = sum(fnum(q['quantity']) for q in qs)
        status = 'OK' if abs(now - tgt) < 0.001 else 'MISMATCH'
        if status != 'OK':
            ok = False
        print(f'  {status} pid {pid} loc {loc}: now {now}, want {tgt}')
    print('ALL OK' if ok else 'HAS FAILURES')


if __name__ == '__main__':
    main()