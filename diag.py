import app, json
from collections import Counter

with app.app.app_context():
    print('Running diagnostics...')
    rows = app.build_rows()
    print('build_rows() rows count:', len(rows))
    print('good count:', sum(1 for r in rows if r.get('status')=='GOOD'))

    cnt = Counter()
    no_sell = 0
    for r in rows:
        for d in r.get('defects', []):
            cnt[d] += 1
        try:
            sell = r.get('sell_raw')
            if sell is None or (isinstance(sell, (int,float)) and sell <= 0):
                no_sell += 1
        except Exception:
            no_sell += 1

    print('defect counts:', dict(cnt))
    print('rows with no sell price:', no_sell)

    cp = app.catalogue_products()
    try:
        data = cp.get_json()
    except Exception:
        try:
            data = json.loads(cp.get_data())
        except Exception:
            data = None

    print('catalogue_products total:', data.get('total_products') if isinstance(data, dict) else 'N/A')
    print('catalogue_products templates:', data.get('templates') if isinstance(data, dict) else 'N/A')
    print('catalogue_products variants:', data.get('variants') if isinstance(data, dict) else 'N/A')

    samples = [r for r in rows if (r.get('sell_raw') in (None, 0))]
    print('sample missing sell (up to 5):')
    for s in samples[:5]:
        print('-', s.get('type'), s.get('id'), s.get('name'), 'sell=', s.get('sell_raw'))

    print('Done')