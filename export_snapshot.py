#!/usr/bin/env python3
"""Export the dashboard numbers to ../snapshot.json (read-only: Excel + a YNAB balance read).

Used by the scheduled phone-dashboard refresh. Reads YNAB_TOKEN from the environment,
falling back to the Windows user environment in the registry, so it also works when
started by a process that was launched before the variable was set.
"""
import json, os, sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), 'snapshot.json')


def _user_env(name):
    if os.environ.get(name):
        return
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment') as k:
            os.environ[name] = winreg.QueryValueEx(k, name)[0]
    except Exception:
        pass


for _n in ('YNAB_TOKEN', 'ANTHROPIC_API_KEY'):
    _user_env(_n)

sys.path.insert(0, HERE)
os.chdir(HERE)
import app  # noqa: E402  (importing does not start the server)

a = app.build_analytics()
keep = ['total', 'budget_civa', 'budget_siva', 'pct', 'available', 'bank', 'cash',
        'balance_source', 'projected', 'projected_siva', 'remaining_to_pay']
out = {k: a[k] for k in keep}
out['table'] = [{k: r[k] for k in ('tag', 'actual', 'budget_civa', 'closed', 'has_budget', 'status')}
                for r in a['table']]
out['monthly'] = list(zip(json.loads(a['monthly_labels']), json.loads(a['monthly_values']),
                          json.loads(a['monthly_cumulative'])))
out['generated_at'] = datetime.now().astimezone().isoformat(timespec='minutes')

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=1, default=str)
# Wrapper document for the phone dashboard's database: snapshot/latest
with open(os.path.join(os.path.dirname(HERE), 'snapshot-doc.json'), 'w', encoding='utf-8') as f:
    json.dump({'data': out, 'updatedAt': out['generated_at']}, f, ensure_ascii=False, default=str)
print(f"ok {OUT} balance_source={out['balance_source']} generated_at={out['generated_at']}")
