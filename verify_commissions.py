#!/usr/bin/env python3
import os
from pathlib import Path

# Simple verification: map filename stems to a contract token and show commission
# This mirrors the loader fallback: symbol = stem.split("_")[0].upper()
FALLBACK_CTYPE = {"ES": "MES", "NQ": "MNQ", "GC": "MGC", "CL": "CL"}
COMMISSION_RATES = {
    "MES": 1.74,
    "MNQ": 1.73,
    "MGC": 2.45,
    "NQ":  5.00,
    "GC":  5.54,
    "CL":  5.31,
}

p = Path('.')
files = sorted([f for f in p.glob('*.xlsx')] + [f for f in p.glob('*.csv')])
if not files:
    print("No .xlsx or .csv files found in cwd.")
    raise SystemExit(0)

rows = []
for f in files:
    stem = f.stem
    # tokenise by underscore or space
    token = stem.split('_')[0].split(' ')[0].upper()
    ctype = FALLBACK_CTYPE.get(token, token)
    comm = COMMISSION_RATES.get(ctype)
    rows.append((stem, token, ctype, comm))

print(f"Found {len(rows)} system files:\n")
print(f"{'Stem':45} {'Token':6} {'CType':6} {'Comm $/ct RT':>12}")
print('-'*74)
for stem, token, ctype, comm in rows:
    comm_s = f"${comm:.2f}" if comm is not None else 'N/A'
    print(f"{stem:45} {token:6} {ctype:6} {comm_s:>12}")
