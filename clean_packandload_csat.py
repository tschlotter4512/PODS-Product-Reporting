#!/usr/bin/env python3
"""
clean_packandload_csat.py — Monthly Pack & Load CSAT file for P&L
Usage:  python clean_packandload_csat.py <medallia_export.xlsx|.csv> [output.xlsx]

Input is the Medallia Post Experience Survey export (residential + commercial),
filtered in Medallia to "I used a PODS recommended labor provider".

Output keeps only responses where Pack & Load was a provider for load, unload,
or both. When a different provider handled the other leg, that provider's name
and score are blanked so P&L never sees competitors. The single shared
"reason for your score" comment is kept, matching prior months' files.

The labor-provider questions are not in Snowflake (MEDALLIA_SAAS.INVITATIONS
has the columns but they are never populated), so this cannot pull from SF.
"""

import os, re, sys
import pandas as pd

PNL = re.compile(r'pack\s*(&|and)\s*load', re.I)

# Matched by header substring so line breaks / nbsp in Medallia headers don't matter.
COLS = {
    'load':        'use to load your container',
    'unload':      'use to unload your container',
    'both':        'for both loading',
    'load_score':  'company that loaded your container',
    'unload_score':'company that unloaded your container',
    'both_score':  'company that loaded & unloaded',
}


def norm(s):
    return re.sub(r'\s+', ' ', str(s).replace('\xa0', ' ')).strip().lower()


def find_cols(df):
    found = {}
    for key, needle in COLS.items():
        hits = [c for c in df.columns if needle in norm(c)]
        if not hits:
            sys.exit(f'Column not found for {key!r} (looking for "{needle}")')
        found[key] = hits[0]
    return found


def is_pnl(v):
    return pd.notna(v) and bool(PNL.search(str(v)))


def clean(df):
    c = find_cols(df)
    df = df.copy()
    keep = df[c['load']].map(is_pnl) | df[c['unload']].map(is_pnl) | df[c['both']].map(is_pnl)
    df = df[keep].copy()

    for prov, score in (('load', 'load_score'), ('unload', 'unload_score'), ('both', 'both_score')):
        other = df[c[prov]].notna() & ~df[c[prov]].map(is_pnl)
        df.loc[other, [c[prov], c[score]]] = None

    return df.reset_index(drop=True)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    df = pd.read_csv(src) if src.lower().endswith('.csv') else pd.read_excel(src)
    out = clean(df)
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(os.path.abspath(src)), 'PackandLoad_CSAT_cleaned.xlsx')
    out.to_excel(dst, index=False)
    print(f'{len(df)} responses in, {len(out)} Pack & Load responses out -> {dst}')


if __name__ == '__main__':
    main()
