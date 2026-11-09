#!/usr/bin/env python3
"""
update_opsat_dashboard.py — Weekly OPSAT dashboard (OBR slides 8-9)
Usage:  python update_opsat_dashboard.py                      # last completed week
        python update_opsat_dashboard.py --week 2026-10-19    # specific week (Monday date)
        python update_opsat_dashboard.py --from-json grid.json --week 2026-10-19

Writes OPSAT_Dashboard_<week>.html next to this script. The file is fully
self-contained (no internet needed) so it can be emailed or dropped in Teams.
It holds confidential survey results: it is git-ignored and is NOT pushed to
GitHub Pages like the other dashboards in this repo.

Snowflake: log in via SSO browser popup. Requires config.json in the same directory:
{
    "snowflake_user": "your.email@pods.com",
    "snowflake_account": "pods.us-east-1"
}
"""

import sys, json, os, html
from datetime import date, timedelta
from collections import defaultdict

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.json')
SF_WH       = 'ELT_WH'

REGIONS = ['Northeast', 'Southeast', 'Central', 'West', 'Canada']
BUCKETS = ['timing', 'driver', 'container', 'placement']
COLORS  = {'Northeast': '#E8692A', 'Southeast': '#6B1F7B', 'Central': '#1B2A9C',
           'West': '#E0409B', 'Canada': '#0F8A5F', 'PODS': '#111111'}
LY_COLORS = {'Northeast': '#F5B48F', 'Southeast': '#B583D6', 'Central': '#8EA2F0',
             'West': '#F3A6CF', 'Canada': '#8FD1B5', 'PODS': '#A8A8A8'}
BASELINES_PATH = os.path.join(SCRIPT_DIR, 'market_baselines.json')
TARGET_COLOR = '#F2B705'

# ── Targets ────────────────────────────────────────────────────────────────────
# Goals V6_NPS and YTD.xlsx, tab "AOP and YTD", OPSAT 2026 AOP row (= FPU2 row 164).
# Monthly targets drive the monthly chart; quarterly targets drive the weekly
# charts (a week belongs to the quarter containing its Thursday).
MONTHLY_TARGET = {
    2026: [84.50, 85.60, 85.30, 87.00, 86.40, 84.20, 83.50, 83.36, 84.31, 85.80, 85.80, 85.80],
}
QUARTERLY_TARGET = {
    2026: [85.11, 85.77, 83.73, 85.80],
}

# ── SQL ────────────────────────────────────────────────────────────────────────
# OPSAT = count of 5s / count of 1-5 answers (6 = N/A) across the questions below.
# "From" questions roll up to the From warehouse's region, "To" questions
# (re-delivery, final pick-up) to the To warehouse's region. Region comes from
# CONFORMED_LOCATION_DIM; Medallia's own region field still has old names.
# Canada markets are their own region (the location dim files them under West/Central).
# Region 'Z' = franchisee, excluded from PODS totals.

def coalesce(*cols):
    cols = [f'i."{c}"' for c in cols]
    return cols[0] if len(cols) == 1 else f"COALESCE({', '.join(cols)})"

QUESTIONS = {
    ('f', 'timing'): [coalesce('q_pods_move_delivery_timing_osat_5scale')],
    ('f', 'driver'): [
        coalesce('q_pods_move_driver_contact_b4_arrival_scale5_na'),
        coalesce('q_pods_move_osat_care_given_parking_lot_5scale',
                 'q_pods_move_comm_delivery_care_parkinglot_scale5_na',
                 'q_pods_move_commercial_driver_del_care_parkinglot_scale5_na'),
        coalesce('q_pods_move_residential_driver_del_care_driveway_scale5_na'),
        coalesce('q_pods_move_delivery_drvr_osat_5scale'),
        coalesce('q_pods_move_residential_driver_explain_feat_5scale_na'),
    ],
    ('f', 'container'): [
        coalesce('q_pods_move_container_condition_osat_5scale'),
        coalesce('q_pods_move_residential_driver_cleanliness_5scale_na'),
    ],
    ('f', 'placement'): [coalesce('q_pods_move_driver_container_placement_scale5_na')],
    ('t', 'timing'): [
        coalesce('q_pods_move_re_delivery_timing_osat_5scale',
                 'q_pods_re_move_delivery_timing_osat_5scale'),
        coalesce('q_pods_move_timing_pickup_osat_5scale'),
    ],
    ('t', 'driver'): [
        coalesce('q_pods_move_re_driver_contact_b4_arrival_scale5_na',
                 'q_pods_re_move_driver_contact_b4_arrival_scale5_na'),
        coalesce('q_pods_move_re_residential_driver_del_care_driveway_scale5_na',
                 'q_pods_re_move_residential_driver_del_care_driveway_scale5_na'),
        coalesce('q_pods_move_re_delivery_drvr_osat_5scale',
                 'q_pods_re_move_delivery_drvr_osat_5scale'),
    ],
    ('t', 'placement'): [
        coalesce('q_pods_move_re_driver_container_placement_scale5_na',
                 'q_pods_re_move_driver_container_placement_scale5_na'),
    ],
}


def grid_sql(start, end):
    """Answers by week x month x region x bucket for responses in [start, end)."""
    measures = []
    for (side, b), qs in QUESTIONS.items():
        five = ' + '.join(f'IFF(TRY_TO_NUMBER({q}::string)=5,1,0)' for q in qs)
        ans  = ' + '.join(f'IFF(TRY_TO_NUMBER({q}::string) BETWEEN 1 AND 5,1,0)' for q in qs)
        measures.append(f'{five} AS {side}_{b}_5,\n           {ans} AS {side}_{b}_a')
    unions = '\n      UNION ALL '.join(
        f"SELECT rdate, reg_{side} reg, mkt_{side} mkt, '{b}' b, {side}_{b}_5 s5, {side}_{b}_a sa FROM r"
        for side, b in QUESTIONS)
    return f"""
WITH loc AS (
    SELECT GL_SITE_ID, MAX(REGION_NAME) region
    FROM PROD_DW.DW_BASE.CONFORMED_LOCATION_DIM
    WHERE CURRENT_RECORD_IND::string IN ('1','Y','TRUE','true')
    GROUP BY 1
),
r AS (
    SELECT TO_DATE(TRY_TO_TIMESTAMP(i."e_responsedate")) rdate,
           IFF(i."e_pods_move_franchisename_from_txt" ILIKE '%Canada%', 'Canada', lf.region) reg_f,
           IFF(i."e_pods_move_franchisename_to_txt" ILIKE '%Canada%', 'Canada', lt.region) reg_t,
           i."e_pods_move_franchisename_from_txt" mkt_f, i."e_pods_move_franchisename_to_txt" mkt_t,
           {(','+chr(10)+'           ').join(measures)}
    FROM PROD_LAKE.MEDALLIA_SAAS.INVITATIONS i
    LEFT JOIN loc lf ON lf.GL_SITE_ID::string = i."e_pods_move_warehouseid_from_txt"::string
    LEFT JOIN loc lt ON lt.GL_SITE_ID::string = i."e_pods_move_warehouseid_to_txt"::string
    WHERE i."a_survey_program" = 'Post Experience Survey'
      AND COALESCE(i.DELETED_IND::string, 'N') <> 'Y'
      AND TO_DATE(TRY_TO_TIMESTAMP(i."e_responsedate")) >= '{start}'
      AND TO_DATE(TRY_TO_TIMESTAMP(i."e_responsedate")) <  '{end}'
),
f AS ({unions})
SELECT TO_CHAR(DATE_TRUNC('week', rdate), 'YYYY-MM-DD') wk,
       TO_CHAR(DATE_TRUNC('month', rdate), 'YYYY-MM-DD') mon,
       COALESCE(reg, '?') reg, COALESCE(mkt, '?') mkt, b, SUM(s5) s5, SUM(sa) sa
FROM f
GROUP BY 1, 2, 3, 4, 5
HAVING SUM(sa) > 0
ORDER BY 1, 2, 3, 4, 5
"""

# ── Data ───────────────────────────────────────────────────────────────────────

def fetch_snowflake(start, end):
    try:
        import snowflake.connector
    except ImportError:
        sys.exit("ERROR: Run: pip install snowflake-connector-python")
    with open(CONFIG_PATH) as fh:
        cfg = json.load(fh)
    conn = snowflake.connector.connect(
        user=cfg['snowflake_user'], account=cfg['snowflake_account'],
        authenticator='externalbrowser', warehouse=SF_WH)
    cur = conn.cursor()
    cur.execute(grid_sql(start, end))
    return [list(r) for r in cur.fetchall()]


def load_json(path):
    with open(path) as fh:
        d = json.load(fh)
    if isinstance(d, dict):
        d = d.get('result_set', d).get('data', d)
    return d


class Grid:
    """Sums of 5s / answers keyed by (week, month, region, bucket)."""

    def __init__(self, rows):
        acc = defaultdict(lambda: [0, 0])
        for wk, mon, reg, _mkt, *v in rows:
            if reg in ('Z', '?'):
                continue
            for i, b in enumerate(BUCKETS):
                a = acc[(str(wk)[:10], str(mon)[:10], reg, b)]
                a[0] += int(v[2 * i])
                a[1] += int(v[2 * i + 1])
        self.rows = [(date.fromisoformat(wk), date.fromisoformat(mon), reg, b, s5, sa)
                     for (wk, mon, reg, b), (s5, sa) in acc.items() if sa]

    def score(self, pred):
        s5 = sa = 0
        for r in self.rows:
            if pred(*r):
                s5 += r[4]
                sa += r[5]
        return (100.0 * s5 / sa if sa else None), sa


def reg_ok(region):
    return (lambda reg: True) if region == 'PODS' else (lambda reg: reg == region)

# ── Market trends (Index Trends workbook logic) ───────────────────────────────
# Mirrors "OPSAT Index Trends" tabs Market Baselines / <metric> Market Gap Trends /
# Market Trends Visual. Flat gaps, frozen σ and settings are frozen history read from
# market_baselines.json; only the 6-week actuals come from the weekly pull.

METRICS = ['opsat'] + BUCKETS
N_TREND = 20


def _score(v, m):
    if m == 'opsat':
        s5, sa = sum(v[0::2]), sum(v[1::2])
    else:
        i = BUCKETS.index(m)
        s5, sa = v[2 * i], v[2 * i + 1]
    return 100.0 * s5 / sa if sa else None


def _slope(ys):
    pts = [(x, y) for x, y in enumerate(ys, 1) if y is not None]
    if len(pts) < 2:
        return None
    mx = sum(x for x, _ in pts) / len(pts)
    my = sum(y for _, y in pts) / len(pts)
    den = sum((x - mx) ** 2 for x, _ in pts)
    return sum((x - mx) * (y - my) for x, y in pts) / den


def _trajectory(pos, head):
    if pos == 'Below LCL':
        return {'rising': 'Below LCL — recovering', 'falling': 'Below LCL — worsening'}.get(head, 'Below LCL — stalled')
    if pos == 'Near LCL':
        return {'falling': 'Approaching LCL', 'rising': 'Near LCL — pulling up'}.get(head, 'Near LCL — holding')
    if pos == 'Centered':
        return {'falling': 'In control — drifting down', 'rising': 'In control — drifting up'}.get(head, 'In control')
    if pos == 'Near UCL':
        return 'Approaching UCL' if head == 'rising' else 'Near UCL — holding'
    return {'falling': 'Above UCL — fading', 'rising': 'Above UCL — improving'}.get(head, 'Above UCL — holding')


class MarketGrid:
    def __init__(self, rows):
        self.mkt = defaultdict(lambda: [0] * 8)
        self.nat = defaultdict(lambda: [0] * 8)
        for wk, _mon, reg, mkt, *v in rows:
            w = date.fromisoformat(str(wk)[:10])
            v = [int(x) for x in v]
            a = self.mkt[(w, mkt)]
            for i in range(8):
                a[i] += v[i]
            if reg != 'Z':
                n = self.nat[w]
                for i in range(8):
                    n[i] += v[i]

    def window(self, end, mkt=None):
        """Counts for the 6 weeks ending with week `end` (national when mkt is None)."""
        tot = [0] * 8
        for k in range(6):
            w = end - timedelta(weeks=k)
            v = self.nat.get(w) if mkt is None else self.mkt.get((w, mkt))
            if v:
                for i in range(8):
                    tot[i] += v[i]
        return tot


def build_markets(mg, anchor, base):
    P = base['params']
    weeks = [anchor - timedelta(weeks=N_TREND - 1 - i) for i in range(N_TREND)]
    nat_w = {w: mg.window(w) for w in weeks}
    nat = {m: _score(nat_w[anchor], m) for m in METRICS}
    out = []
    for b in base['markets']:
        name, sig = b['market'], b['sigma']
        cur = mg.window(anchor, name)
        q = sum(cur[1::2])
        row = {'region': b['region'], 'market': name, 'answers': q, 'sigma': sig, 'm': {}}
        for m in METRICS:
            sc = _score(cur, m)
            gap = sc - nat[m] if sc is not None and nat[m] is not None else None
            trend = []
            for w in weeks:
                s, n = _score(mg.window(w, name), m), _score(nat_w[w], m)
                trend.append(round(s - n, 3) if s is not None and n is not None else None)
            flat, k = b['flat'][m], P['k'][m]
            row['m'][m] = {'score': sc, 'gap': gap, 'trend': trend, 'center': flat,
                           'ucl': flat + k * sig, 'lcl': flat - k * sig}
        out.append(row)

    med = sorted(r['answers'] for r in out)
    n = len(med)
    median = (med[n // 2] if n % 2 else (med[n // 2 - 1] + med[n // 2]) / 2) or 1
    for r in out:
        iw = (r['answers'] / median) ** 0.5 + P['lam']['opsat']
        for m in METRICS:
            d = r['m'][m]
            k, warn, tol = P['k'][m], P['warn'][m], P['slope_tol'][m]
            sig = r['sigma']
            z = (d['gap'] - d['center']) / sig if d['gap'] is not None else None
            slope = _slope(d['trend'][-6:])
            if z is None or slope is None:
                d.update(z=z, slope=slope, priority=None, trajectory='')
                continue
            pos = ('Below LCL' if z <= -k else 'Near LCL' if z <= -warn * k else
                   'Centered' if z < warn * k else 'Near UCL' if z < k else 'Above UCL')
            head = 'falling' if slope < -tol else 'rising' if slope > tol else 'flat'
            sev = max(-P['clamp'][m], min(P['clamp'][m], z + P['h'][m] * slope / sig))
            convex = (1 if sev > 0 else -1 if sev < 0 else 0) * abs(sev) ** P['p'][m]
            pri = convex * iw if r['answers'] >= P['gate'][m] else None
            d.update(z=z, slope=slope, priority=pri, trajectory=_trajectory(pos, head))
        pris = [r['m'][b]['priority'] for b in BUCKETS]
        r['focus'] = '' if any(p is None for p in pris) else BUCKETS[pris.index(min(pris))].title()
    return {'anchor': anchor.isoformat(), 'weeks': [w.isoformat() for w in weeks],
            'national': nat, 'markets': out}

# ── Calculations ───────────────────────────────────────────────────────────────

def quarter_target(wk):
    thu = wk + timedelta(days=3)
    t = QUARTERLY_TARGET.get(thu.year)
    return t[(thu.month - 1) // 3] if t else None


def build(grid, anchor, n_weeks=6, n_months=6):
    weeks = [anchor - timedelta(weeks=n_weeks - 1 - i) for i in range(n_weeks)]
    ly_weeks = [w - timedelta(weeks=52) for w in weeks]
    out = {'anchor': anchor.isoformat(), 'week_end': (anchor + timedelta(days=6)).isoformat(),
           'weeks': [w.isoformat() for w in weeks], 'ly_weeks': [w.isoformat() for w in ly_weeks],
           'weekly_target': [quarter_target(w) for w in weeks]}

    weekly = {}
    for region in ['PODS'] + REGIONS:
        ok = reg_ok(region)
        cy = [grid.score(lambda wk, mon, reg, b, *_: wk == w and ok(reg)) for w in weeks]
        ly = [grid.score(lambda wk, mon, reg, b, *_: wk == w and ok(reg)) for w in ly_weeks]
        weekly[region] = {'cy': [s for s, _ in cy], 'ly': [s for s, _ in ly],
                          'cy_n': [n for _, n in cy], 'ly_n': [n for _, n in ly]}
    out['weekly'] = weekly

    yoy = {}
    for region in ['PODS'] + REGIONS:
        ok = reg_ok(region)
        row = {}
        for b in BUCKETS + ['opsat']:
            bok = (lambda x: True) if b == 'opsat' else (lambda x, b=b: x == b)
            cy, _ = grid.score(lambda wk, mon, reg, bb, *_: wk in weeks and ok(reg) and bok(bb))
            ly, _ = grid.score(lambda wk, mon, reg, bb, *_: wk in ly_weeks and ok(reg) and bok(bb))
            row[b] = {'cy': cy, 'ly': ly, 'delta': (cy - ly) if cy is not None and ly is not None else None}
        yoy[region] = row
    out['yoy'] = yoy

    last_day = anchor + timedelta(days=6)
    months = []
    m = date(last_day.year, last_day.month, 1)
    for _ in range(n_months):
        months.insert(0, m)
        m = date(m.year - (m.month == 1), (m.month - 2) % 12 + 1, 1)
    monthly = {}
    for region in ['PODS'] + REGIONS:
        ok = reg_ok(region)
        monthly[region] = [grid.score(lambda wk, mon, reg, b, *_: mon == mm and wk <= anchor and ok(reg))[0]
                           for mm in months]
    out['months'] = [mm.isoformat() for mm in months]
    out['monthly'] = monthly
    out['monthly_target'] = [MONTHLY_TARGET.get(mm.year, [None] * 12)[mm.month - 1] for mm in months]
    return out

# ── Rendering (inline SVG, no external libraries) ─────────────────────────────

def fmt(v, nd=1):
    return '–' if v is None else f'{v:.{nd}f}'


def line_chart(series, labels, width, height, ymin=None, ymax=None, legend=True, step=None):
    """series: list of dicts {name, values, color, dash, width, tips}."""
    pad_l, pad_r, pad_t, pad_b = 38, 112 if legend else 12, 12, 34
    vals = [v for s in series for v in s['values'] if v is not None]
    lo = ymin if ymin is not None else (int(min(vals)) - 1)
    hi = ymax if ymax is not None else (int(max(vals)) + 2)
    step = step or (1 if hi - lo <= 10 else 2 if hi - lo <= 20 else 5)
    lo, hi = lo - lo % step, hi + (-hi) % step
    pw, ph = width - pad_l - pad_r, height - pad_t - pad_b
    n = len(labels)
    x = lambda i: pad_l + (pw * i / (n - 1) if n > 1 else pw / 2)
    y = lambda v: pad_t + ph * (1 - (v - lo) / (hi - lo))
    p = [f'<svg viewBox="0 0 {width} {height}" class="chart" role="img">']
    t = lo
    while t <= hi + 1e-9:
        p.append(f'<line x1="{pad_l}" x2="{pad_l + pw}" y1="{y(t):.1f}" y2="{y(t):.1f}" class="grid"/>'
                 f'<text x="{pad_l - 6}" y="{y(t) + 4:.1f}" class="ytick">{t:g}</text>')
        t += step
    for i, lab in enumerate(labels):
        p.append(f'<text x="{x(i):.1f}" y="{pad_t + ph + 18}" class="xtick">{html.escape(lab)}</text>')
    for s in series:
        pts = [(x(i), y(v)) for i, v in enumerate(s['values']) if v is not None]
        d = ' '.join(f'{"M" if k == 0 else "L"}{a:.1f},{b:.1f}' for k, (a, b) in enumerate(pts))
        dash = ' stroke-dasharray="7,5"' if s.get('dash') else ''
        p.append(f'<path d="{d}" fill="none" stroke="{s["color"]}" stroke-width="{s.get("width", 2.5)}"{dash}/>')
        if not s.get('dash'):
            for i, v in enumerate(s['values']):
                if v is None:
                    continue
                tip = s.get('tips', [None] * n)[i] or f'{s["name"]} {labels[i]}: {v:.2f}'
                p.append(f'<circle cx="{x(i):.1f}" cy="{y(v):.1f}" r="3.5" fill="{s["color"]}">'
                         f'<title>{html.escape(tip)}</title></circle>')
    if legend:
        for k, s in enumerate(series):
            ly_ = pad_t + 14 + k * 20
            dash = ' stroke-dasharray="6,4"' if s.get('dash') else ''
            p.append(f'<line x1="{width - pad_r + 14}" x2="{width - pad_r + 36}" y1="{ly_}" y2="{ly_}" '
                     f'stroke="{s["color"]}" stroke-width="3"{dash}/>'
                     f'<text x="{width - pad_r + 42}" y="{ly_ + 4}" class="legend">{html.escape(s["name"])}</text>')
    p.append('</svg>')
    return ''.join(p)


def bar_panel(title, row, color):
    vals = [row[b]['delta'] for b in BUCKETS]
    w, h, mid, scale = 360, 120, 46, 10.0
    p = [f'<div class="panel"><h3>{html.escape(title)}</h3>',
         f'<svg viewBox="0 0 {w} {h}" class="chart">']
    cw = w / 4
    for i, b in enumerate(BUCKETS):
        cx = cw * i + cw / 2
        p.append(f'<text x="{cx:.0f}" y="12" class="bucket">{b.upper()}</text>')
        v = vals[i]
        if v is None:
            continue
        bh = min(abs(v), 6) * scale
        top = mid if v < 0 else mid - bh
        tip = (f'{title} {b}: {row[b]["cy"]:.1f} this year vs {row[b]["ly"]:.1f} last year '
               f'({v:+.1f} pts)')
        p.append(f'<rect x="{cx - 22:.0f}" y="{top:.1f}" width="44" height="{bh:.1f}" fill="{color}">'
                 f'<title>{html.escape(tip)}</title></rect>')
        ty = top + bh + 15 if v < 0 else top - 5
        p.append(f'<text x="{cx:.0f}" y="{ty:.1f}" class="barval">{v:+.1f}</text>')
    p.append(f'<line x1="4" x2="{w - 4}" y1="{mid}" y2="{mid}" class="axis"/>')
    for i in range(1, 4):
        p.append(f'<line x1="{cw * i:.0f}" x2="{cw * i:.0f}" y1="20" y2="{h - 4}" class="sep"/>')
    p.append('</svg>')
    o = row['opsat']
    p.append(f'<div class="sub">OPSAT {fmt(o["cy"])} vs {fmt(o["ly"])} LY '
             f'(<b class="{"neg" if (o["delta"] or 0) < 0 else "pos"}">{fmt(o["delta"]) if o["delta"] is None else f"{o["delta"]:+.1f}"}</b>)</div>')
    p.append('</div>')
    return ''.join(p)


def markets_section(md):
    payload = json.dumps(md, separators=(',', ':')).replace('</', '<\\/')
    return MARKETS_HTML.replace('__DATA__', payload).replace(
        '__ANCHOR__', f'{date.fromisoformat(md["anchor"]) - timedelta(weeks=5):%m/%d/%Y} – '
                      f'{date.fromisoformat(md["anchor"]) + timedelta(days=6):%m/%d/%Y}')


MARKETS_HTML = r"""
<section id="markets">
  <h2>OPSAT – Market trends</h2>
  <p class="note">Last 6 weeks (__ANCHOR__). Each market is judged against its own normal gap to national
  (2023–2025) with limits of ±k·σ that move with national. Priority: more negative = needs attention sooner
  (blank = under the minimum answer count). Click a market to see its charts: black = market's 6-week rolling gap
  vs national, gray = national, yellow dotted = the market's normal gap, green / red dashed = upper / lower limit.</p>
  <div class="controls">
    <label>Region <select id="mk-region">
      <option value="">All (excl. franchisee)</option><option>Northeast</option><option>Southeast</option>
      <option>Central</option><option>West</option><option>Canada</option><option>Franchisee</option></select></label>
    <label>Sort by <select id="mk-sort">
      <option value="opsat">OPSAT</option><option value="timing">Timing</option><option value="driver">Driver</option>
      <option value="container">Container</option><option value="placement">Placement</option></select></label>
    <span id="mk-count" class="note"></span>
  </div>
  <div class="mk-split">
    <div class="panel mk-list"><table id="mk-table"><thead><tr><th>Market</th><th>Answers</th>
      <th>OPSAT</th><th>Timing</th><th>Driver</th><th>Container</th><th>Placement</th><th>Focus area</th></tr>
      <tr class="subh"><th></th><th>6 wks</th><th colspan="5">priority · trajectory (last 6 wks)</th><th></th></tr></thead><tbody></tbody></table></div>
    <div class="mk-detail">
      <div class="panel"><h3 id="mk-title"></h3><div id="mk-facts" class="note"></div></div>
      <div id="mk-charts"></div>
    </div>
  </div>
</section>
<script>
(function () {
  var D = __DATA__;
  var M = ['opsat', 'timing', 'driver', 'container', 'placement'];
  var NAME = {opsat: 'OPSAT', timing: 'Timing', driver: 'Driver', container: 'Container', placement: 'Placement'};
  var FOCUS = {Timing: '#FDE2CF', Driver: '#E4D4F4', Container: '#D3DCFB', Placement: '#FAD3E7'};
  var labels = D.weeks.map(function (w) { var p = w.split('-'); return p[1] + '/' + p[2]; });
  var selected = null;
  function f(v, n) { return v === null || v === undefined ? '–' : v.toFixed(n === undefined ? 1 : n); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  function shade(p) {
    if (p === null) return '#f2f2f2';
    var t = Math.min(1, Math.sqrt(Math.abs(p)) / 8);
    return p < 0 ? 'rgba(215,25,32,' + (0.08 + 0.6 * t) + ')' : 'rgba(46,125,50,' + (0.08 + 0.5 * t) + ')';
  }
  function trajClass(t) {
    return /Below LCL/.test(t) ? 'neg b' : /Approaching LCL|Near LCL/.test(t) ? 'neg' :
           /Above UCL/.test(t) ? 'pos b' : /UCL/.test(t) ? 'pos' : '';
  }
  function rows() {
    var reg = document.getElementById('mk-region').value, s = document.getElementById('mk-sort').value;
    var r = D.markets.filter(function (m) { return reg ? m.region === reg : m.region !== 'Franchisee'; });
    r.sort(function (a, b) {
      var x = a.m[s].priority, y = b.m[s].priority;
      if (x === null && y === null) return a.market < b.market ? -1 : 1;
      if (x === null) return 1; if (y === null) return -1; return x - y;
    });
    return r;
  }
  function renderTable() {
    var r = rows(), h = '';
    r.forEach(function (m) {
      h += '<tr data-m="' + esc(m.market) + '"' + (m.market === selected ? ' class="sel"' : '') + '><th>' + esc(m.market) +
           '<div class="rg">' + esc(m.region) + '</div></th><td>' + m.answers.toLocaleString() + '</td>';
      M.forEach(function (k) {
        var d = m.m[k];
        h += '<td class="mc" style="background:' + shade(d.priority) + '"><b>' + f(d.priority) + '</b><div class="traj ' +
             trajClass(d.trajectory) + '">' + esc(d.trajectory || '–') + '</div></td>';
      });
      h += '<td><span class="pill" style="background:' + (FOCUS[m.focus] || '#eee') + '">' + esc(m.focus || '–') + '</span></td></tr>';
    });
    document.querySelector('#mk-table tbody').innerHTML = h;
    document.getElementById('mk-count').textContent = r.length + ' markets';
    if (r.length && !r.some(function (m) { return m.market === selected; })) select(r[0].market);
  }
  function chart(d, title) {
    var W = 520, H = 230, L = 40, R = 12, T = 10, B = 30, n = labels.length;
    var vals = d.trend.filter(function (v) { return v !== null; }).concat([0, d.ucl, d.lcl]);
    var lo = Math.floor(Math.min.apply(null, vals) - 1), hi = Math.ceil(Math.max.apply(null, vals) + 1);
    var step = hi - lo <= 12 ? 2 : hi - lo <= 30 ? 5 : 10;
    lo = Math.floor(lo / step) * step; hi = Math.ceil(hi / step) * step;
    var x = function (i) { return L + (W - L - R) * i / (n - 1); };
    var y = function (v) { return T + (H - T - B) * (1 - (v - lo) / (hi - lo)); };
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" class="chart">';
    for (var t = lo; t <= hi + 1e-9; t += step)
      s += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t) + '" y2="' + y(t) + '" class="grid"/><text x="' + (L - 6) +
           '" y="' + (y(t) + 4) + '" class="ytick">' + t + '</text>';
    labels.forEach(function (lab, i) { if (i % 2 === 1 || i === n - 1) s += '<text x="' + x(i) + '" y="' + (H - 10) + '" class="xtick">' + lab + '</text>'; });
    function flat(v, color, dash, name) {
      return '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="' + color +
             '" stroke-width="2"' + (dash ? ' stroke-dasharray="' + dash + '"' : '') + '><title>' + name + ': ' + f(v, 2) + '</title></line>';
    }
    s += flat(0, '#9e9e9e', '', 'National (0)') + flat(d.ucl, '#2E7D32', '6,4', 'UCL') +
         flat(d.lcl, '#C62828', '6,4', 'LCL') + flat(d.center, '#F2B705', '2,3', 'Center (normal gap)');
    var p = '', first = true;
    d.trend.forEach(function (v, i) { if (v === null) { first = true; return; } p += (first ? 'M' : 'L') + x(i).toFixed(1) + ',' + y(v).toFixed(1); first = false; });
    s += '<path d="' + p + '" fill="none" stroke="#111" stroke-width="2.5"/>';
    d.trend.forEach(function (v, i) {
      if (v === null) return;
      var c = v < d.lcl ? '#C62828' : v > d.ucl ? '#2E7D32' : '#111';
      s += '<circle cx="' + x(i) + '" cy="' + y(v) + '" r="3.5" fill="' + c + '"><title>Wk ' + labels[i] + ': ' + f(v, 2) + ' pts vs national</title></circle>';
    });
    s += '</svg>';
    return '<div class="panel"><h3>' + title + ' <span class="' + trajClass(d.trajectory) + '">' + esc(d.trajectory || '') +
           '</span></h3>' + s + '<div class="sub">6-week rolling gap vs national · center ' + f(d.center) + ' · limits ' +
           f(d.lcl) + ' to ' + f(d.ucl) + ' · priority ' + f(d.priority) + '</div></div>';
  }
  function select(name) {
    selected = name;
    var m = D.markets.filter(function (x) { return x.market === name; })[0];
    document.querySelectorAll('#mk-table tbody tr').forEach(function (tr) { tr.classList.toggle('sel', tr.getAttribute('data-m') === name); });
    document.getElementById('mk-title').textContent = m.market + ' (' + m.region + ')';
    var o = m.m.opsat;
    document.getElementById('mk-facts').innerHTML = 'OPSAT last 6 weeks <b>' + f(o.score) + '</b> vs national ' + f(D.national.opsat) +
      ' (gap ' + f(o.gap) + ', normal ' + f(o.center) + ') · ' + m.answers.toLocaleString() + ' answers · σ ' + f(m.sigma, 2) +
      ' · focus area <b>' + esc(m.focus || '–') + '</b>';
    document.getElementById('mk-charts').innerHTML = M.map(function (k) { return chart(m.m[k], NAME[k]); }).join('');
  }
  document.getElementById('mk-region').addEventListener('change', renderTable);
  document.getElementById('mk-sort').addEventListener('change', renderTable);
  document.querySelector('#mk-table tbody').addEventListener('click', function (e) {
    var tr = e.target.closest('tr'); if (tr) select(tr.getAttribute('data-m'));
  });
  renderTable();
})();
</script>
"""


def wk_label(iso):
    d = date.fromisoformat(iso)
    return f'Wk {d:%m/%d}'


def render(d, md=None):
    anchor = date.fromisoformat(d['anchor'])
    iso_wk = anchor.isocalendar()[1]
    labels = [wk_label(w) for w in d['weeks']]
    tgt = d['weekly_target']

    main = line_chart(
        [{'name': r, 'values': d['weekly'][r]['cy'], 'color': COLORS[r],
          'tips': [f'{r} {labels[i]}: {v:.2f} ({n:,} answers)' if v is not None else None
                   for i, (v, n) in enumerate(zip(d['weekly'][r]['cy'], d['weekly'][r]['cy_n']))]}
         for r in REGIONS] +
        [{'name': 'Target', 'values': tgt, 'color': TARGET_COLOR, 'dash': True}],
        labels, 620, 400)

    smalls = []
    for r in ['PODS'] + REGIONS:
        w = d['weekly'][r]
        ly_labels = [wk_label(x) for x in d['ly_weeks']]
        chart = line_chart(
            [{'name': str(anchor.year - 1), 'values': w['ly'], 'color': LY_COLORS[r],
              'tips': [f'{r} {ly_labels[i]}/{anchor.year - 1}: {fmt(v, 2)}' for i, v in enumerate(w['ly'])]},
             {'name': str(anchor.year), 'values': w['cy'], 'color': COLORS[r], 'width': 3},
             {'name': 'Target', 'values': tgt, 'color': TARGET_COLOR, 'dash': True}],
            labels, 420, 230, ymin=70, ymax=95, step=5)
        cur, prev = w['cy'][-1], w['ly'][-1]
        badge = '' if cur is None or prev is None else \
            f'<span class="{"neg" if cur < prev else "pos"}">{cur:.1f} ({cur - prev:+.1f} YoY)</span>'
        smalls.append(f'<div class="panel"><h3>{"OPSAT – PODS" if r == "PODS" else r} {badge}</h3>{chart}</div>')

    mlabels = [date.fromisoformat(m).strftime('%b') for m in d['months']]
    mlabels[-1] += ' MTD'
    monthly = line_chart(
        [{'name': r, 'values': d['monthly'][r], 'color': COLORS[r]} for r in REGIONS] +
        [{'name': 'Target', 'values': d['monthly_target'], 'color': TARGET_COLOR, 'dash': True}],
        mlabels, 620, 400)

    bars = ''.join(bar_panel('PODS' if r == 'PODS' else r, d['yoy'][r], COLORS[r]) for r in ['PODS'] + REGIONS)

    return TEMPLATE.format(
        title=f'OPSAT – Week {iso_wk} ({anchor:%b %d} – {date.fromisoformat(d["week_end"]):%b %d, %Y})',
        main=main, smalls=''.join(smalls), monthly=monthly, bars=bars,
        markets=markets_section(md) if md else '',
        markets_nav='<a href="#markets">Market trends</a>' if md else '',
        ly_range=f'{date.fromisoformat(d["ly_weeks"][0]):%m/%d/%Y} – '
                 f'{date.fromisoformat(d["ly_weeks"][-1]) + timedelta(days=6):%m/%d/%Y}',
        cy_range=f'{date.fromisoformat(d["weeks"][0]):%m/%d/%Y} – {date.fromisoformat(d["week_end"]):%m/%d/%Y}',
        generated=date.today().strftime('%b %d, %Y'))


TEMPLATE = """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
  :root {{ --red:#D71920; --ink:#1d1d1f; --muted:#6e6e73; --line:#e5e5ea; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; font-family:-apple-system,"Segoe UI",Roboto,Arial,sans-serif; color:var(--ink); background:#f5f5f7; }}
  header {{ background:#fff; border-bottom:4px solid var(--red); padding:18px 28px; display:flex; justify-content:space-between; align-items:baseline; flex-wrap:wrap; gap:8px; }}
  header h1 {{ margin:0; font-size:22px; }}
  header .meta {{ color:var(--muted); font-size:13px; }}
  nav {{ background:#fff; padding:0 28px; border-bottom:1px solid var(--line); display:flex; gap:24px; position:sticky; top:0; z-index:2; }}
  nav a {{ padding:12px 0; color:var(--muted); text-decoration:none; font-weight:600; font-size:14px; border-bottom:3px solid transparent; }}
  nav a:hover {{ color:var(--red); border-color:var(--red); }}
  main {{ padding:24px 28px 48px; max-width:1500px; margin:0 auto; }}
  section {{ margin-bottom:40px; }}
  section > h2 {{ color:var(--red); font-size:24px; margin:0 0 4px; }}
  section > .note {{ color:var(--muted); font-size:13px; margin:0 0 16px; }}
  .split {{ display:grid; grid-template-columns:minmax(380px,1fr) minmax(420px,1.3fr); gap:20px; align-items:start; }}
  .quad {{ display:grid; grid-template-columns:repeat(2,1fr); gap:16px; }}
  .quad .panel:first-child, .quad .panel:nth-child(6) {{ grid-column:1 / -1; max-width:50%; margin:0 auto; width:100%; }}
  .controls {{ display:flex; gap:18px; align-items:center; margin:0 0 12px; font-size:14px; font-weight:600; }}
  .controls select {{ margin-left:6px; padding:4px 6px; font-size:14px; }}
  .mk-split {{ display:grid; grid-template-columns:minmax(0,1.7fr) minmax(380px,1fr); gap:20px; align-items:start; }}
  .mk-list {{ overflow:auto; max-height:1250px; padding:0; }}
  .mk-list table {{ margin:0; font-size:12px; }}
  .mk-list thead th {{ position:sticky; top:0; background:#fff; z-index:1; text-align:center; font-size:11px; vertical-align:bottom; }}
  .mk-list tbody tr {{ cursor:pointer; }} .mk-list tbody tr:hover th {{ color:var(--red); }}
  .mk-list tbody tr.sel {{ outline:2px solid var(--red); outline-offset:-2px; }}
  .mk-list td, .mk-list th {{ white-space:nowrap; }} .mk-list td.mc {{ text-align:left; }}
  .mk-list td.mc {{ white-space:normal; min-width:96px; max-width:118px; padding:4px 6px; }}
  .mk-list .traj {{ font-size:10.5px; line-height:1.2; font-weight:500; color:#555; }} .mk-list .traj.neg {{ color:#C62828; }} .mk-list .traj.pos {{ color:#2E7D32; }}
  .mk-list .rg {{ font-size:10.5px; font-weight:500; color:var(--muted); }} .mk-list tr.subh th {{ font-weight:500; color:var(--muted); top:28px; }}
  .b {{ font-weight:700; }} .pill {{ padding:2px 8px; border-radius:10px; font-weight:600; }}
  .mk-detail {{ display:grid; gap:12px; position:sticky; top:52px; }} #mk-charts {{ display:grid; gap:12px; }}
  .panel {{ background:#fff; border-radius:10px; padding:14px 16px; box-shadow:0 1px 3px rgba(0,0,0,.08); }}
  .panel h3 {{ margin:0 0 6px; font-size:15px; display:flex; justify-content:space-between; gap:8px; }}
  .panel h3 span {{ font-size:13px; font-weight:600; }}
  .sub {{ color:var(--muted); font-size:13px; text-align:center; margin-top:4px; }}
  .chart {{ width:100%; height:auto; display:block; }}
  .chart .grid {{ stroke:var(--line); }}
  .chart .axis {{ stroke:#111; stroke-width:2; }}
  .chart .sep {{ stroke:var(--line); }}
  .chart text {{ font-size:11px; fill:var(--muted); }}
  .chart .ytick {{ text-anchor:end; }}
  .chart .xtick {{ text-anchor:middle; }}
  .chart .legend {{ fill:var(--ink); font-size:12px; }}
  .chart .bucket {{ text-anchor:middle; font-weight:700; fill:var(--ink); font-size:10px; letter-spacing:.5px; }}
  .chart .barval {{ text-anchor:middle; font-weight:700; fill:var(--ink); font-size:12px; }}
  .neg {{ color:#C62828; }} .pos {{ color:#2E7D32; }}
  table {{ border-collapse:collapse; width:100%; font-size:13px; margin-top:12px; background:#fff; }}
  th, td {{ padding:6px 8px; text-align:right; border-bottom:1px solid var(--line); }}
  th:first-child {{ text-align:left; }}
  tr.tgt td, tr.tgt th {{ color:#9a7400; font-weight:600; }}
  footer {{ color:var(--muted); font-size:12px; padding:0 28px 28px; max-width:1500px; margin:0 auto; }}
  @media (max-width:900px) {{ .split, .quad, .mk-split {{ grid-template-columns:1fr; }} .quad .panel:first-child {{ max-width:100%; }} }}
  @media print {{ nav {{ display:none; }} section {{ break-inside:avoid; }} body {{ background:#fff; }} }}
</style></head>
<body>
<header><h1>{title}</h1><div class="meta">Confidential and proprietary · Generated {generated}</div></header>
<nav><a href="#weekly">Last 6 weeks YoY</a><a href="#regional">Regional YoY performance</a>{markets_nav}</nav>
<main>
<section id="weekly">
  <h2>OPSAT – Last 6 weeks, YoY</h2>
  <p class="note">{cy_range} vs same weeks last year ({ly_range}). Target = quarterly AOP. Hover any point for the value.</p>
  <div class="split">
    <div class="panel"><h3>OPSAT – By Region, Weekly</h3>{main}</div>
    <div class="quad">{smalls}</div>
  </div>
</section>
<section id="regional">
  <h2>OPSAT – Last 6 weeks Regional YoY Performance</h2>
  <p class="note">Left: OPSAT by region by month (current month is month-to-date), target = monthly AOP.
  Right: change in % top-box by question group, last 6 weeks vs same weeks last year (points).</p>
  <div class="split">
    <div class="panel"><h3>OPSAT – By Region, MTD</h3>{monthly}</div>
    <div class="quad">{bars}</div>
  </div>
</section>
{markets}
</main>
<footer>OPSAT = share of 5s among 1–5 answers (N/A excluded) across Post Experience Survey timing, driver, container
and placement questions. Weeks run Monday–Sunday by survey response date. Delivery questions count toward the
origin market's region; re-delivery and final pick-up questions toward the destination market's region.
Canada is reported as its own region. Franchisee markets excluded from PODS. Market trends follow the OPSAT Index Trends
workbook (frozen baselines and σ in market_baselines.json). Source: Medallia via Snowflake.</footer>
</body></html>
"""

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args = sys.argv[1:]
    anchor = None
    src = None
    if '--week' in args:
        anchor = date.fromisoformat(args[args.index('--week') + 1])
    if '--from-json' in args:
        src = args[args.index('--from-json') + 1]
    if anchor is None:
        today = date.today()
        anchor = today - timedelta(days=today.weekday() + 7)
    if anchor.weekday() != 0:
        sys.exit('ERROR: --week must be a Monday')

    start = anchor - timedelta(weeks=57)
    end = anchor + timedelta(days=7)
    if '--print-sql' in args:
        print(grid_sql(start, end))
        return
    rows = load_json(src) if src else fetch_snowflake(start, end)
    data = build(Grid(rows), anchor)
    md = None
    if os.path.exists(BASELINES_PATH):
        with open(BASELINES_PATH) as fh:
            md = build_markets(MarketGrid(rows), anchor, json.load(fh))
    out = os.path.join(SCRIPT_DIR, f'OPSAT_Dashboard_{anchor.isoformat()}.html')
    with open(out, 'w') as fh:
        fh.write(render(data, md))
    p = data['weekly']['PODS']['cy']
    print(f'PODS weekly OPSAT: {", ".join(fmt(v, 2) for v in p)}')
    print(f'Wrote {out}')


if __name__ == '__main__':
    main()
