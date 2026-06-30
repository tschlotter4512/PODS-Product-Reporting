#!/usr/bin/env python3
"""
update_ob_dashboard.py — Weekly OB Performance Dashboard automation
Usage:  python update_ob_dashboard.py
        python update_ob_dashboard.py --dry-run          # no GitHub push, saves preview locally
        python update_ob_dashboard.py --week 2026-06-15  # specific week (Monday date)

Snowflake: log in via SSO browser popup on first run. Okta session must be active.
Requires config.json in the same directory:
{
    "github_token": "ghp_...",
    "snowflake_user": "your.email@pods.com",
    "snowflake_account": "pods.us-east-1"
}
"""

import sys, json, re, base64, urllib.request, os
from datetime import date, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Config ─────────────────────────────────────────────────────────────────────

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.json')
LOCAL_PATH  = os.path.join(SCRIPT_DIR, 'ob_data.js')
REPO        = 'tschlotter4512/PODS-Product-Reporting'
DATA_FILE   = 'ob_data.js'
SF_WH       = 'ELT_WH'

# ── Agent roster ───────────────────────────────────────────────────────────────
# (dashboard_name, group, sf_owner_id, network_id, five9_name)
# five9_name = how the agent appears in FIVE9_CALL_DATA."AGENT NAME"
# network_id = Agent Container Booked By in ORDER_CONTAINER_SALES_FACT

AGENTS = [
    # CLW
    ('Sandra Haun',                  'CLW', '0053m00000AkMoVAAV', 'SHaun',        'Sandra Haun'),
    ('Frederick Hutchinson',         'CLW', '0053m00000AkMopAAF', 'FHutchinson',  'Frederick Hutchinson'),
    ('Christopher Monaldi',          'CLW', '0053m00000AkMozAAF', 'CMonaldi',     'Christopher Monaldi'),
    ('Michael Petersen',             'CLW', '0053m00000AkMouAAF', 'MPetersen',    'Michael Petersen'),
    ('Matthew Floming',              'CLW', '0053m000008jv3eAAA', 'MDFloming',    'Matthew Floming'),
    ('Ivy Fichtl',                   'CLW', '0053m00000AkMoBAAV', 'mfichtl',      'Ivy Fichtl'),
    ('Richard Stoffregen',           'CLW', '0053m00000DLY3SAAX', 'RJStoffregen','Richard Stoffregen'),
    ('Dawn Carreiro',                'CLW', '0053m00000AkMpOAAV', 'DCARREIRO',    'Dawn Carreiro'),
    ('Christina Morales',            'CLW', '0053m00000AkMoGAAV', 'cmorales',     'Christina Morales'),
    ('Kimberly Hewitt',              'CLW', '0053m00000AkMo6AAF', 'KHewitt',      'Kimberly Hewitt'),
    ('Dale Besida',                  'CLW', '005Hu00000Q3qA5IAJ', 'DJBesida',     'Dale Besida'),
    ('Erik Goldring',                'CLW', '005Hu00000OpsqXIAR', 'ESGoldring',   'Erik Goldring'),
    ('Rene Dunn',                    'CLW', '0053m00000AkMoQAAV', 'NDunn',        'Rene Dunn'),
    # CNX
    ('Charmaine Bacay',              'CNX', '0053m00000DCQDUAA5', 'CBacay',       'Charmaine Bacay'),
    ('Edison Viray',                 'CNX', '005Hu00000PmyDPIAZ', 'EViray',       'Edison Viray'),
    ('Gimarie Barastas',             'CNX', '005Hu00000Q2TuFIAV', 'GBarastas',    'Gimarie Barastas'),
    ('Jerico Galve',                 'CNX', '005Hu00000SZTxPIAX', 'JGalve',       'Jerico Galve'),
    ('Cristal Apple Barredo',        'CNX', '0053m00000DCQDkAAP', 'CABarredo',    'Cristal Apple Barredo'),
    ('Jessa Rey Belotindos',         'CNX', '005Hu00000Q3q8hIAB', 'JRBelotindos', 'Jessa Rey Belontindos'),
    ('Lienard Bryan Organo',         'CNX', '005Hu00000SZk9GIAT', 'LBOrgano',     'Lienard Bryan Organo'),
    ('Sendy Firmalino',              'CNX', '0053m00000DLYO4AAP', 'SFirmalino',   'Sendy Firmalino'),
    ('Ednalyn Mirandilla',           'CNX', '0053m00000DLY0nAAH', 'EMirandilla',  'Ednalyn Mirandilla'),
    ('Hannay Lozano',                'CNX', '0053m00000DCQDjAAP', 'HLozano',      'Hannay Lozano'),
    ('Ferdinand Carlo Caete',        'CNX', '005Hu00000PmyDtIAJ', 'FCCaete',      'Ferdinand Carlo Canete'),
    ('Eumarbel Dionglay',            'CNX', '005Hu00000Q3q8NIAR', 'EDionglay',    'Eumarbel Dionglay'),
    ('Milkos Malcolm Orven Miguel',  'CNX', '0053m00000DLY28AAH', 'MOMiguel',     'Milkos Malcolm Orven Miguel'),
    ('Jan Malvin Molina',            'CNX', '005Hu00000Q2Tu0IAF', 'JMMolina',     'Jan Malvin Molina'),
    ('May Anthonette Clavel',        'CNX', '005Hu00000Q3q9VIAR', 'MAClavel',     'May Anthonette Clavel'),
    ('Marielle Coleen Placido',      'CNX', '0053m00000DCQDZAA5', 'MCPlacido',    'Marielle Coleen Placido'),
    ('Jake Russel Bernardino',       'CNX', '005Hu00000PmyE3IAJ', 'JRBernardino', 'Jake Russel Bernardino'),
    ('Marijoh Mae Lee',              'CNX', '005Hu00000PmyDKIAZ', 'MMLee',        'Marijoh Mae Lee'),
    ('Jesse Vergara Jr',             'CNX', '005Hu00000SZTxOIAX', 'JVergaraJr',   'Jesse Vergara Jr'),
    ('Erwin Cablao',                 'CNX', '005Hu00000RJHO2IAP', 'ECablao',      'Erwin Cablao'),
    ('Shella Marie Francisco',       'CNX', '005Hu00000Q2TopIAF', 'SMFrancisco',  'Shella Marie Francisco'),
    ('Aldwin Navarro',               'CNX', '005Hu00000Q3fCMIAZ', 'ANavarro',     'Aldwin Navarro'),
    ('James Rebualos',               'CNX', '005Hu00000Q3q8cIAB', 'JRebualos',    'James Rebualos'),
    ('Thaddaeus Paulo Malig',        'CNX', '005Hu00000Q2TouIAF', 'TPMalig',      'Thaddaeus Paulo Malig'),
    ('Brian Valdez',                 'CNX', '0053m00000DCQDiAAP', 'BValdez',      'Brian Valdez'),
    ('Irish Verona',                 'CNX', '005Hu00000SZk94IAD', 'IVerona',      'Irish Verona'),
    ('Jerome Roger Rosites',         'CNX', '005Hu00000SZk8nIAD', 'JRRosites',    'Jerome Roger Rosites'),
    ('Ivan Marc Lanestosa',          'CNX', '005Hu00000SZloMIAT', 'IMLanestosa',  'Ivan Marc Lanestosa'),
    ('Christian jerick Sanchez',     'CNX', '005Hu00000SXJajIAH', 'CjSanchez',    'Christian jerick Sanchez'),
    ('John Rick Elumba',             'CNX', '005Hu00000SZlaqIAD', 'JRElumba',     'John Rick Elumba'),
]

# Derived lookups
BY_FIVE9   = {a[4]: a for a in AGENTS}   # five9_name → agent tuple
BY_NID     = {a[3]: a for a in AGENTS}   # network_id → agent tuple
BY_SF_ID   = {a[2]: a for a in AGENTS}   # sf_owner_id → agent tuple
BY_NAME    = {a[0]: a for a in AGENTS}   # dashboard_name → agent tuple

CLW_FIVE9  = [a[4] for a in AGENTS if a[1] == 'CLW']
CNX_FIVE9  = [a[4] for a in AGENTS if a[1] == 'CNX']

def _sql_list(items):
    return ', '.join(f"'{i}'" for i in items)

DNIS_EXCLUSIONS = (
    '2042986848','2045107637','2052813704','2057887915','2059429111','2083781676',
    '2085593679','2085733227','2085739626','2103473091','2103655251','2107403890',
    '2144554511','2144751091','2149013848','2149187791','2149520953','2169901716',
    '2259373865','2402167931','2513481178','2533029792','2566403748','2703029164',
    '2816080907','3029224720','3125500336','3155060117','3175578964','3184482424',
    '3193608003','3193937637','3213058804','3346144543','3366469413','3372087415',
    '3392277330','3392371915','3465476555','3862900325','4012255817','4037232240',
    '4043664427','404848755','4102526335','4106109026','4143168538','4232901234',
    '4234727511','4235056692','4235939474','4403649545','4407733869','4435454762',
    '4436106942','4436107747','448970003','4695221222','4699340920','4704508531',
    '4842016028','5012865022','5015171814','5025238329','5032470100','5035936436',
    '5052357329','5054293172','5073823717','5086416165','5107506965','5123504180',
    '5123926060','5127484323','5137060679','5145593288','5146662647','5154224729',
    '5182484734','5185309557','5188482383','5188616259','5189446708','5415548464',
    '5419711518','5419711682','5598592450','5632052175','5852692156','6015224385',
    '6034023738','6082791166','6099942347','6102091819','6103602453','6105594568',
    '6122806010','6143320520','6153544695','6162507830','6164774500','6177776244',
    '6177777579','6302735247','6303504911','6304298413','6304700730','6307830088',
    '6308165002','6308576525','6308620856','6313657090','6315326090','631599271',
    '6315992711','6319024393','6613486832','6782386432','6785672345','6787767920',
    '6788731997','6789901971','7023084620','7042005629','7045889011','7064646278',
    '7067559667','7077456300','7193009699','7193777368','7194375299','7205010583',
    '7205765741','7205820757','7276080170','7322211600','7325890706','7346241217',
    '7604294699','7702717901','7702717902','7703643296','7705025251','7744650933',
    '7812550464','7818087144','8033417489','8039204610','8086827637','8125497757',
    '8133829254','8165644874','8182576037','8283204585','8285452415','8285785645',
    '8287747631','8322503126','8327214954','8437358191','8452680571','8454769785',
    '8455026409','8456852828','8478447902','8502642558','8566001203','8573838992',
    '8574888844','8593918909','8644447894','8653045288','8657408102','8777707637',
    '9024011959','9042190226','9103162699','9106518980','9152692929','9162014805',
    '9192141596','9208827448','9512296725','9562271944','9738411714','9738411715',
    '9788863801','9803013505','9804067640','9807226178'
)

# ── Date helpers ───────────────────────────────────────────────────────────────

def last_week_range():
    today  = date.today()
    monday = today - timedelta(days=today.weekday() + 7)
    return monday, monday + timedelta(days=6)

def date_label(monday, sunday):
    return f"{monday.month}/{monday.day}–{sunday.month}/{sunday.day}"

def next_week_num(html):
    nums = re.findall(r'"W(\d+)"', html)
    return max(int(n) for n in nums) + 1 if nums else 72

# ── Snowflake ──────────────────────────────────────────────────────────────────

def get_conn(cfg):
    try:
        import snowflake.connector
    except ImportError:
        sys.exit("ERROR: Run: pip install snowflake-connector-python")
    return snowflake.connector.connect(
        user          = cfg['snowflake_user'],
        account       = cfg['snowflake_account'],
        authenticator = 'externalbrowser',
        warehouse     = SF_WH,
    )

def run_query(conn, sql):
    cur = conn.cursor()
    cur.execute(sql)
    cols = [c[0].lower() for c in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]

# ── SQL ────────────────────────────────────────────────────────────────────────

def sql_ob_calls(start, end):
    dnis_list = _sql_list(DNIS_EXCLUSIONS)
    agent_list = _sql_list([a[4] for a in AGENTS])
    return f"""
SELECT
    input_json:"AGENT NAME"::STRING AS agent,
    COUNT(*) AS ob_calls
FROM PROD_LAKE.FIVE9_SAAS.FIVE9_CALL_DATA
WHERE DATE(TRY_TO_TIMESTAMP(input_json:"TIMESTAMP"::STRING, 'DY, DD MON YYYY HH24:MI:SS'))
      BETWEEN '{start}' AND '{end}'
  AND input_json:"CALL TYPE"::STRING = 'Manual'
  AND input_json:"CAMPAIGN"::STRING LIKE 'DID - Resi OB%'
  AND input_json:"DNIS"::STRING NOT LIKE '599%'
  AND input_json:"DNIS"::STRING NOT IN ({dnis_list})
  AND input_json:"AGENT NAME"::STRING IN ({agent_list})
GROUP BY 1
"""

def sql_ib_calls(start, end):
    clw_list = _sql_list(CLW_FIVE9)
    cnx_list = _sql_list(CNX_FIVE9)
    return f"""
SELECT
    input_json:"AGENT NAME"::STRING AS agent,
    COUNT(*) AS ib_calls
FROM PROD_LAKE.FIVE9_SAAS.FIVE9_CALL_DATA
WHERE DATE(TRY_TO_TIMESTAMP(input_json:"TIMESTAMP"::STRING, 'DY, DD MON YYYY HH24:MI:SS'))
      BETWEEN '{start}' AND '{end}'
  AND input_json:"CALL TYPE"::STRING = 'Inbound'
  AND input_json:"HANDLE TIME"::STRING > '00:00:00'
  AND input_json:"PODS.Line of Business"::STRING != 'Service'
  AND (
      (input_json:"AGENT NAME"::STRING IN ({clw_list})
       AND (input_json:"CAMPAIGN"::STRING LIKE '%Local%'
            OR input_json:"CAMPAIGN"::STRING LIKE '% IF %'))
      OR
      (input_json:"AGENT NAME"::STRING IN ({cnx_list})
       AND input_json:"CAMPAIGN"::STRING NOT LIKE '%Local%'
       AND input_json:"CAMPAIGN"::STRING NOT LIKE '% IF %')
  )
GROUP BY 1
"""

def sql_opps(start, end):
    owner_list = _sql_list([a[2] for a in AGENTS])
    return f"""
SELECT
    o.input_json:"OwnerId"::STRING AS owner_id,
    COUNT(*) AS opps
FROM PROD_LAKE.SALESFORCE_SAAS.OPPORTUNITY_JSON o
WHERE o.input_json:"IsDeleted"::STRING = 'false'
  AND DATE(CONVERT_TIMEZONE('UTC', 'America/New_York',
      TRY_TO_TIMESTAMP_NTZ(LEFT(o.input_json:"CreatedDate"::STRING, 19))))
      BETWEEN '{start}' AND '{end}'
  AND o.input_json:"OwnerId"::STRING IN ({owner_list})
GROUP BY 1
"""

def sql_containers(start, end):
    nid_list = _sql_list([a[3] for a in AGENTS])
    return f"""
SELECT
    "Agent Container Booked By" AS nid,
    SUM("Gross Container Count") AS containers
FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
WHERE "Date Container Booked" BETWEEN '{start}' AND '{end}'
  AND "Super Segment" != 'Commercial'
  AND "Agent Container Booked By" IN ({nid_list})
GROUP BY 1
"""

def sql_cancel_7d(cancel_start, end):
    nid_list = _sql_list([a[3] for a in AGENTS])
    return f"""
SELECT
    "Agent Container Booked By" AS nid,
    SUM("Gross Container Count") AS containers_3wk,
    SUM(CASE WHEN "Date Container Cancelled" <= DATEADD(day, 7, "Date Container Booked")
             THEN "Cancelled Container Count" ELSE 0 END) AS cancelled_7d
FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
WHERE "Date Container Booked" BETWEEN '{cancel_start}' AND '{end}'
  AND "Super Segment" != 'Commercial'
  AND "Agent Container Booked By" IN ({nid_list})
GROUP BY 1
"""

# ── Fetch all data (parallel) ──────────────────────────────────────────────────

def fetch_week_data(conn, monday, sunday):
    start        = monday.isoformat()
    end          = sunday.isoformat()
    cancel_start = (sunday - timedelta(days=20)).isoformat()

    queries = {
        'ob':       sql_ob_calls(start, end),
        'ib':       sql_ib_calls(start, end),
        'opps':     sql_opps(start, end),
        'cont':     sql_containers(start, end),
        'cancel':   sql_cancel_7d(cancel_start, end),
    }

    results = {}

    def run(key, sql):
        cur = conn.cursor()
        cur.execute(sql)
        cols = [c[0].lower() for c in cur.description]
        return key, [dict(zip(cols, row)) for row in cur.fetchall()]

    print('  Running 5 queries in parallel...')
    with ThreadPoolExecutor(max_workers=5) as ex:
        futures = {ex.submit(run, k, v): k for k, v in queries.items()}
        for f in as_completed(futures):
            key, rows = f.result()
            results[key] = rows
            print(f'    ✓ {key} ({len(rows)} rows)')

    # Build lookup dicts
    ob_map     = {r['agent']: r['ob_calls'] for r in results['ob']}
    ib_map     = {r['agent']: r['ib_calls'] for r in results['ib']}
    opps_map   = {r['owner_id']: r['opps'] for r in results['opps']}
    cont_map   = {r['nid']: r['containers'] for r in results['cont']}
    cancel_map = {r['nid']: r for r in results['cancel']}

    agents_out = []
    for dash_name, group, sf_id, nid, five9_name in AGENTS:
        ob     = ob_map.get(five9_name)
        ib     = ib_map.get(five9_name)
        opps   = opps_map.get(sf_id)
        cont   = cont_map.get(nid)
        c_row  = cancel_map.get(nid, {})
        c3     = c_row.get('containers_3wk') or 0
        c7d    = c_row.get('cancelled_7d') or 0

        total_calls    = (ob or 0) + (ib or 0)
        calls_per_opp  = round(total_calls / opps, 4) if opps else None
        conversion     = round(100 * (cont or 0) / opps, 2) if opps else None
        cancel_7d      = round(100 * c7d / c3, 1) if c3 else None

        agents_out.append({
            'name':         dash_name,
            'group':        group,
            'opps':         int(opps) if opps is not None else None,
            'ob_calls':     int(ob)   if ob   is not None else None,
            'ib_calls':     int(ib)   if ib   is not None else None,
            'containers':   int(cont) if cont is not None else None,
            'calls_per_opp': calls_per_opp,
            'conversion':   conversion,
            'cancel_7d':    cancel_7d,
        })

    return agents_out

# ── GitHub helpers ─────────────────────────────────────────────────────────────

def github_get(path, token):
    url = f'https://api.github.com/repos/{REPO}/contents/{path}'
    req = urllib.request.Request(url, headers={
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json',
    })
    data = json.loads(urllib.request.urlopen(req).read())
    return base64.b64decode(data['content']).decode('utf-8'), data['sha']

def github_put(path, content, sha, token, msg):
    payload = json.dumps({
        'message': msg,
        'content': base64.b64encode(content.encode('utf-8')).decode('utf-8'),
        'sha': sha,
    }).encode('utf-8')
    req = urllib.request.Request(
        f'https://api.github.com/repos/{REPO}/contents/{path}',
        data=payload,
        headers={'Authorization': f'token {token}', 'Content-Type': 'application/json'},
        method='PUT',
    )
    return json.loads(urllib.request.urlopen(req).read())['commit']['sha']

# ── HTML update ────────────────────────────────────────────────────────────────

DATA_PAT = re.compile(r'(const D=)(\{.*?\})(;)', re.DOTALL)

def update_js(js, week_key, week_data):
    m = DATA_PAT.search(js)
    if not m:
        raise ValueError("Could not find 'const D={...}' in ob_data.js")
    existing = json.loads(m.group(2))
    existing[week_key] = week_data
    new_json = json.dumps(existing, separators=(',', ':'))
    return DATA_PAT.sub(lambda x: f'{x.group(1)}{new_json}{x.group(3)}', js)

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args    = sys.argv[1:]
    dry_run = '--dry-run' in args
    args    = [a for a in args if not a.startswith('--')]

    monday = None
    if args:
        try:
            monday = date.fromisoformat(args[0])
            if monday.weekday() != 0:
                monday -= timedelta(days=monday.weekday())
                print(f'  Adjusted to Monday: {monday}')
        except ValueError:
            sys.exit(f"ERROR: Invalid date '{args[0]}'. Use YYYY-MM-DD.")

    if monday is None:
        monday, _ = last_week_range()
    sunday = monday + timedelta(days=6)

    print(f'\nOB Dashboard update — {monday} to {sunday}')

    if not os.path.exists(CONFIG_PATH):
        sys.exit(f'ERROR: config.json not found at {CONFIG_PATH}')
    with open(CONFIG_PATH) as f:
        cfg = json.load(f)

    token = cfg.get('github_token', '')
    if not token or 'YOUR_TOKEN' in token:
        sys.exit('ERROR: github_token not set in config.json')

    # ── Step 1: Pull ob_data.js from GitHub ──
    print('\n[1] Downloading ob_data.js from GitHub...')
    js, sha  = github_get(DATA_FILE, token)
    week_num = next_week_num(js)
    week_key = f'W{week_num}'
    label    = date_label(monday, sunday)
    print(f'    Next week: {week_key} ({label})')

    # ── Step 2: Snowflake ──
    print('\n[2] Connecting to Snowflake (SSO browser will open)...')
    conn = get_conn(cfg)
    print('    Connected.')

    print(f'\n[3] Fetching data for {monday} → {sunday}...')
    agents = fetch_week_data(conn, monday, sunday)
    conn.close()

    total_opps = sum(a['opps'] or 0 for a in agents)
    total_cont = sum(a['containers'] or 0 for a in agents)
    print(f'    {len(agents)} agents | {total_opps:,} opps | {total_cont:,} containers')

    # ── Step 3: Build week object ──
    week_data = {
        'num':    week_num,
        'date':   label,
        'agents': agents,
    }

    # ── Step 4: Inject into ob_data.js ──
    print(f'\n[4] Injecting {week_key} into ob_data.js...')
    new_js = update_js(js, week_key, week_data)
    print(f'    Done — {len(new_js):,} chars')

    # ── Dry run ──
    if dry_run:
        preview = os.path.join(SCRIPT_DIR, 'ob_data_preview.js')
        with open(preview, 'w', encoding='utf-8') as f:
            f.write(new_js)
        print(f'\nDRY RUN complete — preview saved to:\n  {preview}')
        return

    # ── Step 5: Push to GitHub ──
    print('\n[5] Pushing to GitHub...')
    commit_msg = f'[auto] OB Dashboard {week_key} ({label}): {total_opps:,} opps, {total_cont:,} containers'
    new_sha = github_put(DATA_FILE, new_js, sha, token, commit_msg)
    print(f'    Committed: {new_sha[:8]}')

    # ── Step 6: Sync locally ──
    with open(LOCAL_PATH, 'w', encoding='utf-8') as f:
        f.write(new_js)
    print(f'    Synced locally: {LOCAL_PATH}')

    print(f'\nDone! https://tschlotter4512.github.io/PODS-Product-Reporting/OB_Performance_Dashboard.html')

if __name__ == '__main__':
    main()
