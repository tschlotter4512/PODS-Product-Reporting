#!/usr/bin/env python3
"""
update_il_dashboard.py
Weekly automation for the Integrated Labor dashboard.
Run every Monday after downloading HaH Quotes and Orders CSVs.

Also generates a per-agent performance CSV (period + all-time quotes/orders).
Run twice a week: Monday (covers prior Fri-Sun) and Thursday (covers this Mon-Thu).
"""

import json, base64, re, os, sys
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import requests
import snowflake.connector
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# ── Agent performance rating thresholds (Period Orders) ────────────────────────
PERF_GREEN_MIN = 2   # 2+ orders in the period = green
PERF_YELLOW_MIN = 1  # 1 order = yellow, 0 = red

# ── Config ─────────────────────────────────────────────────────────────────────
CONFIG_PATH = r"F:\Revenue Management\Team Individual Folders\TSchlotter\PODS-Product-Reporting\config.json"
LOCAL_DIR   = r"F:\Revenue Management\Team Individual Folders\TSchlotter\PODS-Product-Reporting"
GITHUB_REPO = "tschlotter4512/PODS-Product-Reporting"
LAUNCH_DATE = "2026-05-01"
ECR                = 3700   # expected container revenue ($)
# How far a PODS order may sit from the HaH labor booking and still be considered the
# same job. Repeat customers have old orders on the same CID — without a cap, a labor
# job can match an order booked years earlier (gaps of 1036 and 1496 days observed),
# which then measures the WRONG order's cancellation. 83% of matches land within a
# week. Set to 30 for now; tighten to 7 once volume supports it.
MATCH_WINDOW_DAYS  = 30
# Headcount assumed for the full-rollout scenario. This is an ASSUMPTION, not a
# measurement — it drives projected_rollout_savings on IL_Internal_Opportunity_Wins.html
# (removed from the main dashboard 2026-08-03 because the projection multiplies an
# effect measured on ~11 containers by ~30x). Change it here, nowhere else.
ROLLOUT_TARGET_AGENTS = 400
AVG_CTR_PER_ORDER  = 1.3    # avg containers per order — ONLY for weights that are
                            # order counts (legacy Tableau path). The Snowflake cancel
                            # rate weights by actual containers, so it must NOT be used there.

# ── Agent roster: dept + location for each trained agent ──────────────────────
DEPT_MAP = {
    "Aiko Noquillo":               ("SST",     "CNX"),
    "Aldwin Navarro":              ("OB",      "CNX"),
    "Almira Castro":               ("SST",     "CNX"),
    "Amy Sisterman":               ("Sales",   "CLW"),
    "Anna Kathrisse Serrano":      ("SST",     "CNX"),
    "Annalee Bueno":               ("SST",     "CNX"),
    "Brian Valdez":                ("OB",      "CNX"),
    "Carmie Ann Gorospe":          ("SST",     "CNX"),
    "Carmie Nillama":              ("SST",     "CNX"),
    "Carol Brown":                 ("Sales",   "CLW"),
    "Catherine Saitta":            ("Sales",   "CLW"),
    "Charmaine Bacay":             ("OB",      "CNX"),
    "Christian Jerick Sanchez":    ("OB",      "CNX"),
    "Christina Morales":           ("OB",      "CLW"),
    "Christopher Monaldi":         ("OB",      "CLW"),
    "Christopher Washington":      ("Sales",   "CLW"),
    "Cristal Apple Barredo":       ("OB",      "CNX"),
    "Dale Besida":                 ("OB",      "CLW"),
    "Dawn Carreiro":               ("OB",      "CLW"),
    "Debra Springfield":           ("Sales",   "CLW"),
    "Dexter Weaver":               ("Sales",   "CLW"),
    "Donald Alfaro":               ("SST",     "CLW"),
    "Edison Viray":                ("OB",      "CNX"),
    "Ednalyn Mirandilla":          ("OB",      "CNX"),
    "Erik Goldring":               ("OB",      "CLW"),
    "Erwin Cablao":                ("OB",      "CNX"),
    "Eumarbel Dionglay":           ("OB",      "CNX"),
    "Felicia Knight":              ("SST",     "CLW"),
    "Ferdinand Carlo Canete":      ("OB",      "CNX"),
    "Fern Vallido":                ("Service", "—"),
    "Frederick Hutchinson":        ("OB",      "CLW"),
    "Gimarie Barastas":            ("OB",      "CNX"),
    "Gregory D. Bell":             ("Sales",   "CLW"),
    "Hannay Lozano":               ("OB",      "CNX"),
    "Irish Verona":                ("OB",      "CNX"),
    "Ivan Marc Lanestosa":         ("OB",      "CNX"),
    "Ivy Fichtl":                  ("OB",      "CLW"),
    "Jake Russel Bernardino":      ("OB",      "CNX"),
    "James Onnagan":               ("SST",     "CNX"),
    "James Rebualos":              ("OB",      "CNX"),
    "Jan Malvin Molina":           ("OB",      "CNX"),
    "Jeanielle Navarro":           ("SST",     "CNX"),
    "Jerico Galve":                ("OB",      "CNX"),
    "Jerome Roger Rosites":        ("OB",      "CNX"),
    "Jessa Rey Belontindos":       ("OB",      "CNX"),
    "Jesse Vergara Jr":            ("OB",      "CNX"),
    "Johnny Cardinale":            ("SST",     "CLW"),
    "Justin Brown":                ("Sales",   "CLW"),
    "Katleen May Castro":          ("Sales",   "CNX"),
    "Kimberly Hewitt":             ("OB",      "CLW"),
    "Kimberly Thebeau":            ("Sales",   "CLW"),
    "Lienard Bryan Organo":        ("OB",      "CNX"),
    "Maria Concepcion Cauguiran":  ("SST",     "CNX"),
    "Marielle Coleen Placido":     ("OB",      "CNX"),
    "Marijoh Mae Lee":             ("OB",      "CNX"),
    "Mary France Rodriguez":       ("OB",      "CNX"),
    "Matthew Floming":             ("OB",      "CLW"),
    "Matthew Hepburn":             ("Sales",   "CLW"),
    "May Anthonette Clavel":       ("OB",      "CNX"),
    "Mica Mae Maniego":            ("SST",     "CNX"),
    "Micah Joy Abrao":             ("SST",     "CNX"),
    "Michael Petersen":            ("OB",      "CLW"),
    "Milkos Malcolm Orven Miguel": ("OB",      "CNX"),
    "Nicole Slater":               ("Sales",   "CLW"),
    "Princess Heidi Sesma":        ("Service", "CNX"),
    "Rene Dunn":                   ("OB",      "CLW"),
    "Ria Francisco":               ("Sales",   "CNX"),
    "Richard Stoffregen":          ("OB",      "CLW"),
    "Robert Dusenberry":           ("Sales",   "CLW"),
    "Ryan Dale Aldave":            ("SST",     "CNX"),
    "Sandra Haun":                 ("OB",      "CLW"),
    "Sara Messenger":              ("SST",     "CLW"),
    "Sendy Firmalino":             ("OB",      "CNX"),
    "Shannon Glendye":             ("Sales",   "CLW"),
    "Shella Marie Francisco":      ("OB",      "CNX"),
    "Syralee Mae Gamil":           ("SST",     "CNX"),
    "Ted Johnson":                 ("Sales",   "CLW"),
    "Thaddaeus Paulo Malig":       ("OB",      "CNX"),
    "Viviane Peux":                ("SST",     "CLW"),
    # Michelle Fallahee's team (added 2026-07-20)
    "Shayla Brown":                ("Sales",   "CLW"),
    "Ashley Burrell":              ("Sales",   "CLW"),
    "Christopher Carlesimo":       ("Sales",   "CLW"),
    "James P. Dempsey Jr.":        ("Sales",   "CLW"),
    "Heidi Gathen":                ("Sales",   "CLW"),
    "Tracy Heath":                 ("Sales",   "CLW"),
    "Henssy Henriquez":            ("Sales",   "CLW"),
    "Nathaniel Herden":            ("Sales",   "CLW"),
    "Wendy Horowitz":              ("Sales",   "CLW"),
    "Seiha Ly":                    ("Sales",   "CLW"),
    "Gavin Novak":                 ("Sales",   "CLW"),
    "Brian C. Quinn":              ("Sales",   "CLW"),
    "LaVerne Reed":                ("Sales",   "CLW"),
    "Cheryl Rosa":                 ("Sales",   "CLW"),
    "Jalmise Victor":              ("Sales",   "CLW"),
    # Jon'tae Ferguson's team (added 2026-07-20)
    "Michael Alcalde":             ("Sales",   "CLW"),
    "LaQuishia Brown":             ("Sales",   "CLW"),
    "Alexander Cook":              ("Sales",   "CLW"),
    "Gennaro Guardascione":        ("Sales",   "CLW"),
    "Joy Hall":                    ("Sales",   "CLW"),
    "James Jackson":               ("Sales",   "CLW"),
    "Kasey M. Lockhart":           ("Sales",   "CLW"),
    "Donavan O. Mitchell":         ("Sales",   "CLW"),
    "Justin Myers Singh":          ("Sales",   "CLW"),
    "Kelvin Pearce":               ("Sales",   "CLW"),
    "Jordan Rosario Garcia":       ("Sales",   "CLW"),
    "Scott Scalzo":                ("Sales",   "CLW"),
    "Alfred Thornton":             ("Sales",   "CLW"),
    "Blake Waters":                ("Sales",   "CLW"),
    # Remaining "Other new agents" (added 2026-07-20)
    "Stephen Hardwick":            ("Sales",   "CLW"),
    "Alton Nelson":                ("Sales",   "CLW"),
    "Lynna Stites":                ("Sales",   "CLW"),
    "Scot Horowitz":               ("Sales",   "CLW"),
    "Valerie Mabe":                ("Sales",   "CLW"),
    "Leanne Souza":                ("Sales",   "CLW"),
    "Henry Noble":                 ("Sales",   "CLW"),
    "Jorden N. Murray":            ("Sales",   "CLW"),
    "Tanisha Lee":                 ("Sales",   "CLW"),
    "Latoya Jones":                ("Sales",   "CLW"),
    "James Meckfessel":            ("Sales",   "CLW"),
    "Charles Chambliss":           ("Sales",   "CLW"),
    "Amy Ashmore":                 ("Sales",   "CLW"),
    "Sakeema McCray":              ("Sales",   "CLW"),
    "Patricia Paino":              ("Sales",   "CLW"),
    "Gentian Grezda":              ("Sales",   "CLW"),
    "Arlene Wethern":              ("Sales",   "CLW"),
    "Frank Miedzianowski":         ("Sales",   "CLW"),
    "Madeline Quiros":             ("Sales",   "CLW"),
    "Caroline McMahon":            ("Sales",   "CLW"),
    "Scott Schneider":             ("Sales",   "CLW"),
}

# ── Helpers ────────────────────────────────────────────────────────────────────
# ── Agent username map: canonical name → ORDER_CONTAINER_SALES_FACT "Agent Container Booked By" ──
# Derived at runtime (see build_agent_username_map) from the Verint employee directory's
# Email field — "smessenger@pods.com" -> "SMessenger", which is exactly the username
# Snowflake's "Agent Container Booked By" column uses. Confirmed via spot-check that all
# 129 current roster agents' derived usernames actually appear booking orders in
# ORDER_CONTAINER_SALES_FACT. This used to be a hand-maintained list of ~23 agents (built
# for the pitch-rate calc only) — that undercoverage is why it's now derived for the full
# roster instead.
def build_agent_username_map(directory, roster_names):
    username_map = {}
    for name in roster_names:
        info = directory.get(name)
        if not info:
            m = fuzzy_match(name, list(directory.keys()))
            info = directory.get(m) if m else None
        email = info.get("email") if info else None
        if email and "@" in email:
            username_map[name] = email.split("@")[0]
    return username_map

def load_config():
    with open(CONFIG_PATH) as f:
        return json.load(f)

def get_snowflake_conn(cfg):
    return snowflake.connector.connect(
        user=cfg["snowflake_user"],
        account=cfg["snowflake_account"],
        authenticator="externalbrowser",
        warehouse="ELT_WH",
        database="PROD_DW",
        schema="DW_SEMANTIC",
    )

# ── PODS employee directory — used to auto-detect new PODS agents ─────────────
def load_pods_employee_directory(conn):
    """Latest row per employee from the Verint employee feed: name -> {email, organization}."""
    sql = """
    SELECT REPLACE("FirstName",'"','') AS FIRST_NAME,
           REPLACE("LastName",'"','')  AS LAST_NAME,
           REPLACE("Email",'"','')     AS EMAIL,
           REPLACE("Organization",'"','') AS ORGANIZATION
    FROM (
      SELECT "FirstName","LastName","Email","Organization",
             ROW_NUMBER() OVER (PARTITION BY "FirstName","LastName" ORDER BY LOAD_TIMESTAMP_TZ DESC) AS rn
      FROM PROD_LAKE.VERINT_SAAS.VW_EMPLOYEEINFO2_CONSOLIDATED
    )
    WHERE rn = 1
    """
    cur = conn.cursor()
    cur.execute(sql)
    directory = {}
    for first, last, email, org in cur.fetchall():
        full = f"{first} {last}".strip()
        directory[full] = {"email": email, "organization": org or ""}
    return directory

def guess_dept_loc(organization):
    """Best-guess dept/location from a Verint Organization string. Always flagged for human review."""
    org_l = organization.lower()
    if "outbound" in org_l:
        dept = "OB"
    elif "sst" in org_l:
        dept = "SST"
    elif "sales" in org_l or "commercial" in org_l:
        dept = "Sales"
    elif "service" in org_l:
        dept = "Service"
    else:
        dept = "SST"  # ambiguous org strings (e.g. ".. Hawaii City Sales") have skewed SST in practice
    if "clw" in org_l:
        loc = "CLW"
    elif "concentrix philippines" in org_l or "cnx" in org_l:
        loc = "CNX"
    else:
        loc = "—"
    return dept, loc

ROSTER_XLSX_PATH = r"F:\Revenue Management\Team Individual Folders\TSchlotter\Labor\Integrated Labor\IL_Agent_Roster 2.0.xlsx"
ROSTER_XLSX_SHEETS = ["Jontae & Michelle (all set up)", "Other new agents"]

# Supervisors who show up on the roster spreadsheet but don't take quotes/orders themselves —
# exclude from the untracked-agent nag since they'll never need a DEPT_MAP entry.
SUPERVISOR_NAMES = {
    "Charles DeMeo", "Chris Crenshaw", "Cythia Maready", "Jacob Rosenweig",
    "Jon'tae Ferguson", "Karel VanderZee", "Leonard Dixon", "Michelle Fallahee",
    "Patrice Jernigan", "Peter Wozdusiewicz", "Sergio Jewell",
}

def detect_untracked_trained_agents(known_roster_names):
    """Flag agents listed in the IL_Agent_Roster 2.0.xlsx onboarding tabs (the actual
    source of truth for who's trained on IL) who aren't in DEPT_MAP yet. Catches newly
    trained agents before they've shown any HaH activity, which detect_new_pods_agents
    alone would miss. Dept can't be inferred from the roster (only location can) —
    each hit needs a human to fill in Sales/SST/OB before pasting into DEPT_MAP."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(ROSTER_XLSX_PATH, data_only=True)
    except Exception as e:
        print(f"  (skipped — could not open roster workbook: {e})")
        return []

    missing = []
    for sheet_name in ROSTER_XLSX_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            if sheet_name == "Other new agents":
                name, email, supervisor = (row + (None, None, None))[:3]
                loc = None
            else:  # Jontae & Michelle style: Supervisor, Last, First, Last-First, Location, Email
                supervisor, last, first, _, loc, email = (row + (None,) * 6)[:6]
                name = f"{first} {last}".strip() if first and last else None
            if not name:
                continue
            name = str(name).replace("\xa0", " ").strip()
            if fuzzy_match(name, known_roster_names) or fuzzy_match(name, SUPERVISOR_NAMES):
                continue
            missing.append({"name": name, "loc": loc or "?", "email": email, "supervisor": supervisor, "sheet": sheet_name})
    return missing

def detect_new_pods_agents(all_hah_names, directory, known_roster_names):
    """Match HaH CSR names not already in DEPT_MAP against the real PODS employee
    directory. Returns (new_agents, still_unmatched) where new_agents maps the
    employee's canonical directory name -> dept/loc/email/org guess, and
    still_unmatched are names with no PODS employee match at all (true HaH-only)."""
    dir_by_norm = {}
    for full_name, info in directory.items():
        dir_by_norm.setdefault(normalize(full_name), []).append((full_name, info))

    new_agents, still_unmatched = {}, []
    for name in all_hah_names:
        if fuzzy_match(name, known_roster_names):
            continue  # already a known roster agent, handled normally
        norm = normalize(name)
        candidates = dir_by_norm.get(norm)
        if not candidates:
            parts = name.strip().split()
            if len(parts) >= 2:
                rev = normalize(' '.join(parts[1:] + [parts[0]]))
                candidates = dir_by_norm.get(rev)
        if candidates:
            full_name, info = candidates[0]
            org = info["organization"]
            if "inactive" in org.lower():
                still_unmatched.append(name)
                continue
            dept, loc = guess_dept_loc(org)
            new_agents[full_name] = {
                "dept": dept, "loc": loc,
                "email": info["email"], "organization": org, "hah_name": name,
            }
        else:
            still_unmatched.append(name)
    return new_agents, still_unmatched

def gh_get(token, path):
    r = requests.get(
        f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}",
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
    )
    r.raise_for_status()
    return r.json()

def gh_put(token, path, content_bytes, sha, message):
    r = requests.put(
        f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}",
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
        json={"message": message, "content": base64.b64encode(content_bytes).decode(), "sha": sha},
    )
    r.raise_for_status()
    print(f"  ✓ Pushed {path}")

def pick_file(title, filetypes=None):
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    path = filedialog.askopenfilename(
        parent=root, title=title,
        filetypes=filetypes or [("CSV files", "*.csv"), ("All files", "*.*")]
    )
    root.destroy()
    if not path:
        print(f"No file selected for: {title}")
        sys.exit(1)
    return path

def read_csv_auto(path):
    try:
        return pd.read_csv(path, encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-16")

# ── Quotes + orders archives ───────────────────────────────────────────────────
# Both HaH exports are capped at 1000 rows and sorted newest-first, so they
# silently drop the OLDEST rows rather than erroring (the export just ends with an
# "Export limited to 1000 records" line). Quotes hit the cap in Sep 2026 at ~145/week,
# which began cutting into the 8-week trend chart. Orders hit it in Nov 2026: the
# export stopped reaching back to launch, so May/early-June orders vanished, agents
# with early orders were pinned at their prior total by the never-go-below-prior
# floor in build_agents, and the May cancel-rate cohort lost its labor side.
#
# Fix: keep every row we have ever seen in a local archive keyed on its ID, and
# read the dashboard off that instead of off one week's export. Fresh rows win on
# ID collision, so anything still inside the export window keeps its latest values.
#
# Quotes: the only fields used — Quote ID and DateCreated — are immutable.
# Orders: OrderStatus and CSR_Name CAN change after booking (cancellations, HaH mods),
# and those changes stop arriving once an order ages out of the export window
# (~5 months at current volume). That is still strictly better than dropping the
# order outright, which is what the capped export alone does.
QUOTES_ARCHIVE = os.path.join(LOCAL_DIR, "quotes_archive.csv")
QUOTE_ID_COL   = "Quote ID"
QUOTE_DATE_COL = "DateCreated"

ORDERS_ARCHIVE = os.path.join(LOCAL_DIR, "orders_archive.csv")
ORDER_ID_COL   = "Order #"
ORDER_DATE_COL = "Date Booked"

EXPORT_CAP_MARKER = "Export limited to"

ARCHIVES = {
    "quotes": {"path": QUOTES_ARCHIVE, "id": QUOTE_ID_COL, "date": QUOTE_DATE_COL,
               "noun": "quotes", "flag": "--backfill"},
    "orders": {"path": ORDERS_ARCHIVE, "id": ORDER_ID_COL, "date": ORDER_DATE_COL,
               "noun": "orders", "flag": "--backfill-orders"},
}

def _archive_dates(df, spec):
    return pd.to_datetime(df[spec["date"]], format="mixed", errors="coerce")

def _archive_key(df, spec):
    # One export can read IDs as int and the next as str (the cap footer forces the
    # whole column to str), so normalize before de-duplicating.
    return df[spec["id"]].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

def export_was_capped(df):
    """True if the export carries HaH's row-cap footer, i.e. older rows were cut off."""
    head = df[df.columns[:2]].astype(str)
    return bool(head.apply(lambda c: c.str.startswith(EXPORT_CAP_MARKER)).any().any())

def _clean_archive_rows(df, spec, label):
    """Drop rows with no ID and the export's cap-footer line."""
    if spec["id"] not in df.columns or spec["date"] not in df.columns:
        print(f"  ⚠ {label}: missing '{spec['id']}' or '{spec['date']}' column — cannot archive.")
        print(f"    Columns found: {list(df.columns)}")
        sys.exit(1)
    ids = df[spec["id"]]
    keep = ids.notna() & ~ids.astype(str).str.startswith(EXPORT_CAP_MARKER)
    out = df[keep].copy()
    if len(out) < len(df):
        print(f"  {label}: skipped {len(df) - len(out)} row(s) with no {spec['id']} (blank or cap-footer line)")
    return out

def load_archive(kind):
    spec = ARCHIVES[kind]
    if not os.path.exists(spec["path"]):
        return None
    return _clean_archive_rows(read_csv_auto(spec["path"]), spec, f"{kind.title()} archive")

def merge_archive(kind, incoming_df, label=None, allow_gap=False):
    """Upsert one export into the archive and return the full history."""
    spec  = ARCHIVES[kind]
    label = label or f"{kind.title()} export"
    noun  = spec["noun"]
    capped   = export_was_capped(incoming_df)
    incoming = _clean_archive_rows(incoming_df, spec, label)
    archive  = load_archive(kind)
    i_dates  = _archive_dates(incoming, spec).dropna()

    if archive is None:
        # A capped export doesn't reach back to launch, so seeding the archive from
        # it alone would bake the missing history in permanently.
        if capped and not allow_gap and not i_dates.empty and \
                i_dates.min() > pd.Timestamp(LAUNCH_DATE) + pd.Timedelta(days=7):
            print(f"\n  ⚠⚠⚠ NO {noun.upper()} ARCHIVE YET, AND THIS EXPORT IS CAPPED ⚠⚠⚠")
            print(f"  It starts {i_dates.min():%Y-%m-%d}, not at launch ({LAUNCH_DATE}).")
            print(f"  Seed the archive first from exports that together cover {LAUNCH_DATE} -> today")
            print(f"  (each under 1000 rows, e.g. split by date range in the HaH portal):")
            print(f"       python update_il_dashboard.py {spec['flag']}")
            sys.exit(1)
        merged = incoming
        d = _archive_dates(merged, spec).dropna()
        print(f"  Created {noun} archive: {len(merged):,} {noun} "
              f"({d.min():%Y-%m-%d} -> {d.max():%Y-%m-%d})")
    else:
        a_dates = _archive_dates(archive, spec).dropna()
        # The gap guard. The export reaches back a fixed NUMBER OF ROWS, not a fixed
        # date range, so skipping enough weeks means the newest export no longer
        # touches the archive and the rows in between are lost with no error.
        # Refuse rather than write a hole nobody can see afterwards.
        if not a_dates.empty and not i_dates.empty and i_dates.min() > a_dates.max():
            missing_days = (i_dates.min() - a_dates.max()).days
            print(f"\n  ⚠⚠⚠ GAP IN {noun.upper()[:-1]} HISTORY ⚠⚠⚠")
            print(f"  Archive ends   {a_dates.max():%Y-%m-%d}")
            print(f"  This export starts {i_dates.min():%Y-%m-%d}  ({missing_days} days later)")
            print(f"  The export is row-capped, so it no longer reaches back to the archive.")
            print(f"  {noun.title()} in that window would be lost silently.")
            print(f"  Fix: export the missing range separately, then run:")
            print(f"       python update_il_dashboard.py {spec['flag']}")
            if not allow_gap:
                sys.exit(1)
            print(f"  (backfill: proceeding anyway)\n")

        before = len(archive)
        # Incoming last so keep='last' lets fresh rows overwrite archived ones.
        merged = pd.concat([archive, incoming], ignore_index=True)
        merged = merged[~_archive_key(merged, spec).duplicated(keep="last")]
        d = _archive_dates(merged, spec).dropna()
        print(f"  {noun.title()} archive: {before:,} + {len(incoming):,} incoming "
              f"-> {len(merged):,} unique (+{len(merged) - before:,} new)")
        print(f"    history now spans {d.min():%Y-%m-%d} -> {d.max():%Y-%m-%d}")

    merged = merged.sort_values(spec["date"], key=lambda s: pd.to_datetime(s, format="mixed", errors="coerce"))
    merged = merged.reset_index(drop=True)
    merged.to_csv(spec["path"], index=False, encoding="utf-8")
    return merged

def merge_quotes_archive(incoming_df, label="Quotes export", allow_gap=False):
    return merge_archive("quotes", incoming_df, label, allow_gap)

def merge_orders_archive(incoming_df, label="Orders export", allow_gap=False):
    return merge_archive("orders", incoming_df, label, allow_gap)

def backfill_archive(kind, paths=None):
    """One-time (or repeat) seeding of an archive from historical exports."""
    spec, noun = ARCHIVES[kind], ARCHIVES[kind]["noun"]
    print("=" * 50)
    print(f"  {noun.title()} Archive Backfill")
    print("=" * 50)
    existing = load_archive(kind)
    if existing is not None:
        d = _archive_dates(existing, spec).dropna()
        print(f"\nCurrent archive: {len(existing):,} {noun} "
              f"({d.min():%Y-%m-%d} -> {d.max():%Y-%m-%d})")
    else:
        print("\nNo archive yet — this will create it.")

    if not paths:
        print(f"\nSelect one or more historical HaH {noun.title()} CSVs (Ctrl+click for multiple)...")
        root = tk.Tk(); root.attributes('-topmost', True); root.withdraw()
        paths = filedialog.askopenfilenames(
            parent=root, title=f"Select historical HaH {noun.title()} CSV(s)",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        root.destroy()
    if not paths:
        print("Nothing selected — aborting.")
        sys.exit(1)

    # Oldest file first so the newest export's values win any collision.
    ordered = sorted(paths, key=lambda p: _archive_dates(
        _clean_archive_rows(read_csv_auto(p), spec, os.path.basename(p)), spec).max())
    for p in ordered:
        print(f"\n  Merging {os.path.basename(p)} ...")
        merge_archive(kind, read_csv_auto(p), os.path.basename(p), allow_gap=True)

    final = load_archive(kind)
    d = _archive_dates(final, spec).dropna()
    print(f"\n✅ Archive now holds {len(final):,} {noun} ({d.min():%Y-%m-%d} -> {d.max():%Y-%m-%d})")
    print(f"   {spec['path']}")
    print("\n   Check for gaps in the month-by-month counts below — a month that")
    print("   dips sharply against its neighbours may mean a range was missed:")
    for m, c in final.assign(_m=_archive_dates(final, spec).dt.to_period("M")).groupby("_m").size().items():
        print(f"     {m}  {c:>5,}")

def backfill_quotes_archive():
    backfill_archive("quotes")

def backfill_orders_archive():
    backfill_archive("orders")

def last_sunday():
    today = datetime.now()
    # weekday(): Mon=0, Sun=6
    days_back = (today.weekday() + 1) % 7  # days since last Sunday
    if days_back == 0:
        days_back = 7  # if today is Sunday, use the one 7 days ago
    return (today - timedelta(days=days_back)).strftime("%Y-%m-%d")

def fmt_date(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    day = dt.day
    return dt.strftime(f"%B {day}, %Y")

def week_label(end_date_str):
    end   = datetime.strptime(end_date_str, "%Y-%m-%d")
    start = end - timedelta(days=6)
    return f"{start.month}/{start.day}-{end.month}/{end.day}"

# ── Fuzzy name matching ────────────────────────────────────────────────────────
def normalize(name):
    return re.sub(r"[^a-z]", "", str(name).lower())

def fuzzy_match(hah_name, roster_names):
    import difflib
    hn = normalize(hah_name)
    # Exact normalized match
    for rn in roster_names:
        if normalize(rn) == hn:
            return rn
    # Reversed name match (e.g. "Bueno Annalee" → "Annalee Bueno")
    parts = hah_name.strip().split()
    if len(parts) >= 2:
        reversed_name = normalize(' '.join(parts[1:] + [parts[0]]))
        for rn in roster_names:
            if normalize(rn) == reversed_name:
                return rn
    # Close match (handles typos like Mondaldi vs Monaldi)
    close = difflib.get_close_matches(hn, [normalize(r) for r in roster_names], n=1, cutoff=0.88)
    if close:
        for rn in roster_names:
            if normalize(rn) == close[0]:
                return rn
    # Last-name + first initial match
    hl = hah_name.strip().split()
    last_h = normalize(hl[-1]) if hl else ""
    for rn in roster_names:
        rl = rn.strip().split()
        last_r = normalize(rl[-1]) if rl else ""
        if last_h and last_h == last_r and len(last_h) > 3:
            if normalize(hl[0])[0] == normalize(rl[0])[0]:
                return rn
    return None

# ── Re-attribute HaH-touched orders back to the originating PODS agent ────────
def resolve_originating_agent(order_po, order_customer_id, order_date, quotes_df, roster_names):
    """Find the agent who first quoted an order, by joining the order's PO # back
    to the Quotes CSV.

    Agents have been told at different times to put either the Customer ID (CID)
    or the Order Number in the quote's PO # field, so both show up mixed in the
    data. Resolve by:
      1. Try PO # as a CID — match quotes on Customer_ID, and among those pick the
         one whose DateCreated is closest to (and before) the order's date, since
         a repeat customer can have several quotes over time and date narrows it
         to the one that actually led to this order.
      2. If no quote's Customer_ID matches, fall back to treating PO # as a literal
         order/reference number and match quotes whose own PO # field equals it.
      3. If neither matches, return None — no quote history recoverable (order was
         likely started directly as a HaH order).
    Returns the roster agent name of the best-matching quote, or None.
    """
    po = str(order_po).strip()
    if po in ("", ".", "nan", "None"):
        return None

    candidates = pd.DataFrame()
    try:
        cid = int(float(po))
    except ValueError:
        cid = None
    if cid is not None and "Customer_ID" in quotes_df.columns:
        cid_matches = quotes_df[quotes_df["Customer_ID"] == cid]
        if not cid_matches.empty:
            candidates = cid_matches

    if candidates.empty and "PO #" in quotes_df.columns:
        candidates = quotes_df[quotes_df["PO #"].astype(str).str.strip() == po]

    if candidates.empty:
        return None

    candidates = candidates.copy()
    candidates["_qdt"] = pd.to_datetime(candidates["DateCreated"], format="mixed", errors="coerce")
    candidates = candidates.dropna(subset=["_qdt"])
    if candidates.empty:
        return None

    if pd.notna(order_date):
        before = candidates[candidates["_qdt"] <= order_date]
        best = before.loc[before["_qdt"].idxmax()] if not before.empty else candidates.loc[candidates["_qdt"].idxmin()]
    else:
        best = candidates.loc[candidates["_qdt"].idxmin()]

    return fuzzy_match(str(best["CSR_Name"]), roster_names)


def reattribute_hah_orders(quotes_df, orders_df, roster_names):
    """Re-point orders whose CSR_Name isn't a known roster agent (i.e. would roll up
    into the generic 'HaH Agent' bucket) back to the PODS agent who originated the
    quote, when recoverable via the PO#/CID join in resolve_originating_agent.

    Returns (modified orders_df, stats dict) — CSR_Name is overwritten in place for
    reattributed rows so every downstream consumer (dashboard KPIs, weekly trends,
    WoW, agent performance CSV) picks up the correction automatically.
    """
    orders_df = orders_df.copy()
    order_dates = (
        pd.to_datetime(orders_df["Date Booked"], format="mixed", errors="coerce")
        if "Date Booked" in orders_df.columns else pd.Series([pd.NaT] * len(orders_df), index=orders_df.index)
    )

    reattributed, unresolved = 0, 0
    for idx, row in orders_df.iterrows():
        name = str(row.get("CSR_Name", "")).strip()
        if not name or fuzzy_match(name, roster_names):
            continue  # already a known roster agent — nothing to fix
        resolved = resolve_originating_agent(
            row.get("PO #"), row.get("Customer_ID"), order_dates.loc[idx], quotes_df, roster_names
        )
        if resolved:
            orders_df.at[idx, "CSR_Name"] = resolved
            reattributed += 1
        else:
            unresolved += 1

    return orders_df, {"reattributed": reattributed, "unresolved": unresolved}


# ── HaH CSV processing ─────────────────────────────────────────────────────────
def check_csv_freshness(quotes_df, orders_df):
    """Warn loudly if one CSV's most recent date lags well behind the other's —
    usually means a stale/cached export got picked instead of a fresh download,
    which silently undercounts recent quotes/orders without erroring."""
    date_cols_q = [c for c in quotes_df.columns if "date" in c.lower() and "creat" in c.lower()]
    date_cols_o = [c for c in orders_df.columns if "date" in c.lower() and "book" in c.lower()]
    if not date_cols_q or not date_cols_o:
        return
    q_max = pd.to_datetime(quotes_df[date_cols_q[0]], format="mixed").max()
    o_max = pd.to_datetime(orders_df[date_cols_o[0]], format="mixed").max()
    gap = abs((q_max - o_max).total_seconds()) / 3600
    if gap > 24:
        older, newer, older_max, newer_max = (
            ("Orders", "Quotes", o_max, q_max) if o_max < q_max else ("Quotes", "Orders", q_max, o_max)
        )
        print(f"\n  ⚠⚠⚠ DATA FRESHNESS WARNING ⚠⚠⚠")
        print(f"  {older} CSV's latest date is {older_max} — {newer} CSV goes up to {newer_max}.")
        print(f"  That's a {gap:.0f}-hour gap. The {older} export looks stale (an old/cached download)")
        print(f"  rather than a fresh pull — recent {older.lower()} will be undercounted in this report.")
        print(f"  Recommend re-downloading a fresh {older} CSV before trusting these numbers.\n")

def process_hah_csvs(quotes_df, orders_df, end_date):
    # Both CSVs are exported "as of today", so a mid-week pull carries quotes and
    # orders dated after the report's week-ending Sunday. Everything else on the
    # dashboard (weekly trends, WoW, every Snowflake query) is bounded at end_date,
    # so leaving these two unbounded made the headline KPIs run ahead of the charts
    # sitting underneath them. Nothing is lost — the overhang lands in next week's run.
    def clamp_to_end_date(df, label, kw):
        cols = [c for c in df.columns if "date" in c.lower() and kw in c.lower()]
        if not cols:
            print(f"  ⚠ No {label} date column found — {label} NOT clamped to {end_date}")
            return df
        dt = pd.to_datetime(df[cols[0]], format="mixed", errors="coerce")
        # An unparseable date becomes NaT, and every NaT comparison is False — so a
        # naive `dt <= cutoff` would silently discard those rows as if they were in
        # the future. Keep them and say so; dropping a real quote over a malformed
        # date would undercount an agent with no visible sign it happened.
        undated = dt.isna()
        keep = (dt <= pd.Timestamp(end_date + " 23:59:59")) | undated
        dropped = int((~keep).sum())
        if dropped:
            print(f"  {label}: excluded {dropped} row(s) dated after {end_date} (via {cols[0]})")
        if undated.any():
            print(f"  {label}: {int(undated.sum())} row(s) have no parseable {cols[0]} — kept, not clamped")
        return df[keep]

    quotes_df = clamp_to_end_date(quotes_df, "Quotes", "creat")
    orders_df = clamp_to_end_date(orders_df, "Orders", "book")

    def find_col(df, keywords):
        for kw in keywords:
            for c in df.columns:
                if kw.lower() in c.lower():
                    return c
        return None

    q_name_col   = find_col(quotes_df, ["csr_name", "csr", "agent name", "agent"])
    q_id_col     = find_col(quotes_df, ["quote id", "quote_id", "quoteid"])
    o_name_col   = find_col(orders_df, ["csr_name", "csr", "agent name", "agent"])
    o_id_col     = find_col(orders_df, ["order #", "order_number", "order id", "order_id", "orderid"])
    o_status_col = find_col(orders_df, ["status"])

    if not q_name_col or not o_name_col:
        print("  ⚠ Could not detect agent name column. Available columns:")
        print("    Quotes:", list(quotes_df.columns))
        print("    Orders:", list(orders_df.columns))
        sys.exit(1)

    print(f"  Quotes: {q_name_col} / {q_id_col}")
    print(f"  Orders: {o_name_col} / {o_id_col} / status: {o_status_col}")

    # Quotes — count unique quote IDs per agent
    if q_id_col:
        q_by_agent = quotes_df.groupby(q_name_col)[q_id_col].nunique().to_dict()
    else:
        q_by_agent = quotes_df.groupby(q_name_col).size().to_dict()

    # Orders — exclude cancelled
    if o_status_col:
        orders_active = orders_df[~orders_df[o_status_col].astype(str).str.strip().str.lower().str.contains("cancel", na=False)]
    else:
        orders_active = orders_df

    if o_id_col:
        o_by_agent = orders_active.groupby(o_name_col)[o_id_col].nunique().to_dict()
    else:
        o_by_agent = orders_active.groupby(o_name_col).size().to_dict()

    all_hah_names = set(q_by_agent.keys()) | set(o_by_agent.keys())
    return q_by_agent, o_by_agent, all_hah_names

# ── Cancel rate from Tableau Labor Impact Table CSV ────────────────────────────
def read_cancel_rate_from_tableau_csv(csv_path, o_by_roster):
    """Compute weighted cancel rate delta from the Tableau 'Labor Impact Table' CSV.

    Weights each IL agent's (labor - regular) cancel rate delta by their labor
    order count from the HaH Orders CSV, matching the Excel methodology.
    """
    for enc in ("utf-16", "utf-8", "latin-1"):
        try:
            df = pd.read_csv(csv_path, encoding=enc, sep="\t")
            if len(df.columns) >= 6:
                break
        except Exception:
            continue
    df.columns = [c.strip() for c in df.columns]

    name_col = next((c for c in df.columns if "agent" in c.lower() and "name" in c.lower()), None)
    r7_col   = next((c for c in df.columns if "regular" in c.lower() and "7" in c), None)
    l7_col   = next((c for c in df.columns if "labor" in c.lower() and "7" in c.lower()), None)
    r30_col  = next((c for c in df.columns if "regular" in c.lower() and "30" in c), None)
    l30_col  = next((c for c in df.columns if "labor" in c.lower() and "30" in c.lower()), None)

    if not all([name_col, r7_col, l7_col, r30_col, l30_col]):
        raise ValueError(
            f"Could not find required columns in {csv_path}.\n"
            f"Found columns: {list(df.columns)}\n"
            "Expected: AGENT_NAME, Regular Container CR 7-Day, Labor Container CR 7-Day, "
            "Regular Container CR 30-Day, Labor Container Cancel Rate 30"
        )

    def pct(v):
        if pd.isna(v) or str(v).strip() in ("", "nan"):
            return None
        try:
            return float(str(v).strip().replace("%", "")) / 100
        except Exception:
            return None

    roster_names = list(DEPT_MAP.keys())
    w7 = w30 = total_weight = 0.0
    missing_data_weight = 0.0
    used, no_orders, missing_data, not_found = [], [], [], []

    matched_names = set()
    for _, row in df.iterrows():
        agent_name = str(row[name_col]).strip()
        matched = fuzzy_match(agent_name, roster_names)
        if not matched:
            continue
        matched_names.add(matched)
        weight = o_by_roster.get(matched, 0)
        if weight == 0:
            no_orders.append(matched)  # no labor orders to measure a delta against — exclude, don't fake a weight
            continue
        l7  = pct(row[l7_col])
        r7  = pct(row[r7_col])
        l30 = pct(row[l30_col])
        r30 = pct(row[r30_col])
        if l7 is None or r7 is None or l30 is None or r30 is None:
            missing_data.append(matched)  # usually a missing PODS CID in the PO# field — Tableau can't link the order
            missing_data_weight += weight
            continue
        used.append(matched)
        w7           += (l7  - r7)  * weight
        w30          += (l30 - r30) * weight
        total_weight += weight

    not_found = [r for r in roster_names if r not in matched_names and o_by_roster.get(r, 0) > 0]

    print(f"  Cancel rate calc: {len(used)} of {len(roster_names)} roster agents have usable Tableau data")
    if missing_data:
        print(f"    {len(missing_data)} imputed with group-avg delta - matched but missing labor cancel-rate data (likely no PODS CID in PO# field):")
        for a in sorted(missing_data):
            print(f"      - {a}")
    if not_found:
        print(f"    {len(not_found)} excluded - have booked labor orders but no row in the Tableau CSV at all:")
        for a in sorted(not_found):
            print(f"      - {a}")
    if no_orders:
        print(f"    {len(no_orders)} excluded - no booked labor orders yet (name matched, 0 weight)")

    if total_weight == 0:
        raise ValueError("No IL agents with labor cancel data found in the CSV.")

    # Agents with booked labor orders but no usable Tableau cancel data (usually a
    # missing PODS CID in the PO# field) get the group-weighted-average delta
    # imputed, so their order volume still counts toward total savings instead of
    # being silently dropped. This doesn't change delta_7d/delta_30d (imputing the
    # mean leaves the weighted mean unchanged) but grows total_weight, which the
    # savings dollar figure scales on.
    if missing_data_weight > 0:
        avg_delta_7  = w7  / total_weight
        avg_delta_30 = w30 / total_weight
        w7           += avg_delta_7  * missing_data_weight
        w30          += avg_delta_30 * missing_data_weight
        total_weight += missing_data_weight

    delta_7d  = round(w7  / total_weight * 100, 1)
    delta_30d = round(w30 / total_weight * 100, 1)

    # Savings are driven by whichever window (7d or 30d) shows the larger-magnitude
    # delta, since that reflects the greater current financial impact.
    primary_period = 7 if abs(delta_7d) >= abs(delta_30d) else 30
    primary_delta  = delta_7d if primary_period == 7 else delta_30d
    savings        = round(abs(primary_delta) / 100 * total_weight * AVG_CTR_PER_ORDER * ECR)
    print(f"  Savings driven by {primary_period}-day delta ({primary_delta}pp) — larger magnitude than the other window")

    return {"delta_7d": delta_7d, "delta_30d": delta_30d, "savings": savings, "primary_period": primary_period}

# ── Cancel rate from Snowflake (replaces the Tableau "Labor Impact Table" pull) ──
def compute_cancel_rate_from_snowflake(conn, directory, roster_names, orders_df, end_date):
    """Compute the weighted 7d/30d labor-vs-regular cancel-rate delta straight from
    ORDER_CONTAINER_SALES_FACT, replacing the weekly Tableau CSV pull.

    For each HaH order, joins to Snowflake via its PO # field. Agents have been told
    at different times to put either the Customer ID (CID) or the Order Number there,
    so both show up mixed in the data — resolve by:
      1. Try PO # as ORDER_CUSTOMER_DIM_KEY (the real CID). A repeat customer can have
         several orders, so pick whichever is closest before (or nearest to) the HaH
         order's own date.
      2. If no CID match, fall back to PO # as a literal Order Number.
      3. If neither matches, repeat steps 1-2 against the Reference ID field instead
         (same CID-then-OrderNumber order) — a fallback for rows where PO # is blank
         or garbage but Reference ID happens to hold a valid CID/order number.
      4. If nothing matches, the order is unresolved and excluded — same as Tableau's
         "missing PODS CID" bucket, except this recovers more of them since Tableau
         only ever tries step 1 (CID-only).
    'Regular' (baseline) cancel rate per agent is that agent's NON-LABOR Snowflake
    bookings in the same window — the orders resolved above are excluded so the two
    sides of the comparison are mutually exclusive. (Before 2026-08-03 the baseline
    was ALL bookings, labor included, which put labor orders on both sides and
    understated the delta.)

    AGENT ATTRIBUTION is a fallback chain, because neither source is reliable alone:
      1. HaH CSR_Name (after reattribute_hah_orders) if it lands on the roster — that
         is the rep who actually sold the labor.
      2. Otherwise Snowflake's "Agent Container Booked By".
      3. Otherwise unattributed, and PRINTED so a missing roster member gets noticed.
    Why both are needed: HaH overwrites CSR_Name with whoever last touched the order,
    so ~28% of labor orders showed a HaH agent rather than the PODS seller. But going
    PODS-only was worse in a different way — 20 orders were booked by a colleague who
    merely keyed the container while a trained agent did the selling, and every one of
    them named a roster agent on the HaH side.

    The baseline still keys off "Agent Container Booked By", so an agent can be credited
    with a labor container a colleague booked. That is deliberate: we are measuring the
    effect of SELLING labor, not of keying the order.

    NOT "Agent Original Order Booked By": one PODS order can accumulate containers
    booked by different agents on different days. "Original" credits whoever opened
    the order — sometimes years earlier (observed gaps of 185, 1036 and 1496 days on
    repeat customers) — so an agent who adds a container to an existing order, and
    sells the labor on it, gets no credit.
    Weighting: each agent's delta is weighted by their labor order count IN THE SAME
    WINDOW being measured (not cumulative since launch). Weighting by cumulative volume
    while measuring the rate on a trailing window let agents with a big historical
    footprint but little/no recent activity distort the average — the sample the rate
    is computed from should be the same sample that determines its weight.
    """
    orders_df = orders_df.copy()
    orders_df["_dt"] = pd.to_datetime(orders_df["Date Booked"], format="mixed", errors="coerce")
    # NOTE: HaH-cancelled orders are deliberately KEPT here. If a customer was sold
    # labor, they belong in the labor group regardless of whether the labor job later
    # fell through — dropping them filters on the outcome, and those are exactly the
    # customers most likely to also cancel their container. Excluding them removed the
    # worst cases from the labor side and flattered the program. (Other call sites in
    # this script still exclude cancelled orders; that is correct for counting bookings,
    # but wrong here.)
    # _roster (the fuzzy-matched HaH CSR_Name) is kept only to report how often PODS
    # disagrees with it. Rows that DON'T match the roster are deliberately NOT dropped:
    # those are mostly orders a HaH agent modded, which overwrote CSR_Name with a
    # non-PODS name. The PODS lookup below recovers the real selling agent, so filtering
    # on the HaH name here would throw away exactly the orders we can now attribute.
    orders_df["_roster"] = orders_df["CSR_Name"].apply(lambda n: fuzzy_match(str(n), roster_names))
    orders_df["_po"] = orders_df["PO #"].astype(str).str.strip()
    orders_df = orders_df[~orders_df["_po"].isin(["", ".", "nan", "None"])]
    orders_df["_ref"] = orders_df["Reference ID"].astype(str).str.strip()
    orders_df.loc[orders_df["_ref"].isin(["", ".", "nan", "None"]), "_ref"] = None

    username_map = build_agent_username_map(directory, roster_names)

    cur = conn.cursor()
    def fetch_df(sql):
        cur.execute(sql)
        cols = [c[0] for c in cur.description]
        return pd.DataFrame(cur.fetchall(), columns=cols)

    cids, ordnums = set(), set()
    for v in list(orders_df["_po"]) + [v for v in orders_df["_ref"] if v is not None]:
        try:
            n = int(float(v))
            cids.add(n)
            ordnums.add(n)
        except ValueError:
            pass
    if not cids:
        raise ValueError("No linkable PO # values found in Orders CSV")

    cid_list = ",".join(str(c) for c in cids)
    snow_by_cid = fetch_df(f"""
        SELECT "ORDER_CUSTOMER_DIM_KEY" AS CID, "Order Number" AS ORDER_NUMBER,
               UPPER("Agent Container Booked By") AS PODS_AGENT,
               "Date Order Booked" AS DATE_BOOKED, "Date Container Cancelled" AS CANCEL_DATE,
               "Cancelled Container Count" AS CANCELLED, "Gross Container Count" AS GROSS
        FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
        WHERE "ORDER_CUSTOMER_DIM_KEY" IN ({cid_list})
    """)
    ordnum_list = ",".join(str(o) for o in ordnums)
    snow_by_ordnum = fetch_df(f"""
        SELECT "Order Number" AS ORDER_NUMBER,
               UPPER("Agent Container Booked By") AS PODS_AGENT,
               "Date Order Booked" AS DATE_BOOKED, "Date Container Cancelled" AS CANCEL_DATE,
               "Cancelled Container Count" AS CANCELLED, "Gross Container Count" AS GROSS
        FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
        WHERE "Order Number" IN ({ordnum_list})
    """)
    usernames = list(username_map.values())
    username_list = ", ".join(f"'{u.upper()}'" for u in usernames)
    snow_regular = fetch_df(f"""
        SELECT UPPER("Agent Container Booked By") AS USERNAME, "Order Number" AS ORDER_NUMBER,
               "Date Order Booked" AS DATE_BOOKED, "Date Container Cancelled" AS CANCEL_DATE,
               "Cancelled Container Count" AS CANCELLED, "Gross Container Count" AS GROSS
        FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
        WHERE UPPER("Agent Container Booked By") IN ({username_list})
          AND "Date Order Booked" >= '{LAUNCH_DATE}' AND "Date Order Booked" <= '{end_date}'
    """)

    for _df in (snow_by_cid, snow_by_ordnum, snow_regular):
        _df["DATE_BOOKED"] = pd.to_datetime(_df["DATE_BOOKED"])
        _df["CANCEL_DATE"] = pd.to_datetime(_df["CANCEL_DATE"])
    cid_groups = {k: v for k, v in snow_by_cid.groupby("CID")}
    ordnum_groups = {k: v for k, v in snow_by_ordnum.groupby("ORDER_NUMBER")}
    username_to_roster = {v.upper(): k for k, v in username_map.items()}
    snow_regular["agent"] = snow_regular["USERNAME"].map(username_to_roster)

    def resolve_one(id_str, hah_date):
        """Return EVERY container row of the matched PODS order.

        ORDER_CONTAINER_SALES_FACT is container-grain — one row per container, with
        "Gross Container Count" always 1. Returning a single row (as this did before
        2026-08-03) counted only one container of a multi-container order, so the
        labor side was effectively measured in orders while the non-labor baseline
        was measured in containers. Labor orders average ~1.44 containers, so that
        silently dropped roughly a third of the labor population.
        """
        try:
            cid = int(float(id_str))
        except ValueError:
            return None
        cand = cid_groups.get(cid)
        if cand is None or cand.empty:
            cand = ordnum_groups.get(cid)
        if cand is None or cand.empty:
            return None
        # Choose the closest ORDER (a repeat customer may have several), then keep
        # all of that order's containers.
        per_order = cand.groupby("ORDER_NUMBER")["DATE_BOOKED"].min()
        if pd.notna(hah_date):
            # Discard candidates too far from the labor booking to be the same job.
            within = per_order[(per_order - hah_date).abs().dt.days <= MATCH_WINDOW_DAYS]
            if within.empty:
                return None      # unresolved rather than matched to an unrelated order
            before = within[within <= hah_date]
            pick = before.idxmax() if not before.empty else within.idxmin()
        else:
            pick = per_order.idxmin()
        return cand[cand["ORDER_NUMBER"] == pick]

    def resolve(po_str, ref_str, hah_date):
        grp = resolve_one(po_str, hah_date)
        if grp is not None and not grp.empty:
            return grp
        if ref_str is not None:
            return resolve_one(ref_str, hah_date)
        return None

    results = []
    labor_order_numbers = set()   # PODS orders that ARE labor — excluded from the non-labor baseline
    by_hah_name = 0          # credited from the HaH CSR (post-reattribution)
    recovered_by_pods = 0    # HaH name unusable, PODS booking agent rescued it
    split_credit = 0         # both known but they disagree — labor seller != container booker
    unattributed = []        # neither source lands on the roster; reported, never silent
    for _, row in orders_df.iterrows():
        grp = resolve(row["_po"], row["_ref"], row["_dt"])
        if grp is None or grp.empty:
            continue
        ordnum = int(grp["ORDER_NUMBER"].iloc[0])
        if ordnum in labor_order_numbers:
            continue   # two HaH rows resolved to the same PODS order — count its containers once
        labor_order_numbers.add(ordnum)

        # ── Attribution fallback chain ───────────────────────────────────────
        # Neither source is trustworthy on its own:
        #   * HaH CSR_Name gets overwritten when a HaH agent mods the order, so ~28% of
        #     labor orders showed a HaH agent instead of the PODS rep who sold it.
        #   * The PODS booking agent is often a colleague who merely keyed the container
        #     while a trained agent did the selling — 20 such orders observed, every one
        #     with a roster agent named on the HaH side.
        # So: believe the HaH name when it resolves to a roster agent (that IS the person
        # who sold the labor), and fall back to the PODS booking agent when it doesn't.
        # Note this can credit an agent for a container someone else booked — intentional,
        # since we're measuring the effect of SELLING labor, not of keying the order.
        pods_user  = str(grp["PODS_AGENT"].iloc[0]).upper().strip()
        pods_agent = username_to_roster.get(pods_user)
        hah_agent  = None if pd.isna(row["_roster"]) else row["_roster"]

        if hah_agent is not None:
            agent = hah_agent
            by_hah_name += 1
            if pods_agent is not None and pods_agent != hah_agent:
                split_credit += 1
        elif pods_agent is not None:
            agent = pods_agent
            recovered_by_pods += 1
        else:
            unattributed.append((pods_user, str(row.get("CSR_Name", "?"))))
            continue

        for _, c in grp.iterrows():
            results.append({
                "agent":       agent,
                "ordnum":      ordnum,
                "cancelled":   int(c["CANCELLED"]),
                "gross":       int(c["GROSS"]),
                # PODS booking/cancel dates, not the HaH ones — days-to-cancel is a PODS event
                "booked":      c["DATE_BOOKED"],
                "cancel_date": c["CANCEL_DATE"],
            })
    res_df = pd.DataFrame(results, columns=["agent", "ordnum", "cancelled", "gross", "booked", "cancel_date"])
    print(f"  Labor: {len(labor_order_numbers)} PODS orders -> {len(res_df)} containers "
          f"({len(res_df)/max(len(labor_order_numbers),1):.2f} per order)")
    print(f"  Attribution: {by_hah_name} via HaH CSR name, {recovered_by_pods} recovered via the PODS "
          f"booking agent, {len(unattributed)} unattributed")
    if split_credit:
        print(f"    ({split_credit} of those had a DIFFERENT roster agent booking the container than "
              f"selling the labor — credited to the seller)")
    if unattributed:
        # Never drop these silently: if a real IL agent is missing from DEPT_MAP this is
        # the only place it shows up, since detect_new_pods_agents only reads HaH names.
        print(f"    Unattributed labor orders — neither the HaH name nor the PODS booking agent is on "
              f"the roster. Add to DEPT_MAP if any is a trained IL agent:")
        from collections import Counter
        for (pods_u, hah_n), cnt in Counter(unattributed).most_common():
            print(f"      PODS '{pods_u}' / HaH '{hah_n}'  x{cnt}")

    # Baseline must be NON-labor orders, not all orders. Leaving the labor orders in
    # here puts them on both sides of the comparison, which pulls the two rates
    # together and understates the delta.
    before_n = len(snow_regular)
    snow_regular = snow_regular[~snow_regular["ORDER_NUMBER"].isin(labor_order_numbers)]
    print(f"  Baseline: excluded {before_n - len(snow_regular)} labor orders -> {len(snow_regular)} non-labor orders")

    end_dt = pd.Timestamp(end_date + " 23:59:59")
    res_df["booked"]      = pd.to_datetime(res_df["booked"])
    res_df["cancel_date"] = pd.to_datetime(res_df["cancel_date"])

    # Drop containers on PODS orders booked before launch. A labor job sold in May can
    # legitimately match an order booked in late April (inside MATCH_WINDOW_DAYS), but
    # the non-labor baseline is already bounded at LAUNCH_DATE by its query — so keeping
    # them would put containers on the labor side from a period the baseline can't see.
    _pre = len(res_df)
    res_df = res_df[res_df["booked"] >= pd.Timestamp(LAUNCH_DATE)]
    if _pre != len(res_df):
        print(f"  Dropped {_pre - len(res_df)} labor container(s) on pre-launch PODS orders")

    # ── Labor attach rate ─────────────────────────────────────────────────────
    # Deliberately NOT maturity-filtered: whether an order has labor is known the day
    # it's booked, so the 30-day cutoff the cancel rate needs would just discard a
    # month of perfectly good data here. Measured in ORDERS because labor is sold per
    # move (one crew loads all of a customer's containers), not per container.
    attach_win = res_df[res_df["booked"] <= end_dt]
    labor_orders = int(attach_win["ordnum"].nunique())
    tot_orders = fetch_df(f"""
        SELECT COUNT(DISTINCT "Order Number") AS ORDERS
        FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
        WHERE UPPER("Agent Container Booked By") IN ({username_list})
          AND "Date Order Booked" >= '{LAUNCH_DATE}' AND "Date Order Booked" <= '{end_date}'
    """)["ORDERS"].iloc[0]
    attach_rate = round(100.0 * labor_orders / int(tot_orders), 2) if tot_orders else 0.0
    print(f"  Attach rate: {labor_orders} labor orders / {int(tot_orders):,} roster orders = {attach_rate}%")

    def _cancelled_within(df, booked_col, cancel_col, cancelled_col, days):
        """Containers that cancelled within `days` of being booked. Rows that never
        cancelled, or cancelled later than the threshold, contribute 0."""
        lag = (df[cancel_col] - df[booked_col]).dt.days
        return df[cancelled_col].where(lag.notna() & (lag <= days), 0)

    def weighted_delta(days, collect_rows=None):
        # "Cancelled within N days of booking", measured across every order old enough
        # to have had the full N days to cancel.
        #
        # The old version asked "of orders BOOKED in the last N days, how many
        # cancelled?" — which put brand-new orders in the denominator before they'd
        # had time to go bad, systematically understating the rate and making the
        # 7d and 30d figures non-comparable. The maturity cutoff below is one step
        # stricter than Ross's model, which leaves immature orders in the denominator.
        cutoff = end_dt - pd.Timedelta(days=days)
        lab_win = res_df[res_df["booked"] <= cutoff].copy()
        reg_win = snow_regular[snow_regular["DATE_BOOKED"] <= cutoff].copy()
        lab_win["_c"] = _cancelled_within(lab_win, "booked", "cancel_date", "cancelled", days)
        reg_win["_c"] = _cancelled_within(reg_win, "DATE_BOOKED", "CANCEL_DATE", "CANCELLED", days)
        w_sum, total_weight, used, no_data, no_orders = 0.0, 0.0, 0, 0, 0
        # Running totals across the agents that actually count, so the dashboard can
        # show the raw numbers behind the delta instead of just the delta.
        tot = {"labor_ctr": 0, "labor_cx": 0, "reg_ctr": 0, "reg_cx": 0}
        for agent in roster_names:
            lab = lab_win[lab_win["agent"] == agent]
            lg, lc = int(lab["gross"].sum()), int(lab["_c"].sum())
            if lg == 0:
                no_orders += 1
                continue
            reg = reg_win[reg_win["agent"] == agent]
            rg, rc = int(reg["GROSS"].sum()), int(reg["_c"].sum())
            if rg == 0:
                no_data += 1
                continue
            tot["labor_ctr"] += lg; tot["labor_cx"] += lc
            tot["reg_ctr"]   += rg; tot["reg_cx"]   += rc
            weight = lg  # weight by this agent's mature labor container count
            labor_cr = lc / lg
            regular_cr = rc / rg
            delta = labor_cr - regular_cr
            w_sum += delta * weight
            total_weight += weight
            used += 1
            if collect_rows is not None:
                # Same units as the aggregate savings formula (abs(delta_30d/100) * total_weight
                # * ECR): this agent's *signed* dollar slice of that total. Summing this column
                # across all rows reproduces the reported $ savings exactly (as long as most
                # deltas share the same sign, which they do in practice).
                savings_contribution = delta * weight * ECR
                collect_rows.append({
                    "agent": agent,
                    "labor_gross": lg, "labor_cancelled": lc, "labor_cr": round(labor_cr * 100, 1),
                    "regular_gross": rg, "regular_cancelled": rc, "regular_cr": round(regular_cr * 100, 1),
                    "delta_pp": round(delta * 100, 1),
                    "weight": weight,
                    "savings_contribution": round(savings_contribution),
                })
        return w_sum, total_weight, used, no_data, no_orders, tot

    w7, tw7, used7, nodata7, noord7, tot7 = weighted_delta(7)
    agent_rows_30d = []
    w30, tw30, used30, nodata30, noord30, tot30 = weighted_delta(30, collect_rows=agent_rows_30d)

    if agent_rows_30d:
        total_weight_30d = sum(r["weight"] for r in agent_rows_30d)
        total_savings_30d = sum(r["savings_contribution"] for r in agent_rows_30d)
        for r in agent_rows_30d:
            r["weight_pct"] = round(r["weight"] / total_weight_30d * 100, 1) if total_weight_30d else 0.0
            r["pct_of_total_savings"] = (
                round(r["savings_contribution"] / total_savings_30d * 100, 1) if total_savings_30d else 0.0
            )
        agent_rows_30d.sort(key=lambda r: abs(r["savings_contribution"]), reverse=True)
        cols = ["agent", "labor_gross", "labor_cancelled", "labor_cr",
                "regular_gross", "regular_cancelled", "regular_cr", "delta_pp",
                "weight", "weight_pct", "savings_contribution", "pct_of_total_savings"]
        out_path = os.path.join(LOCAL_DIR, f"il_cancel_rate_by_agent_30d_{end_date}.csv")
        pd.DataFrame(agent_rows_30d, columns=cols).to_csv(out_path, index=False)
        print(f"  Per-agent 30d labor-vs-regular breakdown written to {out_path}")
        print(f"  Sum of savings_contribution = ${round(total_savings_30d):,} (should match reported 30d savings)")

    if tw7 == 0 and tw30 == 0:
        raise ValueError("No IL agents with usable Snowflake labor+regular cancel data found.")

    delta_7d  = round(w7  / tw7  * 100, 1) if tw7  else 0.0
    delta_30d = round(w30 / tw30 * 100, 1) if tw30 else 0.0

    print(f"  7d:  {used7} agents used, {nodata7} matched-but-missing-labor/regular data, {noord7} no labor orders yet")
    print(f"  30d: {used30} agents used, {nodata30} matched-but-missing-labor/regular data, {noord30} no labor orders yet")

    # Always drive savings off the 30-day window — more sample size, less week-to-week
    # whiplash than letting a noisy 7-day spike take over (used to pick whichever
    # window had the larger-magnitude delta; that let a single volatile week swing
    # the $ figure).
    primary_period = 30
    primary_delta  = delta_30d
    total_weight   = tw30
    # No AVG_CTR_PER_ORDER here: total_weight is already a CONTAINER count (summed
    # "Gross Container Count"), so applying an orders->containers factor would count
    # containers twice. The constant still applies in the legacy Tableau path, whose
    # weights are order counts.
    savings        = round(abs(primary_delta) / 100 * total_weight * ECR)
    print(f"  Savings driven by 30-day delta ({primary_delta}pp)")

    # ── Uncertainty band on the 30-day delta ─────────────────────────────────
    # The delta rests on a small number of labor containers, so a point estimate on its
    # own is misleading. Standard error of the labor rate dominates — the non-labor
    # baseline has thousands of containers and contributes almost nothing. Effective n is
    # the labor ORDER count, not the container count: containers on one order cancel
    # together, so treating them as independent would understate the interval.
    lg30, lc30 = tot30["labor_ctr"], tot30["labor_cx"]
    cutoff30 = end_dt - pd.Timedelta(days=30)
    eff_n = int(res_df[res_df["booked"] <= cutoff30]["ordnum"].nunique()) or 1
    p = (lc30 / lg30) if lg30 else 0.0
    se_pp = 100.0 * ((p * (1 - p) / eff_n) ** 0.5)
    margin = round(1.96 * se_pp, 1)
    delta_low  = round(delta_30d - margin, 1)   # more negative = bigger benefit
    delta_high = round(delta_30d + margin, 1)
    savings_low  = round(max(abs(delta_high), 0) / 100 * total_weight * ECR) if delta_high < 0 else 0
    savings_high = round(abs(delta_low) / 100 * total_weight * ECR)
    print(f"  95% band on 30d delta: {delta_low}pp .. {delta_high}pp "
          f"(±{margin}pp, effective n = {eff_n} labor orders)")
    print(f"  Savings range: ${savings_low:,} .. ${savings_high:,}")

    # ── Cancel rate by BOOKING MONTH cohort ───────────────────────────────────
    # The blended figure above hides drift: a strong early month keeps propping it
    # up for months after the effect has faded. This groups orders by the month they
    # were booked and asks the same question of each cohort. Pooled (SUM/SUM) rather
    # than the per-agent weighted delta — monthly slices are too thin per agent.
    # A month is only shown once every order in it has had the full 30 days.
    def cancel_by_month(days=30):
        cutoff = end_dt - pd.Timedelta(days=days)
        lab = res_df.copy()
        reg = snow_regular.copy()
        lab["_m"] = lab["booked"].dt.to_period("M")
        reg["_m"] = reg["DATE_BOOKED"].dt.to_period("M")
        # Nothing before the launch month: a labor job booked in early May can match a
        # PODS order booked in late April (inside MATCH_WINDOW_DAYS), which produced a
        # stray 1-container "Apr" cohort. The program didn't exist then.
        launch_m = pd.Period(LAUNCH_DATE, freq="M")
        out = []
        for mth in sorted(set(lab["_m"].dropna()) | set(reg["_m"].dropna())):
            if mth < launch_m:
                continue
            month_end = mth.to_timestamp(how="end")
            mature    = month_end <= cutoff
            lm = lab[(lab["_m"] == mth) & (lab["booked"]      <= cutoff)]
            rm = reg[(reg["_m"] == mth) & (reg["DATE_BOOKED"] <= cutoff)]
            lg, rg = int(lm["gross"].sum()), int(rm["GROSS"].sum())
            # Date this cohort is finally complete = month end + the full cancel window.
            complete_after = (month_end + pd.Timedelta(days=days)).strftime("%b %d").replace(" 0", " ")
            row = {"month": str(mth), "label": mth.strftime("%b"),
                   "labor_containers": lg, "mature": bool(mature),
                   "complete_after": complete_after}
            if lg and rg:
                lc = int(_cancelled_within(lm, "booked", "cancel_date", "cancelled", days).sum())
                rc = int(_cancelled_within(rm, "DATE_BOOKED", "CANCEL_DATE", "CANCELLED", days).sum())
                row["labor_cr"]   = round(100 * lc / lg, 1)
                row["regular_cr"] = round(100 * rc / rg, 1)
                row["delta"]      = round(100 * (lc / lg - rc / rg), 1)
            else:
                row["labor_cr"] = row["regular_cr"] = row["delta"] = None
            out.append(row)
        return out

    by_month = cancel_by_month(30)
    print("  Cancel rate by booking month (30d):")
    for r in by_month:
        if r["delta"] is None:
            print(f"    {r['label']}: no comparable data yet ({r['labor_containers']} labor ctrs)")
        else:
            flag = "" if r["mature"] else "  <- still maturing"
            print(f"    {r['label']}: labor {r['labor_cr']}% vs regular {r['regular_cr']}% "
                  f"= {r['delta']}pp on {r['labor_containers']} labor ctrs{flag}")

    # Raw numbers behind each delta, so the dashboard can show its own working.
    def _detail(days, tot, agents_used):
        lg, lc, rg, rc = tot["labor_ctr"], tot["labor_cx"], tot["reg_ctr"], tot["reg_cx"]
        return {
            "days": days,
            "labor_containers":   lg,
            "labor_cancelled":    lc,
            "labor_cr":           round(100 * lc / lg, 1) if lg else None,
            "regular_containers": rg,
            "regular_cancelled":  rc,
            "regular_cr":         round(100 * rc / rg, 1) if rg else None,
            "agents":             agents_used,
            "cutoff":             (end_dt - pd.Timedelta(days=days)).strftime("%Y-%m-%d"),
            # Human-readable booking window this figure actually covers. Ends at the
            # maturity cutoff, not the report date — the last N days aren't eligible yet.
            "covers_from":        pd.Timestamp(LAUNCH_DATE).strftime("%b %d").replace(" 0", " "),
            "covers_to":          (end_dt - pd.Timedelta(days=days)).strftime("%b %d").replace(" 0", " "),
        }

    return {"delta_7d": delta_7d, "delta_30d": delta_30d, "savings": savings,
            "primary_period": primary_period, "by_month": by_month,
            "range": {"delta_low": delta_low, "delta_high": delta_high,
                      "savings_low": savings_low, "savings_high": savings_high,
                      "margin_pp": margin, "effective_n": eff_n},
            "detail_7d":  _detail(7,  tot7,  used7),
            "detail_30d": _detail(30, tot30, used30),
            "attach_rate": attach_rate,
            "attach_detail": {
                "labor_orders": labor_orders,
                "labor_containers": int(len(attach_win)),
                "total_orders": int(tot_orders),
                "agents":       int(attach_win["agent"].nunique()),
                "roster":       len(roster_names),
                "from":         pd.Timestamp(LAUNCH_DATE).strftime("%b %d").replace(" 0", " "),
                "to":           end_dt.strftime("%b %d").replace(" 0", " "),
            }}


# ── Load previous agent data from il_data.js ──────────────────────────────────
def load_existing_agents():
    path = os.path.join(LOCAL_DIR, "il_data.js")
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        content = f.read()
    m = re.search(r"agents:\s*(\[[\s\S]*?\]),\s*\}", content)
    if m:
        try:
            arr = json.loads(m.group(1))
            return {a["name"]: a for a in arr}
        except Exception:
            pass
    return {}

# ── Build agent list ───────────────────────────────────────────────────────────
HAH_AGENT_NAME = "HaH Agent"

def build_agents(q_by_hah, o_by_hah, all_hah_names, prev_agents):
    roster_names = list(DEPT_MAP.keys())
    unmatched = []

    # Map HaH name → roster canonical name; non-roster names roll up under HAH_AGENT_NAME
    hah_to_roster = {}
    for name in all_hah_names:
        matched = fuzzy_match(name, roster_names)
        hah_to_roster[name] = matched if matched else HAH_AGENT_NAME
        if not matched:
            unmatched.append(name)

    if unmatched:
        print(f"  Rolled up {len(unmatched)} unrecognized names into '{HAH_AGENT_NAME}':")
        for s in sorted(unmatched):
            print(f"    - {s}")
        print("  To add a new PODS agent to the roster instead, add them to DEPT_MAP in the script.")

    # Accumulate quotes/orders to roster canonical names (+ HaH Agent rollup)
    q_totals, o_totals = {}, {}
    for hah_name, roster_name in hah_to_roster.items():
        q_totals[roster_name] = q_totals.get(roster_name, 0) + q_by_hah.get(hah_name, 0)
        o_totals[roster_name] = o_totals.get(roster_name, 0) + o_by_hah.get(hah_name, 0)

    agents = []
    ratchet_engaged = []
    for name, (dept, loc) in DEPT_MAP.items():
        prev = prev_agents.get(name, {})
        # Quotes never go below prior — a created quote stays created forever, even
        # if the deal falls through, so the CSV's cumulative count should only grow.
        # Kept as a safety net, but now that quotes come from the archive rather than
        # a row-capped export it should never actually engage. When it does, say so:
        # this silently masked seven weeks of missing history before the archive
        # existed, and a silent mask is how that went unnoticed.
        from_data = q_totals.get(name, 0)
        quoted = max(from_data, prev.get("quoted", 0))
        if from_data < prev.get("quoted", 0):
            ratchet_engaged.append((name, from_data, prev["quoted"]))
        # Orders are NOT clamped to prior — unlike quotes, a booked order can be
        # legitimately cancelled after the fact, which should reduce this count.
        # Trust the CSV's current active (non-cancelled) total directly.
        booked = o_totals.get(name, 0)
        agents.append({"name": name, "dept": dept, "loc": loc, "quoted": quoted, "booked": booked})

    if ratchet_engaged:
        total_masked = sum(p - f for _, f, p in ratchet_engaged)
        print(f"\n  ⚠ Quote ratchet engaged for {len(ratchet_engaged)} agent(s) — "
              f"data shows {total_masked} fewer quote(s) than last week, prior values kept.")
        print(f"    With the archive in place this should not happen. Likely causes:")
        print(f"    archive was reset/corrupted, or quotes were deleted in HaH.")
        for n, f, p in sorted(ratchet_engaged, key=lambda x: x[2] - x[1], reverse=True)[:5]:
            print(f"      - {n}: data {f}, prior {p}")

    # HaH Agent rollup row — quotes/orders from HaH agents not on the PODS roster
    prev_hah = prev_agents.get(HAH_AGENT_NAME, {})
    hah_quoted = max(q_totals.get(HAH_AGENT_NAME, 0), prev_hah.get("quoted", 0))
    hah_booked = o_totals.get(HAH_AGENT_NAME, 0)
    if hah_quoted or hah_booked:
        agents.append({"name": HAH_AGENT_NAME, "dept": "HaH", "loc": "—", "quoted": hah_quoted, "booked": hah_booked})

    agents.sort(key=lambda a: (-a["booked"], -a["quoted"], a["name"]))
    return agents

def q_total_orders_by_agent(conn, end_date, username_map):
    """Total PODS orders AND containers per roster agent since launch.

    Returns (orders_by_agent, containers_by_agent). Both are needed: pitch rate and
    attach rate work in orders, while opportunity sizing works in containers.
    """
    usernames = list(username_map.values())
    username_list = ", ".join(f"'{u.upper()}'" for u in usernames)
    sql = f"""
    SELECT UPPER("Agent Container Booked By") AS USERNAME,
           COUNT(DISTINCT "Order Number") AS total_orders,
           COUNT(*)                        AS total_containers
    FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
    WHERE "Date Order Booked" >= '{LAUNCH_DATE}'
      AND "Date Order Booked" <= '{end_date}'
      AND UPPER("Agent Container Booked By") IN ({username_list})
    GROUP BY 1
    """
    cur = conn.cursor()
    cur.execute(sql)
    rows = cur.fetchall()
    username_to_roster = {v.upper(): k for k, v in username_map.items()}
    orders, containers = {}, {}
    for user, o, c in rows:
        name = username_to_roster.get(user)
        if name:
            orders[name] = o
            containers[name] = c
    return orders, containers

def compute_dept_stats(agents, total_orders_by_agent=None):
    stats = {}
    for dept in ("Sales", "SST", "OB"):
        grp = [a for a in agents if a["dept"] == dept]
        total  = len(grp)
        active = sum(1 for a in grp if a["quoted"] > 0)
        orders = sum(a["booked"] for a in grp)
        quotes = sum(a["quoted"] for a in grp)
        if total_orders_by_agent:
            dept_total_orders = sum(total_orders_by_agent.get(a["name"], 0) for a in grp)
            pitch_rate = round(quotes / dept_total_orders * 100, 1) if dept_total_orders else 0
        else:
            pitch_rate = 0
        stats[dept] = {
            "total": total, "active": active,
            "orders": orders, "quotes": quotes,
            "conv_rate": round(orders / quotes * 100, 1) if quotes else 0,
            "pitch_rate": pitch_rate,
        }
    return stats

# ── Load existing weekly_trends from local il_data.js ─────────────────────────
def load_existing_weekly_trends():
    path = os.path.join(LOCAL_DIR, "il_data.js")
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        content = f.read()
    m = re.search(r"weekly_trends:\s*(\[[\s\S]*?\])\s*,", content)
    if not m:
        return []
    raw = m.group(1)
    # Normalize JS object syntax to JSON (quote unquoted keys)
    raw = re.sub(r'(\{|,)\s*(\w+)\s*:', r'\1 "\2":', raw)
    # Remove trailing commas before } or ]
    raw = re.sub(r',\s*([}\]])', r'\1', raw)
    try:
        return json.loads(raw)
    except Exception:
        return []

def load_existing_cancel_data():
    """Fall back to cancel rate values already in il_data.js if Snowflake query fails."""
    path = os.path.join(LOCAL_DIR, "il_data.js")
    if not os.path.exists(path):
        return {"delta_7d": 0, "delta_30d": 0, "savings": 0}
    with open(path, encoding="utf-8") as f:
        content = f.read()
    def extract_float(key):
        m = re.search(rf"{key}:\s*(-?[\d.]+)", content)
        return float(m.group(1)) if m else 0
    def extract_int(key):
        m = re.search(rf"{key}:\s*([\d]+)", content)
        return int(m.group(1)) if m else 0
    delta_7d  = extract_float("cancel_delta_7d")
    delta_30d = extract_float("cancel_delta_30d")
    m = re.search(r"cancel_delta_primary_period:\s*(\d+)", content)
    primary_period = int(m.group(1)) if m else (7 if abs(delta_7d) >= abs(delta_30d) else 30)
    # json.dumps(indent=2) closes the array with "]" at column 0, so anchor on that
    # rather than an indented bracket.
    m = re.search(r"cancel_by_month:\s*(\[[\s\S]*?\n\])\s*,", content)
    try:
        by_month = json.loads(m.group(1)) if m else []
    except Exception:
        by_month = []
    m2 = re.search(r"labor_attach_detail:\s*(\{[\s\S]*?\n\})\s*,", content)
    try:
        attach_detail = json.loads(m2.group(1)) if m2 else {}
    except Exception:
        attach_detail = {}
    return {
        "delta_7d":       delta_7d,
        "delta_30d":      delta_30d,
        "savings":        extract_int("cancel_savings"),
        "primary_period": primary_period,
        "by_month":       by_month,
        "attach_rate":    extract_float("labor_attach_rate"),
        "attach_detail":  attach_detail,
        "range":          {},
    }

def load_prev_totals():
    """Read prev_orders, prev_quotes, prev_savings, and cum_* from current il_data.js."""
    path = os.path.join(LOCAL_DIR, "il_data.js")
    if not os.path.exists(path):
        return 0, 0, 0, 0, 0
    with open(path, encoding="utf-8") as f:
        content = f.read()
    def extract(key):
        m = re.search(rf"{key}:\s*([\d]+)", content)
        return int(m.group(1)) if m else 0
    return (
        extract("total_orders"),
        extract("total_quotes"),
        extract("cancel_savings"),
        extract("cum_quotes"),
        extract("cum_orders"),
    )

# ── Generate il_data.js content ────────────────────────────────────────────────
def build_weekly_trends(quotes_df, orders_df, hah_to_roster, end_date, existing_trends):
    """Compute weekly quote/order counts directly from date-filtered CSVs."""
    import pandas as pd

    end_dt = pd.Timestamp(end_date + " 23:59:59")

    # Map HaH names to roster names
    qdf = quotes_df[quotes_df.iloc[:, 0].isin(hah_to_roster) if False else quotes_df["CSR_Name"].isin(hah_to_roster)].copy()
    odf = orders_df[orders_df["CSR_Name"].isin(hah_to_roster) & ~orders_df["OrderStatus"].str.lower().str.contains("cancel", na=False)].copy()

    # Parse dates
    date_cols_q = [c for c in qdf.columns if "date" in c.lower() and "creat" in c.lower()]
    date_cols_o = [c for c in odf.columns if "date" in c.lower() and "book" in c.lower()]
    if not date_cols_q or not date_cols_o:
        return existing_trends  # fallback

    qdf["_dt"] = pd.to_datetime(qdf[date_cols_q[0]], format="mixed")
    odf["_dt"] = pd.to_datetime(odf[date_cols_o[0]], format="mixed")

    qdf = qdf[qdf["_dt"] <= end_dt]
    odf = odf[odf["_dt"] <= end_dt]

    qdf["week"] = qdf["_dt"].dt.to_period("W-SUN")
    odf["week"] = odf["_dt"].dt.to_period("W-SUN")

    q_id_col = next((c for c in qdf.columns if "quote" in c.lower() and "id" in c.lower()), None)
    o_id_col = next((c for c in odf.columns if "order" in c.lower() and "#" in c.lower()), None)

    wq = qdf.groupby("week")[q_id_col].nunique() if q_id_col else qdf.groupby("week").size()
    wo = odf.groupby("week")[o_id_col].nunique() if o_id_col else odf.groupby("week").size()

    all_weeks = sorted(set(wq.index) | set(wo.index))
    last8 = all_weeks[-8:]
    weekly_trends = []
    for w in last8:
        s = w.start_time; e = w.end_time
        lbl = f"{s.month}/{s.day}-{e.month}/{e.day}"
        weekly_trends.append({"week": lbl, "quotes": int(wq.get(w, 0)), "orders": int(wo.get(w, 0))})
    return weekly_trends


def build_wow_data(quotes_df, orders_df, hah_to_roster, end_date):
    """Compute this-week-vs-last-week quotes/orders per agent and per dept.

    Uses the same date-filter + roster-mapping approach as build_weekly_trends,
    just windowed to exactly two 7-day periods instead of grouped by week.
    """
    import pandas as pd

    end_dt   = pd.Timestamp(end_date + " 23:59:59")
    cur_start  = end_dt - timedelta(days=6, hours=23, minutes=59, seconds=59)
    prev_end   = end_dt - timedelta(days=7)
    prev_start = prev_end - timedelta(days=6, hours=23, minutes=59, seconds=59)

    qdf = quotes_df[quotes_df["CSR_Name"].isin(hah_to_roster)].copy()
    odf = orders_df[orders_df["CSR_Name"].isin(hah_to_roster) & ~orders_df["OrderStatus"].str.lower().str.contains("cancel", na=False)].copy()

    date_cols_q = [c for c in qdf.columns if "date" in c.lower() and "creat" in c.lower()]
    date_cols_o = [c for c in odf.columns if "date" in c.lower() and "book" in c.lower()]
    if not date_cols_q or not date_cols_o:
        return {}, []

    qdf["_dt"] = pd.to_datetime(qdf[date_cols_q[0]], format="mixed")
    odf["_dt"] = pd.to_datetime(odf[date_cols_o[0]], format="mixed")

    qdf["_roster"] = qdf["CSR_Name"].map(hah_to_roster)
    odf["_roster"] = odf["CSR_Name"].map(hah_to_roster)

    q_id_col = next((c for c in qdf.columns if "quote" in c.lower() and "id" in c.lower()), None)
    o_id_col = next((c for c in odf.columns if "order" in c.lower() and "#" in c.lower()), None)

    def counts_in_window(df, id_col, start, end):
        sub = df[(df["_dt"] >= start) & (df["_dt"] <= end)]
        if id_col:
            return sub.groupby("_roster")[id_col].nunique().to_dict()
        return sub.groupby("_roster").size().to_dict()

    q_cur  = counts_in_window(qdf, q_id_col, cur_start, end_dt)
    q_prev = counts_in_window(qdf, q_id_col, prev_start, prev_end)
    o_cur  = counts_in_window(odf, o_id_col, cur_start, end_dt)
    o_prev = counts_in_window(odf, o_id_col, prev_start, prev_end)

    agent_names = list(DEPT_MAP.keys()) + [HAH_AGENT_NAME]
    agent_wow = []
    for name in agent_names:
        dept, loc = DEPT_MAP.get(name, ("HaH", "—"))
        row = {
            "name": name, "dept": dept, "loc": loc,
            "quotes_cur":  int(q_cur.get(name, 0)),
            "quotes_prev": int(q_prev.get(name, 0)),
            "orders_cur":  int(o_cur.get(name, 0)),
            "orders_prev": int(o_prev.get(name, 0)),
        }
        if row["quotes_cur"] or row["quotes_prev"] or row["orders_cur"] or row["orders_prev"]:
            agent_wow.append(row)

    dept_wow = {}
    for dept in ("Sales", "SST", "OB"):
        grp = [a for a in agent_wow if a["dept"] == dept]
        dept_wow[dept] = {
            "quotes_cur":  sum(a["quotes_cur"]  for a in grp),
            "quotes_prev": sum(a["quotes_prev"] for a in grp),
            "orders_cur":  sum(a["orders_cur"]  for a in grp),
            "orders_prev": sum(a["orders_prev"] for a in grp),
        }

    return dept_wow, agent_wow


# ── Agent performance CSV: trailing 7 days + all-time, per agent ──────────────
def report_window(today=None):
    """Return (start, end) timestamps for a rolling trailing-7-day window ending
    today, regardless of which day of the week the script is run. Simple and
    consistent — always answers "who's been active in the last week," not tied
    to a Mon/Thu run cadence."""
    today = today or datetime.now()
    end   = today.replace(hour=23, minute=59, second=59, microsecond=0)
    start = (end - timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)
    return pd.Timestamp(start), pd.Timestamp(end)


def build_agent_performance_csv(quotes_df, orders_df, hah_to_roster, agents, report_end_date):
    """Write a per-agent CSV: trailing-7-day quotes/orders (per report_window) + all-time quotes/orders.

    report_end_date (the report's week-ending Sunday, same as il_data.js) names the
    output file, so re-running for the same reporting week overwrites one file
    instead of piling up a new file per calendar day the script happens to run.
    """
    start_dt, end_dt = report_window()

    qdf = quotes_df[quotes_df["CSR_Name"].isin(hah_to_roster)].copy()
    odf = orders_df[orders_df["CSR_Name"].isin(hah_to_roster) & ~orders_df["OrderStatus"].str.lower().str.contains("cancel", na=False)].copy()

    date_cols_q = [c for c in qdf.columns if "date" in c.lower() and "creat" in c.lower()]
    date_cols_o = [c for c in odf.columns if "date" in c.lower() and "book" in c.lower()]
    if not date_cols_q or not date_cols_o:
        print("  WARNING: could not detect date columns — skipping agent performance CSV")
        return None

    qdf["_dt"] = pd.to_datetime(qdf[date_cols_q[0]], format="mixed")
    odf["_dt"] = pd.to_datetime(odf[date_cols_o[0]], format="mixed")
    qdf["_roster"] = qdf["CSR_Name"].map(hah_to_roster)
    odf["_roster"] = odf["CSR_Name"].map(hah_to_roster)

    q_id_col = next((c for c in qdf.columns if "quote" in c.lower() and "id" in c.lower()), None)
    o_id_col = next((c for c in odf.columns if "order" in c.lower() and "#" in c.lower()), None)

    qsub = qdf[(qdf["_dt"] >= start_dt) & (qdf["_dt"] <= end_dt)]
    osub = odf[(odf["_dt"] >= start_dt) & (odf["_dt"] <= end_dt)]

    q_period = qsub.groupby("_roster")[q_id_col].nunique().to_dict() if q_id_col else qsub.groupby("_roster").size().to_dict()
    o_period = osub.groupby("_roster")[o_id_col].nunique().to_dict() if o_id_col else osub.groupby("_roster").size().to_dict()

    rows = []
    for a in agents:
        name = a["name"]
        rows.append({
            "Agent":               name,
            "Dept":                a["dept"],
            "Location":            a["loc"],
            "Last 7 Days Quotes":  int(q_period.get(name, 0)),
            "Last 7 Days Orders":  int(o_period.get(name, 0)),
            "All-Time Quotes":     a["quoted"],
            "All-Time Orders":     a["booked"],
        })
    rows.sort(key=lambda r: (-r["Last 7 Days Orders"], -r["Last 7 Days Quotes"], r["Agent"]))

    fname = f"agent_performance_{report_end_date}.xlsx"
    out_path = os.path.join(LOCAL_DIR, fname)
    write_agent_performance_xlsx(rows, out_path)
    print(f"  ✓ Agent performance XLSX ({start_dt.date()} to {end_dt.date()}): {out_path}")
    return out_path


def write_agent_performance_xlsx(rows, out_path):
    """Write agent performance rows to .xlsx, coloring the Last 7 Days Orders cell
    red/yellow/green based on PERF_GREEN_MIN / PERF_YELLOW_MIN thresholds."""
    RED    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    GREEN  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    headers = ["Agent", "Dept", "Location", "Last 7 Days Quotes", "Last 7 Days Orders", "All-Time Quotes", "All-Time Orders"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Agent Performance"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    orders_col = headers.index("Last 7 Days Orders") + 1
    for row in rows:
        ws.append([row[h] for h in headers])
        r = ws.max_row
        orders = row["Last 7 Days Orders"]
        fill = RED if orders < PERF_YELLOW_MIN else (YELLOW if orders < PERF_GREEN_MIN else GREEN)
        ws.cell(row=r, column=orders_col).fill = fill

    for i, h in enumerate(headers, start=1):
        width = max(len(h), max((len(str(row[h])) for row in rows), default=0)) + 2
        ws.column_dimensions[get_column_letter(i)].width = width

    wb.save(out_path)


def build_il_data_js(end_date, agents, dept_stats, cancel_data, weekly_trends,
                     prev_orders, prev_quotes, prev_savings, dept_wow, agent_wow,
                     total_containers_by_agent=None):
    total_containers_by_agent = total_containers_by_agent or {}
    total_quotes  = sum(a["quoted"] for a in agents)
    total_orders  = sum(a["booked"] for a in agents)
    conv_rate     = round(total_orders / total_quotes * 100, 1) if total_quotes else 0
    # Adoption counts must exclude the synthetic "HaH Agent" rollup row — it is a bucket
    # for non-PODS names, not a person. Including it made total_agents 130 against a
    # 129-agent roster, so "X of 130 active" was measured against a fake denominator.
    real_agents   = [a for a in agents if a["name"] != HAH_AGENT_NAME]
    total_agents  = len(real_agents)
    active_agents = sum(1 for a in real_agents if a["quoted"] > 0)

    savings        = cancel_data["savings"]
    delta_7d       = cancel_data["delta_7d"]
    delta_30d      = cancel_data["delta_30d"]
    primary_period = cancel_data["primary_period"]
    by_month       = cancel_data.get("by_month", [])
    detail_7d      = cancel_data.get("detail_7d")  or {}
    detail_30d     = cancel_data.get("detail_30d") or {}
    attach_rate    = cancel_data.get("attach_rate")
    attach_detail  = cancel_data.get("attach_detail") or {}
    cancel_range   = cancel_data.get("range") or {}
    if attach_rate is None:
        attach_rate = 0.0            # Snowflake path failed; loader below supplies the old value

    # NOTE: annualized_savings was removed 2026-08-03. It divided cumulative savings by
    # the full elapsed period, but `savings` only covers bookings up to the 30-day
    # maturity cutoff — so the divisor spanned about a month more than the numerator.
    # It also annualised an effect measured on ~11 containers. Not worth repairing.
    launch_dt         = datetime.strptime(LAUNCH_DATE, "%Y-%m-%d")
    end_dt             = datetime.strptime(end_date, "%Y-%m-%d")
    months_elapsed     = max((end_dt - launch_dt).days / 30.44, 1 / 30.44)
    monthly_savings    = savings / months_elapsed
    monthly_per_agent  = monthly_savings / active_agents if active_agents else 0
    projected_total    = round(monthly_per_agent * ROLLOUT_TARGET_AGENTS * 12)
    projected_str      = f"~${projected_total:,}"
    # Inputs behind the projection, so the card can show its own working. Two known
    # weaknesses are deliberately surfaced rather than hidden: months_elapsed spans the
    # full period while `savings` only covers bookings up to the 30-day maturity cutoff,
    # and active_agents counts anyone who QUOTED, not the smaller set that actually
    # produced the savings. Both understate the per-agent rate; correcting them roughly
    # triples the projection, which is itself a signal the linear model is fragile.
    projected_detail = {
        "savings":            savings,
        "months_elapsed":     round(months_elapsed, 2),
        "monthly_savings":    round(monthly_savings),
        "active_agents":      active_agents,
        "monthly_per_agent":  round(monthly_per_agent),
        "rollout_agents":     ROLLOUT_TARGET_AGENTS,
        "containers_actual":  round(savings / ECR, 1),
        "containers_implied": round(projected_total / ECR),
    }

    # Silent grid from agents
    silent_grid = []
    for loc in ("CLW", "CNX"):
        for dept in ("Sales", "SST", "OB"):
            grp = [a for a in agents if a["dept"] == dept and a["loc"] == loc]
            silent_grid.append({
                "loc": loc, "dept": dept,
                "total":  len(grp),
                "silent": sum(1 for a in grp if a["quoted"] == 0),
            })

    # ── Opportunity table — COMPUTED (was hardcoded "~3,320"/"~442" strings) ──────
    # For each department: how many containers its SILENT agents booked, and how many
    # would have carried labor if they attached at the same rate as agents who already
    # sell it. Benchmarks against active agents, not the overall roster — the roster
    # rate includes the silent agents themselves and would understate the gap.
    active_attach = None
    if attach_detail.get("labor_containers") and total_containers_by_agent:
        selling = {a["name"] for a in agents if a.get("booked", 0) > 0 and a["name"] != HAH_AGENT_NAME}
        sell_ctr = sum(total_containers_by_agent.get(n, 0) for n in selling)
        if sell_ctr:
            active_attach = attach_detail["labor_containers"] / sell_ctr

    opp_table = []
    for dept, priority in (("OB", "Critical"), ("Sales", "High"), ("SST", "Moderate")):
        ds = dept_stats.get(dept, {})
        total, active = ds.get("total", 0), ds.get("active", 0)
        silent = total - active
        silent_names = [a["name"] for a in agents
                        if a["dept"] == dept and a["quoted"] == 0 and a["name"] != HAH_AGENT_NAME]
        silent_ctr = sum(total_containers_by_agent.get(n, 0) for n in silent_names)
        missed = round(silent_ctr * active_attach) if active_attach else None
        opp_table.append({
            "dept": dept,
            "desc": f"{silent} of {total} silent" if silent != total else f"All {total} silent",
            "silent": silent,
            "est_containers": f"{silent_ctr:,}" if silent_ctr else "—",
            "est_missed": f"{missed:,}" if missed is not None else "—",
            "priority": priority,
        })
    if active_attach:
        print(f"  Opportunity table: benchmarked at {active_attach*100:.2f}% container attach "
              f"(agents who already sell labor)")

    sales = dept_stats.get("Sales", {})
    sst   = dept_stats.get("SST",   {})
    ob    = dept_stats.get("OB",    {})

    content = f"""// Integrated Labor Dashboard Data — auto-generated by update_il_dashboard.py
// Report week ending: {end_date}

const IL_DATA = {{
  updated_date: "{fmt_date(end_date)}",

  // Program totals (cumulative since launch May 1, 2026)
  total_quotes:  {total_quotes},
  total_orders:  {total_orders},
  conv_rate:     {conv_rate},

  // Cancel rate impact
  cancel_delta_7d:  {delta_7d},
  cancel_delta_30d: {delta_30d},
  cancel_delta_primary_period: {primary_period},   // always 30 — more sample, less week-to-week whiplash than the 7-day window
  cancel_savings:   {savings},
  ecr:              {ECR},   // avg revenue from one container — converts saved containers to $

  // 95% band on the 30-day delta and the savings that follow from it. Effective sample
  // is labor ORDERS, not containers — containers on the same order cancel together.
  cancel_range: {json.dumps(cancel_range, indent=2)},

  // Cancel rate by BOOKING-MONTH cohort. The blended delta above hides drift —
  // a strong early month keeps propping it up long after the effect has faded.
  // mature=false means that month's orders haven't all had 30 days to cancel yet.
  cancel_by_month: {json.dumps(by_month, indent=2)},

  // Raw numbers behind each delta — powers the hover tooltips on the cancel cards.
  cancel_detail_7d:  {json.dumps(detail_7d, indent=2)},
  cancel_detail_30d: {json.dumps(detail_30d, indent=2)},
  projected_rollout_savings: "{projected_str}",
  projected_rollout_agents:  {ROLLOUT_TARGET_AGENTS},
  projected_rollout_detail: {json.dumps(projected_detail, indent=2)},

  // Computed, not manual. Distinct PODS orders with labor / all orders booked by roster
  // agents, since launch. NOT maturity-filtered — attachment is known at booking.
  labor_attach_rate: {attach_rate},
  labor_attach_detail: {json.dumps(attach_detail, indent=2)},

  // Adoption
  active_agents: {active_agents},
  total_agents:  {total_agents},

  // Prior week values (for trajectory narrative)
  prev_orders:  {prev_orders},
  prev_quotes:  {prev_quotes},
  prev_savings: {prev_savings},

  // Department breakdown
  depts: {{
    Sales: {{ total: {sales.get("total",0)}, active: {sales.get("active",0)}, orders: {sales.get("orders",0)}, quotes: {sales.get("quotes",0)}, pitch_rate: {sales.get("pitch_rate",0)}, conv_rate: {sales.get("conv_rate",0)} }},
    SST:   {{ total: {sst.get("total",0)},   active: {sst.get("active",0)},   orders: {sst.get("orders",0)},   quotes: {sst.get("quotes",0)},   pitch_rate: {sst.get("pitch_rate",0)},   conv_rate: {sst.get("conv_rate",0)}   }},
    OB:    {{ total: {ob.get("total",0)},    active: {ob.get("active",0)},    orders: {ob.get("orders",0)},    quotes: {ob.get("quotes",0)},    pitch_rate: {ob.get("pitch_rate",0)},    conv_rate: {ob.get("conv_rate",0)}    }},
  }},

  // Silent agent breakdown by location + dept
  silent_grid: {json.dumps(silent_grid, indent=2)},

  // Weekly new quotes & orders (last 8 weeks) — for stacked bar chart
  weekly_trends: {json.dumps(weekly_trends, indent=2)},

  // This week vs last week, by department
  dept_wow: {json.dumps(dept_wow, indent=2)},

  // This week vs last week, by agent (only agents with activity in either week)
  agent_wow: {json.dumps(agent_wow, indent=2)},

  // Opportunity table
  opp_table: {json.dumps(opp_table, indent=2)},

  // Agent roster (cumulative since launch)
  agents: {json.dumps(agents, indent=2)},
}};
"""
    return content

# ── Main ────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 50)
    print("  IL Dashboard Update")
    print("=" * 50)
    cfg = load_config()

    end_date = last_sunday()
    print(f"\nReport week ending: {end_date}\n")

    # Pick files
    print("Select the HaH Quotes CSV (cumulative since launch)...")
    quotes_path = pick_file("Select HaH Quotes CSV — cumulative since 5/1/26")
    print(f"  ✓ {os.path.basename(quotes_path)}")

    print("Select the HaH Orders CSV (cumulative since launch)...")
    orders_path = pick_file("Select HaH Orders CSV — cumulative since 5/1/26")
    print(f"  ✓ {os.path.basename(orders_path)}")

    # Process HaH CSVs
    print("\n[1/4] Processing HaH CSVs...")
    quotes_export = read_csv_auto(quotes_path)
    orders_df_raw = read_csv_auto(orders_path)
    check_csv_freshness(quotes_export, orders_df_raw)

    # Merge this week's exports into the archives and work from full history, not the
    # capped 1000-row exports. Everything downstream (KPIs, trends, WoW, agent CSV,
    # cancel rate) reads quotes_df_raw / orders_df_raw, so they all get the complete series.
    quotes_df_raw = merge_quotes_archive(quotes_export)
    orders_df_raw = merge_orders_archive(orders_df_raw)

    roster_names = list(DEPT_MAP.keys())
    orders_df_raw, reattr_stats = reattribute_hah_orders(quotes_df_raw, orders_df_raw, roster_names)
    if reattr_stats["reattributed"]:
        print(f"  Re-attributed {reattr_stats['reattributed']} HaH-touched order(s) to their originating "
              f"PODS agent via PO#/CID match ({reattr_stats['unresolved']} unresolved, stay bucketed as HaH Agent)")

    q_by_agent, o_by_agent, all_hah_names = process_hah_csvs(quotes_df_raw, orders_df_raw, end_date)
    print(f"  {len(all_hah_names)} unique agents across quotes + orders")

    # Build agent data
    print("\n[2/4] Building agent data...")
    prev_agents = load_existing_agents()
    agents      = build_agents(q_by_agent, o_by_agent, all_hah_names, prev_agents)
    active     = sum(1 for a in agents if a["quoted"] > 0)
    total      = len(agents)
    print(f"  {active} active / {total} total agents")

    # Snowflake — total orders per agent, new-agent detection, and cancel rate
    # (labor vs. regular, via the PO#/CID-then-OrderNumber join — see
    # compute_cancel_rate_from_snowflake). This replaced the weekly Tableau
    # "Labor Impact Table" CSV pull.
    print("\n[3/4] Connecting to Snowflake (browser SSO will open)...")
    total_orders_by_agent = {}
    total_containers_by_agent = {}
    cancel_data = None
    try:
        conn = get_snowflake_conn(cfg)
        print("  Connected.")
        directory = load_pods_employee_directory(conn)
        agent_username_map = build_agent_username_map(directory, roster_names)
        total_orders_by_agent, total_containers_by_agent = q_total_orders_by_agent(conn, end_date, agent_username_map)
        print(f"  Total orders by agent: {len(total_orders_by_agent)} agents mapped "
              f"(usernames resolved for {len(agent_username_map)}/{len(roster_names)} roster agents)")

        print("  Checking for new PODS agents not yet in DEPT_MAP...")
        new_agents, _ = detect_new_pods_agents(all_hah_names, directory, list(DEPT_MAP.keys()))
        if new_agents:
            print(f"\n  🆕 {len(new_agents)} NEW PODS AGENT(S) DETECTED — auto-classified for this run only:")
            for name, info in sorted(new_agents.items()):
                DEPT_MAP[name] = (info["dept"], info["loc"])
                print(f'      "{name}": ("{info["dept"]}", "{info["loc"]}"),  # {info["email"]} | org: {info["organization"]}')
            print("  ⚠ Review the dept/location guesses above (esp. Sales vs SST) and paste any that are")
            print("    correct into DEPT_MAP in update_il_dashboard.py to make them permanent.")
        else:
            print("  No new PODS agents detected.")

        print("  Checking roster spreadsheet for trained agents missing from DEPT_MAP...")
        untracked = detect_untracked_trained_agents(list(DEPT_MAP.keys()))
        if untracked:
            print(f"\n  ⚠ {len(untracked)} agent(s) on the roster spreadsheet not in DEPT_MAP yet (dept unknown — fill in Sales/SST/OB):")
            for a in sorted(untracked, key=lambda x: x["name"]):
                sup = f" | supervisor: {a['supervisor']}" if a.get("supervisor") else ""
                print(f'      "{a["name"]}": ("???", "{a["loc"]}"),  # {a["email"]}{sup} | from "{a["sheet"]}"')
            print("  ⚠ Confirm department for each, then paste into DEPT_MAP in update_il_dashboard.py.")
        else:
            print("  No gaps found — DEPT_MAP matches the roster spreadsheet.")

        print("\n[4/5] Computing cancel rate from Snowflake (PO#/CID join)...")
        cancel_data = compute_cancel_rate_from_snowflake(conn, directory, roster_names, orders_df_raw, end_date)
        print(f"  7d delta: {cancel_data['delta_7d']}pp | 30d delta: {cancel_data['delta_30d']}pp | savings: ${cancel_data['savings']:,}")

        conn.close()
    except Exception:
        import traceback
        traceback.print_exc()
        print("  WARNING: Snowflake query failed — pitch rate will show 0, new-agent detection skipped")

    if cancel_data is None:
        print("  Keeping existing cancel rate values from il_data.js")
        cancel_data = load_existing_cancel_data()

    dept_stats = compute_dept_stats(agents, total_orders_by_agent)
    for dept, ds in dept_stats.items():
        print(f"    {dept}: {ds['active']}/{ds['total']} active, {ds['quotes']} quotes, {ds['orders']} orders")

    # Load previous totals for trajectory narrative
    prev_orders, prev_quotes, prev_savings, _, _ = load_prev_totals()
    existing_trends = load_existing_weekly_trends()

    # Build weekly trends directly from date-filtered CSVs (no cumulative delta math)
    import pandas as pd
    # quotes_df_raw / orders_df_raw already loaded + reattributed above
    # Rebuild hah_to_roster for weekly calc
    all_names_raw = set(quotes_df_raw.get("CSR_Name", pd.Series()).dropna()) | \
                    set(orders_df_raw.get("CSR_Name", pd.Series()).dropna())
    hah_to_roster_map = {}
    for n in all_names_raw:
        m = fuzzy_match(n, roster_names)
        hah_to_roster_map[n] = m if m else HAH_AGENT_NAME
    weekly_trends = build_weekly_trends(quotes_df_raw, orders_df_raw, hah_to_roster_map, end_date, existing_trends)
    print(f"  Weekly trends: {len(weekly_trends)} weeks")
    for w in weekly_trends:
        print(f"    {w['week']}: {w['quotes']} quotes, {w['orders']} orders")

    # Week-over-week breakdown (dept + agent level)
    dept_wow, agent_wow = build_wow_data(quotes_df_raw, orders_df_raw, hah_to_roster_map, end_date)
    print(f"  WoW (this week vs last week):")
    for dept, dw in dept_wow.items():
        print(f"    {dept}: quotes {dw['quotes_prev']} -> {dw['quotes_cur']} | orders {dw['orders_prev']} -> {dw['orders_cur']}")

    # Agent performance CSV (2x/week: Monday = Fri-Sun, Thursday = Mon-Thu)
    print("\n  Building agent performance CSV...")
    build_agent_performance_csv(quotes_df_raw, orders_df_raw, hah_to_roster_map, agents, end_date)

    # Generate content
    print("\n[5/5] Writing and pushing...")
    il_data_content = build_il_data_js(
        end_date, agents, dept_stats, cancel_data,
        weekly_trends, prev_orders, prev_quotes, prev_savings,
        dept_wow, agent_wow,
        total_containers_by_agent=total_containers_by_agent
    )

    # Write il_data.js locally
    local_path = os.path.join(LOCAL_DIR, "il_data.js")
    with open(local_path, "w", encoding="utf-8") as f:
        f.write(il_data_content)
    print(f"  ✓ Written locally: {local_path}")

    # Push to GitHub
    token = cfg["github_token"]
    existing = gh_get(token, "il_data.js")
    gh_put(token, "il_data.js", il_data_content.encode("utf-8"), existing["sha"],
           f"Update IL dashboard data — week ending {end_date}")

    total_q = sum(a["quoted"] for a in agents)
    total_o = sum(a["booked"] for a in agents)
    print(f"\n✅ Done!")
    print(f"   Quotes: {total_q:,}  |  Orders: {total_o:,}  |  Active agents: {active}/{total}")
    print(f"   Cancel savings: ${cancel_data['savings']:,}")
    print(f"   GitHub: https://github.com/{GITHUB_REPO}/blob/main/il_data.js")

if __name__ == "__main__":
    if "--backfill-orders" in sys.argv:
        backfill_orders_archive()
    elif "--backfill" in sys.argv:
        backfill_quotes_archive()
    else:
        main()
