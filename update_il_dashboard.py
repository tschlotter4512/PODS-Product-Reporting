#!/usr/bin/env python3
"""
update_il_dashboard.py
Weekly automation for the Integrated Labor dashboard.
Run every Monday after downloading HaH Quotes and Orders CSVs.
"""

import json, base64, re, os, sys
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import requests
import snowflake.connector

# ── Config ─────────────────────────────────────────────────────────────────────
CONFIG_PATH = r"F:\Revenue Management\Team Individual Folders\TSchlotter\PODS-Product-Reporting\config.json"
LOCAL_DIR   = r"F:\Revenue Management\Team Individual Folders\TSchlotter\PODS-Product-Reporting"
GITHUB_REPO = "tschlotter4512/PODS-Product-Reporting"
LAUNCH_DATE = "2026-05-01"
ECR                = 3700   # expected container revenue ($)
AVG_CTR_PER_ORDER  = 1.3    # average containers per order

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
}

# ── Helpers ────────────────────────────────────────────────────────────────────
# ── Agent username map: canonical name → ORDER_CONTAINER_SALES_FACT "Agent Container Booked By" ──
AGENT_USERNAME_MAP = {
    "Amy Sisterman":          "AMSisterman",
    "Catherine Saitta":       "CSaitta",
    "Ted Johnson":            "TWJohnson",
    "Carmie Nillama":         "CNillama",
    "Jeanielle Navarro":      "JNavarro2",
    "Johnny Cardinale":       "JCardinale",
    "Donald Alfaro":          "dalfaro",
    "Nicole Slater":          "NHSlater",
    "Felicia Knight":         "fknight",
    "Kimberly Thebeau":       "KThebeau",
    "Ednalyn Mirandilla":     "EMirandilla",
    "Almira Castro":          "ACastro3",
    "Katleen May Castro":     "KMCastro",
    "Charmaine Bacay":        "CBacay",
    "Sara Messenger":         "SMessenger",
    "Christina Morales":      "cmorales",
    "Christopher Monaldi":    "CMonaldi",
    "Jake Russel Bernardino": "JRBernardino",
    "Irish Verona":           "IVerona",
    "James Onnagan":          "JOnnagan",
    "James Rebualos":         "JRebualos",
    "Jerico Galve":           "JGalve",
    "Princess Heidi Sesma":   "PHSesma",
}

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

# ── HaH CSV processing ─────────────────────────────────────────────────────────
def process_hah_csvs(quotes_path, orders_path):
    quotes_df = read_csv_auto(quotes_path)
    orders_df = read_csv_auto(orders_path)

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

    for _, row in df.iterrows():
        agent_name = str(row[name_col]).strip()
        matched = fuzzy_match(agent_name, roster_names)
        if not matched:
            continue
        l7  = pct(row[l7_col])
        r7  = pct(row[r7_col])
        l30 = pct(row[l30_col])
        r30 = pct(row[r30_col])
        if l7 is None or r7 is None or l30 is None or r30 is None:
            continue
        weight = o_by_roster.get(matched, 0) or 1  # fall back to 1 so agent isn't dropped
        w7           += (l7  - r7)  * weight
        w30          += (l30 - r30) * weight
        total_weight += weight

    if total_weight == 0:
        raise ValueError("No IL agents with labor cancel data found in the CSV.")

    delta_7d  = round(w7  / total_weight * 100, 1)
    delta_30d = round(w30 / total_weight * 100, 1)
    savings   = round(abs(delta_30d) / 100 * total_weight * AVG_CTR_PER_ORDER * ECR)

    return {"delta_7d": delta_7d, "delta_30d": delta_30d, "savings": savings}

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
def build_agents(q_by_hah, o_by_hah, all_hah_names, prev_agents):
    roster_names = list(DEPT_MAP.keys())
    skipped = []

    # Map HaH name → roster canonical name; skip non-roster names (HaH agents)
    hah_to_roster = {}
    for name in all_hah_names:
        matched = fuzzy_match(name, roster_names)
        if matched:
            hah_to_roster[name] = matched
        else:
            skipped.append(name)

    if skipped:
        print(f"  Skipped {len(skipped)} unrecognized names (HaH agents or new trainees):")
        for s in sorted(skipped):
            print(f"    - {s}")
        print("  To add a new PODS agent, add them to DEPT_MAP in the script.")

    # Accumulate quotes/orders to roster canonical names
    q_totals, o_totals = {}, {}
    for hah_name, roster_name in hah_to_roster.items():
        q_totals[roster_name] = q_totals.get(roster_name, 0) + q_by_hah.get(hah_name, 0)
        o_totals[roster_name] = o_totals.get(roster_name, 0) + o_by_hah.get(hah_name, 0)

    agents = []
    for name, (dept, loc) in DEPT_MAP.items():
        prev = prev_agents.get(name, {})
        # Use CSV value if present; fall back to previous value (never go below prior)
        quoted = max(q_totals.get(name, 0), prev.get("quoted", 0))
        booked = max(o_totals.get(name, 0), prev.get("booked", 0))
        agents.append({"name": name, "dept": dept, "loc": loc, "quoted": quoted, "booked": booked})

    agents.sort(key=lambda a: (-a["booked"], -a["quoted"], a["name"]))
    return agents

def q_total_orders_by_agent(conn, end_date):
    """Query total orders per IL agent from ORDER_CONTAINER_SALES_FACT using username map."""
    usernames = list(AGENT_USERNAME_MAP.values())
    username_list = ", ".join(f"'{u}'" for u in usernames)
    sql = f"""
    SELECT "Agent Container Booked By", COUNT(DISTINCT "Order Number") AS total_orders
    FROM PROD_DW.DW_SEMANTIC.ORDER_CONTAINER_SALES_FACT
    WHERE "Date Order Booked" >= '{LAUNCH_DATE}'
      AND "Date Order Booked" <= '{end_date}'
      AND "Agent Container Booked By" IN ({username_list})
    GROUP BY "Agent Container Booked By"
    """
    cur = conn.cursor()
    cur.execute(sql)
    username_to_orders = {row[0]: row[1] for row in cur.fetchall()}
    # Reverse map: username → roster name
    username_to_roster = {v: k for k, v in AGENT_USERNAME_MAP.items()}
    return {username_to_roster[u]: cnt for u, cnt in username_to_orders.items() if u in username_to_roster}

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
    return {
        "delta_7d":  extract_float("cancel_delta_7d"),
        "delta_30d": extract_float("cancel_delta_30d"),
        "savings":   extract_int("cancel_savings"),
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


def build_il_data_js(end_date, agents, dept_stats, cancel_data, weekly_trends,
                     prev_orders, prev_quotes, prev_savings):
    total_quotes  = sum(a["quoted"] for a in agents)
    total_orders  = sum(a["booked"] for a in agents)
    conv_rate     = round(total_orders / total_quotes * 100, 1) if total_quotes else 0
    total_agents  = len(agents)
    active_agents = sum(1 for a in agents if a["quoted"] > 0)

    savings       = cancel_data["savings"]
    delta_7d      = cancel_data["delta_7d"]
    delta_30d     = cancel_data["delta_30d"]

    # Annualized = current monthly run rate × 12
    monthly_savings = savings / max(1, active_agents)
    annualized      = round(monthly_savings * 12 * active_agents)
    proj_per_agent  = savings / active_agents if active_agents else 0
    projected_total = round(proj_per_agent * 220 * 12)
    projected_str   = f"~${projected_total:,}"

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

    # Opportunity table (static estimates — update if you have container denominators)
    opp_table = [
        {"dept": "OB",    "desc": f"All {dept_stats.get('OB',{}).get('total',43)} fully trained",
         "silent": dept_stats.get("OB",{}).get("total",43) - dept_stats.get("OB",{}).get("active",0),
         "est_containers": "~3,320", "est_missed": "~442", "priority": "Critical"},
        {"dept": "Sales", "desc": f"{dept_stats.get('Sales',{}).get('total',15) - dept_stats.get('Sales',{}).get('active',0)} silent",
         "silent": dept_stats.get("Sales",{}).get("total",15) - dept_stats.get("Sales",{}).get("active",0),
         "est_containers": "~1,401", "est_missed": "~187", "priority": "High"},
        {"dept": "SST",   "desc": f"{dept_stats.get('SST',{}).get('total',18) - dept_stats.get('SST',{}).get('active',0)} silent",
         "silent": dept_stats.get("SST",{}).get("total",18) - dept_stats.get("SST",{}).get("active",0),
         "est_containers": "~361", "est_missed": "~48", "priority": "Moderate"},
    ]

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
  cancel_savings:   {savings},
  annualized_savings:        {annualized},
  projected_rollout_savings: "{projected_str}",
  projected_rollout_agents:  220,

  labor_attach_rate: 1.7,   // manual — update when Snowflake container denominator available

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

    print("Select the Tableau Labor Impact Table CSV...")
    tableau_csv_path = pick_file(
        "Select Tableau Labor Impact Table CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    print(f"  ✓ {os.path.basename(tableau_csv_path)}")

    # Process HaH CSVs
    print("\n[1/4] Processing HaH CSVs...")
    q_by_agent, o_by_agent, all_hah_names = process_hah_csvs(quotes_path, orders_path)
    print(f"  {len(all_hah_names)} unique agents across quotes + orders")

    # Snowflake — total orders per agent for pitch rate
    print("\n[2/4] Connecting to Snowflake for pitch rate (browser SSO will open)...")
    total_orders_by_agent = {}
    try:
        conn = get_snowflake_conn(cfg)
        print("  Connected.")
        total_orders_by_agent = q_total_orders_by_agent(conn, end_date)
        print(f"  Total orders by agent: {len(total_orders_by_agent)} agents mapped")
        conn.close()
    except Exception:
        import traceback
        traceback.print_exc()
        print("  WARNING: Snowflake query failed — pitch rate will show 0")

    # Build agent data
    print("\n[3/4] Building agent data...")
    prev_agents = load_existing_agents()
    agents      = build_agents(q_by_agent, o_by_agent, all_hah_names, prev_agents)
    dept_stats  = compute_dept_stats(agents, total_orders_by_agent)
    active     = sum(1 for a in agents if a["quoted"] > 0)
    total      = len(agents)
    print(f"  {active} active / {total} total agents")
    for dept, ds in dept_stats.items():
        print(f"    {dept}: {ds['active']}/{ds['total']} active, {ds['quotes']} quotes, {ds['orders']} orders")

    # Cancel rate from Tableau CSV (weighted by each agent's labor order count)
    print("\n[4/5] Computing cancel rate from Tableau CSV...")
    o_by_roster = {a["name"]: a["booked"] for a in agents}
    try:
        cancel_data = read_cancel_rate_from_tableau_csv(tableau_csv_path, o_by_roster)
        print(f"  7d delta: {cancel_data['delta_7d']}pp | 30d delta: {cancel_data['delta_30d']}pp | savings: ${cancel_data['savings']:,}")
    except Exception:
        import traceback
        traceback.print_exc()
        print("  Keeping existing cancel rate values from il_data.js")
        cancel_data = load_existing_cancel_data()

    # Load previous totals for trajectory narrative
    prev_orders, prev_quotes, prev_savings, _, _ = load_prev_totals()
    existing_trends = load_existing_weekly_trends()

    # Build weekly trends directly from date-filtered CSVs (no cumulative delta math)
    import pandas as pd
    quotes_df_raw = read_csv_auto(quotes_path)
    orders_df_raw = read_csv_auto(orders_path)
    # Rebuild hah_to_roster for weekly calc
    roster_names = list(DEPT_MAP.keys())
    all_names_raw = set(quotes_df_raw.get("CSR_Name", pd.Series()).dropna()) | \
                    set(orders_df_raw.get("CSR_Name", pd.Series()).dropna())
    hah_to_roster_map = {}
    for n in all_names_raw:
        m = fuzzy_match(n, roster_names)
        if m: hah_to_roster_map[n] = m
    weekly_trends = build_weekly_trends(quotes_df_raw, orders_df_raw, hah_to_roster_map, end_date, existing_trends)
    print(f"  Weekly trends: {len(weekly_trends)} weeks")
    for w in weekly_trends:
        print(f"    {w['week']}: {w['quotes']} quotes, {w['orders']} orders")

    # Generate content
    print("\n[5/5] Writing and pushing...")
    il_data_content = build_il_data_js(
        end_date, agents, dept_stats, cancel_data,
        weekly_trends, prev_orders, prev_quotes, prev_savings
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
    main()
