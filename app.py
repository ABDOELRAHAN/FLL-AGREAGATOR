import hmac
import html
import math
import numbers
import re
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================
# COURIER OPS — منظومة متابعة المناديب اليومية
# هنجر ستيشن / جاهز / ذا شيفز
# ============================================================

DB_PATH = Path(__file__).parent / "daily_tracker.db"
ADMIN_PASSWORD = "admin 444"

PLATFORMS = {
    "hunger": {"name": "هنجر ستيشن", "icon": "🟠", "target": 17, "color": "#f97316"},
    "jahez": {"name": "جاهز", "icon": "🔴", "target": 7, "color": "#ef4444"},
    "chefz": {"name": "ذا شيفز", "icon": "🟣", "target": 7, "color": "#8b5cf6"},
}

NAV = [
    ("overview", "🏠", "نظرة عامة"),
    ("hunger", "🟠", "هنجر ستيشن"),
    ("jahez", "🔴", "جاهز"),
    ("chefz", "🟣", "ذا شيفز"),
    ("upload", "📥", "رفع البيانات"),
]

NAME_HEADERS = {"name", "ridername", "drivername", "couriername"}
ID_KEYS = ("driverid", "riderid", "courierid")
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}

esc = html.escape


# ------------------------------------------------------------
# Parsing
# ------------------------------------------------------------
def norm(text):
    return re.sub(r"[\s_\-\.]+", "", str(text).strip().lower())


def is_id_header(x):
    n = norm(x)
    return n == "id" or any(k in n for k in ID_KEYS)


def parse_date_header(v, default_year):
    """يفهم: 1-Oct / 01/10 / تاريخ حقيقي / رقم تاريخ Excel."""
    try:
        if v is None or pd.isna(v):
            return None
    except Exception:
        pass
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, numbers.Real) and not isinstance(v, bool):
        if 40000 < float(v) < 60000:
            return date(1899, 12, 30) + timedelta(days=int(v))
        return None
    if not isinstance(v, str):
        return None

    s = v.strip()
    try:
        m = re.match(r"^(\d{1,2})[\s\-/]([A-Za-z]{3,9})\.?(?:[\s\-/,]*(\d{2,4}))?$", s)
        if m:
            month = MONTHS.get(m.group(2)[:3].lower())
            if month:
                year = int(m.group(3)) if m.group(3) else default_year
                year = year + 2000 if year < 100 else year
                return date(year, month, int(m.group(1)))
        m = re.match(r"^(\d{1,2})[/\.\-](\d{1,2})(?:[/\.\-](\d{2,4}))?$", s)
        if m:  # dd/mm
            year = int(m.group(3)) if m.group(3) else default_year
            year = year + 2000 if year < 100 else year
            return date(year, int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None
    return None


def clean_id(v):
    try:
        if v is None or pd.isna(v):
            return None
    except Exception:
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = re.sub(r"\.0$", "", str(v).strip())
    if s.lower() in ("", "nan", "none", "total", "grand total", "المجموع", "الإجمالي"):
        return None
    return s


def clean_name(v):
    try:
        if v is None or pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() in ("nan", "#n/a", "n/a", "#na", "none") else s


def parse_cell(v):
    """('on', طلبات) / ('off', None) / (None, None) لو فاضية."""
    try:
        if v is None or pd.isna(v):
            return None, None
    except Exception:
        pass
    if isinstance(v, str):
        s = v.strip().lower()
        if s in ("off", "o", "leave", "اجازة", "إجازة"):
            return "off", None
        try:
            return "on", int(float(s.replace(",", "")))
        except ValueError:
            return None, None
    try:
        return "on", int(float(v))
    except Exception:
        return None, None


def describe(raw, n_rows=6, n_cols=9):
    lines = []
    for i in range(min(len(raw), n_rows)):
        cells = " | ".join(str(x)[:16] for x in raw.iloc[i].tolist()[:n_cols])
        lines.append(f"صف {i + 1}: {cells}")
    return "\n".join(lines)


def parse_sheet(raw, default_year, default_month=None):
    raw = raw.reset_index(drop=True)

    header_idx = None
    for i in range(min(len(raw), 40)):
        if any(is_id_header(x) for x in raw.iloc[i].tolist()):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("مش لاقي صف العناوين (لازم يكون فيه Rider ID أو DRIVER ID). أول الصفوف اللي شفتها:\n"
                         + describe(raw, 8))

    header = raw.iloc[header_idx].tolist()
    id_col = next(j for j, v in enumerate(header) if is_id_header(v))
    name_col = next((j for j, v in enumerate(header) if norm(v) in NAME_HEADERS), None)

    date_cols = {}
    for j, v in enumerate(header):
        d = parse_date_header(v, default_year)
        if d:
            date_cols[j] = d

    # لو العناوين أرقام أيام بس (1,2,3...) نستخدم الشهر والسنة المختارين
    if not date_cols and default_month:
        for j, v in enumerate(header):
            if j <= id_col:
                continue
            try:
                n = float(v)
            except (TypeError, ValueError):
                continue
            if n == int(n) and 1 <= int(n) <= 31:
                try:
                    date_cols[j] = date(default_year, default_month, int(n))
                except ValueError:
                    pass
        if len(date_cols) < 3:
            date_cols = {}

    if not date_cols:
        cells = " | ".join(str(x)[:16] for x in header[:14])
        raise ValueError("لقيت صف العناوين لكن مفيش أعمدة تواريخ مفهومة. عناوين الصف:\n" + cells)

    rows = []
    for k in range(header_idx + 1, len(raw)):
        r = raw.iloc[k]
        rid = clean_id(r.iloc[id_col])
        if rid is None:
            continue
        name = clean_name(r.iloc[name_col]) if name_col is not None else ""
        for j, d in date_cols.items():
            status, orders = parse_cell(r.iloc[j])
            if status is not None:
                rows.append((rid, name, d, status, orders))

    if not rows:
        raise ValueError("لقيت العناوين لكن مفيش بيانات مناديب تحتها.")

    df = pd.DataFrame(rows, columns=["rider_id", "rider_name", "date", "status", "orders"])
    df = df.drop_duplicates(subset=["rider_id", "date"], keep="last").reset_index(drop=True)
    df["month"] = df["date"].map(lambda d: d.strftime("%Y-%m"))
    return df


def read_best(up, year, month):
    """يجرّب كل الشيتات في الملف ويرجّع اللي اتقرا صح + أخطاء الباقي."""
    ok, errors = {}, {}
    if up.name.lower().endswith(".csv"):
        ok["CSV"] = parse_sheet(pd.read_csv(up, header=None), year, month)
        return ok, errors
    xl = pd.ExcelFile(up)
    for sh in xl.sheet_names:
        try:
            ok[sh] = parse_sheet(pd.read_excel(xl, sheet_name=sh, header=None), year, month)
        except Exception as e:
            errors[sh] = str(e)
    return ok, errors


# ------------------------------------------------------------
# Database
# ------------------------------------------------------------
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS daily (
            platform TEXT NOT NULL, month TEXT NOT NULL, date TEXT NOT NULL,
            rider_id TEXT NOT NULL, rider_name TEXT, status TEXT NOT NULL, orders INTEGER
        )
    """)
    conn.commit()
    conn.close()


def save_platform(platform, df):
    """يمسح الشهور اللي في الملف بس لنفس المنصة ويحط الجديد."""
    out = df.copy()
    out["date"] = out["date"].map(lambda d: d.isoformat())
    out["platform"] = platform
    months = sorted(out["month"].unique())
    conn = get_conn()
    conn.executemany("DELETE FROM daily WHERE platform = ? AND month = ?",
                     [(platform, m) for m in months])
    out[["platform", "month", "date", "rider_id", "rider_name", "status", "orders"]].to_sql(
        "daily", conn, if_exists="append", index=False)
    conn.commit()
    conn.close()
    return months


def list_months(platform):
    conn = get_conn()
    rows = conn.execute("SELECT DISTINCT month FROM daily WHERE platform = ? ORDER BY month DESC",
                        (platform,)).fetchall()
    conn.close()
    return [r[0] for r in rows]


def load_month(platform, month):
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT rider_id, rider_name, date, status, orders FROM daily WHERE platform = ? AND month = ?",
        conn, params=[platform, month])
    conn.close()
    df["date"] = pd.to_datetime(df["date"]).dt.date
    return df


# ------------------------------------------------------------
# Analysis
# ------------------------------------------------------------
def analyze(df, target):
    d = df.copy()
    d["orders"] = pd.to_numeric(d["orders"], errors="coerce")
    d["is_on"] = d["status"].eq("on")
    d["is_off"] = d["status"].eq("off")
    d["hit"] = d["is_on"] & (d["orders"] >= target)

    daily = d.groupby("date").agg(
        on_duty=("is_on", "sum"), off_duty=("is_off", "sum"),
        orders=("orders", "sum"), hit=("hit", "sum"),
    ).reset_index().sort_values("date")
    daily["total"] = daily["on_duty"] + daily["off_duty"]
    daily["below"] = daily["on_duty"] - daily["hit"]
    daily["avg_per_rider"] = (daily["orders"] / daily["on_duty"].where(daily["on_duty"] > 0)).round(2)
    daily["day"] = daily["date"].map(lambda x: x.strftime("%d-%b"))

    riders = d.groupby("rider_id").agg(
        name=("rider_name", lambda s: next((x for x in s if x), "")),
        days_worked=("is_on", "sum"), days_off=("is_off", "sum"),
        total_orders=("orders", "sum"), days_hit=("hit", "sum"),
    ).reset_index()
    worked = riders["days_worked"].where(riders["days_worked"] > 0)
    riders["avg"] = (riders["total_orders"] / worked).round(2)
    riders["hit_pct"] = (riders["days_hit"] / worked * 100).round(0)
    riders["status"] = riders.apply(
        lambda r: "none" if r["days_worked"] == 0 else ("achieved" if r["avg"] >= target else "not"), axis=1)
    riders["consistent"] = (riders["days_worked"] > 0) & (riders["days_hit"] == riders["days_worked"])
    return daily, riders


def summary(daily, riders):
    last = daily.iloc[-1] if len(daily) else None
    return {
        "total": int(len(riders)),
        "last_day": last["day"] if last is not None else "—",
        "on_duty": int(last["on_duty"]) if last is not None else 0,
        "off_duty": int(last["off_duty"]) if last is not None else 0,
        "achieved": int((riders["status"] == "achieved").sum()),
        "not": int((riders["status"] == "not").sum()),
        "consistent": int(riders["consistent"].sum()),
        "none": int((riders["status"] == "none").sum()),
    }


# ------------------------------------------------------------
# Styling
# ------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
html, body, .stApp, [class*="st-"], button, input, textarea {font-family:'Cairo','Segoe UI',sans-serif;}
.stApp {background:#f3f5fb;}
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {display:none !important;}
header[data-testid="stHeader"] {background:transparent; pointer-events:none;}
header[data-testid="stHeader"] button, header[data-testid="stHeader"] a, [data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapseButton"] {pointer-events:auto;}
.block-container {max-width:1320px; padding:1.2rem 2.2rem 4rem;}

/* sidebar */
[data-testid="stSidebar"] {background:linear-gradient(185deg,#070d1c 0%,#0f1a35 60%,#14213f 100%); border-right:1px solid rgba(255,255,255,.06);}
[data-testid="stSidebar"] .stButton>button {width:100%; justify-content:flex-start; background:transparent; border:1px solid transparent;
  color:#cbd5e1; font-weight:700; font-size:15px; padding:.65rem 1rem; border-radius:14px; transition:.15s;}
[data-testid="stSidebar"] .stButton>button:hover {background:rgba(255,255,255,.07); color:#fff; border-color:rgba(255,255,255,.08);}
[data-testid="stSidebar"] button[kind="primary"], [data-testid="stSidebar"] [data-testid="stBaseButton-primary"] {
  background:linear-gradient(120deg,#d4a437,#f3d27a) !important; color:#0b1220 !important; border:none !important;
  box-shadow:0 8px 22px rgba(212,164,55,.35);}
.brand {padding:10px 6px 18px; direction:rtl; text-align:right;}
.brand .logo {width:46px; height:46px; border-radius:14px; background:linear-gradient(135deg,#d4a437,#f3d27a);
  display:flex; align-items:center; justify-content:center; font-size:24px; margin-bottom:10px;}
.brand .t1 {color:#fff; font-weight:800; font-size:20px; letter-spacing:.5px;}
.brand .t2 {color:#94a3b8; font-size:12px; margin-top:2px;}
.nav-label {color:#64748b; font-size:11px; font-weight:700; letter-spacing:1.5px; margin:10px 8px 6px; direction:rtl; text-align:right;}

/* hero */
.hero {border-radius:26px; padding:28px 34px; color:#fff; direction:rtl; margin-bottom:18px;
  background:linear-gradient(115deg,#0a1226 0%,#14213f 45%,var(--c) 160%); box-shadow:0 18px 50px rgba(10,18,38,.25);
  display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; position:relative; overflow:hidden;}
.hero:after {content:""; position:absolute; width:340px; height:340px; border-radius:50%; left:-90px; top:-140px; background:rgba(255,255,255,.06);}
.hero h1 {margin:0; font-size:32px; font-weight:800; line-height:1.3; padding:0;}
.hero h1 a, [data-testid="stHeaderActionElements"] {display:none !important;}
.hero .sub {opacity:.8; margin-top:4px; font-size:14px;}
.hero .pills {display:flex; gap:8px; flex-wrap:wrap; position:relative; z-index:1;}
.pill {background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.18); border-radius:999px; padding:7px 15px; font-size:13px; font-weight:700;}

/* KPI */
.kpi-grid {display:grid; grid-template-columns:repeat(auto-fit,minmax(185px,1fr)); gap:14px; margin-bottom:18px; direction:rtl;}
.kpi {background:#fff; border-radius:20px; padding:18px 20px; position:relative; overflow:hidden; border:1px solid #eceff7;
  box-shadow:0 8px 28px rgba(15,23,42,.06); text-align:right;}
.kpi:before {content:""; position:absolute; right:0; top:0; bottom:0; width:5px; background:var(--c);}
.kpi .k-icon {position:absolute; left:16px; top:14px; font-size:26px;}
.kpi .k-label {color:#64748b; font-size:13px; font-weight:700;}
.kpi .k-value {font-size:38px; font-weight:800; color:#0f172a; line-height:1.15; margin:4px 0 2px;}
.kpi .k-sub {font-size:12px; color:#94a3b8; font-weight:600;}

/* panels */
.panel {background:#fff; border-radius:22px; padding:20px 22px; border:1px solid #eceff7; margin-bottom:16px;
  box-shadow:0 8px 28px rgba(15,23,42,.06); direction:rtl; text-align:right;}
.panel h3 {margin:0 0 14px; font-size:16px; font-weight:800; color:#0f172a;}
.sec {font-size:20px; font-weight:800; color:#0f172a; margin:26px 0 12px; direction:rtl; text-align:right;}
.empty {color:#94a3b8; text-align:center; padding:18px 0; font-size:14px;}

/* workforce chart */
.legend-row {display:flex; gap:16px; font-size:12px; color:#64748b; margin:-6px 0 8px; font-weight:700;}
.legend-row i {display:inline-block; width:10px; height:10px; border-radius:3px; margin-left:6px;}
.cols {display:flex; align-items:flex-end; gap:6px; overflow-x:auto; padding-top:6px; direction:ltr;}
.col {flex:1; min-width:24px; display:flex; flex-direction:column; align-items:center;}
.col .num {font-size:11px; font-weight:800; color:#0f172a; margin-bottom:3px;}
.col .bar {width:100%; max-width:36px; height:170px; display:flex; flex-direction:column; justify-content:flex-end;
  border-radius:9px; overflow:hidden; background:#f1f5f9;}
.seg.on {background:linear-gradient(180deg,#34d399,#16a34a);} .seg.off {background:#cbd5e1;}
.col .day {font-size:11px; color:#64748b; margin-top:5px; font-weight:700;}

/* donut */
.donut-wrap {display:flex; flex-direction:column; align-items:center; gap:14px;}
.donut {width:170px; height:170px; border-radius:50%; display:flex; align-items:center; justify-content:center;}
.donut-in {width:122px; height:122px; border-radius:50%; background:#fff; display:flex; flex-direction:column; align-items:center; justify-content:center;
  box-shadow:inset 0 0 0 1px #eef0f6;}
.donut-in .d-val {font-size:30px; font-weight:800; color:#0f172a; line-height:1.1;} .donut-in .d-sub {font-size:11px; color:#64748b; font-weight:700;}
.legend {display:flex; gap:14px; flex-wrap:wrap; justify-content:center; font-size:12px; font-weight:700; color:#475569;}
.legend i {display:inline-block; width:10px; height:10px; border-radius:50%; margin-left:5px;}

/* leaderboard */
.lb-row {display:flex; align-items:center; gap:12px; padding:10px 0; border-bottom:1px solid #f1f5f9;}
.lb-row:last-child {border-bottom:none;}
.lb-rank {width:28px; height:28px; border-radius:50%; background:#0f172a; color:#fff; font-size:12px; font-weight:800;
  display:flex; align-items:center; justify-content:center; flex:0 0 28px;}
.lb-main {flex:1; min-width:0;} .lb-name {font-weight:800; color:#0f172a; font-size:14px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.lb-id {font-size:11px; color:#94a3b8; font-weight:600;}
.lb-bar {height:6px; border-radius:6px; background:#eef2f7; margin-top:5px; overflow:hidden;} .lb-bar span {display:block; height:100%; border-radius:6px;}
.lb-val {font-weight:800; font-size:18px; color:#0f172a;}

/* overview platform cards */
.pgrid {display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,300px),1fr)); gap:16px; direction:rtl;}
.pcard {background:#fff; border-radius:24px; padding:22px; border:1px solid #eceff7; box-shadow:0 10px 32px rgba(15,23,42,.07);
  border-top:6px solid var(--c); text-align:right;}
.pc-head {display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;}
.pc-head b {font-size:19px; font-weight:800; color:#0f172a;} .pc-head span {font-size:12px; color:#94a3b8; font-weight:700;}
.pc-stats {display:grid; grid-template-columns:repeat(3,1fr); gap:8px; margin-bottom:16px;}
.pc-stats div {background:#f8fafc; border-radius:14px; padding:10px; text-align:center;}
.pc-stats b {display:block; font-size:24px; font-weight:800; color:#0f172a;} .pc-stats small {color:#64748b; font-weight:700; font-size:11px;}
.pc-foot {display:flex; justify-content:space-around; margin-top:14px; font-weight:800; font-size:13px; color:#334155;}
.mini-donut {width:110px; height:110px; margin:0 auto; border-radius:50%; display:flex; align-items:center; justify-content:center;}
.mini-donut .donut-in {width:78px; height:78px;} .mini-donut .d-val {font-size:20px;}
.up-head {font-size:18px; font-weight:800; padding:12px 16px; border-radius:16px; background:#fff; border:1px solid #eceff7;
  border-right:6px solid var(--c); margin-bottom:8px; direction:rtl; text-align:right;}
[data-testid="stDataFrame"] {border-radius:16px; overflow:hidden; box-shadow:0 8px 28px rgba(15,23,42,.06);}

/* mobile top nav (hidden on desktop) */
.st-key-mobilenav {display:none;}
.svg-wrap {overflow-x:auto; -webkit-overflow-scrolling:touch;}
.cols {-webkit-overflow-scrolling:touch;}

@media (max-width: 768px) {
  .block-container {padding:3.2rem .75rem 3.5rem !important;}
  .st-key-mobilenav {display:block; padding:0 0 8px;}
  .st-key-mobilenav [data-baseweb="select"] > div {min-height:48px; border-radius:14px; font-weight:800; font-size:16px;
      background:#0f1a35; color:#fff; border:none;}
  .st-key-mobilenav [data-baseweb="select"] * {color:#fff !important;}
  .st-key-mobilenav svg {fill:#f3d27a !important;}

  .hero {padding:18px 18px; border-radius:20px; margin-bottom:12px;}
  .hero:after {display:none;}
  .hero h1 {font-size:22px;}
  .hero .sub {font-size:12px;}
  .pill {font-size:11px; padding:5px 11px;}

  .kpi-grid {grid-template-columns:repeat(2, minmax(0,1fr)); gap:10px; margin-bottom:12px;}
  .kpi-grid .kpi:last-child:nth-child(odd) {grid-column:1 / -1;}
  .kpi {padding:13px 14px; border-radius:16px;}
  .kpi .k-icon {font-size:18px; left:10px; top:10px;}
  .kpi .k-label {font-size:11px;}
  .kpi .k-value {font-size:27px;}
  .kpi .k-sub {font-size:10px;}

  .panel {padding:14px; border-radius:18px; margin-bottom:12px;}
  .panel h3 {font-size:14px; margin-bottom:10px;}
  .sec {font-size:17px; margin:18px 0 10px;}

  .cols {gap:4px; padding-bottom:4px;}
  .col {min-width:26px; flex:1 0 26px;}
  .col .bar {height:130px;}
  .svg-wrap svg {min-width:520px;}

  .donut {width:140px; height:140px;} .donut-in {width:100px; height:100px;} .donut-in .d-val {font-size:24px;}
  .lb-val {font-size:16px;} .lb-name {font-size:13px;}

  .pgrid {gap:12px;}
  .pcard {padding:16px; border-radius:20px;}
  .pc-head b {font-size:16px;} .pc-stats b {font-size:20px;}

  .stButton>button, .stDownloadButton>button {min-height:46px;}
  [data-testid="stSidebar"] .stButton>button {min-height:50px;}
}
@media (max-width: 380px) {
  .kpi .k-value {font-size:23px;} .hero h1 {font-size:19px;}
}
</style>
"""


def hero(title, sub, pills, color="#d4a437"):
    pills_html = "".join(f'<span class="pill">{esc(p)}</span>' for p in pills)
    st.markdown(
        f'<div class="hero" style="--c:{color}"><div><h1>{title}</h1><div class="sub">{esc(sub)}</div></div>'
        f'<div class="pills">{pills_html}</div></div>', unsafe_allow_html=True)


def kpi_grid(items):
    cards = "".join(
        f'<div class="kpi" style="--c:{c}"><div class="k-icon">{ic}</div><div class="k-label">{esc(l)}</div>'
        f'<div class="k-value">{v:,}</div><div class="k-sub">{esc(s)}</div></div>'
        for l, v, s, c, ic in items)
    st.markdown(f'<div class="kpi-grid">{cards}</div>', unsafe_allow_html=True)


def section(text):
    st.markdown(f'<div class="sec">{esc(text)}</div>', unsafe_allow_html=True)


def donut_css(a, n, x):
    tot = (a + n + x) or 1
    pa = a / tot * 100
    pn = pa + n / tot * 100
    return (f"background:conic-gradient(#16a34a 0 {pa:.2f}%,#ef4444 {pa:.2f}% {pn:.2f}%,#cbd5e1 {pn:.2f}% 100%)")


def donut_panel(s, target):
    a, n, x = s["achieved"], s["not"], s["none"]
    worked = a + n
    center = f"{a / worked * 100:.0f}%" if worked else "—"
    st.markdown(
        f'<div class="panel"><h3>🎯 تحقيق التارجت ({target} طلب/يوم)</h3><div class="donut-wrap">'
        f'<div class="donut" style="{donut_css(a, n, x)}"><div class="donut-in"><div class="d-val">{center}</div>'
        f'<div class="d-sub">حققوا التارجت</div></div></div>'
        f'<div class="legend"><span><i style="background:#16a34a"></i>حقق {a}</span>'
        f'<span><i style="background:#ef4444"></i>لم يحقق {n}</span>'
        f'<span><i style="background:#cbd5e1"></i>بدون شغل {x}</span></div></div></div>',
        unsafe_allow_html=True)


def workforce_chart(daily):
    mx = max(int(daily["total"].max()), 1)
    cols = ""
    for r in daily.itertuples():
        tip = f"{r.day}: دوام {int(r.on_duty)} · إجازة {int(r.off_duty)}"
        cols += (f'<div class="col" title="{esc(tip)}"><div class="num">{int(r.on_duty)}</div>'
                 f'<div class="bar"><div class="seg off" style="height:{r.off_duty / mx * 100:.1f}%"></div>'
                 f'<div class="seg on" style="height:{r.on_duty / mx * 100:.1f}%"></div></div>'
                 f'<div class="day">{r.day[:2]}</div></div>')
    st.markdown(
        '<div class="panel"><h3>👥 الدوام والإجازة يوم بيوم</h3>'
        '<div class="legend-row"><span><i style="background:#16a34a"></i>دوام</span>'
        '<span><i style="background:#cbd5e1"></i>إجازة</span></div>'
        f'<div class="cols">{cols}</div></div>', unsafe_allow_html=True)


def avg_chart(daily, target, color):
    pts = daily[daily["avg_per_rider"].notna()].reset_index(drop=True)
    if pts.empty:
        st.markdown('<div class="panel"><h3>📈 المتوسط اليومي</h3><div class="empty">لا توجد بيانات</div></div>',
                    unsafe_allow_html=True)
        return
    W, H, PL, PR, PT, PB = 640, 240, 40, 16, 16, 30
    ymax = max(float(pts["avg_per_rider"].max()), float(target)) * 1.2
    n = len(daily)
    idx = {d: i for i, d in enumerate(daily["day"])}

    def x(i):
        return PL + (W - PL - PR) * (i / max(n - 1, 1))

    def y(v):
        return PT + (H - PT - PB) * (1 - v / ymax)

    grid = ""
    for k in range(5):
        v = ymax * k / 4
        grid += (f'<line x1="{PL}" x2="{W - PR}" y1="{y(v):.1f}" y2="{y(v):.1f}" stroke="#eef2f7"/>'
                 f'<text x="{PL - 6}" y="{y(v) + 4:.1f}" font-size="10" fill="#94a3b8" text-anchor="end">{v:.0f}</text>')
    xy = [(x(idx[r.day]), y(r.avg_per_rider)) for r in pts.itertuples()]
    poly = " ".join(f"{a:.1f},{b:.1f}" for a, b in xy)
    area = f"{xy[0][0]:.1f},{y(0):.1f} {poly} {xy[-1][0]:.1f},{y(0):.1f}"
    dots = "".join(f'<circle cx="{a:.1f}" cy="{b:.1f}" r="3.5" fill="#fff" stroke="{color}" stroke-width="2.2"/>' for a, b in xy)
    step = max(1, math.ceil(n / 12))
    xl = "".join(f'<text x="{x(i):.1f}" y="{H - 8}" font-size="10" fill="#94a3b8" text-anchor="middle">{d[:2]}</text>'
                 for d, i in idx.items() if i % step == 0)
    ty = y(target)
    svg = (f'<svg viewBox="0 0 {W} {H}" width="100%" xmlns="http://www.w3.org/2000/svg">'
           f'<defs><linearGradient id="g{abs(hash(color)) % 9999}" x1="0" x2="0" y1="0" y2="1">'
           f'<stop offset="0" stop-color="{color}" stop-opacity=".28"/><stop offset="1" stop-color="{color}" stop-opacity="0"/></linearGradient></defs>'
           f'{grid}<polygon points="{area}" fill="url(#g{abs(hash(color)) % 9999})"/>'
           f'<polyline points="{poly}" fill="none" stroke="{color}" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>'
           f'{dots}<line x1="{PL}" x2="{W - PR}" y1="{ty:.1f}" y2="{ty:.1f}" stroke="#16a34a" stroke-width="2" stroke-dasharray="6 5"/>'
           f'<text x="{W - PR}" y="{ty - 6:.1f}" font-size="11" font-weight="700" fill="#16a34a" text-anchor="end">التارجت {target}</text>'
           f'{xl}</svg>')
    st.markdown(f'<div class="panel"><h3>📈 متوسط الطلبات للمندوب (يوم بيوم)</h3><div class="svg-wrap">{svg}</div></div>', unsafe_allow_html=True)


def leaderboard(title, icon, df, target, color):
    rows = ""
    for k, r in enumerate(df.itertuples(), 1):
        nm = r.name or r.rider_id
        w = min(r.avg / target * 100, 100)
        rows += (f'<div class="lb-row"><div class="lb-rank">{k}</div><div class="lb-main">'
                 f'<div class="lb-name">{esc(nm)}</div><div class="lb-id">{esc(r.rider_id)} · {int(r.days_worked)} يوم</div>'
                 f'<div class="lb-bar"><span style="width:{w:.0f}%;background:{color}"></span></div></div>'
                 f'<div class="lb-val">{r.avg:.1f}</div></div>')
    if not rows:
        rows = '<div class="empty">لا يوجد</div>'
    st.markdown(f'<div class="panel"><h3>{icon} {esc(title)}</h3>{rows}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------
# Pages
# ------------------------------------------------------------
def overview_page():
    hero("نظرة عامة", "ملخص المنصات التلاتة في مكان واحد", [date.today().strftime("%d %b %Y")])

    data = {}
    for key, cfg in PLATFORMS.items():
        months = list_months(key)
        if months:
            daily, riders = analyze(load_month(key, months[0]), cfg["target"])
            data[key] = (months[0], summary(daily, riders))

    if data:
        tot = lambda f: sum(v[1][f] for v in data.values())
        kpi_grid([
            ("إجمالي المناديب", tot("total"), "كل المنصات", "#0ea5e9", "👥"),
            ("دوام (آخر يوم)", tot("on_duty"), "مجموع المنصات", "#16a34a", "✅"),
            ("إجازة (آخر يوم)", tot("off_duty"), "مجموع المنصات", "#94a3b8", "🏖️"),
            ("حقق التارجت", tot("achieved"), "بمتوسط الشهر", "#16a34a", "🎯"),
            ("لم يحقق", tot("not"), "تحت التارجت", "#ef4444", "⚠️"),
        ])

    cards = ""
    for key, cfg in PLATFORMS.items():
        head = f'<div class="pc-head"><b>{cfg["icon"]} {cfg["name"]}</b>'
        if key not in data:
            cards += (f'<div class="pcard" style="--c:{cfg["color"]}">{head}<span>—</span></div>'
                      f'<div class="empty">لا توجد بيانات بعد — ارفع الشيت من صفحة رفع البيانات</div></div>')
            continue
        month, s = data[key]
        worked = s["achieved"] + s["not"]
        center = f'{s["achieved"] / worked * 100:.0f}%' if worked else "—"
        cards += (
            f'<div class="pcard" style="--c:{cfg["color"]}">{head}<span>{month} · {s["last_day"]}</span></div>'
            f'<div class="pc-stats"><div><b>{s["total"]}</b><small>مندوب</small></div>'
            f'<div><b>{s["on_duty"]}</b><small>دوام</small></div><div><b>{s["off_duty"]}</b><small>إجازة</small></div></div>'
            f'<div class="mini-donut" style="{donut_css(s["achieved"], s["not"], s["none"])}"><div class="donut-in">'
            f'<div class="d-val">{center}</div><div class="d-sub">التارجت {cfg["target"]}</div></div></div>'
            f'<div class="pc-foot"><span>✅ {s["achieved"]}</span><span>❌ {s["not"]}</span><span>⭐ {s["consistent"]}</span></div></div>')
    st.markdown(f'<div class="pgrid">{cards}</div>', unsafe_allow_html=True)
    st.caption("✅ حقق التارجت   ❌ لم يحقق   ⭐ منتظم (حقق كل يوم). الأرقام بالتارجت الافتراضي، ولو غيّرته في صفحة المنصة الأرقام هناك هي الأدق.")


def platform_page(key):
    cfg = PLATFORMS[key]
    color = cfg["color"]
    months = list_months(key)
    if not months:
        hero(f'{cfg["icon"]} {cfg["name"]}', "لا توجد بيانات بعد", [], color)
        st.info("لسه مفيش بيانات للمنصة دي. روح لصفحة «رفع البيانات».")
        return

    month = st.session_state.get(f"month_{key}", months[0])
    if month not in months:
        month = months[0]
    target = int(st.session_state.get(f"target_{key}", cfg["target"]))

    df = load_month(key, month)
    daily, riders = analyze(df, target)
    s = summary(daily, riders)

    hero(f'{cfg["icon"]} {cfg["name"]}', f"التارجت اليومي {target} طلب للمندوب (متوسط الأيام اللي اشتغلها)",
         [f"الشهر {month}", f'آخر يوم: {s["last_day"]}', f'{s["total"]} مندوب'], color)

    f1, f2, _ = st.columns([1, 1, 3])
    f1.selectbox("الشهر", months, key=f"month_{key}")
    f2.number_input("التارجت اليومي", min_value=1, value=cfg["target"], step=1, key=f"target_{key}")

    day_total = (s["on_duty"] + s["off_duty"]) or 1
    kpi_grid([
        ("إجمالي المناديب", s["total"], "في شيت الشهر", "#0ea5e9", "👥"),
        (f'دوام · {s["last_day"]}', s["on_duty"], f'{s["on_duty"] / day_total * 100:.0f}% من المناديب', "#16a34a", "✅"),
        (f'إجازة · {s["last_day"]}', s["off_duty"], f'{s["off_duty"] / day_total * 100:.0f}% من المناديب', "#94a3b8", "🏖️"),
        ("حقق التارجت", s["achieved"], f"متوسط ≥ {target}", "#16a34a", "🎯"),
        ("لم يحقق", s["not"], "تحت التارجت", "#ef4444", "⚠️"),
        ("منتظم", s["consistent"], "حقق التارجت كل يوم", "#f59e0b", "⭐"),
    ])

    c1, c2 = st.columns([2, 1])
    with c1:
        workforce_chart(daily)
    with c2:
        donut_panel(s, target)

    c3, c4, c5 = st.columns([2, 1, 1])
    with c3:
        avg_chart(daily, target, color)
    with c4:
        top = riders[riders["status"] == "achieved"].sort_values("avg", ascending=False).head(5)
        leaderboard("الأفضل أداءً", "🏆", top, target, "#16a34a")
    with c5:
        low = riders[riders["status"] == "not"].sort_values("avg").head(5)
        leaderboard("تحتاج متابعة", "🚨", low, target, "#ef4444")

    section("📅 التفاصيل يوم بيوم")
    table = daily[["day", "total", "on_duty", "off_duty", "hit", "below", "orders", "avg_per_rider"]].rename(columns={
        "day": "اليوم", "total": "الإجمالي", "on_duty": "دوام", "off_duty": "إجازة",
        "hit": f"حقق ≥{target}", "below": "أقل من التارجت", "orders": "إجمالي الطلبات", "avg_per_rider": "متوسط/مندوب"})
    st.dataframe(table, hide_index=True)

    section("🏍️ المناديب")
    f1, f2 = st.columns([3, 2])
    choice = f1.radio("عرض", ["الكل", "حقق التارجت", "لم يحقق", "منتظم (حقق كل يوم)", "بدون شغل"],
                      horizontal=True, key=f"filter_{key}", label_visibility="collapsed")
    q = f2.text_input("بحث", key=f"search_{key}", placeholder="🔍 ابحث بالـ ID أو الاسم",
                      label_visibility="collapsed").strip().lower()

    view = riders.copy()
    if choice == "حقق التارجت":
        view = view[view["status"] == "achieved"]
    elif choice == "لم يحقق":
        view = view[view["status"] == "not"]
    elif choice == "منتظم (حقق كل يوم)":
        view = view[view["consistent"]]
    elif choice == "بدون شغل":
        view = view[view["status"] == "none"]
    if q:
        view = view[view["rider_id"].str.lower().str.contains(q, regex=False)
                    | view["name"].str.lower().str.contains(q, regex=False)]
    view = view.sort_values("avg", ascending=False, na_position="last")

    show = pd.DataFrame({
        "Rider ID": view["rider_id"],
        "الاسم": view["name"].where(view["name"] != "", view["rider_id"]),
        "الحالة": view["status"].map({"achieved": "✅ حقق", "not": "❌ لم يحقق", "none": "— بدون شغل"}),
        "المتوسط اليومي": view["avg"],
        "التارجت": target,
        "أيام الشغل": view["days_worked"],
        "أيام الإجازة": view["days_off"],
        "إجمالي الطلبات": view["total_orders"].astype(int),
        "أيام حقق فيها": view["days_hit"],
        "نسبة الأيام": view["hit_pct"],
        "منتظم": view["consistent"].map({True: "⭐", False: ""}),
    })
    mx = view["avg"].max()
    mx = max(float(target) * 1.5, float(mx) if pd.notna(mx) else 0.0)
    st.dataframe(show, hide_index=True, column_config={
        "المتوسط اليومي": st.column_config.ProgressColumn("المتوسط اليومي", min_value=0.0, max_value=mx, format="%.1f"),
        "نسبة الأيام": st.column_config.ProgressColumn("نسبة أيام التارجت", min_value=0, max_value=100, format="%.0f%%"),
    })
    st.download_button("📥 تحميل الجدول CSV", show.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{key}_{month}_riders.csv", mime="text/csv", key=f"dl_{key}")


def upload_card(key, up, year, month):
    try:
        ok, errors = read_best(up, year, month)
    except Exception as e:
        st.error(f"مقدرتش أفتح الملف: {e}")
        return

    if not ok:
        st.error("مقدرتش أفهم الملف.")
        for sh, msg in errors.items():
            with st.expander(f"الشيت: {sh}", expanded=len(errors) == 1):
                st.code(msg)
        st.caption("لو الرسالة مش واضحة ابعتهالي (صورة أو نسخ) وهظبط القارئ على ملفك.")
        return

    sheet = next(iter(ok))
    if len(ok) > 1:
        sheet = st.selectbox("الشيت", list(ok.keys()), key=f"sheet_{key}")
    df = ok[sheet]

    d0, d1 = df["date"].min(), df["date"].max()
    st.success(f"✔ {df['rider_id'].nunique():,} مندوب · {df['date'].nunique()} يوم\n\n{d0:%d-%b} ← {d1:%d-%b-%Y}")
    st.caption("الشهور اللي في الملف بس بتتحدّث، والشهور القديمة بتفضل.")
    if st.button("💾 حفظ / تحديث", type="primary", key=f"save_{key}"):
        months = save_platform(key, df)
        st.session_state["flash"] = f"✅ تم حفظ {PLATFORMS[key]['name']}: " + "، ".join(months)
        st.rerun()


def upload_page():
    hero("📥 رفع البيانات", "ارفع الشيت المجمّع كل يوم، والشهر اللي فيه بيتحدّث تلقائيًا", [])
    c1, c2, _ = st.columns([1, 2, 2])
    year = c1.number_input("السنة", min_value=2020, max_value=2100, value=date.today().year, step=1)
    month = c2.selectbox("الشهر (بيُستخدم بس لو الشيت بيكتب أرقام الأيام 1، 2، 3 من غير شهر)",
                         list(range(1, 13)), index=date.today().month - 1)

    for col, (key, cfg) in zip(st.columns(3), PLATFORMS.items()):
        with col:
            st.markdown(f'<div class="up-head" style="--c:{cfg["color"]}">{cfg["icon"]} {cfg["name"]}</div>',
                        unsafe_allow_html=True)
            saved = list_months(key)
            st.caption("المحفوظ: " + ("، ".join(saved) if saved else "لا شيء"))
            up = st.file_uploader("الملف", type=["xlsx", "xls", "csv"], key=f"file_{key}",
                                  label_visibility="collapsed")
            if up:
                upload_card(key, up, int(year), int(month))


# ------------------------------------------------------------
# Shell
# ------------------------------------------------------------
def login():
    if st.session_state.get("ok"):
        return True
    st.markdown('<div class="brand" style="max-width:420px;margin:12vh auto 0;text-align:center">'
                '<div class="logo" style="margin:0 auto 12px">📊</div>'
                '<div class="t1" style="color:#0f172a;font-size:26px">COURIER OPS</div>'
                '<div class="t2">منظومة متابعة المناديب</div></div>', unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        with st.form("login"):
            pwd = st.text_input("الباسورد", type="password")
            go = st.form_submit_button("دخول", type="primary")
        if go:
            if hmac.compare_digest(pwd.encode(), ADMIN_PASSWORD.encode()):
                st.session_state["ok"] = True
                st.rerun()
            else:
                st.error("الباسورد غلط.")
    return False


def sidebar():
    with st.sidebar:
        st.markdown('<div class="brand"><div class="logo">📊</div><div class="t1">COURIER OPS</div>'
                    '<div class="t2">منظومة متابعة المناديب</div></div><div class="nav-label">القائمة</div>',
                    unsafe_allow_html=True)
        for key, icon, label in NAV:
            active = st.session_state["page"] == key
            if st.button(f"{icon}   {label}", key=f"nav_{key}", type="primary" if active else "secondary"):
                st.session_state["page"] = key
                st.rerun()
        st.markdown('<div style="height:40px"></div>', unsafe_allow_html=True)
        if st.button("🚪   تسجيل خروج", key="logout"):
            st.session_state["ok"] = False
            st.rerun()


def mobile_nav():
    """قايمة تنقل ظاهرة فوق في الموبايل بس (CSS بيخبّيها في الشاشات الكبيرة)."""
    keys = [k for k, _, _ in NAV]
    labels = {k: f"{i}  {l}" for k, i, l in NAV}
    with st.container(key="mobilenav"):
        st.session_state["mnav"] = st.session_state["page"]
        st.selectbox("القائمة", keys, format_func=labels.get, key="mnav", label_visibility="collapsed",
                     on_change=lambda: st.session_state.update(page=st.session_state["mnav"]))


def main():
    st.set_page_config(page_title="Courier Ops", page_icon="📊", layout="wide",
                       initial_sidebar_state="auto")
    st.markdown(CSS, unsafe_allow_html=True)
    init_db()
    if not login():
        st.stop()

    st.session_state.setdefault("page", "overview")
    sidebar()
    mobile_nav()

    flash = st.session_state.pop("flash", None)
    if flash:
        st.success(flash)

    page = st.session_state["page"]
    if page == "overview":
        overview_page()
    elif page == "upload":
        upload_page()
    else:
        platform_page(page)


if __name__ == "__main__":
    main()
