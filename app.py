import hmac
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================
# DAILY COURIER TRACKER
# 3 منصات: هنجر ستيشن / جاهز / ذا شيفز
# - ترفع الشيت كل يوم (فيه كل أيام الشهر) وبيتحدّث الشهر ده بس
# - عدد المناديب الدوام والإجازة كل يوم
# - مين حقق التارجت (بمتوسط الأيام اللي اشتغلها) ومين منتظم
# ============================================================

DB_PATH = Path(__file__).parent / "daily_tracker.db"
ADMIN_PASSWORD = "admin 444"

PLATFORMS = {
    "hunger": {"name": "هنجر ستيشن", "icon": "🟠", "target": 17},
    "jahez": {"name": "جاهز", "icon": "🔴", "target": 7},
    "chefz": {"name": "ذا شيفز", "icon": "🟣", "target": 7},
}

ID_HEADERS = {"riderid", "driverid", "courierid", "id"}
NAME_HEADERS = {"name", "ridername", "drivername", "couriername"}
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


# ------------------------------------------------------------
# Parsing helpers
# ------------------------------------------------------------
def norm(text):
    return re.sub(r"[\s_\-\.]+", "", str(text).strip().lower())


def parse_date_header(v, default_year):
    """يفهم عناوين زي: 1-Oct / 01/10 / تاريخ حقيقي."""
    try:
        if v is None or pd.isna(v):
            return None
    except Exception:
        pass
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
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
    """يرجّع ('on', عدد_الطلبات) أو ('off', None) أو (None, None) لو فاضية."""
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


def parse_sheet(raw, default_year):
    """يقرا الشيت بأي من الشكلين (جاهز / هنجر): يلاقي صف العناوين لوحده."""
    raw = raw.reset_index(drop=True)

    header_idx = None
    for i in range(min(len(raw), 40)):
        if any(norm(x) in ID_HEADERS for x in raw.iloc[i].tolist()):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("مش لاقي صف العناوين (لازم يكون فيه Rider ID أو DRIVER ID).")

    header = raw.iloc[header_idx].tolist()
    id_col = next(j for j, v in enumerate(header) if norm(v) in ID_HEADERS)
    name_col = next((j for j, v in enumerate(header) if norm(v) in NAME_HEADERS), None)

    date_cols = {}
    for j, v in enumerate(header):
        d = parse_date_header(v, default_year)
        if d:
            date_cols[j] = d
    if not date_cols:
        raise ValueError("مش لاقي أعمدة التواريخ في صف العناوين.")

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
        raise ValueError("لقيت العناوين لكن مفيش بيانات تحتها.")

    df = pd.DataFrame(rows, columns=["rider_id", "rider_name", "date", "status", "orders"])
    df = df.drop_duplicates(subset=["rider_id", "date"], keep="last").reset_index(drop=True)
    df["month"] = df["date"].map(lambda d: d.strftime("%Y-%m"))
    return df


# ------------------------------------------------------------
# Database
# ------------------------------------------------------------
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS daily (
            platform TEXT NOT NULL,
            month TEXT NOT NULL,
            date TEXT NOT NULL,
            rider_id TEXT NOT NULL,
            rider_name TEXT,
            status TEXT NOT NULL,
            orders INTEGER
        )
    """)
    conn.commit()
    conn.close()


def save_platform(platform, df):
    """يمسح الشهور اللي في الملف بس لنفس المنصة ويحط الجديد (الشهور القديمة بتفضل)."""
    out = df.copy()
    out["date"] = out["date"].map(lambda d: d.isoformat())
    out["platform"] = platform
    months = sorted(out["month"].unique())

    conn = get_conn()
    conn.executemany(
        "DELETE FROM daily WHERE platform = ? AND month = ?",
        [(platform, m) for m in months],
    )
    out[["platform", "month", "date", "rider_id", "rider_name", "status", "orders"]].to_sql(
        "daily", conn, if_exists="append", index=False
    )
    conn.commit()
    conn.close()
    return months


def list_months(platform):
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT month FROM daily WHERE platform = ? ORDER BY month DESC", (platform,)
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


def load_month(platform, month):
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT rider_id, rider_name, date, status, orders FROM daily WHERE platform = ? AND month = ?",
        conn, params=[platform, month],
    )
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
        on_duty=("is_on", "sum"),
        off_duty=("is_off", "sum"),
        orders=("orders", "sum"),
        hit=("hit", "sum"),
    ).reset_index().sort_values("date")
    daily["total"] = daily["on_duty"] + daily["off_duty"]
    daily["below"] = daily["on_duty"] - daily["hit"]
    daily["avg_per_rider"] = (daily["orders"] / daily["on_duty"].where(daily["on_duty"] > 0)).round(2)
    daily["day"] = daily["date"].map(lambda x: x.strftime("%d-%b"))

    riders = d.groupby("rider_id").agg(
        name=("rider_name", lambda s: next((x for x in s if x), "")),
        days_worked=("is_on", "sum"),
        days_off=("is_off", "sum"),
        total_orders=("orders", "sum"),
        days_hit=("hit", "sum"),
    ).reset_index()
    riders["avg"] = (riders["total_orders"] / riders["days_worked"].where(riders["days_worked"] > 0)).round(2)
    riders["hit_pct"] = (riders["days_hit"] / riders["days_worked"].where(riders["days_worked"] > 0) * 100).round(0)

    def status(row):
        if row["days_worked"] == 0:
            return "none"
        return "achieved" if row["avg"] >= target else "not"

    riders["status"] = riders.apply(status, axis=1)
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
# Auth
# ------------------------------------------------------------
def login():
    if st.session_state.get("ok"):
        return True
    st.markdown("## 🔐 تسجيل الدخول")
    with st.form("login"):
        pwd = st.text_input("الباسورد", type="password")
        go = st.form_submit_button("دخول")
    if go:
        if hmac.compare_digest(pwd.encode(), ADMIN_PASSWORD.encode()):
            st.session_state["ok"] = True
            st.rerun()
        else:
            st.error("الباسورد غلط.")
    return False


# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
def upload_section(key, cfg, has_data):
    with st.expander("📥 رفع الشيت", expanded=not has_data):
        year = st.number_input(
            "السنة (بتتستخدم لو التواريخ في الشيت مكتوبة من غير سنة)",
            min_value=2020, max_value=2100, value=date.today().year, step=1, key=f"year_{key}",
        )
        up = st.file_uploader("ملف Excel / CSV", type=["xlsx", "xls", "csv"], key=f"file_{key}")
        if not up:
            return

        try:
            if up.name.lower().endswith(".csv"):
                raw = pd.read_csv(up, header=None)
            else:
                xl = pd.ExcelFile(up)
                sheet = xl.sheet_names[0]
                if len(xl.sheet_names) > 1:
                    sheet = st.selectbox("الشيت", xl.sheet_names, key=f"sheet_{key}")
                raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            df = parse_sheet(raw, int(year))
        except Exception as e:
            st.error(f"مقدرتش أقرا الملف: {e}")
            return

        d0, d1 = df["date"].min(), df["date"].max()
        st.success(
            f"لقيت {df['rider_id'].nunique():,} مندوب، "
            f"{df['date'].nunique()} يوم ({d0:%d-%b} → {d1:%d-%b-%Y})."
        )
        st.caption("الشهور اللي في الملف بس هتتمسح وتتحدّث، والشهور القديمة بتفضل زي ما هي.")

        if st.button("💾 حفظ / تحديث", type="primary", key=f"save_{key}"):
            months = save_platform(key, df)
            st.success("✅ تم الحفظ: " + "، ".join(months))
            st.rerun()


def kpi_row(s, target):
    c = st.columns(6)
    c[0].metric("إجمالي المناديب", f"{s['total']:,}")
    c[1].metric(f"دوام ({s['last_day']})", f"{s['on_duty']:,}")
    c[2].metric(f"إجازة ({s['last_day']})", f"{s['off_duty']:,}")
    c[3].metric(f"حقق التارجت ({target})", f"{s['achieved']:,}")
    c[4].metric("لم يحقق", f"{s['not']:,}")
    c[5].metric("منتظم (كل يوم)", f"{s['consistent']:,}")
    if s["none"]:
        st.caption(f"ملحوظة: {s['none']} مندوب مفيش لهم أي يوم شغل في الشهر ده (مش محسوبين في حقق/لم يحقق).")


def platform_tab(key):
    cfg = PLATFORMS[key]
    months = list_months(key)
    upload_section(key, cfg, has_data=bool(months))

    if not months:
        st.info("لسه مفيش بيانات للمنصة دي. ارفع الشيت من فوق.")
        return

    c1, c2 = st.columns([1, 1])
    month = c1.selectbox("الشهر", months, key=f"month_{key}")
    target = c2.number_input(
        "التارجت اليومي (متوسط طلبات/يوم)", min_value=1, value=cfg["target"], step=1, key=f"target_{key}"
    )

    df = load_month(key, month)
    daily, riders = analyze(df, target)
    kpi_row(summary(daily, riders), target)

    st.markdown("#### 📅 يوم بيوم")
    table = daily[["day", "total", "on_duty", "off_duty", "hit", "below", "orders", "avg_per_rider"]].rename(columns={
        "day": "اليوم", "total": "الإجمالي", "on_duty": "دوام", "off_duty": "إجازة",
        "hit": f"حقق ≥{target} (في اليوم)", "below": "أقل من التارجت", "orders": "إجمالي الطلبات",
        "avg_per_rider": "متوسط/مندوب",
    })
    st.dataframe(table, hide_index=True)

    ch1, ch2 = st.columns(2)
    with ch1:
        st.caption("عدد الدوام والإجازة كل يوم")
        st.bar_chart(daily.set_index("day")[["on_duty", "off_duty"]])
    with ch2:
        st.caption("متوسط الطلبات للمندوب كل يوم مقابل التارجت")
        line = daily.set_index("day")[["avg_per_rider"]].copy()
        line["التارجت"] = target
        st.line_chart(line)

    st.markdown("#### 🏍️ المناديب")
    f1, f2 = st.columns([2, 3])
    choice = f1.radio(
        "عرض", ["الكل", "حقق التارجت", "لم يحقق", "منتظم (حقق كل يوم)", "بدون شغل"],
        horizontal=True, key=f"filter_{key}", label_visibility="collapsed",
    )
    q = f2.text_input("بحث بالـ ID أو الاسم", key=f"search_{key}", label_visibility="collapsed",
                      placeholder="بحث بالـ ID أو الاسم").strip().lower()

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
        "أيام الشغل": view["days_worked"],
        "أيام الإجازة": view["days_off"],
        "إجمالي الطلبات": view["total_orders"].astype(int),
        "المتوسط اليومي": view["avg"],
        "التارجت": target,
        "الحالة": view["status"].map({"achieved": "✅ حقق", "not": "❌ لم يحقق", "none": "— بدون شغل"}),
        "أيام حقق فيها": view["days_hit"],
        "نسبة الأيام %": view["hit_pct"],
        "منتظم": view["consistent"].map({True: "⭐", False: ""}),
    })
    st.dataframe(show, hide_index=True)
    st.download_button(
        "📥 تحميل الجدول CSV", show.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{key}_{month}_riders.csv", mime="text/csv", key=f"dl_{key}",
    )


def overview_tab():
    rows = []
    for key, cfg in PLATFORMS.items():
        months = list_months(key)
        if not months:
            rows.append({"المنصة": f"{cfg['icon']} {cfg['name']}", "الشهر": "—", "التارجت": cfg["target"]})
            continue
        df = load_month(key, months[0])
        daily, riders = analyze(df, cfg["target"])
        s = summary(daily, riders)
        rows.append({
            "المنصة": f"{cfg['icon']} {cfg['name']}", "الشهر": months[0], "آخر يوم": s["last_day"],
            "التارجت": cfg["target"], "إجمالي المناديب": s["total"], "دوام": s["on_duty"],
            "إجازة": s["off_duty"], "حقق التارجت": s["achieved"], "لم يحقق": s["not"],
            "منتظم": s["consistent"],
        })
    st.markdown("#### نظرة عامة (آخر شهر لكل منصة، بالتارجت الافتراضي)")
    st.dataframe(pd.DataFrame(rows), hide_index=True)
    st.caption("لو غيّرت التارجت جوه تاب المنصة، الأرقام هناك هي الأدق.")


def main():
    st.set_page_config(page_title="Daily Courier Tracker", page_icon="📊", layout="wide")
    init_db()
    if not login():
        st.stop()

    st.title("📊 متابعة المناديب اليومية")
    tabs = st.tabs(["🏠 نظرة عامة"] + [f"{c['icon']} {c['name']}" for c in PLATFORMS.values()])
    with tabs[0]:
        overview_tab()
    for tab, key in zip(tabs[1:], PLATFORMS.keys()):
        with tab:
            platform_tab(key)


if __name__ == "__main__":
    main()
