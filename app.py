import io
import time
import requests
import pandas as pd
import streamlit as st

# ---- Kobo settings ----
KOBO_TOKEN = "9c1986eed989fb1aa3c7872de07aa6d03aeb125f"
KOBO_ASSET_UID = "aqFb9MJU62tPZC2KaLgjfQ"
KOBO_SERVER = "https://eu.kobotoolbox.org"  # EU server

HEADERS = {"Authorization": f"Token {KOBO_TOKEN}"}
BASE = f"{KOBO_SERVER.rstrip('/')}/api/v2/assets/{KOBO_ASSET_UID}"

FRIENDLY = {
    "department": "Department", "year": "Year", "student": "Student", "guide_name": "Guide",
    "icmr_approved_26": "ICMR approved (2026)", "budget_required": "Budget required",
    "budget_amount": "Budget amount (INR)", "iec_approval": "IEC approval",
    "work_started": "Work started", "icmr_approved_old": "ICMR approved (2024/25)",
    "project_started": "Project started", "report_ready": "Report ready",
    "project_stage": "Project stage", "_id": "Submission ID",
    "_submission_time": "Submitted on", "_submitted_by": "Submitted by", "_status": "Kobo status",
}
FIRST = ["Department", "Year", "Student", "Guide"]

st.set_page_config(page_title="Kobo Live Dashboard", page_icon="📊", layout="wide")


# ---------------- Kobo API ----------------
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_choices():
    """Choice lists from the form (department / student / yes-no labels, guide)."""
    r = requests.get(BASE + "/", headers=HEADERS, params={"format": "json"}, timeout=60)
    r.raise_for_status()
    ch = pd.DataFrame(r.json()["content"]["choices"])
    if "label" in ch:
        ch["label"] = ch["label"].apply(lambda x: x[0] if isinstance(x, list) and x else x)
    return ch


@st.cache_data(ttl=60, show_spinner="Fetching data from KoboToolbox…")
def fetch_submissions():
    url, params, rows = BASE + "/data/", {"format": "json", "limit": 1000}, []
    while url:
        r = requests.get(url, headers=HEADERS, params=params, timeout=60)
        r.raise_for_status()
        j = r.json()
        rows += j.get("results", [])
        url, params = j.get("next"), None
    return pd.DataFrame(rows)


def prepare(raw, ch):
    df = raw.copy()
    df = df[[c for c in df.columns if not c.startswith(("meta", "formhub"))]]
    df.columns = [c.split("/")[-1] for c in df.columns]      # drop group prefix
    df = df.loc[:, ~df.columns.duplicated()]
    # keep only simple values (drop attachments, geolocation, etc.)
    for c in list(df.columns):
        if df[c].map(lambda v: isinstance(v, (list, dict))).any():
            df = df.drop(columns=c)
    keep_meta = {"_id", "_submission_time", "_submitted_by", "_status"}
    df = df[[c for c in df.columns if not c.startswith(("_", "meta", "formhub", "__")) or c in keep_meta]]

    # codes -> labels using the form's choice lists
    if ch is not None and not ch.empty:
        def lab(list_name):
            sub = ch[ch.list_name == list_name]
            return dict(zip(sub["name"], sub["label"]))
        maps = {"department": lab("departments"), "student": lab("students"),
                "iec_approval": lab("iec"), "project_stage": lab("stage")}
        for c in ["icmr_approved_26", "budget_required", "work_started",
                  "icmr_approved_old", "project_started", "report_ready"]:
            maps[c] = lab("yesno")
        for c, m in maps.items():
            if c in df:
                df[c] = df[c].map(m).fillna(df[c])
    df = df.rename(columns=FRIENDLY)
    if "Submitted on" in df:
        df["Submitted on"] = pd.to_datetime(df["Submitted on"], errors="coerce").dt.tz_localize(None)
    # numeric-looking text -> numbers
    for c in df.columns:
        if df[c].dtype == object or str(df[c].dtype).startswith("str"):
            num = pd.to_numeric(df[c], errors="coerce")
            if df[c].notna().any() and num.notna().sum() == df[c].notna().sum() and c not in FIRST:
                df[c] = num
    order = [c for c in FIRST if c in df] + [c for c in df.columns if c not in FIRST]
    return df[order].sort_values("Submitted on", ascending=False) if "Submitted on" in df else df[order]


def to_excel(df, per_dept=False):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out = df.copy()
        for c in out.select_dtypes("datetimetz").columns:
            out[c] = out[c].dt.tz_localize(None)
        out.to_excel(w, index=False, sheet_name="Report")
        if per_dept and "Department" in out:
            for name, g in out.groupby("Department"):
                g.dropna(axis=1, how="all").to_excel(w, index=False, sheet_name=str(name)[:31])
        for ws in w.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    50, max(len(str(c.value or "")) for c in col) + 3)
    return buf.getvalue()


# ---------------- Generic filters for every column ----------------
def column_filter(df, c):
    s = df[c]
    if pd.api.types.is_datetime64_any_dtype(s):
        if s.notna().any():
            lo, hi = s.min().date(), s.max().date()
            if lo < hi:
                rng = st.date_input(c, (lo, hi), min_value=lo, max_value=hi, key=f"f_{c}")
                if isinstance(rng, tuple) and len(rng) == 2:
                    d = s.dt.date
                    return df[(d >= rng[0]) & (d <= rng[1]) | s.isna()]
        return df
    if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
        if s.notna().any() and s.min() < s.max():
            lo, hi = float(s.min()), float(s.max())
            a, b = st.slider(c, lo, hi, (lo, hi), key=f"f_{c}")
            if (a, b) != (lo, hi):
                return df[s.between(a, b)]
        return df
    vals = s.dropna().astype(str)
    if vals.nunique() <= 60:
        sel = st.multiselect(c, sorted(vals.unique()), key=f"f_{c}", placeholder="All")
        return df[s.astype(str).isin(sel)] if sel else df
    txt = st.text_input(f"{c} (contains)", key=f"f_{c}")
    return df[s.astype(str).str.contains(txt, case=False, na=False, regex=False)] if txt else df


# ---------------- App ----------------
st.title("📊 Kobo Live Dashboard")

with st.sidebar:
    st.header("Data")
    if st.button("🔄 Refresh now"):
        st.cache_data.clear()
    auto = st.checkbox("Auto-refresh every 60 seconds", value=True)

st.session_state["last_full_run"] = time.time()


@st.fragment(run_every=60)
def _auto_refresh():
    if time.time() - st.session_state.get("last_full_run", 0) >= 55:
        st.cache_data.clear()
        st.rerun()


if auto:
    _auto_refresh()

try:
    raw = fetch_submissions()
except requests.HTTPError as e:
    st.error(f"Kobo returned error {e.response.status_code}. Check the token, form ID and server in app.py.")
    st.stop()
except Exception as e:
    st.error(f"Could not reach KoboToolbox: {e}")
    st.stop()

if raw.empty:
    st.info("No submissions found in this Kobo form yet.")
    st.stop()

try:
    choices = fetch_choices()
except Exception:
    choices = None

df = prepare(raw, choices)
st.sidebar.success(f"🟢 Live: {len(df)} submissions · updated {time.strftime('%H:%M:%S')}")

# ---- filters: every column ----
with st.sidebar:
    st.header("Filters")
    fdf = df
    top = [c for c in FIRST if c in df.columns]
    for c in top:
        fdf = column_filter(fdf, c)
    rest = [c for c in df.columns if c not in top]
    with st.expander(f"More column filters ({len(rest)})"):
        for c in rest:
            fdf = column_filter(fdf, c)

# ---- KPIs ----
c1, c2, c3 = st.columns(3)
c1.metric("Records", len(fdf), f"of {len(df)}", delta_color="off")
if "Department" in fdf:
    c2.metric("Departments", fdf["Department"].nunique())
if "Student" in fdf:
    c3.metric("Students", fdf["Student"].nunique())

if fdf.empty:
    st.warning("No records match the selected filters.")
    st.stop()

view = fdf.dropna(axis=1, how="all")  # hide columns with no data for this selection

tab_data, tab_chart = st.tabs(["📄 Data (view & download)", "📈 Chart"])

with tab_data:
    st.dataframe(view, width="stretch", hide_index=True)
    multi = "Department" in view and view["Department"].nunique() > 1
    per_dept = st.checkbox("Excel: also add one sheet per department", value=multi)
    b1, b2, _ = st.columns([1, 1, 4])
    b1.download_button("⬇️ Download Excel", to_excel(view, per_dept), "kobo_report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    b2.download_button("⬇️ Download CSV", view.to_csv(index=False).encode("utf-8-sig"),
                       "kobo_report.csv", "text/csv")

with tab_chart:
    cat_cols = [c for c in view.columns if view[c].nunique() <= 60 and c != "Submission ID"]
    if cat_cols:
        pick = st.selectbox("Count records by", cat_cols,
                            index=cat_cols.index("Department") if "Department" in cat_cols else 0)
        st.bar_chart(view[pick].value_counts())
