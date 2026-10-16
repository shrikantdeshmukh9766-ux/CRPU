import io
import time
import requests
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Student Project Dashboard", page_icon="📊", layout="wide")

ROSTER_FILE = "data.xlsx"  # XLSForm: used for department / student / guide names

# Final column names -> question names in the Kobo form
Q_COLS = {
    "icmr_approved_26": "ICMR approved (2026)",
    "budget_required": "Budget required",
    "budget_amount": "Budget amount (INR)",
    "iec_approval": "IEC approval",
    "work_started": "Work started",
    "icmr_approved_old": "ICMR approved (2024/25)",
    "project_started": "Project started",
    "report_ready": "Report ready",
    "project_stage": "Project stage",
}


def secret(name, default=""):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


# ---------------- Data loading ----------------
@st.cache_data
def load_roster(file):
    xl = pd.ExcelFile(file)
    ch = xl.parse("choices", dtype=str)
    depts = ch[ch.list_name == "departments"][["name", "label"]]
    depts.columns = ["dept", "Department"]
    stu = ch[ch.list_name == "students"][["name", "label", "dept", "yr", "guide"]].copy()
    stu.columns = ["student_id", "Student", "dept", "Year", "Guide"]
    stu["Year"] = stu["Year"].str.replace(r"\.0$", "", regex=True)
    roster = stu.merge(depts, on="dept", how="left").drop(columns="dept")
    labels = {}
    for lst in ("yesno", "iec", "stage"):
        sub = ch[ch.list_name == lst]
        labels[lst] = dict(zip(sub["name"], sub["label"]))
    return roster, labels


@st.cache_data(ttl=60, show_spinner="Fetching data from KoboToolbox…")
def fetch_kobo(server, uid, token):
    url = f"{server.rstrip('/')}/api/v2/assets/{uid}/data/"
    headers = {"Authorization": f"Token {token}"}
    params = {"format": "json", "limit": 1000}
    rows = []
    while url:
        r = requests.get(url, headers=headers, params=params, timeout=60)
        r.raise_for_status()
        j = r.json()
        rows += j.get("results", [])
        url, params = j.get("next"), None
    df = pd.DataFrame(rows)
    # group questions come as "grp_2026/question" -> keep only the question name
    df.columns = [c.split("/")[-1] for c in df.columns]
    return df


def build_report(roster, labels, subs):
    """Roster (all students) left-joined with their latest Kobo submission."""
    if subs is None or subs.empty:
        out = roster.copy()
        out["Status"] = "Pending"
        return out
    subs = subs.copy()
    if "_submission_time" in subs:
        subs = subs.sort_values("_submission_time")
    subs = subs.drop_duplicates("student", keep="last")  # latest per student
    keep = ["student", "_submission_time"] + [c for c in Q_COLS if c in subs.columns]
    subs = subs[keep].rename(columns={"student": "student_id", "_submission_time": "Submitted on"})
    out = roster.merge(subs, on="student_id", how="left")
    out["Status"] = out["Submitted on"].notna().map({True: "Submitted", False: "Pending"})
    out["Submitted on"] = pd.to_datetime(out["Submitted on"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    yn = labels["yesno"]
    for c in ["icmr_approved_26", "budget_required", "work_started",
              "icmr_approved_old", "project_started", "report_ready"]:
        if c in out:
            out[c] = out[c].map(yn).fillna(out[c])
    if "iec_approval" in out:
        out["iec_approval"] = out["iec_approval"].map(labels["iec"]).fillna(out["iec_approval"])
    if "project_stage" in out:
        out["project_stage"] = out["project_stage"].map(labels["stage"]).fillna(out["project_stage"])
    return out.rename(columns=Q_COLS)


def to_excel(df, per_dept=False):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Report")
        if per_dept:
            for name, g in df.groupby("Department"):
                g.dropna(axis=1, how="all").to_excel(w, index=False, sheet_name=str(name)[:31])
        df.groupby(["Department", "Year", "Status"]).size().reset_index(name="Students") \
          .to_excel(w, index=False, sheet_name="Summary")
        ws = w.sheets["Report"]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value or "")) for c in col) + 3
    return buf.getvalue()


# ---------------- Sidebar: source ----------------
st.title("📊 Student Project Dashboard")
roster, labels = load_roster(ROSTER_FILE)

with st.sidebar:
    st.header("Data source")
    server = st.text_input("Kobo server", secret("KOBO_SERVER", "https://kf.kobotoolbox.org"))
    uid = st.text_input("Form (asset) UID", secret("KOBO_ASSET_UID", "aqFb9MJU62tPZC2KaLgjfQ"))
    token = secret("KOBO_TOKEN") or st.text_input("API token", type="password")
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

subs = None
if token:
    try:
        subs = fetch_kobo(server, uid, token)
        st.sidebar.success(f"🟢 Live: {len(subs)} submissions · updated {time.strftime('%H:%M:%S')}")
    except requests.HTTPError as e:
        st.sidebar.error(f"Kobo error {e.response.status_code}. Check token / server / UID.")
    except Exception as e:
        st.sidebar.error(f"Could not reach Kobo: {e}")
else:
    st.sidebar.info("Add your API token to see submitted answers. Showing roster only.")

df = build_report(roster, labels, subs)
base_cols = ["student_id", "Student", "Department", "Year", "Guide", "Status"]
other = [c for c in df.columns if c not in base_cols]
df = df[base_cols + other].rename(columns={"student_id": "Student ID"})
df = df.sort_values(["Department", "Year", "Student"]).reset_index(drop=True)

# ---------------- Filters ----------------
with st.sidebar:
    st.header("Filters")
    sel_dept = st.multiselect("Department", sorted(df["Department"].dropna().unique()),
                              placeholder="All departments (select one or many)")
    d = df[df["Department"].isin(sel_dept)] if sel_dept else df
    sel_year = st.multiselect("Year", sorted(d["Year"].unique(), reverse=True), placeholder="All years")
    d = d[d["Year"].isin(sel_year)] if sel_year else d
    sel_guide = st.multiselect("Guide", sorted(d["Guide"].dropna().unique()), placeholder="All guides")
    d = d[d["Guide"].isin(sel_guide)] if sel_guide else d
    sel_status = st.multiselect("Status", ["Submitted", "Pending"], placeholder="All")
    d = d[d["Status"].isin(sel_status)] if sel_status else d
    q = st.text_input("Search student name")
fdf = d[d["Student"].str.contains(q, case=False, na=False)] if q else d

# ---------------- KPIs ----------------
n_sub = (fdf["Status"] == "Submitted").sum()
c1, c2, c3, c4 = st.columns(4)
c1.metric("Students", len(fdf))
c2.metric("Submitted", n_sub)
c3.metric("Pending", len(fdf) - n_sub)
c4.metric("Completion", f"{(n_sub / len(fdf) * 100 if len(fdf) else 0):.0f}%")

if fdf.empty:
    st.warning("No records match the selected filters.")
    st.stop()

# Drop columns that are entirely empty for this filter (e.g. 2026-only questions)
view = fdf.dropna(axis=1, how="all")

tab_data, tab_charts = st.tabs(["📄 Report (view & download)", "📈 Charts"])

with tab_data:
    st.caption(f"Showing {len(view)} of {len(df)} records")
    st.dataframe(view, width="stretch", hide_index=True)
    per_dept = st.checkbox("Excel: also add one sheet per department", value=len(sel_dept) > 1)
    b1, b2, _ = st.columns([1, 1, 4])
    b1.download_button("⬇️ Download Excel", to_excel(view, per_dept), "student_report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    b2.download_button("⬇️ Download CSV", view.to_csv(index=False).encode("utf-8-sig"),
                       "student_report.csv", "text/csv")

with tab_charts:
    a, b = st.columns(2)
    with a:
        st.subheader("Submitted vs Pending by department")
        st.bar_chart(pd.crosstab(fdf["Department"], fdf["Status"]))
    with b:
        st.subheader("Students by year")
        st.bar_chart(fdf["Year"].value_counts().sort_index())
    for col in ["Project stage", "Work started", "Report ready", "ICMR approved (2026)", "IEC approval"]:
        if col in fdf and fdf[col].notna().any():
            st.subheader(col)
            st.bar_chart(fdf[col].value_counts())
