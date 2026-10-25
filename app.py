import os
import io
import hashlib
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

from result import (
    load_workbook_data, normalize_text, completeness_report, quality_issues,
    safe_numeric, format_money, to_csv_bytes, to_excel_bytes, make_summary
)

st.set_page_config(
    page_title="Client & Caregiver Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Consistent high-contrast dark theme across every dashboard page and chart.
pio.templates.default = "plotly_dark"
st.markdown("""
<style>
:root { color-scheme: dark; --bg:#0B0F14; --panel:#151B23; --panel2:#1C2430; --line:#303A48; --text:#F5F7FA; --muted:#B6C0CD; --blue:#75A7FF; --teal:#54D6C5; }
html, body, [class*="css"], .stApp, [data-testid="stAppViewContainer"] { background-color:var(--bg); color:var(--text); }
.stApp { background:var(--bg); }
[data-testid="stHeader"] { background:rgba(11,15,20,.96); }
.block-container { padding-top:1.25rem; padding-bottom:2rem; max-width:1600px; }
h1,h2,h3,h4,h5,h6,p,li,label,span,div { color:inherit; }
[data-testid="stMetric"] { background:var(--panel); border:1px solid var(--line); padding:16px 18px; border-radius:14px; box-shadow:0 2px 10px rgba(0,0,0,.18); }
[data-testid="stMetricLabel"], [data-testid="stMetricValue"], [data-testid="stMetricDelta"] { color:var(--text) !important; }
[data-testid="stSidebar"] { background:#101720; border-right:1px solid var(--line); }
[data-testid="stSidebar"] * { color:#F5F7FA !important; }
[data-testid="stDataFrame"], [data-testid="stTable"] { border:1px solid var(--line); border-radius:10px; }
[data-testid="stDataFrame"] * { color:#F5F7FA; }
[data-baseweb="select"] > div, [data-baseweb="input"] > div, [data-baseweb="textarea"] > div { background:#151B23 !important; border-color:#465365 !important; color:#F5F7FA !important; }
input, textarea { color:#F5F7FA !important; }
div.stButton > button, div.stDownloadButton > button, button[kind] { border-radius:9px; font-weight:600; background:#263548; color:#FFFFFF; border:1px solid #52657D; }
div.stButton > button:hover, div.stDownloadButton > button:hover { background:#354B65; border-color:#75A7FF; color:#FFFFFF; }
[data-testid="stAlert"] { background:#18212C; color:#F5F7FA; border:1px solid var(--line); }
.stMarkdown, .stCaption, small { color:#D3DBE5; }
.small-muted { color:#B6C0CD; font-size:.88rem; }
.status-pill { display:inline-block; padding:5px 10px; border-radius:20px; background:#17382F; color:#8EF0C6; font-size:12px; font-weight:700; }
hr { border-color:var(--line); }
a { color:#8BB8FF !important; }
@media (max-width: 700px) {
  .block-container { padding-left: .8rem; padding-right: .8rem; }
  [data-testid="stMetric"] { padding:10px; }
}
</style>
""", unsafe_allow_html=True)

DEFAULT_PATHS = [
    r"D:\pew pew\Clients and Caregiver List.xlsx",
    os.path.join(os.path.dirname(__file__), "Clients and Caregiver List.xlsx"),
    "/mnt/data/Clients and Caregiver List.xlsx",
    "/mnt/data/Clients and Caregiver List(1).xlsx",
]

def find_default_file():
    for path in DEFAULT_PATHS:
        if os.path.isfile(path):
            return path
    return None

@st.cache_data(show_spinner=False, ttl=60)
def load_from_path(path, modified):
    return load_workbook_data(path=path)

@st.cache_data(show_spinner=False, ttl=60)
def load_from_bytes(file_bytes, digest):
    return load_workbook_data(file_bytes=file_bytes)

def get_data():
    uploaded = st.session_state.get("workbook_upload")
    if uploaded is not None:
        content = uploaded.getvalue()
        digest = hashlib.sha256(content).hexdigest()
        return load_from_bytes(content, digest), f"Uploaded workbook · {uploaded.name}"
    path = find_default_file()
    if path:
        modified = os.path.getmtime(path)
        return load_from_path(path, modified), path
    return None, None

with st.sidebar:
    st.markdown("## ◈ CareOps")
    st.caption("CLIENT & CAREGIVER INTELLIGENCE")
    page = st.radio(
        "NAVIGATION",
        ["Executive Overview", "Client Management", "Caregiver Management",
         "Services & Demand", "Financial Overview", "Geographic Analysis",
         "Data Quality & Records", "Reports & Export", "Settings & Connection"],
        label_visibility="visible",
    )
    st.divider()
    st.markdown("**Data source**")
    st.file_uploader("Upload Excel workbook", type=["xlsx", "xls"], key="workbook_upload",
                     help="Use your current workbook. If running locally, the app also checks the configured default path.")
    if st.button("↻ Refresh data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.caption("Refresh re-reads the workbook. Excel files do not sync automatically unless you refresh or upload a new version.")

data, source_label = get_data()
if data is None:
    st.title("Client & Caregiver Intelligence")
    st.warning("Excel workbook not found. Upload **Clients and Caregiver List.xlsx** in the sidebar, or place it beside app.py.")
    st.code(r"D:\pew pew\Clients and Caregiver List.xlsx", language="text")
    st.info("Expected sheets: client records (headers such as Customer ID, Client Name) and caregiver records (headers such as Name, Age, Mobile).")
    st.stop()

clients = data["clients"].copy()
caregivers = data["caregivers"].copy()
metadata = data["metadata"]
last_loaded = datetime.now().strftime("%d %b %Y, %I:%M:%S %p")

def metric(label, value, help_text=None):
    st.metric(label, value, help=help_text)

def safe_unique(series):
    vals = series.dropna().astype(str).str.strip()
    vals = vals[vals.ne("")]
    return int(vals.nunique())

def filter_select(label, series, key):
    options = sorted(series.dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
    return st.selectbox(label, ["All"] + options, key=key)

def filtered_clients(frame, search="", service_type="All", category="All", area="All"):
    out = frame.copy()
    if search.strip():
        needle = search.strip().lower()
        fields = [c for c in ["Client Name", "Customer ID"] if c in out.columns]
        mask = pd.Series(False, index=out.index)
        for c in fields:
            mask |= out[c].astype(str).str.lower().str.contains(needle, na=False)
        out = out[mask]
    for col, val in [("Service Type", service_type), ("Service Categories", category), ("Service Area", area)]:
        if val != "All" and col in out.columns:
            out = out[out[col].astype(str).str.strip() == val]
    return out

def table_view(df, height=420):
    if df.empty:
        st.info("No records match the current filters.")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True, height=height)

def export_buttons(df, base_name):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Download CSV", to_csv_bytes(df), f"{base_name}.csv", "text/csv", use_container_width=True)
    with c2:
        st.download_button("Download Excel", to_excel_bytes(df), f"{base_name}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

# Shared derived values; source values are not written back to Excel.
for frame in (clients, caregivers):
    for col in frame.select_dtypes(include="object").columns:
        frame[col] = frame[col].map(lambda x: x.strip() if isinstance(x, str) else x)

client_pay_col = "Client Pay" if "Client Pay" in clients.columns else None
care_pay_col = "Caregiver Payment" if "Caregiver Payment" in clients.columns else None
if client_pay_col:
    clients[client_pay_col] = safe_numeric(clients[client_pay_col])
if care_pay_col:
    clients[care_pay_col] = safe_numeric(clients[care_pay_col])
care_pay_metric_col = "Caregiver Payment Recorded" if "Caregiver Payment Recorded" in clients.columns else care_pay_col
if care_pay_metric_col:
    clients[care_pay_metric_col] = safe_numeric(clients[care_pay_metric_col])

st.title(page)
st.markdown('<span class="status-pill">WORKBOOK LOADED</span> &nbsp; <span class="small-muted">Last loaded: '+last_loaded+' · Source: '+str(source_label)+'</span>', unsafe_allow_html=True)
st.caption("All figures are calculated from the loaded workbook. Refresh after changing the Excel file.")

if page == "Executive Overview":
    st.subheader("Executive snapshot")
    complete_c, incomplete_c = completeness_report(clients)
    complete_g, incomplete_g = completeness_report(caregivers)
    c1,c2,c3,c4 = st.columns(4)
    with c1: metric("Total clients", f"{len(clients):,}")
    with c2: metric("Total caregivers", f"{len(caregivers):,}")
    with c3: metric("Client records complete", f"{complete_c:,}", "Complete means all available fields in that row are populated.")
    with c4: metric("Caregiver records complete", f"{complete_g:,}", "Complete means all available fields in that row are populated.")
    c5,c6,c7,c8 = st.columns(4)
    with c5: metric("Service types", f"{safe_unique(clients['Service Type']) if 'Service Type' in clients else 0:,}")
    with c6: metric("Service areas", f"{safe_unique(clients['Service Area']) if 'Service Area' in clients else 0:,}")
    with c7:
        metric("Recorded client payments", format_money(clients[client_pay_col].sum(min_count=1)) if client_pay_col else "N/A")
    with c8:
        metric("Recorded caregiver payments", format_money(clients[care_pay_metric_col].sum(min_count=1)) if care_pay_metric_col else "N/A")
    st.divider()
    left,right = st.columns(2)
    with left:
        st.subheader("Client distribution by service type")
        if "Service Type" in clients and clients["Service Type"].notna().any():
            vc = clients["Service Type"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts().rename_axis("Service Type").reset_index(name="Records")
            fig = px.pie(vc, names="Service Type", values="Records", hole=.58, color_discrete_sequence=["#2864DC","#0F9D91","#6C8CE5","#8DCFC7","#D99019","#91A0B7"])
            fig.update_layout(margin=dict(t=10,b=10,l=5,r=5), legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("Service Type column has no usable values.")
    with right:
        st.subheader("Clients by service area")
        if "Service Area" in clients and clients["Service Area"].notna().any():
            area = clients["Service Area"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts().head(15).rename_axis("Service Area").reset_index(name="Clients")
            fig = px.bar(area.sort_values("Clients"), x="Clients", y="Service Area", orientation="h", text="Clients", color_discrete_sequence=["#2864DC"])
            fig.update_layout(margin=dict(t=10,b=10,l=5,r=5), yaxis_title="", xaxis_title="Client records")
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("Service Area column has no usable values.")
    left,right = st.columns(2)
    with left:
        st.subheader("Record completeness")
        comp = pd.DataFrame({"Dataset":["Clients","Caregivers"],"Complete":[complete_c,complete_g],"Incomplete":[incomplete_c,incomplete_g]})
        fig = px.bar(comp, x="Dataset", y=["Complete","Incomplete"], barmode="stack",
                     color_discrete_map={"Complete":"#0F9D91","Incomplete":"#D99019"})
        fig.update_layout(margin=dict(t=10,b=10,l=5,r=5), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        st.subheader("Caregiver experience")
        if "Work Experience" in caregivers and caregivers["Work Experience"].notna().any():
            exp = caregivers["Work Experience"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts().head(12).rename_axis("Experience").reset_index(name="Caregivers")
            fig = px.bar(exp, x="Experience", y="Caregivers", text="Caregivers", color_discrete_sequence=["#0F9D91"])
            fig.update_layout(margin=dict(t=10,b=10,l=5,r=5), xaxis_title="", yaxis_title="Caregiver records")
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("No Work Experience data available.")
    st.subheader("Management insights")
    insights = []
    if "Service Type" in clients and clients["Service Type"].notna().any():
        top = clients["Service Type"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().value_counts()
        if len(top): insights.append(f"Most common recorded service type: **{top.index[0]}** ({top.iloc[0]} client records).")
    if "Service Area" in clients and clients["Service Area"].notna().any():
        top = clients["Service Area"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().value_counts()
        if len(top): insights.append(f"Area with the highest recorded client count: **{top.index[0]}** ({top.iloc[0]} records).")
    if incomplete_c: insights.append(f"**{incomplete_c} client records** have at least one blank field.")
    if incomplete_g: insights.append(f"**{incomplete_g} caregiver records** have at least one blank field.")
    if care_pay_col:
        missing_pay = int(clients[care_pay_col].isna().sum())
        insights.append(f"Caregiver payment is blank or non-numeric in **{missing_pay} client rows**.")
    if not insights: insights = ["No major insights can be calculated from the available fields."]
    for item in insights: st.markdown("- " + item)
    st.caption("Payment totals are recorded amounts only. They are not profit or audited accounting figures.")

elif page == "Client Management":
    st.subheader("Client directory")
    a,b,c,d = st.columns([2,1,1,1])
    with a: search = st.text_input("Search name or customer ID", placeholder="Type to search...")
    with b: typ = filter_select("Service type", clients["Service Type"] if "Service Type" in clients else pd.Series(dtype=str), "client_type")
    with c: cat = filter_select("Service category", clients["Service Categories"] if "Service Categories" in clients else pd.Series(dtype=str), "client_cat")
    with d: area = filter_select("Service area", clients["Service Area"] if "Service Area" in clients else pd.Series(dtype=str), "client_area")
    result = filtered_clients(clients, search, typ, cat, area)
    st.caption(f"{len(result):,} matching records · {len(clients):,} total client records")
    client_display = result.drop(columns=["Caregiver Payment Recorded"], errors="ignore")
    table_view(client_display)
    export_buttons(client_display, "filtered_clients")
    with st.expander("Client record details"):
        if not result.empty:
            chosen = st.selectbox("Choose a client", result.index, format_func=lambda i: str(result.loc[i].get("Client Name", f"Record {i+1}")))
            detail = result.loc[chosen].drop(labels=["Caregiver Payment Recorded"], errors="ignore")
            st.json({k: (None if pd.isna(v) else v.item() if hasattr(v,"item") else v) for k,v in detail.to_dict().items()})

elif page == "Caregiver Management":
    st.subheader("Caregiver directory")
    search = st.text_input("Search caregiver name or address", placeholder="Type a name or location...")
    g = caregivers.copy()
    if search.strip():
        mask = pd.Series(False,index=g.index)
        for col in ["Name","Present Address","Education Qualification","Work Experience"]:
            if col in g: mask |= g[col].astype(str).str.contains(search.strip(), case=False, na=False)
        g = g[mask]
    f1,f2,f3 = st.columns(3)
    with f1:
        if "Present Address" in g: addr = filter_select("Present address", g["Present Address"], "care_addr")
        else: addr = "All"
    with f2:
        if "Education Qualification" in g: edu = filter_select("Education", g["Education Qualification"], "care_edu")
        else: edu = "All"
    with f3:
        if "Work Experience" in g: expf = filter_select("Work experience", g["Work Experience"], "care_exp")
        else: expf = "All"
    for col,val in [("Present Address",addr),("Education Qualification",edu),("Work Experience",expf)]:
        if val != "All" and col in g: g = g[g[col].astype(str).str.strip() == val]
    st.caption(f"{len(g):,} matching records · phone numbers are masked in this general view.")
    shown = g.copy()
    if "Mobile" in shown:
        shown["Mobile"] = shown["Mobile"].map(lambda x: ("••••" + str(x)[-4:]) if pd.notna(x) and str(x).strip() else x)
    table_view(shown)
    export_buttons(g, "caregiver_records_authorized")
    with st.expander("Caregiver profile"):
        if not g.empty:
            idx = st.selectbox("Choose caregiver", g.index, format_func=lambda i: str(g.loc[i].get("Name", f"Record {i+1}")))
            row = g.loc[idx].to_dict()
            if "Mobile" in row and pd.notna(row["Mobile"]):
                row["Mobile"] = "••••" + str(row["Mobile"])[-4:]
            st.json({k: (None if pd.isna(v) else v.item() if hasattr(v,"item") else v) for k,v in row.items()})

elif page == "Services & Demand":
    st.subheader("Services & recorded client distribution")
    a,b = st.columns(2)
    with a:
        if "Service Type" in clients:
            x = clients["Service Type"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts().rename_axis("Service Type").reset_index(name="Records")
            st.plotly_chart(px.bar(x.sort_values("Records"),x="Records",y="Service Type",orientation="h",text="Records",color_discrete_sequence=["#2864DC"]),use_container_width=True)
    with b:
        if "Service Categories" in clients:
            x = clients["Service Categories"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts().rename_axis("Service Category").reset_index(name="Records")
            st.plotly_chart(px.pie(x,names="Service Category",values="Records",hole=.5),use_container_width=True)
    if "Service Area" in clients and "Service Type" in clients:
        st.subheader("Service type by area")
        cross = clients.assign(**{"Service Area":clients["Service Area"].fillna("Missing"),"Service Type":clients["Service Type"].fillna("Missing")}).groupby(["Service Area","Service Type"]).size().reset_index(name="Client records")
        st.plotly_chart(px.bar(cross,x="Service Area",y="Client records",color="Service Type",barmode="stack"),use_container_width=True)
    st.subheader("Ranked service records")
    ranked = clients.copy()
    cols = [c for c in ["Service Type","Service Categories","Service Area"] if c in ranked]
    if cols:
        st.dataframe(ranked[cols].value_counts(dropna=False).reset_index(name="Records").sort_values("Records",ascending=False),use_container_width=True,hide_index=True)

elif page == "Financial Overview":
    st.subheader("Recorded payment analysis")
    if not client_pay_col and not care_pay_col:
        st.warning("No recognized payment columns were found in the client worksheet.")
    else:
        p1,p2,p3,p4 = st.columns(4)
        with p1: metric("Recorded client payments",format_money(clients[client_pay_col].sum(min_count=1)) if client_pay_col else "N/A")
        with p2: metric("Recorded caregiver payments",format_money(clients[care_pay_metric_col].sum(min_count=1)) if care_pay_metric_col else "N/A")
        with p3: metric("Average client payment",format_money(clients[client_pay_col].mean()) if client_pay_col else "N/A")
        with p4: metric("Average caregiver payment",format_money(clients[care_pay_metric_col].mean()) if care_pay_metric_col else "N/A")
        if client_pay_col and care_pay_col:
            compare = pd.DataFrame({"Payment Type":["Client Pay","Caregiver Payment"],"Recorded amount":[clients[client_pay_col].sum(min_count=1),clients[care_pay_metric_col].sum(min_count=1)]})
            st.plotly_chart(px.bar(compare,x="Payment Type",y="Recorded amount",text_auto=".2s",color="Payment Type",color_discrete_map={"Client Pay":"#2864DC","Caregiver Payment":"#0F9D91"}),use_container_width=True)
            st.caption("Merged caregiver-payment cells are counted once in recorded totals to avoid double-counting. The difference between client payments and caregiver payments is not presented as profit.")
        if "Service Type" in clients:
            st.subheader("Payments by service type")
            agg_cols = [c for c in [client_pay_col,care_pay_metric_col] if c]
            agg = clients.groupby("Service Type",dropna=False)[agg_cols].sum(min_count=1).reset_index()
            rename_map = {}
            if client_pay_col and client_pay_col in agg:
                rename_map[client_pay_col] = "Client Pay"
            if care_pay_metric_col and care_pay_metric_col in agg:
                rename_map[care_pay_metric_col] = "Caregiver Payment"
            agg = agg.rename(columns=rename_map)
            chart_y = [c for c in ["Client Pay", "Caregiver Payment"] if c in agg.columns]
            st.plotly_chart(px.bar(agg,x="Service Type",y=chart_y,barmode="group"),use_container_width=True)
        st.subheader("Payment record detail")
        cols = [c for c in ["Sl No","Customer ID","Client Name","Service Type","Service Area",client_pay_col,care_pay_col,care_pay_metric_col] if c and c in clients]
        cols = list(dict.fromkeys(cols))
        table_view(clients[cols])
        missing = sum(int(clients[c].isna().sum()) for c in [client_pay_col,care_pay_col] if c)
        st.info(f"Blank or non-numeric payment cells across available payment fields: {missing:,}. Blank values are not treated as zero.")

elif page == "Geographic Analysis":
    st.subheader("Area-level distribution")
    left,right = st.columns(2)
    with left:
        if "Service Area" in clients:
            x = clients["Service Area"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts().rename_axis("Area").reset_index(name="Client records")
            st.plotly_chart(px.bar(x.sort_values("Client records"),x="Client records",y="Area",orientation="h",text="Client records",color_discrete_sequence=["#2864DC"]),use_container_width=True)
    with right:
        if "Present Address" in caregivers:
            x = caregivers["Present Address"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts().head(15).rename_axis("Caregiver address").reset_index(name="Caregiver records")
            st.plotly_chart(px.bar(x.sort_values("Caregiver records"),x="Caregiver records",y="Caregiver address",orientation="h",text="Caregiver records",color_discrete_sequence=["#0F9D91"]),use_container_width=True)
    st.subheader("Area summary")
    c_area = clients["Service Area"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts() if "Service Area" in clients else pd.Series(dtype=int)
    g_area = caregivers["Present Address"].fillna("Missing").astype(str).str.strip().replace("", "Missing").value_counts() if "Present Address" in caregivers else pd.Series(dtype=int)
    areas = sorted(set(c_area.index)|set(g_area.index))
    summary = pd.DataFrame({"Area":areas,"Client records":[int(c_area.get(a,0)) for a in areas],"Caregiver records at matching address label":[int(g_area.get(a,0)) for a in areas]})
    st.dataframe(summary,use_container_width=True,hide_index=True)
    st.caption("Client service areas and caregiver addresses are separate fields; this is not a client-to-caregiver assignment or coverage ratio.")

elif page == "Data Quality & Records":
    st.subheader("Data quality monitor")
    issues = quality_issues(clients, caregivers, metadata)
    q1,q2,q3 = st.columns(3)
    with q1: metric("Client records",f"{len(clients):,}")
    with q2: metric("Caregiver records",f"{len(caregivers):,}")
    with q3: metric("Quality issues",f"{len(issues):,}")
    if issues.empty:
        st.success("No issues detected by the current basic validation rules.")
    else:
        c1,c2 = st.columns([1,2])
        with c1:
            issue_type = st.selectbox("Issue type",["All"]+sorted(issues["Issue"].dropna().unique().tolist()))
        with c2:
            query = st.text_input("Search record ID, worksheet, or details")
        view = issues.copy()
        if issue_type != "All": view = view[view["Issue"] == issue_type]
        if query.strip(): view = view[view.astype(str).apply(lambda col: col.str.contains(query.strip(),case=False,na=False)).any(axis=1)]
        table_view(view, 500)
        export_buttons(view,"data_quality_issues")
    st.caption("This page flags issues for review only. It never edits the source workbook.")

elif page == "Reports & Export":
    st.subheader("Management reports")
    st.write("Exports are generated from the currently loaded workbook. Apply filters on the relevant page before exporting that page's table.")
    st.markdown("#### Executive summary")
    summary = make_summary(clients,caregivers,metadata)
    st.dataframe(summary,use_container_width=True,hide_index=True)
    st.download_button("Download executive summary (CSV)",to_csv_bytes(summary),"executive_summary.csv","text/csv",use_container_width=True)
    st.download_button("Download complete workbook copy (Excel)",to_excel_bytes({"Clients":clients.drop(columns=["Caregiver Payment Recorded"], errors="ignore"),"Caregivers":caregivers,"Summary":summary}),"management_report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.markdown("#### Individual reports")
    r1,r2 = st.columns(2)
    with r1:
        st.download_button("Client directory CSV",to_csv_bytes(clients.drop(columns=["Caregiver Payment Recorded"], errors="ignore")),"client_directory.csv","text/csv",use_container_width=True)
        st.download_button("Financial detail CSV",to_csv_bytes(clients[[c for c in ["Customer ID","Client Name","Service Type","Service Area",client_pay_col,care_pay_col] if c and c in clients]]),"financial_detail.csv","text/csv",use_container_width=True)
    with r2:
        st.download_button("Caregiver directory CSV (phone masked)",to_csv_bytes(caregivers.assign(**({"Mobile":caregivers["Mobile"].map(lambda x: "••••"+str(x)[-4:] if pd.notna(x) and str(x).strip() else x)} if "Mobile" in caregivers else {}))),"caregiver_directory_masked.csv","text/csv",use_container_width=True)
        st.download_button("Data quality report CSV",to_csv_bytes(quality_issues(clients,caregivers,metadata)),"data_quality_report.csv","text/csv",use_container_width=True)
    st.caption(f"Generated: {datetime.now().strftime('%d %b %Y %H:%M:%S')} · Data loaded: {last_loaded}")

else:
    st.subheader("Settings & data connection")
    st.markdown("### Workbook connection")
    st.write("**Current source:**",source_label)
    st.write("**Last loaded:**",last_loaded)
    st.write("**Refresh method:** manual refresh / rerun, with a 60-second Streamlit cache.")
    st.write("**Sheets discovered:**")
    st.dataframe(pd.DataFrame([{"Worksheet":k,"Rows":v} for k,v in metadata.get("sheet_rows",{}).items()]),use_container_width=True,hide_index=True)
    st.markdown("### Detected columns")
    st.write("Client worksheet:",", ".join(clients.columns.astype(str)))
    st.write("Caregiver worksheet:",", ".join(caregivers.columns.astype(str)))
    st.markdown("### Configuration notes")
    st.info("This Excel-based version does not require Google API credentials or live hosting. For local use, keep the workbook beside app.py or update DEFAULT_PATHS. Uploading a workbook overrides the local default.")
    st.warning("Because Excel is a file, changes are not pushed automatically from another application. Save the workbook, then click Refresh data or rerun the app. If you need automatic folder monitoring, it can be added for a local deployment.")
    st.markdown("### Data privacy")
    st.write("Caregiver phone numbers are masked in tables and standard exports. The workbook and downloaded reports may contain personal and financial information; store them securely.")
