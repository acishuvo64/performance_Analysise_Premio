# ==============================================================================
#  FIELD FORCE PERFORMANCE ANALYSIS DASHBOARD  (ACI Premio Plastics Ltd.)
#  Build: Streamlit + Plotly
#  Purpose: Compare 2 employees' May & June performance, detect fake/"idle"
#           TP visits, GPS movement analysis, sales achievement and generate
#           an auto-written low-achiever comparison report.
# ==============================================================================

import os
import io
import base64
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------------------------
# 0. PAGE CONFIG + UI CHROME (hides Streamlit's default menu/footer/"view source")
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Field Force Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"Get help": None, "Report a bug": None, "About": None},
)

HIDE_STREAMLIT_CHROME = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&family=Inter:wght@400;500;600&display=swap');

html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
h1, h2, h3, h4, .section-title { font-family: 'Poppins', sans-serif; }

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stDeployButton {display:none;}
[data-testid="stToolbar"] {display:none;}
[data-testid="stDecoration"] {display:none;}
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 1300px;}

/* App background */
[data-testid="stAppViewContainer"] > .main {
    background: linear-gradient(180deg, #F4FAF7 0%, #FFFFFF 260px);
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0E3D2A 0%, #0E8F5B 100%);
}
[data-testid="stSidebar"] * { color: #F2FBF6 !important; }
[data-testid="stSidebar"] input, [data-testid="stSidebar"] .stNumberInput input {
    color: #0E3D2A !important;
}

/* Header banner */
.hero-banner {
    background: linear-gradient(120deg, #0E3D2A 0%, #0E8F5B 55%, #C2107A 130%);
    border-radius: 18px; padding: 22px 28px; margin-bottom: 18px;
    box-shadow: 0 8px 24px rgba(14,63,42,0.25);
}
.hero-title { color:#FFFFFF; font-size: 28px; font-weight:700; margin:0; }
.hero-sub { color:#E7F5ED; font-size: 14px; margin-top:4px; }

/* KPI cards */
.kpi-card {
    background: #ffffff; border-radius: 16px; padding: 16px 18px;
    box-shadow: 0 4px 16px rgba(14,63,42,0.10); border-left: 6px solid #0E8F5B;
    transition: transform .15s ease;
}
.kpi-card:hover { transform: translateY(-2px); }
.kpi-title {font-size: 12.5px; letter-spacing:.3px; text-transform:uppercase; color:#6B7A72; margin-bottom:4px; font-weight:600;}
.kpi-value {font-size: 27px; font-weight:700; color:#0E3D2A; font-family:'Poppins',sans-serif;}
.flag-bad {color:#B00020; font-weight:700;}
.flag-good {color:#0E8F5B; font-weight:700;}

.section-title {
    font-size: 19px; font-weight:700; color:#0E3D2A; margin-top: 6px;
    padding: 6px 0 8px 12px; margin-bottom: 12px;
    border-left: 5px solid #C2107A;
    background: linear-gradient(90deg, rgba(14,143,91,0.08), rgba(194,16,122,0.02));
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] {
    background-color: #EAF6EF; border-radius: 10px 10px 0 0; padding: 8px 14px;
    font-weight:600; color:#0E3D2A;
}
.stTabs [aria-selected="true"] {
    background-color: #0E8F5B !important; color: #FFFFFF !important;
}

/* Metric widgets */
[data-testid="stMetric"] {
    background: #ffffff; border-radius: 12px; padding: 10px 14px;
    box-shadow: 0 2px 10px rgba(14,63,42,0.08);
}
</style>
"""
st.markdown(HIDE_STREAMLIT_CHROME, unsafe_allow_html=True)

PLOTLY_TEMPLATE = "plotly_white"
PLOTLY_COLORWAY = ["#0E8F5B", "#C2107A", "#F2A93B", "#3E6BE0", "#6B7A72"]
px.defaults.template = PLOTLY_TEMPLATE
px.defaults.color_discrete_sequence = PLOTLY_COLORWAY

ACI_GREEN = "#0E8F5B"
PREMIO_PINK = "#C2107A"

# ------------------------------------------------------------------------------
# 1. CONFIG / ASSUMPTIONS (editable from the sidebar — see README for guidance)
# ------------------------------------------------------------------------------
DEFAULTS = {
    "daily_target_visits": 20,      # required effective TP visits per working day
    "minutes_per_visit": 20,        # policy: 1 effective visit credited per N idle minutes
    # dwelling this long (or more) at ONE location = NOT effective
    "same_place_flag_hr": 1.0,
    # flag "Travelling" rows with 0 m movement as wasted/fishing time
    "zero_km_flag": True,
}

MONTH_ORDER = ["May", "June", "July"]

# ------------------------------------------------------------------------------
# 2. DATA LOADING
# ------------------------------------------------------------------------------
REQUIRED_SHEETS = [
    "Productivity Perfromance",
    "Highest_Place Found Data",
    "Image_URL",
    "03 month Sales performance",
    "Raw",
]


def _clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [" ".join(str(c).replace("\n", " ").split())
                  for c in df.columns]
    return df


@st.cache_data(show_spinner="Loading workbook...")
def load_workbook(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    missing = [s for s in REQUIRED_SHEETS if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"Missing sheet(s) in workbook: {missing}")

    prod = _clean_cols(xls.parse("Productivity Perfromance"))
    highest = _clean_cols(xls.parse("Highest_Place Found Data"))
    images = _clean_cols(xls.parse("Image_URL"))
    sales = _clean_cols(xls.parse("03 month Sales performance"))
    raw = _clean_cols(xls.parse("Raw"))

    for df in (prod, highest, raw):
        df["Date"] = pd.to_datetime(df["Date"])
        df["Month"] = df["Date"].dt.strftime("%B")

    # de-duplicate rows that are exact duplicates (source export sometimes repeats a row)
    prod = prod.drop_duplicates()

    return prod, highest, images, sales, raw


def parse_hms(t):
    """Robustly turn an Excel time value / string into a datetime.time-like -> minutes since midnight."""
    if pd.isna(t):
        return np.nan
    if isinstance(t, str):
        try:
            t = pd.to_datetime(t).time()
        except Exception:
            return np.nan
    try:
        return t.hour * 60 + t.minute + t.second / 60
    except Exception:
        return np.nan


def parse_distance_m(s):
    if pd.isna(s):
        return 0.0
    s = str(s).strip()
    try:
        if s.endswith("km"):
            return float(s.replace("km", "").strip()) * 1000
        if s.endswith("m"):
            return float(s.replace("m", "").strip())
    except Exception:
        return 0.0
    return 0.0


# ------------------------------------------------------------------------------
# 3. DERIVED / BUSINESS-LOGIC CALCULATIONS
# ------------------------------------------------------------------------------
def build_dwell_episodes(raw_emp: pd.DataFrame, same_place_flag_hr: float, minutes_per_visit: int):
    """
    Collapse consecutive 'Idling / Stop' rows that share the same location into single
    dwell episodes (an employee sitting still at ONE spot for a stretch of time), then
    decide, episode by episode, whether that stretch should count as a genuine effective
    visit or should be flagged (parked too long in one place to be real market work).
    """
    df = raw_emp.sort_values(["Date", "Start Time"]).reset_index(drop=True)
    df["start_min"] = df["Start Time"].apply(parse_hms)
    df["end_min"] = df["End Time"].apply(parse_hms)
    df["duration_min"] = (df["end_min"] - df["start_min"]).clip(lower=0)
    df["loc_key"] = df["Latitude"].round(3).astype(
        str) + "_" + df["Longitude"].round(3).astype(str)

    idle = df[df["Status"] == "Idling / Stop"].copy()
    if idle.empty:
        return pd.DataFrame()

    idle["new_episode"] = (
        (idle["Date"] != idle["Date"].shift()) | (
            idle["loc_key"] != idle["loc_key"].shift())
    ).cumsum()

    episodes = (
        idle.groupby("new_episode")
        .agg(
            Employee_Name=("Employee Name", "first"),
            Date=("Date", "first"),
            Day=("Day", "first"),
            Month=("Month", "first"),
            Location=("Address / Diagnostic report", "first"),
            Latitude=("Latitude", "first"),
            Longitude=("Longitude", "first"),
            Start=("Start Time", "first"),
            End=("End Time", "last"),
            Duration_Min=("duration_min", "sum"),
            Reported_TP_Visit=("Total Tp Visit ( GP Approx)", "sum"),
        )
        .reset_index(drop=True)
    )

    episodes["Duration_Hr"] = episodes["Duration_Min"] / 60
    episodes["Flag_Excessive_Dwell"] = episodes["Duration_Hr"] >= same_place_flag_hr
    # policy-based effective-visit credit: 1 credited visit per `minutes_per_visit` minutes,
    # but a flagged (too-long, same-spot) episode earns ZERO effective visit credit.
    episodes["Effective_Visit_Credit"] = np.where(
        episodes["Flag_Excessive_Dwell"],
        0,
        (episodes["Duration_Min"] / minutes_per_visit).round(1),
    )
    return episodes


def build_wastage(raw_emp: pd.DataFrame, episodes: pd.DataFrame, zero_km_flag: bool):
    """Wasted / 'fishing' time = excessive same-place dwelling + travel legs claimed with 0 m movement."""
    df = raw_emp.copy()
    df["dist_m"] = df["Distance Travelled From Last Location"].apply(
        parse_distance_m)
    zero_travel = df[(df["Status"] == "Travelling") & (
        df["dist_m"] <= 0)] if zero_km_flag else df.iloc[0:0]

    wasted_dwell_hr = episodes.loc[episodes["Flag_Excessive_Dwell"], "Duration_Hr"].sum(
    ) if not episodes.empty else 0
    zero_travel_count = len(zero_travel)

    return {
        "wasted_dwell_hr": round(wasted_dwell_hr, 2),
        "zero_travel_legs": zero_travel_count,
        "zero_travel_rows": zero_travel,
    }


def monthwise_prod_summary(prod: pd.DataFrame) -> pd.DataFrame:
    work_prod_col = "Working Productivity Based of Apprx TP Visit"
    agg_dict = {
        "Total_Working_Hr": ("Total Working Hr", "sum"),
        "Total_Idle_Hr": ("Total Idle Hr", "sum"),
        "Total_Travel_Hr": ("Total Travel Hr", "sum"),
        "Total_Tp_Visit": ("Total Tp Visit ( GP Approx)", "sum"),
        "Avg_Tp_Visit_Productivity": ("TP visit Productivity", "mean"),
        "Avg_Travel_Pct": ("Travel %", "mean"),
        "Avg_Idle_Pct": ("Total Idle %", "mean"),
        "Days_Logged": ("Date", "nunique"),
    }
    if work_prod_col in prod.columns:
        agg_dict["Avg_Working_Productivity"] = (work_prod_col, "mean")
    g = prod.groupby(["Employee Name", "Month"]).agg(**agg_dict).reset_index()
    return g


def style_verdict_column(df: pd.DataFrame, col: str = "Verdict"):
    """pandas 2.1+ renamed Styler.applymap -> Styler.map, and pandas 3.0 removed applymap
    entirely. This helper works on any pandas version."""
    def color_fn(v): return "color:#B00020;font-weight:bold" if "❌" in str(
        v) else "color:#0E8F5B;font-weight:bold"
    styler = df.style
    if hasattr(styler, "map"):
        return styler.map(color_fn, subset=[col])
    return styler.applymap(color_fn, subset=[col])


def touchpoint_pct(row, target_per_day=20):
    return row["Total_Tp_Visit"] / (row["Days_Logged"] * target_per_day) if row["Days_Logged"] else 0


# ------------------------------------------------------------------------------
# 4. SIDEBAR — DATA SOURCE + POLICY CONTROLS
# ------------------------------------------------------------------------------
st.sidebar.markdown("## ⚙️ Setup")

uploaded = st.sidebar.file_uploader(
    "Upload the monthly workbook (.xlsx)", type=["xlsx"],
    help="Same structure as 'Bottom_performar_Data_02_month.xlsx' — sheets: "
         "Productivity Perfromance, Highest_Place Found Data, Image_URL, "
         "03 month Sales performance, Raw.",
)

if uploaded is None:
    st.title("📊 Field Force Performance Dashboard")
    st.info("👈 Upload the monthly data workbook (.xlsx) from the sidebar to begin.")
    st.stop()

try:
    prod, highest, images, sales, raw = load_workbook(uploaded.getvalue())
except Exception as e:
    st.error(f"Could not read the workbook: {e}")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🎯 Policy assumptions")
daily_target_visits = st.sidebar.number_input(
    "Daily effective-visit target", 1, 100, DEFAULTS["daily_target_visits"])
minutes_per_visit = st.sidebar.number_input(
    "Minutes of idle time = 1 credited visit", 1, 120, DEFAULTS["minutes_per_visit"])
same_place_flag_hr = st.sidebar.slider(
    "Flag as 'not effective' if same spot ≥ (hours)", 0.25, 4.0, DEFAULTS["same_place_flag_hr"], 0.25)
zero_km_flag = st.sidebar.checkbox(
    "Flag 0-distance 'Travelling' legs as wasted time", value=True)
st.sidebar.caption(
    "These are adjustable interpretations of your rule (e.g. 'every 20 min idle → "
    "credited visits, but 1.5 hr at one spot is never effective'). Tune them to match "
    "your exact SOP."
)

employees = sorted(prod["Employee Name"].unique().tolist())

st.sidebar.markdown("---")
st.sidebar.markdown("### 🤖 AI narrative (optional)")
use_ai = st.sidebar.checkbox("Generate AI-written comparison report")
api_key = st.sidebar.text_input(
    "Anthropic API key", type="password") if use_ai else None
st.sidebar.caption("Key is used only in this session, never stored or logged.")

# ------------------------------------------------------------------------------
# 5. PRECOMPUTE — everything the tabs need
# ------------------------------------------------------------------------------
month_summary = monthwise_prod_summary(prod)
month_summary["Tp_Visit_Pct"] = month_summary.apply(
    lambda r: touchpoint_pct(r, daily_target_visits), axis=1)

episodes_by_emp = {}
wastage_by_emp = {}
for emp in employees:
    ep = build_dwell_episodes(
        raw[raw["Employee Name"] == emp], same_place_flag_hr, minutes_per_visit)
    episodes_by_emp[emp] = ep
    wastage_by_emp[emp] = build_wastage(
        raw[raw["Employee Name"] == emp], ep, zero_km_flag)

# ------------------------------------------------------------------------------
# 6. HEADER
# ------------------------------------------------------------------------------

col_logo1, col_title, col_logo2 = st.columns([1, 5, 1])
with col_logo1:
    if os.path.exists("ACI.png"):
        st.image("ACI.png", width=90)
with col_title:
    st.markdown(
        "<div class='hero-banner'>"
        "<p class='hero-title'>📊 Field Force Performance & Low-Achiever Analysis</p>"
        "<p class='hero-sub'>ACI Premio Plastics Ltd. &nbsp;•&nbsp; May vs June comparison &nbsp;•&nbsp; GPS-verified productivity audit</p>"
        "</div>",
        unsafe_allow_html=True,
    )
with col_logo2:
    if os.path.exists("ACI_Premio_Plastics_Ltd_.jpg"):
        st.image("ACI_Premio_Plastics_Ltd_.jpg", width=110)

tabs = st.tabs([
    "🏠 Overview",
    "📅 Month-wise Cards",
    "🕒 Day/Area Time Analysis",
    "✅ Effective Visit Audit",
    "🎣 Wastage & 'Fishing' Movement",
    "🗺️ GPS Map & History",
    "🧠 AI Low-Achiever Report",
    "🗂️ Raw Data",
])

# ==============================================================================
# TAB 1 — OVERVIEW
# ==============================================================================
with tabs[0]:
    st.markdown("<div class='section-title'>Employee Summary Cards</div>",
                unsafe_allow_html=True)
    cols = st.columns(len(employees))
    for i, emp in enumerate(employees):
        with cols[i]:
            img_row = images[images["Employee Name"] == emp]
            if not img_row.empty and pd.notna(img_row.iloc[0].get("Image")):
                try:
                    st.image(img_row.iloc[0]["Image"], width=140)
                except Exception:
                    st.markdown(
                        "🧑‍💼 *(photo link could not be rendered — see README)*")
            deg = img_row.iloc[0]["Designation"] if not img_row.empty else ""
            grp = img_row.iloc[0]["Group"] if not img_row.empty else ""
            st.markdown(f"### {emp}")
            st.caption(f"{deg} · Group: {grp}")

            srow = sales[sales["Employee Name"] == emp]
            if not srow.empty:
                avg_sales = srow.iloc[0]["03 Month Avg Sales"]
                achv = srow.iloc[0]["03 Month achv %"]
                st.markdown(
                    f"<div class='kpi-card'><div class='kpi-title'>3-Month Avg Sales</div>"
                    f"<div class='kpi-value'>{avg_sales:,.0f}</div></div>", unsafe_allow_html=True,
                )
                st.markdown(
                    f"<div class='kpi-card' style='margin-top:8px;'><div class='kpi-title'>3-Month Achievement %</div>"
                    f"<div class='kpi-value'>{achv*100:.1f}%</div></div>", unsafe_allow_html=True,
                )

            ep = episodes_by_emp[emp]
            flagged_hr = ep.loc[ep["Flag_Excessive_Dwell"],
                                "Duration_Hr"].sum() if not ep.empty else 0
            st.markdown(
                f"<div class='kpi-card' style='margin-top:8px;'><div class='kpi-title'>Flagged Non-Effective Dwell Time</div>"
                f"<div class='kpi-value flag-bad'>{flagged_hr:.1f} hr</div></div>", unsafe_allow_html=True,
            )

    st.markdown("<div class='section-title'>Monthly KPI Comparison (May vs June)</div>",
                unsafe_allow_html=True)
    metric_choice = st.selectbox(
        "Metric",
        ["Total_Working_Hr", "Total_Idle_Hr", "Total_Travel_Hr",
            "Total_Tp_Visit", "Tp_Visit_Pct"],
        format_func=lambda x: x.replace("_", " "),
    )
    fig = px.bar(
        month_summary, x="Employee Name", y=metric_choice, color="Month", barmode="group",
        color_discrete_sequence=[ACI_GREEN, PREMIO_PINK], text_auto=".2s",
        title=f"{metric_choice.replace('_',' ')} — May vs June",
    )
    st.plotly_chart(fig, use_container_width=True)

# ==============================================================================
# TAB 2 — MONTH-WISE CARDS (separate card per employee per month, all metrics)
# ==============================================================================
with tabs[1]:
    st.markdown("<div class='section-title'>Separate Month-wise Performance Cards</div>",
                unsafe_allow_html=True)
    for emp in employees:
        st.markdown(f"#### 👤 {emp}")
        emp_rows = month_summary[month_summary["Employee Name"] == emp].set_index(
            "Month")
        srow = sales[sales["Employee Name"] ==
                     emp].iloc[0] if not sales[sales["Employee Name"] == emp].empty else None
        month_cols = st.columns(len(emp_rows) + 1)
        for j, m in enumerate([m for m in MONTH_ORDER if m in emp_rows.index]):
            r = emp_rows.loc[m]
            with month_cols[j]:
                st.markdown(f"**{m} 2026**")
                st.metric("Working Hr (total)", f"{r['Total_Working_Hr']:.0f}")
                st.metric("Idle Hr (total)", f"{r['Total_Idle_Hr']:.0f}")
                st.metric("Travel Hr (total)", f"{r['Total_Travel_Hr']:.0f}")
                st.metric("TP Visits (approx)", f"{r['Total_Tp_Visit']:.0f}")
                st.metric("TP Visit Achv %", f"{r['Tp_Visit_Pct']*100:.1f}%")
                if srow is not None:
                    sales_col_candidates = [
                        f"{m[:3]}'26 Sales", f"{m}'26 Sales"]
                    sales_col = next(
                        (c for c in sales_col_candidates if c in sales.columns), None)
                    if sales_col:
                        st.metric(f"{m} Sales Achv %",
                                  f"{srow[sales_col]*100:.1f}%")
        with month_cols[-1]:
            st.markdown("**3-Month Avg**")
            if srow is not None:
                st.metric("Avg Sales", f"{srow['03 Month Avg Sales']:,.0f}")
                st.metric("Avg Achv %", f"{srow['03 Month achv %']*100:.1f}%")
        st.markdown("---")

# ==============================================================================
# TAB 3 — DAY-WISE / HIGH-PLACE TIME ANALYSIS
# ==============================================================================
with tabs[2]:
    emp_pick = st.selectbox("Employee", employees, key="dayarea_emp")
    ep = episodes_by_emp[emp_pick]

    st.markdown("<div class='section-title'>Day-wise Time Spent (all dwell episodes)</div>",
                unsafe_allow_html=True)
    if not ep.empty:
        day_time = ep.groupby("Date", as_index=False)["Duration_Hr"].sum()
        fig = px.bar(day_time, x="Date", y="Duration_Hr", color_discrete_sequence=[ACI_GREEN],
                     title="Total dwell (idle) hours per day")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown(
            "<div class='section-title'>Highest Single-Place Time Spend (per Highest_Place sheet)</div>", unsafe_allow_html=True)
        hdf = highest[highest["Employee Name"] == emp_pick].sort_values(
            "Avg Duration Hr", ascending=False)
        st.dataframe(hdf[["Date", "Day", "Start Time", "End Time", "Avg Duration Hr", "Address / Diagnostic report"]],
                     use_container_width=True, hide_index=True)

        st.markdown(
            "<div class='section-title'>Which Area Was Covered / Visited the Most</div>", unsafe_allow_html=True)
        area_freq = (
            ep.groupby("Location", as_index=False)
            .agg(Times_Visited=("Location", "count"), Total_Hours=("Duration_Hr", "sum"))
            .sort_values("Total_Hours", ascending=False)
            .head(15)
        )
        fig2 = px.bar(area_freq, x="Total_Hours", y="Location", orientation="h",
                      color_discrete_sequence=[PREMIO_PINK], title="Top locations by cumulative time spent")
        fig2.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.warning("No idle/dwell episodes found for this employee.")

# ==============================================================================
# TAB 4 — EFFECTIVE VISIT AUDIT (core rule the user described)
# ==============================================================================
with tabs[3]:
    st.markdown(
        "<div class='section-title'>Effective Visit Audit — 'Did she really do the "
        f"{daily_target_visits} visits?'</div>", unsafe_allow_html=True,
    )
    st.info(
        f"**Rule applied:** every {minutes_per_visit} idle minutes credits 1 effective visit — "
        f"**but** any single stop of **{same_place_flag_hr}+ hours at one exact spot is auto-flagged "
        "and earns ZERO credit**, no matter how many visits were logged there. Adjust in the sidebar."
    )
    emp_pick2 = st.selectbox("Employee", employees, key="audit_emp")
    ep = episodes_by_emp[emp_pick2]
    if ep.empty:
        st.warning("No data.")
    else:
        daily = (
            ep.groupby("Date", as_index=False)
            .agg(
                Reported_TP_Visit=("Reported_TP_Visit", "sum"),
                Credited_Effective_Visits=("Effective_Visit_Credit", "sum"),
                Flagged_Episodes=("Flag_Excessive_Dwell", "sum"),
                Max_Single_Dwell_Hr=("Duration_Hr", "max"),
            )
        )
        daily["Target"] = daily_target_visits
        daily["Gap_vs_Target"] = daily["Credited_Effective_Visits"] - \
            daily["Target"]
        daily["Verdict"] = np.where(
            daily["Gap_vs_Target"] >= 0, "✅ Met Target", "❌ Below Target")

        fig = go.Figure()
        fig.add_bar(x=daily["Date"], y=daily["Reported_TP_Visit"],
                    name="Reported (raw) TP Visits", marker_color="#B0BEC5")
        fig.add_bar(x=daily["Date"], y=daily["Credited_Effective_Visits"],
                    name="Credited Effective Visits", marker_color=ACI_GREEN)
        fig.add_hline(y=daily_target_visits, line_dash="dash",
                      line_color="red", annotation_text="Daily Target")
        fig.update_layout(
            barmode="overlay", title="Reported vs Truly-Credited Effective Visits per Day")
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            style_verdict_column(daily, "Verdict"),
            use_container_width=True, hide_index=True,
        )

        st.markdown(
            "##### 🚩 Flagged (excessive same-place) episodes — these inflated the raw numbers")
        st.dataframe(
            ep[ep["Flag_Excessive_Dwell"]][["Date", "Day", "Location",
                                            "Start", "End", "Duration_Hr", "Reported_TP_Visit"]]
            .sort_values("Duration_Hr", ascending=False),
            use_container_width=True, hide_index=True,
        )

# ==============================================================================
# TAB 5 — WASTAGE / "FISHING" MOVEMENT ANALYSIS
# ==============================================================================
with tabs[4]:
    st.markdown("<div class='section-title'>Wastage Time & Suspicious ('Fishing') Movement</div>",
                unsafe_allow_html=True)
    st.caption(
        "'Fishing' movement = logging travel legs with 0 m real displacement (phone moved, "
        "person didn't), or parking at one spot far longer than any real visit needs."
    )
    for emp in employees:
        ep = episodes_by_emp[emp]
        w = wastage_by_emp[emp]
        st.markdown(f"#### 👤 {emp}")
        c1, c2, c3 = st.columns(3)
        c1.metric("Wasted dwell time (flagged)",
                  f"{w['wasted_dwell_hr']:.1f} hr")
        c2.metric("0-distance 'Travelling' legs", f"{w['zero_travel_legs']}")
        worst_day = ep.loc[ep["Duration_Hr"].idxmax()
                           ] if not ep.empty else None
        if worst_day is not None:
            c3.metric("Longest single same-place stop",
                      f"{worst_day['Duration_Hr']:.1f} hr", help=str(worst_day["Location"]))

        if worst_day is not None:
            st.write(
                f"📍 **Day spent most time in one place:** {worst_day['Date'].date()} "
                f"({worst_day['Day']}) — **{worst_day['Duration_Hr']:.1f} hr** at *{worst_day['Location']}*."
            )
            top_area = (
                ep.groupby("Location")["Duration_Hr"].sum(
                ).sort_values(ascending=False)
            )
            if not top_area.empty:
                st.write(
                    f"🗺️ **Area repeated the most across the month:** *{top_area.index[0]}* "
                    f"— total **{top_area.iloc[0]:.1f} hr** spent there across the period."
                )
        st.markdown("---")

# ==============================================================================
# TAB 6 — GPS MAP & CLICK-THROUGH HISTORY
# ==============================================================================
with tabs[5]:
    st.markdown("<div class='section-title'>GPS Mapping — Location & Date-wise</div>",
                unsafe_allow_html=True)
    emp_pick3 = st.selectbox("Employee", employees, key="map_emp")
    emp_raw = raw[raw["Employee Name"] == emp_pick3].dropna(
        subset=["Latitude", "Longitude"])
    dates_available = sorted(emp_raw["Date"].dt.date.unique())
    date_pick = st.selectbox(
        "Pick a date to see that day's movement history", dates_available)

    day_df = emp_raw[emp_raw["Date"].dt.date ==
                     date_pick].sort_values("Start Time")
    if day_df.empty:
        st.warning("No GPS points for this date.")
    else:
        fig = px.scatter_mapbox(
            day_df, lat="Latitude", lon="Longitude", color="Status",
            hover_name="Address / Diagnostic report",
            hover_data={"Start Time": True, "End Time": True,
                        "Latitude": False, "Longitude": False},
            zoom=9, height=550, color_discrete_map={"Idling / Stop": "#B00020", "Travelling": ACI_GREEN},
        )
        fig.update_layout(mapbox_style="open-street-map",
                          margin=dict(l=0, r=0, t=0, b=0))
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### 📜 Click-through history for the selected date")
        st.dataframe(
            day_df[["Start Time", "End Time", "Status",
                    "Total Tp Visit ( GP Approx)", "Distance Travelled From Last Location", "Address / Diagnostic report"]],
            use_container_width=True, hide_index=True,
        )

# ==============================================================================
# TAB 7 — AI / AUTOMATED LOW-ACHIEVER COMPARISON REPORT
# ==============================================================================
with tabs[6]:
    st.markdown("<div class='section-title'>Two-Month Comparison & Low-Achiever Report</div>",
                unsafe_allow_html=True)

    def month_row(emp, month):
        r = month_summary[(month_summary["Employee Name"] == emp)
                          & (month_summary["Month"] == month)]
        return r.iloc[0] if not r.empty else None

    report_blocks = []
    for emp in employees:
        may_r, jun_r = month_row(emp, "May"), month_row(emp, "June")
        ep = episodes_by_emp[emp]
        flagged_hr = ep.loc[ep["Flag_Excessive_Dwell"],
                            "Duration_Hr"].sum() if not ep.empty else 0
        srow = sales[sales["Employee Name"] == emp].iloc[0]

        may_sales_col = "May'26 Sales"
        jun_sales_col = "June'26 Sales"
        may_sales_achv = srow[may_sales_col] * 100
        jun_sales_achv = srow[jun_sales_col] * 100

        st.markdown(f"#### 👤 {emp}")
        colA, colB = st.columns(2)
        with colA:
            st.markdown("**May**")
            if may_r is not None:
                st.write(
                    f"- Working hr: {may_r['Total_Working_Hr']:.0f} | Idle hr: {may_r['Total_Idle_Hr']:.0f} | Travel hr: {may_r['Total_Travel_Hr']:.0f}")
                st.write(
                    f"- TP visits: {may_r['Total_Tp_Visit']:.0f} ({may_r['Tp_Visit_Pct']*100:.1f}% of target)")
            st.write(f"- Sales achv: {may_sales_achv:.1f}%")
        with colB:
            st.markdown("**June**")
            if jun_r is not None:
                st.write(
                    f"- Working hr: {jun_r['Total_Working_Hr']:.0f} | Idle hr: {jun_r['Total_Idle_Hr']:.0f} | Travel hr: {jun_r['Total_Travel_Hr']:.0f}")
                st.write(
                    f"- TP visits: {jun_r['Total_Tp_Visit']:.0f} ({jun_r['Tp_Visit_Pct']*100:.1f}% of target)")
            st.write(f"- Sales achv: {jun_sales_achv:.1f}%")

        trend = "declined" if (
            jun_r is not None and may_r is not None and jun_r["Tp_Visit_Pct"] < may_r["Tp_Visit_Pct"]) else "improved/stable"
        narrative = (
            f"**Rule-based summary:** {emp}'s effective touch-point productivity **{trend}** from May to June. "
            f"A total of **{flagged_hr:.1f} hours** were flagged as spent parked at a single location well beyond "
            f"what any genuine visit needs, artificially inflating the raw TP-visit count. "
            f"3-month average sales achievement stands at **{srow['03 Month achv %']*100:.1f}%**, "
            "which — combined with the flagged idle-dwelling pattern — is the primary driver behind the low-achiever classification."
        )
        st.markdown(narrative)
        report_blocks.append(f"{emp}: {narrative}")
        st.markdown("---")

    if use_ai and api_key:
        if st.button("🤖 Generate richer AI narrative (Anthropic Claude)"):
            try:
                import anthropic
                client = anthropic.Anthropic(api_key=api_key)
                prompt = (
                    "You are a field-force operations analyst. Using ONLY the structured facts below, "
                    "write a concise, professional 2-month (May vs June) performance comparison report "
                    "explaining, for each employee, why they are classified as a low achiever. Be specific, "
                    "reference the numbers given, and keep it under 350 words per employee.\n\n"
                    + "\n".join(report_blocks)
                )
                resp = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1200,
                    messages=[{"role": "user", "content": prompt}],
                )
                st.markdown("### 🧠 AI-Generated Report")
                st.write(resp.content[0].text)
            except ImportError:
                st.error("Install the SDK first: `pip install anthropic`")
            except Exception as e:
                st.error(f"AI generation failed: {e}")
    elif use_ai and not api_key:
        st.warning(
            "Enter your Anthropic API key in the sidebar to generate the AI narrative.")

# ==============================================================================
# TAB 8 — RAW DATA EXPLORER
# ==============================================================================
with tabs[7]:
    st.markdown("<div class='section-title'>Raw Data (for verification)</div>",
                unsafe_allow_html=True)
    which = st.selectbox("Sheet", ["Productivity Perfromance",
                         "Highest_Place Found Data", "Raw", "03 month Sales performance", "Image_URL"])
    mapping = {
        "Productivity Perfromance": prod, "Highest_Place Found Data": highest,
        "Raw": raw, "03 month Sales performance": sales, "Image_URL": images,
    }
    st.dataframe(mapping[which], use_container_width=True, hide_index=True)

st.markdown("---")
st.caption("Built for ACI Premio Plastics Ltd. — assumptions on effective-visit rules are configurable in the sidebar and should be validated against your official SOP before using this for HR decisions.")
