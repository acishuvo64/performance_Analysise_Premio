# ===============================================================
# AI FIELD FORCE TRACKING DASHBOARD  -  ACI Premio Plastics Ltd.
# Version : 9.0 (Clean / No-login Edition)  |  October 2026
# Run     : streamlit run Tracking_Dashboard_Oct_26.py
#
# Data    : Tracking_Dashboard_Oct_26.xlsx  (same folder as this file)
#           Sheets used -> Summarize_Data, Image_URL,
#                          Same_Place_Found, Employee_Data
# Next month: just upload the new Excel with the same sheet/column
#           names. If the file name changes, any "Tracking_Dashboard_*.xlsx"
#           in this folder is picked up automatically (newest one wins).
# ===============================================================

import re
from dataclasses import dataclass
from pathlib import Path

import folium
import pandas as pd
import plotly.graph_objects as go
import pydeck as pdk
import streamlit as st
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium

# ---------------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------------
APP_TITLE = "AI Field Force Tracking Dashboard"
EXCEL_FILE = "Tracking_Dashboard_Oct_26.xlsx"
EXCEL_GLOB = "Tracking_Dashboard_*.xlsx"
BASE_DIR = Path(__file__).resolve().parent

SHEET_SUMMARY = "Summarize_Data"
SHEET_IMAGES = "Image_URL"
SHEET_SAME_PLACE = "Same_Place_Found"
SHEET_GPS = "Employee_Data"

COL_PRODUCTIVITY = "Work Productivity"
COL_ADDRESS = "Address / Diagnostic report"
COL_IDLE = "Total Idle Hr"

STATUS_IDLE = "Idling / Stop"
STATUS_TRAVEL = "Travelling"

TARGET_PCT = 50   # productivity target  (green  at/above)
LOW_PCT = 30      # low productivity     (red    below)

DEG_COLOR_MAP = {"RSM": "#4F46E5", "Sr. RSM": "#8B5CF6", "TSM": "#06B6D4",
                 "ASM": "#F59E0B", "SM": "#10B981"}
GROUP_PALETTE = ["#6366F1", "#10B981", "#F59E0B", "#EC4899", "#06B6D4"]
AV_COLORS = ["#6366F1", "#10B981", "#F59E0B", "#EC4899",
             "#06B6D4", "#8B5CF6", "#EF4444"]
C_TRAVEL, C_IDLE, C_WORK = "#06B6D4", "#F59E0B", "#6366F1"
WEEKDAY_ORDER = ["Saturday", "Sunday", "Monday", "Tuesday",
                 "Wednesday", "Thursday", "Friday"]

REQUIRED_COLUMNS = {
    SHEET_SUMMARY: ["Employee Name", "Deg", "Group", "Date", "Day", "Total Working Hr",
                    "Total Idle Hr", "Total Travel Hr", "Total Tp Visit",
                    COL_PRODUCTIVITY, "Remarks", "Auditing"],
    SHEET_IMAGES: ["Employee Name", "Designation", "Image"],
    SHEET_SAME_PLACE: ["Employee Name", "Deg", "Group", "Date", "Day", "Start Time",
                       "End Time", "Total Idle Hr", COL_ADDRESS, "Lat", "Lon"],
    SHEET_GPS: ["Employee Name", "Deg", "Group", "Date", "Day", "Start Time",
                "End Time", "Total Idle Hr", "Status", COL_ADDRESS, "Lat", "Lon"],
}

# ---------------------------------------------------------------
# CSS
# ---------------------------------------------------------------
APP_CSS = """
@keyframes shine { 0%{background-position:0% 50%} 100%{background-position:300% 50%} }
.stApp { background: radial-gradient(circle at 8% 0%, #E0E7FF 0%, #F8FAFC 38%, #DCFCE7 100%) !important; }
[data-testid="stSidebar"] { background:#0F172A !important; }
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] span { color:white !important; font-weight:600; }
[data-testid="stSidebar"] [data-baseweb="select"] *,
[data-testid="stSidebar"] input { color:#0F172A !important; }
[data-testid="stSidebar"] [data-baseweb="tag"] span { color:white !important; }

.kpi-blue   { background:linear-gradient(135deg,#4F46E5,#818CF8); box-shadow:0 6px 18px rgba(79,70,229,.35); }
.kpi-green  { background:linear-gradient(135deg,#059669,#34D399); box-shadow:0 6px 18px rgba(5,150,105,.35); }
.kpi-cyan   { background:linear-gradient(135deg,#0284C7,#38BDF8); box-shadow:0 6px 18px rgba(2,132,199,.35); }
.kpi-orange { background:linear-gradient(135deg,#D97706,#FCD34D); box-shadow:0 6px 18px rgba(217,119,6,.35); }
.kpi-pink   { background:linear-gradient(135deg,#DB2777,#F9A8D4); box-shadow:0 6px 18px rgba(219,39,119,.35); }
.kpi-blue,.kpi-green,.kpi-cyan,.kpi-orange,.kpi-pink {
    position:relative; padding:22px 10px; color:white; text-align:center; border-radius:22px;
    min-height:132px; overflow:hidden; transform:perspective(700px) rotateX(9deg);
    transition:transform .35s ease, box-shadow .35s ease; border-bottom:6px solid rgba(0,0,0,.28);
}
.kpi-blue::after,.kpi-green::after,.kpi-cyan::after,.kpi-orange::after,.kpi-pink::after {
    content:""; position:absolute; inset:0; pointer-events:none;
    background:linear-gradient(155deg,rgba(255,255,255,.42) 0%,rgba(255,255,255,0) 46%);
}
.kpi-blue:hover,.kpi-green:hover,.kpi-cyan:hover,.kpi-orange:hover,.kpi-pink:hover {
    transform:perspective(700px) rotateX(0deg) translateY(-10px) scale(1.04);
    box-shadow:0 26px 36px -10px rgba(15,23,42,.45) !important;
}
.kpi-blue h4,.kpi-green h4,.kpi-cyan h4,.kpi-orange h4,.kpi-pink h4 { margin:0 0 8px 0; font-size:13px; font-weight:600; opacity:.95; }
.kpi-blue h2,.kpi-green h2,.kpi-cyan h2,.kpi-orange h2,.kpi-pink h2 { margin:0; font-size:30px; font-weight:900; text-shadow:0 3px 8px rgba(0,0,0,.28); }
.kpi-blue small,.kpi-green small,.kpi-cyan small,.kpi-orange small,.kpi-pink small { display:block; margin-top:4px; opacity:.9; font-size:12px; }

.photo-card { background:linear-gradient(135deg,#1E293B,#334155); border-radius:16px; padding:20px; text-align:center;
    color:white; transform:perspective(800px) rotateY(-4deg); box-shadow:0 18px 34px -10px rgba(0,0,0,.5); }
.loc-card { background:linear-gradient(135deg,#0F172A,#1E293B); border-radius:14px; padding:16px 18px; color:#E2E8F0;
    box-shadow:0 4px 14px rgba(0,0,0,.25); margin-top:10px; }
.loc-card a { color:#6EE7B7 !important; font-weight:700; }
.stay-row { background:#FFFFFF; border-left:6px solid #F59E0B; border-radius:12px; padding:10px 14px; margin:8px 0;
    box-shadow:0 6px 14px -6px rgba(15,23,42,.25); font-size:14px; color:#1E293B; }
.stay-row b.big { color:#DC2626; font-size:16px; }
.stay-row a { color:#4F46E5 !important; font-weight:700; }
.gal-card { background:#FFFFFF; border-radius:18px; padding:14px 8px 10px 8px; text-align:center;
    box-shadow:0 10px 22px -10px rgba(15,23,42,.35); border-bottom:5px solid #6366F1; margin-bottom:6px; }
.gal-card b { display:block; font-size:13px; color:#0F172A; }
.gal-card small { display:block; color:#64748B; font-size:11.5px; }
.gal-card span { font-weight:800; font-size:14px; }

.chip3d { background:linear-gradient(145deg,#FFFFFF,#EEF2FF); border-radius:18px; padding:14px 16px; min-height:96px;
    box-shadow:0 12px 24px -8px rgba(30,41,59,.35), inset 0 2px 0 #fff; border-bottom:5px solid var(--c); transition:transform .3s ease; }
.chip3d:hover { transform:translateY(-6px); }
.chip3d small { color:#64748B; font-weight:700; letter-spacing:.4px; text-transform:uppercase; font-size:11px; }
.chip3d b { display:block; font-size:16px; color:#0F172A; margin:4px 0 2px 0; }
.chip3d span { font-size:12.5px; color:#475569; }
.insight-note { background:linear-gradient(90deg,rgba(16,185,129,.14),rgba(99,102,241,.08)); border-left:5px solid #10B981;
    border-radius:10px; padding:10px 14px; font-size:13.5px; color:#1E293B; margin:4px 0 10px 0; }
.guide { background:linear-gradient(90deg,#FEF9C3,#ECFCCB); border-left:6px solid #F59E0B; border-radius:12px;
    padding:10px 14px; font-size:14px; color:#1E293B; margin:6px 0 14px 0; }

[data-testid="stMain"] h3 { border-left:7px solid #6366F1; padding:6px 0 6px 14px;
    background:linear-gradient(90deg,rgba(99,102,241,.16),transparent 70%); border-radius:8px; }
[data-testid="stPlotlyChart"], [data-testid="stDeckGlJsonChart"] { border-radius:18px; overflow:hidden; background:#FFFFFF;
    box-shadow:0 14px 28px -12px rgba(15,23,42,.35); }

.hero { display:flex; justify-content:space-between; align-items:center;
    background:linear-gradient(120deg,#0F172A,#1E1B4B 55%,#064E3B); border-radius:24px; padding:22px 30px; color:#fff;
    box-shadow:0 24px 40px -14px rgba(15,23,42,.6); margin-bottom:16px; border-bottom:6px solid #10B981; }
.hero-title { font-size:34px; font-weight:900; background:linear-gradient(90deg,#A5B4FC,#F9A8D4,#FCD34D,#6EE7B7);
    background-size:300% 100%; animation:shine 8s linear infinite;
    -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; }
.hero-sub { color:#94A3B8; font-size:15px; margin-top:2px; }
.hero-steps span { display:inline-block; background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.18);
    border-radius:999px; padding:5px 14px; margin:12px 8px 0 0; font-size:13px; }
.stTabs [data-baseweb="tab-list"] { gap:8px; }
.stTabs [data-baseweb="tab"] { background:#FFFFFF; border-radius:14px 14px 0 0; padding:10px 18px; font-weight:800;
    box-shadow:0 4px 10px rgba(0,0,0,.08); }
.stTabs [aria-selected="true"] { background:linear-gradient(135deg,#4F46E5,#7C3AED) !important; }
.stTabs [aria-selected="true"] p { color:#FFFFFF !important; }
/* ---- hide Streamlit's top-right toolbar (Edit / Deploy / menu) and footer ---- */
[data-testid="stToolbar"],
[data-testid="stToolbarActions"],
[data-testid="stAppDeployButton"],
[data-testid="stMainMenu"],
[data-testid="stDecoration"],
[data-testid="stAppViewerBadge"],
[class*="viewerBadge"],
.stDeployButton,
#MainMenu,
footer { display:none !important; visibility:hidden !important; }
[data-testid="stHeader"] { background:transparent !important; }
"""


# ===============================================================
# SMALL HELPERS
# ===============================================================
def clean_cols(d):
    """Remove line-breaks / extra spaces from column names."""
    d.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in d.columns]
    return d


def extract_hour(t):
    try:
        return int(str(t).strip().split(":")[0])
    except Exception:
        return None


def tmin(t):
    """time / 'HH:MM' -> minutes since midnight (for sorting)."""
    try:
        return t.hour * 60 + t.minute
    except AttributeError:
        try:
            hh, mm = str(t).split(":")[:2]
            return int(hh) * 60 + int(mm)
        except Exception:
            return 0


def short(name, n=14):
    name = str(name)
    return name if len(name) <= n else name[:n - 1] + "…"


def fmt_hm(h):
    """4.6667 -> '4 hr 40 min'"""
    try:
        m = int(round(float(h) * 60))
    except Exception:
        return "-"
    if m < 60:
        return f"{m} min"
    return f"{m // 60} hr {m % 60} min" if m % 60 else f"{m // 60} hr"


def initials(name):
    skip = {"md", "md.", "mst", "mst.", "mrs", "mr", "mr."}
    parts = [p for p in re.split(r"\s+", str(name).strip()) if p and p.lower() not in skip]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][0].upper()
    return (parts[0][0] + parts[1][0]).upper()


def gmap_link(lat, lon):
    return f"https://www.google.com/maps?q={lat},{lon}"


def hex_rgba(h, a):
    h = h.lstrip("#")
    return f"rgba({int(h[0:2], 16)},{int(h[2:4], 16)},{int(h[4:6], 16)},{a})"


def prod_color(v):
    return "#10B981" if v >= TARGET_PCT else ("#F59E0B" if v >= LOW_PCT else "#EF4444")


def deg_label(series):
    """One person can hold two designations (e.g. ASM + RSM) -> 'ASM + RSM'."""
    return " + ".join(dict.fromkeys(series.astype(str)))


def guide(text):
    st.markdown(f'<div class="guide">📖 <b>How to read:</b> {text}</div>', unsafe_allow_html=True)


def insight(text):
    st.markdown(f'<div class="insight-note">💡 <b>What this means:</b> {text}</div>',
                unsafe_allow_html=True)


def kpi_card(css_class, title, value, sub=""):
    sub_html = f"<small>{sub}</small>" if sub != "" else ""
    st.markdown(f'<div class="{css_class}"><h4>{title}</h4><h2>{value}</h2>{sub_html}</div>',
                unsafe_allow_html=True)


class Photos:
    """Employee photo lookup. Matches Name + Designation first, then Name only."""

    def __init__(self, image_df):
        self._by_key, self._by_name = {}, {}
        for name, deg, url in zip(image_df["Employee Name"], image_df["Designation"],
                                  image_df["Image"]):
            if isinstance(url, str) and url.strip() and url.strip().lower() != "nan":
                self._by_key[(name, deg)] = url.strip()
                self._by_name.setdefault(name, url.strip())

    def get(self, name, deg=None):
        return self._by_key.get((name, deg)) or self._by_name.get(name)


def avatar_html(name, url, size=120, border="#6366F1"):
    """Round employee photo. Falls back to coloured initials if no / blocked image."""
    col = AV_COLORS[sum(ord(c) for c in str(name)) % len(AV_COLORS)]
    fb_style = (f"width:{size}px;height:{size}px;border-radius:50%;background:{col};color:white;"
                f"font-weight:800;font-size:{int(size * 0.38)}px;line-height:{size - 8}px;"
                f"text-align:center;margin:0 auto 10px auto;border:4px solid {border};"
                "box-sizing:border-box;")
    fallback = f'<div style="{fb_style}">{initials(name)}</div>'
    if not (isinstance(url, str) and url.strip() and url.strip().lower() != "nan"):
        return fallback
    u = url.strip()
    if "?" not in u:
        u += "?width=400&height=400"
    hidden = fallback.replace('style="', 'style="display:none;', 1)
    return (f'<img src="{u}" width="{size}" height="{size}" referrerpolicy="no-referrer" loading="lazy" '
            f'style="display:block;border-radius:50%;object-fit:cover;border:4px solid {border};'
            f'margin:0 auto 10px auto;" '
            f'onerror="this.style.display=\'none\';this.nextElementSibling.style.display=\'block\';">'
            + hidden)


# ===============================================================
# CHART BUILDERS  (clean, colourful, big labels)
# ===============================================================
def show(fig):
    st.plotly_chart(fig, width="stretch", theme=None, config={"displayModeBar": False})


def style(fig, title="", height=420, legend=True):
    fig.update_layout(
        template="plotly_white", height=height,
        title=dict(text=title, x=0.02, xanchor="left", font=dict(size=18, color="#0F172A")),
        paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family="Segoe UI, Arial, sans-serif", size=13, color="#334155"),
        margin=dict(l=20, r=30, t=70, b=30),
        legend=dict(orientation="h", y=1.02, yanchor="bottom", x=1, xanchor="right", title_text=""),
        hoverlabel=dict(bgcolor="#0F172A", font_color="white", font_size=13),
        showlegend=legend, bargap=0.35,
    )
    fig.update_xaxes(showgrid=False, linecolor="#CBD5E1")
    fig.update_yaxes(gridcolor="#EEF2F7", zeroline=False)
    try:
        fig.update_layout(barcornerradius=8)   # rounded bars (newer Plotly only)
    except Exception:
        pass
    return fig


def vbar(x, y, colors, title, fmt="{:.1f}%", ytitle="", xtitle="", height=420, full=None):
    y = [float(v) for v in y]
    fig = go.Figure(go.Bar(
        x=x, y=y, marker=dict(color=colors), text=[fmt.format(v) for v in y],
        textposition="outside", cliponaxis=False, textfont=dict(size=14, color="#0F172A"),
        customdata=[(full or {}).get(a, a) for a in x],
        hovertemplate="%{customdata}: <b>%{text}</b><extra></extra>"))
    style(fig, title, height, legend=False)
    top = max(y) if y else 1
    fig.update_yaxes(range=[0, (top if top > 0 else 1) * 1.2], title=ytitle)
    fig.update_xaxes(title=xtitle)
    return fig


def hbar(names, values, colors, title, fmt="{:.1f}%", xtitle="", full=None, height=None):
    values = [float(v) for v in values]
    height = height or max(320, 110 + len(names) * 42)
    fig = go.Figure(go.Bar(
        x=values, y=names, orientation="h", marker=dict(color=colors),
        text=[fmt.format(v) for v in values], textposition="outside", cliponaxis=False,
        textfont=dict(size=13, color="#0F172A"),
        customdata=[(full or {}).get(a, a) for a in names],
        hovertemplate="%{customdata}: <b>%{text}</b><extra></extra>"))
    style(fig, title, height, legend=False)
    top = max(values) if values else 1
    fig.update_xaxes(range=[0, (top if top > 0 else 1) * 1.22], title=xtitle,
                     showgrid=True, gridcolor="#EEF2F7")
    fig.update_yaxes(autorange="reversed", showgrid=False)
    return fig


def grouped_hbar(names, series, title, fmt="{:.0f}", xtitle="", full=None, height=None):
    """series = [(label, values, colour), ...]"""
    height = height or max(360, 130 + len(names) * 62)
    fig = go.Figure()
    for lab, vals, col in series:
        vals = [float(v) for v in vals]
        fig.add_trace(go.Bar(
            x=vals, y=names, name=lab, orientation="h", marker=dict(color=col),
            text=[fmt.format(v) for v in vals], textposition="outside", cliponaxis=False,
            customdata=[(full or {}).get(a, a) for a in names],
            hovertemplate="%{customdata} | " + lab + ": <b>%{text} hr</b><extra></extra>"))
    style(fig, title, height)
    fig.update_layout(barmode="group", bargap=0.25)
    fig.update_xaxes(title=xtitle, showgrid=True, gridcolor="#EEF2F7")
    fig.update_yaxes(autorange="reversed", showgrid=False)
    return fig


def stacked_vbar(x, layers, title, fmt="{:.0f}", ytitle="", height=420):
    """layers = [(label, {x: value}, colour)] stacked bottom to top, total on top."""
    fig = go.Figure()
    totals = {a: 0.0 for a in x}
    for lab, dv, col in layers:
        vals = [float(dv.get(a, 0)) for a in x]
        for a, v in zip(x, vals):
            totals[a] += v
        fig.add_trace(go.Bar(
            x=x, y=vals, name=lab, marker=dict(color=col),
            text=[fmt.format(v) for v in vals], textposition="inside",
            textfont=dict(color="white", size=13),
            hovertemplate="%{x} | " + lab + ": <b>%{y:.1f} hr</b><extra></extra>"))
    top = max(totals.values()) if totals else 1
    fig.add_trace(go.Scatter(
        x=x, y=[totals[a] + top * 0.06 for a in x], mode="text",
        text=[f"<b>{fmt.format(totals[a])}</b>" for a in x],
        textfont=dict(size=15, color="#0F172A"), showlegend=False, hoverinfo="skip"))
    style(fig, title, height)
    fig.update_layout(barmode="stack")
    fig.update_yaxes(title=ytitle, range=[0, top * 1.18])
    return fig


def line_chart(series, title, yfmt="%", ytitle="", target=None, fill=False, height=420):
    """series = [(label, {date: value}, colour)]"""
    fig = go.Figure()
    unit = "%" if yfmt == "%" else " hr"
    for lab, dv, col in series:
        xs = sorted(dv)
        ys = [float(dv[k]) for k in xs]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, name=lab, mode="lines+markers",
            line=dict(color=col, width=4, shape="spline", smoothing=0.6),
            marker=dict(size=9, color="white", line=dict(color=col, width=3)),
            fill="tozeroy" if fill else None,
            fillcolor=hex_rgba(col, 0.12) if fill else None,
            hovertemplate="%{x|%d %b}: <b>%{y:.1f}" + unit + "</b><extra>" + lab + "</extra>"))
    style(fig, title, height)
    fig.update_xaxes(tickformat="%d %b", title="")
    fig.update_yaxes(title=ytitle, rangemode="tozero")
    if target is not None:
        fig.add_hline(y=target, line_dash="dash", line_color="#94A3B8", line_width=2,
                      annotation_text=f"{target}% target", annotation_position="top left")
    return fig


def donut(labels, values, colors, title, height=380):
    total = float(sum(values)) or 1.0
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.6, sort=False,
        marker=dict(colors=colors, line=dict(color="white", width=4)),
        textinfo="percent", textfont=dict(size=16, color="white"),
        hovertemplate="%{label}: <b>%{value:.1f} hr</b> (%{percent})<extra></extra>"))
    style(fig, title, height)
    fig.update_layout(annotations=[dict(text=f"<b>{total:.0f}</b><br>hours", x=0.5, y=0.5,
                                        showarrow=False, font=dict(size=20, color="#0F172A"))])
    return fig


def gauge(value, title, height=330):
    value = float(value)
    col = prod_color(value)
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=value,
        number=dict(suffix="%", font=dict(size=46, color=col)),
        title=dict(text=title, font=dict(size=17, color="#0F172A")),
        gauge=dict(axis=dict(range=[0, 100], tickwidth=1), bar=dict(color=col, thickness=0.3),
                   bgcolor="white", borderwidth=0,
                   steps=[dict(range=[0, LOW_PCT], color="#FEE2E2"),
                          dict(range=[LOW_PCT, TARGET_PCT], color="#FEF3C7"),
                          dict(range=[TARGET_PCT, 100], color="#D1FAE5")],
                   threshold=dict(line=dict(color="#0F172A", width=4), thickness=0.8,
                                  value=TARGET_PCT))))
    fig.update_layout(height=height, paper_bgcolor="white", margin=dict(l=30, r=30, t=70, b=10),
                      font=dict(family="Segoe UI, Arial, sans-serif"))
    return fig


# ===============================================================
# DATA LOADING
# ===============================================================
def find_excel():
    """Configured file first, else the newest Tracking_Dashboard_*.xlsx in this folder."""
    exact = BASE_DIR / EXCEL_FILE
    if exact.exists():
        return exact
    candidates = sorted(BASE_DIR.glob(EXCEL_GLOB), key=lambda p: p.stat().st_mtime, reverse=True)
    if candidates:
        return candidates[0]
    raise FileNotFoundError(f"'{EXCEL_FILE}' not found in {BASE_DIR}")


@st.cache_data(show_spinner="Loading Excel data…")
def load_data(path, mtime):
    """Read the 4 sheets and tidy them. `mtime` makes the cache refresh when the file changes."""
    sheets = pd.read_excel(path, sheet_name=list(REQUIRED_COLUMNS))
    for name, d in sheets.items():
        clean_cols(d)
        missing = [c for c in REQUIRED_COLUMNS[name] if c not in d.columns]
        if missing:
            raise ValueError(f"Sheet '{name}' is missing column(s): {', '.join(missing)}")

    summary, image_df = sheets[SHEET_SUMMARY], sheets[SHEET_IMAGES]
    same_place, gps = sheets[SHEET_SAME_PLACE], sheets[SHEET_GPS]

    for d in (summary, same_place, gps):
        d["Date"] = pd.to_datetime(d["Date"], errors="coerce")
        for col in ("Employee Name", "Group", "Deg", "Day"):
            d[col] = d[col].astype(str).str.strip()
    summary.dropna(subset=["Date"], inplace=True)
    same_place.dropna(subset=["Date"], inplace=True)
    gps.dropna(subset=["Date"], inplace=True)

    for d in (same_place, gps):
        for col in ("Lat", "Lon"):
            d[col] = pd.to_numeric(d[col], errors="coerce")
        d["Hour"] = d["Start Time"].apply(extract_hour)

    for col in ("Employee Name", "Designation"):
        image_df[col] = image_df[col].astype(str).str.strip()
    return summary, image_df, same_place, gps


# ===============================================================
# SIDEBAR + FILTERS
# ===============================================================
@dataclass
class Filters:
    start: object
    end: object
    days: list
    hour_min: int
    hour_max: int
    group: str
    deg: str
    employee: str


def render_sidebar(summary_df):
    sb = st.sidebar
    sb.markdown(
        '<div style="background:white;border-radius:12px;padding:10px;margin-bottom:10px;text-align:center;">'
        f'<img src="data:image/jpeg;base64,{PREMIO_B64}" style="width:100%;max-width:200px;border-radius:8px;"></div>'
        '<div style="text-align:center;margin-bottom:18px;">'
        f'<img src="data:image/png;base64,{ACI_B64}" style="width:64px;height:64px;border-radius:50%;'
        'border:3px solid #10B981;"></div>',
        unsafe_allow_html=True)
    sb.title("🎛 Dashboard Filters")

    min_date, max_date = summary_df["Date"].min().date(), summary_df["Date"].max().date()
    start = sb.date_input("📅 Start Date", min_date, min_value=min_date,
                          max_value=max_date, format="YYYY/MM/DD")
    end = sb.date_input("📅 End Date", max_date, min_value=min_date,
                        max_value=max_date, format="YYYY/MM/DD")

    days_in_data = set(summary_df["Day"].unique())
    available_days = [d for d in WEEKDAY_ORDER if d in days_in_data]
    days = sb.multiselect("📆 Day Filter", available_days, default=available_days)

    hour_min, hour_max = sb.slider("⏰ Time Range (Hour of Day)", 0, 23, (0, 23),
                                   help="Applies to the Maps and the Same-Place section")

    group = sb.selectbox("🗺 Select Area / Group", ["All"] + sorted(summary_df["Group"].dropna().unique()))
    deg = sb.selectbox("🏷 Select Designation", ["All"] + sorted(summary_df["Deg"].dropna().unique()))
    employee = sb.selectbox("👤 Select Employee",
                            ["All"] + sorted(summary_df["Employee Name"].dropna().unique()),
                            key="sel_emp")
    return Filters(start, end, days, hour_min, hour_max, group, deg, employee)


def apply_filters(data, f):
    out = data[(data["Date"] >= pd.to_datetime(f.start)) & (data["Date"] <= pd.to_datetime(f.end))]
    if f.days:
        out = out[out["Day"].isin(f.days)]
    if f.group != "All":
        out = out[out["Group"] == f.group]
    if f.deg != "All":
        out = out[out["Deg"] == f.deg]
    if f.employee != "All":
        out = out[out["Employee Name"] == f.employee]
    return out.copy()


def filter_hours(d, f):
    return d[d["Hour"].isna() | ((d["Hour"] >= f.hour_min) & (d["Hour"] <= f.hour_max))]


def pick_employee(name):
    st.session_state["sel_emp"] = name


def clear_employee():
    st.session_state["sel_emp"] = "All"


@dataclass
class Ctx:
    """Everything the tabs need, already filtered by the sidebar."""
    f: Filters
    df: pd.DataFrame        # Summarize_Data (filtered)
    gps: pd.DataFrame       # Employee_Data  (filtered)
    sp: pd.DataFrame        # Same_Place_Found (filtered)
    photos: Photos
    month_label: str
    total_emp: int
    total_working: float
    total_travel: float
    total_idle: float
    avg_prod: float

    @property
    def employee(self):
        return self.f.employee


# ===============================================================
# SAME-PLACE HELPERS
# ===============================================================
def build_stays(src):
    """Tidy table of same-place records (index reset so row N = position N)."""
    t = src.dropna(subset=["Lat", "Lon"]).copy()
    if t.empty:
        return pd.DataFrame()
    t["tm"] = t["Start Time"].apply(tmin)
    t = t.sort_values(["Date", "tm"]).reset_index(drop=True)
    return pd.DataFrame({
        "Employee": t["Employee Name"], "Designation": t["Deg"], "Area": t["Group"],
        "Date": t["Date"].dt.strftime("%d %b %Y"), "Day": t["Day"],
        "From": t["Start Time"].astype(str).str[:5], "To": t["End Time"].astype(str).str[:5],
        "Stayed": t[COL_IDLE].apply(fmt_hm), "Hours": t[COL_IDLE].astype(float).round(1),
        "Address": t[COL_ADDRESS].astype(str).str.strip(), "lat": t["Lat"], "lon": t["Lon"],
    })


def stay_card(r, photos):
    link = gmap_link(r["lat"], r["lon"])
    photo = ('<div style="flex:0 0 96px;">'
             + avatar_html(r["Employee"], photos.get(r["Employee"], r["Designation"]), 84, "#10B981")
             + "</div>")
    st.markdown(
        '<div class="loc-card" style="display:flex;gap:16px;align-items:center;">' + photo +
        '<div style="flex:1;">'
        f'<b style="font-size:17px;">{r["Employee"]}</b> '
        f'<span style="opacity:.75;">({r["Designation"]} | {r["Area"]})</span><br>'
        f'📅 {r["Date"]} ({r["Day"]}) &nbsp; ⏰ {r["From"]} → {r["To"]}<br>'
        f'⏳ Stayed at the same place: <b style="color:#FCA5A5;font-size:18px;">{r["Stayed"]}</b><br>'
        f'📍 {r["Address"][:160]}<br>'
        f'<a href="{link}" target="_blank">🗺 Open in Google Maps</a></div></div>',
        unsafe_allow_html=True)


def stay_rows(stays):
    for _, r in stays.iterrows():
        st.markdown(
            '<div class="stay-row">'
            f'📅 <b>{r["Date"]}</b> ({r["Day"]}) &nbsp;⏰ {r["From"]} → {r["To"]} &nbsp;'
            f'⏳ <b class="big">{r["Stayed"]}</b><br>📍 {r["Address"][:130]} &nbsp;'
            f'<a href="{gmap_link(r["lat"], r["lon"])}" target="_blank">🗺 Map</a></div>',
            unsafe_allow_html=True)


# ===============================================================
# EMPLOYEE PROFILE BLOCK (photo + full summary)
# ===============================================================
def render_profile(ctx, name, tag):
    ei = ctx.df[ctx.df["Employee Name"] == name]
    if ei.empty:
        st.warning("No records for this employee with the current filters.")
        return
    first_deg = ei["Deg"].iloc[0]
    deg_val, group_val = deg_label(ei["Deg"]), ei["Group"].iloc[0]

    pc, mc = st.columns([1, 3])
    with pc:
        st.markdown(
            f'<div class="photo-card">{avatar_html(name, ctx.photos.get(name, first_deg), 150)}'
            f'<h3 style="margin:6px 0 4px 0;font-size:16px;border:none;background:none;color:white;">{name}</h3>'
            f'<p style="margin:0;font-size:13px;opacity:.8;">{deg_val} | {group_val}</p></div>',
            unsafe_allow_html=True)
        st.button("⬅ Back to all employees", key=f"back_{tag}", on_click=clear_employee)

    stays = build_stays(ctx.sp[ctx.sp["Employee Name"] == name])
    with mc:
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Working Days", ei["Date"].nunique())
        m2.metric("Working Hr", round(ei["Total Working Hr"].sum(), 1))
        m3.metric("Travel Hr", round(ei["Total Travel Hr"].sum(), 1))
        m4.metric("Idle Hr", round(ei["Total Idle Hr"].sum(), 1))
        m5.metric("Productivity", f"{round(ei[COL_PRODUCTIVITY].mean() * 100, 1)}%")

        low_days = int(ei["Remarks"].astype(str).str.lower().str.contains("low").sum())
        flag_days = int((~ei["Auditing"].astype(str).str.lower().str.contains("no major")).sum())
        n1, n2, n3, n4 = st.columns(4)
        n1.metric("✅ Productive Days", len(ei) - low_days)
        n2.metric("⚠ Low-Productive Days", low_days)
        n3.metric("🚩 Audit-Flagged Days", flag_days)
        n4.metric("📍 Same-Place Hr", round(float(stays["Hours"].sum()), 1) if not stays.empty else 0)

        locs = ctx.gps[ctx.gps["Employee Name"] == name].dropna(subset=["Lat", "Lon"]).copy()
        if not locs.empty:
            locs["tm"] = locs["Start Time"].apply(tmin)
            lr = locs.sort_values(["Date", "tm"]).iloc[-1]
            st.markdown(
                '<div class="loc-card"><b>📍 Last Known Location</b><br>'
                f'{str(lr.get(COL_ADDRESS, "-"))[:140]}<br>'
                f'<span style="opacity:.75;">Date: {lr["Date"].date()} | Time: {str(lr.get("Start Time", "-"))[:5]} | '
                f'Status: {lr.get("Status", "-")}</span><br>'
                f'<a href="{gmap_link(lr["Lat"], lr["Lon"])}" target="_blank">🗺 Open in Google Maps</a></div>',
                unsafe_allow_html=True)

    st.markdown("##### 📍 Same-place stays (how long this person stayed at one location)")
    if stays.empty:
        st.success("No same-place stay found for this employee with the current filters. 👍")
    else:
        stay_rows(stays)


# ===============================================================
# TAB 1 : OVERVIEW
# ===============================================================
def render_overview(ctx):
    df = ctx.df
    cards = (("kpi-blue", "👥 Employee", ctx.total_emp),
             ("kpi-green", "⏱ Working Hr", ctx.total_working),
             ("kpi-cyan", "🚗 Travel Hr", ctx.total_travel),
             ("kpi-orange", "🧍 Idle Hr", ctx.total_idle),
             ("kpi-pink", "📈 Productivity %", f"{ctx.avg_prod}%"))
    for col, (cls, title, val) in zip(st.columns(5), cards):
        with col:
            kpi_card(cls, title, val)
    st.markdown("<br>", unsafe_allow_html=True)

    st.subheader("📷 Employee Profile")
    if ctx.employee == "All":
        st.info("👈 Select an employee from **Select Employee** on the left (or click a photo in the "
                "**🧑‍💼 Employee 360°** tab) — photo, full summary and same-place stays will appear here.")
    else:
        render_profile(ctx, ctx.employee, "ov")
    st.markdown("---")

    prod_emp = df.groupby("Employee Name")[COL_PRODUCTIVITY].mean()
    idle_emp = df.groupby("Employee Name")["Total Idle Hr"].sum()
    day_prod = df.groupby("Day")[COL_PRODUCTIVITY].mean()

    g_col, chip_col = st.columns([1, 2])
    with g_col:
        show(gauge(ctx.avg_prod, "Team Productivity"))
        guide(f"The needle shows the team's Productivity %. The <b>black line is the {TARGET_PCT}% target</b>. "
              "Green = good, yellow = average, red = low.")
    with chip_col:
        k1, k2 = st.columns(2)
        k3, k4 = st.columns(2)
        chips = [
            (k1, "#10B981", "🏆 Top Performer", prod_emp.idxmax(), f"{prod_emp.max() * 100:.1f}% productivity"),
            (k2, "#EF4444", "⚠ Needs Attention", prod_emp.idxmin(), f"{prod_emp.min() * 100:.1f}% productivity"),
            (k3, "#F59E0B", "🧍 Highest Idle Time", idle_emp.idxmax(), f"{idle_emp.max():.1f} idle hrs in period"),
            (k4, "#6366F1", "📅 Best Day", day_prod.idxmax(), f"{day_prod.max() * 100:.1f}% avg productivity"),
        ]
        for col, color, title, nm, sub in chips:
            with col:
                st.markdown(f'<div class="chip3d" style="--c:{color};"><small>{title}</small>'
                            f'<b>{nm}</b><span>{sub}</span></div>', unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(f'<div class="insight-note">📍 <b>{len(ctx.sp)}</b> same-place (long stay at one location) '
                    f'records found. Open the <b>📍 Same Place</b> tab to see who, how long and where.</div>',
                    unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("🍩 Team Time Split  |  🏢 Area Performance")
    d1, d2 = st.columns(2)
    with d1:
        show(donut(["Travel", "Idle"], [ctx.total_travel, ctx.total_idle], [C_TRAVEL, C_IDLE],
                   "Working Hours: Travel vs Idle"))
        guide("Each slice is a share of working time. <b>Blue = travelling</b>, <b>orange = idle</b>. "
              "The smaller the orange slice, the better.")
    with d2:
        gd = df.groupby(["Group", "Deg"], as_index=False)[COL_PRODUCTIVITY].mean()
        gd["P"] = gd[COL_PRODUCTIVITY] * 100
        fig = go.Figure()
        for deg in sorted(gd["Deg"].unique()):
            sub = gd[gd["Deg"] == deg]
            fig.add_trace(go.Bar(
                x=sub["Group"], y=sub["P"], name=deg,
                marker=dict(color=DEG_COLOR_MAP.get(deg, "#64748B")),
                text=[f"{v:.1f}%" for v in sub["P"]], textposition="outside", cliponaxis=False,
                hovertemplate="%{x} | " + deg + ": <b>%{y:.1f}%</b><extra></extra>"))
        style(fig, "Productivity % — Area x Designation", 380)
        fig.update_layout(barmode="group", bargap=0.25)
        fig.update_yaxes(range=[0, max(gd["P"].max() * 1.25, 10)], title="Productivity %")
        show(fig)
        guide("One bar for each Designation inside each Area. <b>Taller bar = higher Productivity.</b>")

    st.subheader("📈 Daily Productivity Trend")
    series = []
    for gi, g in enumerate(sorted(df["Group"].unique())):
        dg = df[df["Group"] == g].groupby(df["Date"].dt.date)[COL_PRODUCTIVITY].mean() * 100
        series.append((g, {d: float(v) for d, v in dg.items()}, GROUP_PALETTE[gi % len(GROUP_PALETTE)]))
    show(line_chart(series, "Daily Avg Productivity % by Area", target=TARGET_PCT, ytitle="Productivity %"))
    guide("Each coloured line is one Area, day by day. <b>Higher line = better day.</b> "
          f"The dashed line is the {TARGET_PCT}% target — try to stay above it.")


# ===============================================================
# TAB 2 : EMPLOYEE 360
# ===============================================================
def render_employee360(ctx):
    df = ctx.df
    if ctx.employee == "All":
        st.subheader("👥 Click an employee to open the full profile")
        sort_by = st.radio("Sort by", ["Name A → Z", "Productivity high → low"], horizontal=True)
        roster = df.groupby("Employee Name").agg(
            Deg=("Deg", deg_label), FirstDeg=("Deg", "first"), Group=("Group", "first"),
            P=(COL_PRODUCTIVITY, "mean")).reset_index().rename(columns={"Employee Name": "Name"})
        roster["P"] = (roster["P"] * 100).round(1)
        roster = (roster.sort_values("P", ascending=False) if sort_by.startswith("Prod")
                  else roster.sort_values("Name"))
        rows = list(roster.itertuples(index=False))
        per_row = 6
        for i in range(0, len(rows), per_row):
            for col, r in zip(st.columns(per_row), rows[i:i + per_row]):
                with col:
                    st.markdown(
                        f'<div class="gal-card">{avatar_html(r.Name, ctx.photos.get(r.Name, r.FirstDeg), 84)}'
                        f'<b>{short(r.Name, 20)}</b><small>{r.Deg} • {r.Group}</small>'
                        f'<span style="color:{prod_color(r.P)};">{r.P:.1f}%</span></div>',
                        unsafe_allow_html=True)
                    st.button("View", key=f"pick_{r.Name}", on_click=pick_employee,
                              args=(r.Name,), width="stretch")
        guide("Click <b>View</b> under any photo — the whole dashboard filters to that person and the "
              "<b>🏠 Overview</b> tab shows the photo + full summary. Come back here for their charts.")
        return

    name = ctx.employee
    render_profile(ctx, name, "e360")
    ei = df[df["Employee Name"] == name]
    emp_prod = round(ei[COL_PRODUCTIVITY].mean() * 100, 1)
    st.markdown("---")

    e1, e2 = st.columns(2)
    with e1:
        show(gauge(emp_prod, short(name, 24)))
        guide(f"Needle = this person's Productivity %. <b>Black line = {TARGET_PCT}% target</b>.")
    with e2:
        show(donut(["Travel", "Idle"],
                   [float(ei["Total Travel Hr"].sum()), float(ei["Total Idle Hr"].sum())],
                   [C_TRAVEL, C_IDLE], "Travel vs Idle"))
        guide("Blue = time travelling, orange = time idle.")

    ed = ei.groupby(ei["Date"].dt.date)[COL_PRODUCTIVITY].mean() * 100
    show(line_chart([(short(name, 20), {d: float(v) for d, v in ed.items()}, C_WORK)],
                    "Daily Productivity %", target=TARGET_PCT, fill=True, ytitle="Productivity %"))
    guide("Each dot is one working day. A dip = a low-productivity day.")

    dd = ei.groupby(ei["Date"].dt.date)[["Total Travel Hr", "Total Idle Hr"]].sum()
    xl = [f"{d:%d %b}" for d in dd.index]
    show(stacked_vbar(xl, [("Travel", dict(zip(xl, dd["Total Travel Hr"])), C_TRAVEL),
                           ("Idle", dict(zip(xl, dd["Total Idle Hr"])), C_IDLE)],
                      "Daily Hours — Travel + Idle", fmt="{:.0f}", ytitle="Hours"))
    guide("Each bar is one day: <b>blue = travel</b>, <b>orange = idle</b>. The number on top is the total hours.")

    st.subheader("🗓 Day-by-day table")
    tbl = ei.sort_values("Date")[["Date", "Day", "Deg", "Total Working Hr", "Total Travel Hr", "Total Idle Hr",
                                  "Total Tp Visit", COL_PRODUCTIVITY, "Remarks", "Auditing"]].copy()
    tbl["Date"] = tbl["Date"].dt.date
    tbl[COL_PRODUCTIVITY] = (tbl[COL_PRODUCTIVITY] * 100).round(1)
    st.dataframe(tbl.rename(columns={COL_PRODUCTIVITY: "Productivity %"}), width="stretch", hide_index=True)
    st.info("🗺 To see this person's full-day route, open the **🗺 Maps** tab and choose *Employee Route*.")


# ===============================================================
# TAB 3 : CHARTS
# ===============================================================
def render_charts(ctx):
    df = ctx.df

    st.subheader("📆 Productivity by Day of Week")
    wd = (df.groupby("Day")[COL_PRODUCTIVITY].mean() * 100).reindex(WEEKDAY_ORDER).dropna()
    show(vbar(list(wd.index), wd.values, [prod_color(v) for v in wd.values],
              "Which day performs best?", ytitle="Productivity %"))
    guide(f"One bar per weekday. <b>Green = {TARGET_PCT}%+ (good)</b>, <b>yellow = {LOW_PCT}–{TARGET_PCT}%</b>, "
          f"<b>red = below {LOW_PCT}%</b>.")
    if len(wd):
        insight(f"Best day: <b>{wd.idxmax()}</b> ({wd.max():.1f}%). Weakest day: <b>{wd.idxmin()}</b> ({wd.min():.1f}%).")

    st.subheader("🏷 Productivity by Designation  |  🏢 by Area")
    a1, a2 = st.columns(2)
    with a1:
        dg = (df.groupby("Deg")[COL_PRODUCTIVITY].mean() * 100).sort_values(ascending=False)
        show(vbar(list(dg.index), dg.values, [DEG_COLOR_MAP.get(k, "#64748B") for k in dg.index],
                  "By Designation", ytitle="Productivity %", height=380))
    with a2:
        ag = (df.groupby("Group")[COL_PRODUCTIVITY].mean() * 100).sort_values(ascending=False)
        show(vbar(list(ag.index), ag.values, GROUP_PALETTE[:len(ag)], "By Area",
                  ytitle="Productivity %", height=380))
    guide("A taller bar means that Designation / Area is more productive.")

    st.subheader("🚗 Travel vs Idle — Top 10 by Idle Hours")
    ti = (df.groupby("Employee Name")[["Total Travel Hr", "Total Idle Hr"]].sum()
          .sort_values("Total Idle Hr", ascending=False).head(10))
    nm = [short(n, 22) for n in ti.index]
    show(grouped_hbar(nm, [("Travel", ti["Total Travel Hr"].tolist(), C_TRAVEL),
                           ("Idle", ti["Total Idle Hr"].tolist(), C_IDLE)],
                      "Hours per employee", xtitle="Hours", full=dict(zip(nm, ti.index))))
    guide("<b>Blue bar = travel hours</b>, <b>orange bar = idle hours</b>. "
          "A long orange bar = that person was idle a lot.")

    st.subheader("🧱 Designation Hours (Travel + Idle)")
    lv = df.groupby("Deg")[["Total Travel Hr", "Total Idle Hr"]].sum()
    show(stacked_vbar(list(lv.index), [("Travel", lv["Total Travel Hr"].to_dict(), C_TRAVEL),
                                       ("Idle", lv["Total Idle Hr"].to_dict(), C_IDLE)],
                      "Total hours by Designation", ytitle="Hours"))
    guide("Each tower = total working time of a Designation: <b>blue (travel)</b> + <b>orange (idle)</b>. "
          "The number on top is the total hours.")

    st.subheader("🛒 Average TP Visits per Day by Designation")
    tp = df.groupby("Deg")["Total Tp Visit"].mean().sort_values(ascending=False)
    show(vbar(list(tp.index), tp.values, [DEG_COLOR_MAP.get(k, "#64748B") for k in tp.index],
              "Avg TP Visit per day", fmt="{:.1f}", ytitle="TP Visits", height=380))
    guide("More TP visits per day = more field activity.")

    st.subheader("📈 Daily Travel & Idle Hours (whole team)")
    tr = df.groupby(df["Date"].dt.date)[["Total Travel Hr", "Total Idle Hr"]].sum()
    show(line_chart([("Travel", {d: float(v) for d, v in tr["Total Travel Hr"].items()}, C_TRAVEL),
                     ("Idle", {d: float(v) for d, v in tr["Total Idle Hr"].items()}, C_IDLE)],
                    "Team hours per day", yfmt="hr", ytitle="Hours"))
    guide("Blue line = travel, orange line = idle, day by day. You want orange <b>going down</b> over time.")


# ===============================================================
# TAB 4 : SAME PLACE
# ===============================================================
def render_same_place(ctx):
    st.subheader("📍 Same Place Summary — who stayed how long, and where")
    stays = build_stays(ctx.sp)
    if stays.empty:
        st.info("No same-place records for the current filters.")
        return

    longest = stays.loc[stays["Hours"].idxmax()]
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        kpi_card("kpi-blue", "📍 Same-Place Records", len(stays))
    with s2:
        kpi_card("kpi-green", "👥 Employees", stays["Employee"].nunique())
    with s3:
        kpi_card("kpi-orange", "⏳ Total Stay Hr", round(float(stays["Hours"].sum()), 1))
    with s4:
        kpi_card("kpi-pink", "🔥 Longest Stay", longest["Stayed"], short(longest["Employee"], 22))
    st.markdown("<br>", unsafe_allow_html=True)

    per = (stays.groupby("Employee")
           .agg(Hours=("Hours", "sum"), Stays=("Hours", "count"))
           .sort_values("Hours", ascending=False))
    nm = [short(n, 22) for n in per.index]
    full = dict(zip(nm, per.index))
    b1, b2 = st.columns(2)
    with b1:
        show(hbar(nm, per["Hours"].tolist(), ["#EF4444" if h >= 8 else "#F59E0B" for h in per["Hours"]],
                  "Total hours stayed at the same place", fmt="{:.1f} hr", xtitle="Hours", full=full))
    with b2:
        show(hbar(nm, per["Stays"].tolist(), ["#6366F1"] * len(nm),
                  "How many times it happened", fmt="{:.0f}", xtitle="Times", full=full))
    guide("<b>Left:</b> total hours each person stayed at one place (red = 8+ hrs). "
          "<b>Right:</b> how many times it happened.")

    st.markdown("##### 👆 Click a row to see exactly where it happened")
    show_cols = ["Employee", "Designation", "Area", "Date", "Day", "From", "To", "Stayed", "Address"]
    event = st.dataframe(
        stays[show_cols], width="stretch", hide_index=True, on_select="rerun",
        selection_mode="single-row", key="stay_table",
        column_config={"Address": st.column_config.TextColumn("Address", width="large")})
    sel_rows = event.selection.rows if event and event.selection else []
    sel_idx = sel_rows[0] if sel_rows and sel_rows[0] < len(stays) else None

    if sel_idx is not None:
        stay_card(stays.iloc[sel_idx], ctx.photos)
        c_lat, c_lon, zoom = float(stays.iloc[sel_idx]["lat"]), float(stays.iloc[sel_idx]["lon"]), 16
    else:
        st.caption("No row selected — showing all locations. Select a row above to zoom in on it.")
        c_lat, c_lon, zoom = float(stays["lat"].mean()), float(stays["lon"].mean()), 7

    fmap = folium.Map(location=[c_lat, c_lon], zoom_start=zoom, tiles="OpenStreetMap")
    for i, r in stays.iterrows():
        is_sel = sel_idx is not None and i == sel_idx
        colr = "darkred" if r["Hours"] >= 8 else ("red" if r["Hours"] >= 5 else "orange")
        popup = (f'<b>{r["Employee"]}</b> ({r["Designation"]}, {r["Area"]})<br>'
                 f'Date: {r["Date"]} ({r["Day"]})<br>Time: {r["From"]} - {r["To"]}<br>'
                 f'Stayed: <b>{r["Stayed"]}</b><br>Address: {r["Address"][:110]}<br>'
                 f'<a href="{gmap_link(r["lat"], r["lon"])}" target="_blank">Open in Google Maps</a>')
        folium.Marker(
            [r["lat"], r["lon"]], popup=folium.Popup(popup, max_width=300),
            tooltip=f'{r["Employee"]} — {r["Stayed"]}',
            icon=folium.Icon(color="green" if is_sel else colr,
                             icon="star" if is_sel else "exclamation-sign")).add_to(fmap)
    if sel_idx is None and len(stays) > 1:
        fmap.fit_bounds([[stays["lat"].min(), stays["lon"].min()],
                         [stays["lat"].max(), stays["lon"].max()]])

    out = st_folium(fmap, height=520, use_container_width=True, key="same_place_map",
                    returned_objects=["last_object_clicked"])
    clicked = out.get("last_object_clicked") if out else None
    if clicked:
        hit = stays[((stays["lat"] - clicked["lat"]).abs() < 2e-4)
                    & ((stays["lon"] - clicked["lng"]).abs() < 2e-4)]
        if not hit.empty:
            st.markdown("**You clicked on the map:**")
            stay_card(hit.iloc[0], ctx.photos)
    guide("<b>Click any pin</b> on the map — a popup shows the employee, time, how long they stayed and the address. "
          "Or click a row in the table above and the map jumps straight to that place.")


# ===============================================================
# TAB 5 : MAPS
# ===============================================================
def map_idle_density(ctx):
    pts = ctx.gps[ctx.gps["Status"] == STATUS_IDLE].dropna(subset=["Lat", "Lon"])
    pts = pts[["Lat", "Lon"]].rename(columns={"Lat": "lat", "Lon": "lng"})
    if pts.empty:
        st.info("No idle GPS data for the current filters.")
        return
    layer = pdk.Layer("HexagonLayer", data=pts, get_position="[lng, lat]", radius=2500,
                      elevation_scale=45, elevation_range=[0, 3000], extruded=True, coverage=0.9,
                      pickable=True, auto_highlight=True)
    view = pdk.ViewState(latitude=float(pts["lat"].mean()), longitude=float(pts["lng"].mean()),
                         zoom=6.3, pitch=55, bearing=12)
    st.pydeck_chart(pdk.Deck(layers=[layer], initial_view_state=view,
                             tooltip={"text": "Stops in this area: {elevationValue}"}))
    guide("<b>Taller tower = more stops / idle time in that area.</b> "
          "Hold Ctrl and drag the mouse to tilt and rotate.")


def map_same_place_hotspots(ctx):
    hs = ctx.sp.dropna(subset=["Lat", "Lon"])
    if hs.empty:
        st.info("No same-place data for the current filters.")
        return
    hs = pd.DataFrame({
        "lat": hs["Lat"], "lng": hs["Lon"], "idle": hs[COL_IDLE].astype(float).round(1),
        "name": hs["Employee Name"], "date": hs["Date"].dt.strftime("%d %b %Y"),
        "addr": hs[COL_ADDRESS].astype(str).str[:90]})
    layer = pdk.Layer("ColumnLayer", data=hs, get_position="[lng, lat]", get_elevation="idle",
                      elevation_scale=5000, radius=4000, get_fill_color="[239, 68, 68, 210]",
                      extruded=True, pickable=True, auto_highlight=True)
    view = pdk.ViewState(latitude=float(hs["lat"].mean()), longitude=float(hs["lng"].mean()),
                         zoom=6.3, pitch=55, bearing=12)
    st.pydeck_chart(pdk.Deck(layers=[layer], initial_view_state=view, tooltip={
        "html": "<b>{name}</b><br/>{idle} hours at the same place<br/>{date}<br/>{addr}",
        "style": {"backgroundColor": "#0F172A", "color": "white"}}))
    guide("<b>Taller red column = more hours at the same place.</b> Hover over a column to see "
          "<b>who, when, how many hours and the address</b>. For click-to-zoom use the <b>📍 Same Place</b> tab.")


def map_employee_route(ctx):
    name = ctx.employee
    if name == "All":
        st.info("👈 Select an employee on the left to see the route.")
        return
    rt = ctx.gps[ctx.gps["Employee Name"] == name].dropna(subset=["Lat", "Lon"]).copy()
    if rt.empty:
        st.warning("No GPS data for the current filters.")
        return

    sel_d = st.selectbox("Select date", sorted(rt["Date"].dt.date.unique()), key="route_date")
    rt = rt[rt["Date"].dt.date == sel_d].copy()
    rt["tm"] = rt["Start Time"].apply(tmin)
    rt = rt.sort_values("tm")
    rt["time"] = rt["Start Time"].astype(str).str[:5] + " - " + rt["End Time"].astype(str).str[:5]
    rt["addr"] = rt[COL_ADDRESS].astype(str).str[:90]
    rt["name"] = name

    path = [[float(a), float(b)] for a, b in zip(rt["Lon"], rt["Lat"])]
    stops_src = rt[rt["Status"] == STATUS_IDLE]
    stops = pd.DataFrame({
        "lat": stops_src["Lat"], "lng": stops_src["Lon"],
        "idle": stops_src[COL_IDLE].astype(float).clip(lower=0.3),
        "name": stops_src["name"], "time": stops_src["time"], "addr": stops_src["addr"],
        "status": "Idle: " + stops_src[COL_IDLE].astype(float).round(1).astype(str) + " hr"})
    trav_src = rt[rt["Status"] == STATUS_TRAVEL]
    trav = pd.DataFrame({"lat": trav_src["Lat"], "lng": trav_src["Lon"], "name": trav_src["name"],
                         "time": trav_src["time"], "addr": trav_src["addr"], "status": "Travelling"})
    ends = pd.DataFrame({
        "lat": [rt.iloc[0]["Lat"], rt.iloc[-1]["Lat"]], "lng": [rt.iloc[0]["Lon"], rt.iloc[-1]["Lon"]],
        "name": name, "time": [rt.iloc[0]["time"], rt.iloc[-1]["time"]],
        "addr": [rt.iloc[0]["addr"], rt.iloc[-1]["addr"]],
        "status": ["🟢 Start of day", "🔴 End of day"]})

    layers = [
        pdk.Layer("PathLayer", data=[{"path": path}], get_path="path", get_width=60,
                  width_min_pixels=3, get_color=[99, 102, 241]),
        pdk.Layer("ScatterplotLayer", data=trav, get_position="[lng, lat]", get_radius=180,
                  get_fill_color=[6, 182, 212, 220], pickable=True),
        pdk.Layer("ColumnLayer", data=stops, get_position="[lng, lat]", get_elevation="idle",
                  elevation_scale=1200, radius=350, get_fill_color=[245, 158, 11, 230],
                  extruded=True, pickable=True, auto_highlight=True),
        pdk.Layer("ScatterplotLayer", data=ends, get_position="[lng, lat]", get_radius=450,
                  get_fill_color=[16, 185, 129, 255], pickable=True),
    ]
    view = pdk.ViewState(latitude=float(rt["Lat"].mean()), longitude=float(rt["Lon"].mean()),
                         zoom=11, pitch=55, bearing=10)
    r1, r2, r3 = st.columns(3)
    r1.metric("GPS Points", len(rt))
    r2.metric("Travelling", len(trav))
    r3.metric("Idle Stops", len(stops))
    st.pydeck_chart(pdk.Deck(layers=layers, initial_view_state=view, tooltip={
        "html": "<b>{name}</b><br/>{status}<br/>Time: {time}<br/>{addr}",
        "style": {"backgroundColor": "#0F172A", "color": "white"}}))
    guide("<b>Blue line = path travelled</b>, <b>blue dots = travelling</b>, <b>orange columns = where the "
          "person stopped (taller = longer)</b>, <b>green dots = start and end of day</b>. "
          "Hover over a dot or column to see the address and time.")


def map_classic(ctx):
    if ctx.employee == "All":
        st.info("Select a specific employee from the sidebar to view the GPS route map.")
        return
    src = ctx.gps[ctx.gps["Employee Name"] == ctx.employee].dropna(subset=["Lat", "Lon"]).copy()
    if src.empty:
        st.warning("No GPS data found for this employee in the selected filters.")
        return

    dates = sorted(src["Date"].dt.date.unique())
    chosen = st.selectbox("Select Date for GPS Map", ["All Dates"] + [str(d) for d in dates], key="map_date")
    if chosen != "All Dates":
        src = src[src["Date"].dt.date == pd.to_datetime(chosen).date()]
    if src.empty:
        return

    src["tm"] = src["Start Time"].apply(tmin)
    src = src.sort_values(["Date", "tm"])
    m = folium.Map(location=[src["Lat"].mean(), src["Lon"].mean()], zoom_start=12)
    coords = list(zip(src["Lat"], src["Lon"]))
    if len(coords) > 1:
        folium.PolyLine(coords, color="#6366F1", weight=3, opacity=0.7).add_to(m)
    cluster = MarkerCluster().add_to(m)
    status_color = {STATUS_TRAVEL: "blue", STATUS_IDLE: "red"}
    for _, row in src.iterrows():
        color = status_color.get(str(row["Status"]), "gray")
        popup_text = (f"<b>{row['Employee Name']}</b><br>Date: {row['Date'].date()}<br>"
                      f"Time: {str(row.get('Start Time', '-'))[:5]} - {str(row.get('End Time', '-'))[:5]}<br>"
                      f"Address: {str(row.get(COL_ADDRESS, '')).strip()[:90]}<br>Status: {row['Status']}")
        folium.CircleMarker(location=[row["Lat"], row["Lon"]], radius=5, color=color, fill=True,
                            fill_color=color, fill_opacity=0.8,
                            popup=folium.Popup(popup_text, max_width=280),
                            tooltip=str(row["Status"])).add_to(cluster)
    folium.Marker([src.iloc[0]["Lat"], src.iloc[0]["Lon"]], popup="Start",
                  icon=folium.Icon(color="green", icon="play")).add_to(m)
    folium.Marker([src.iloc[-1]["Lat"], src.iloc[-1]["Lon"]], popup="End",
                  icon=folium.Icon(color="red", icon="stop")).add_to(m)
    c1, c2, c3 = st.columns(3)
    c1.metric("Total GPS Points", len(src))
    c2.metric("Travelling", int((src["Status"] == STATUS_TRAVEL).sum()))
    c3.metric("Idling", int((src["Status"] == STATUS_IDLE).sum()))
    st_folium(m, height=550, use_container_width=True, key="gps_route_map", returned_objects=[])


def render_maps(ctx):
    st.subheader("🗺 Maps — where did things happen")
    mode = st.radio("Which map do you want to see?",
                    ["🔥 Idle Density (where people stopped most)",
                     "📌 Same-Place Hotspots (long stay at one spot)",
                     "🛣 Employee Route (one person's full-day route)"], horizontal=True)
    if mode.startswith("🔥"):
        map_idle_density(ctx)
    elif mode.startswith("📌"):
        map_same_place_hotspots(ctx)
    else:
        map_employee_route(ctx)

    st.markdown("---")
    with st.expander("🗺 Classic flat map (GPS route with address popups)"):
        map_classic(ctx)


# ===============================================================
# TAB 6 : RANKINGS
# ===============================================================
def render_rankings(ctx):
    st.subheader("🏆 Top Performer Ranking")
    rk = ctx.df.groupby("Employee Name")[COL_PRODUCTIVITY].mean().mul(100).round(1)
    top10 = rk.sort_values(ascending=False).head(10)
    bot10 = rk.sort_values(ascending=True).head(10)

    def ranking_block(data, color, title, note):
        nm = [short(n, 22) for n in data.index]
        show(hbar(nm, data.values, [color] * len(nm), title, xtitle="Productivity %",
                  full=dict(zip(nm, data.index))))
        guide(note)
        tbl = data.reset_index()
        tbl.columns = ["Employee", "Productivity %"]
        tbl.insert(0, "Rank", range(1, len(tbl) + 1))
        st.dataframe(tbl, width="stretch", hide_index=True)

    t_col, b_col = st.columns(2)
    with t_col:
        ranking_block(top10, "#10B981", "🥇 Top 10 Performers",
                      "A longer green bar = higher Productivity. Rank #1 is at the top.")
    with b_col:
        ranking_block(bot10, "#EF4444", "⚠ Bottom 10 — Needs Support",
                      "Each red bar is one of the 10 lowest performers. A shorter bar = lower Productivity.")
    insight(f"The gap between the best and the lowest performer is <b>{round(top10.max() - bot10.min(), 1)} points</b> — "
            "pairing top performers with those behind can help improve results.")


# ===============================================================
# TAB 7 : REPORTS
# ===============================================================
def render_reports(ctx):
    df, f = ctx.df, ctx.f

    st.subheader("🗓 Monthly Performance Table")
    monthly = df.copy()
    monthly["MonthKey"] = monthly["Date"].dt.to_period("M")
    monthly_tbl = (monthly.groupby(["MonthKey", "Employee Name", "Deg", "Group"])
                   .agg(Working_Hr=("Total Working Hr", "sum"), Travel_Hr=("Total Travel Hr", "sum"),
                        Idle_Hr=("Total Idle Hr", "sum"), TP_Visit=("Total Tp Visit", "sum"),
                        Avg_Productivity=(COL_PRODUCTIVITY, "mean")).reset_index())
    monthly_tbl["Avg_Productivity"] = (monthly_tbl["Avg_Productivity"] * 100).round(1)
    for c in ("Working_Hr", "Travel_Hr", "Idle_Hr", "TP_Visit"):
        monthly_tbl[c] = monthly_tbl[c].round(1)
    monthly_tbl = monthly_tbl.sort_values(["MonthKey", "Avg_Productivity"], ascending=[True, False])
    monthly_tbl.insert(0, "Month", monthly_tbl.pop("MonthKey").dt.strftime("%b-%Y"))
    monthly_tbl = monthly_tbl.rename(columns={
        "Working_Hr": "Working Hr", "Travel_Hr": "Travel Hr", "Idle_Hr": "Idle Hr",
        "TP_Visit": "TP Visit", "Avg_Productivity": "Avg Productivity %"})
    st.dataframe(monthly_tbl, width="stretch", hide_index=True)

    month_summary = (monthly.groupby("MonthKey")[COL_PRODUCTIVITY].mean().mul(100).round(1))
    if len(month_summary) > 1:
        labels = [p.strftime("%b-%Y") for p in month_summary.index]
        show(vbar(labels, month_summary.values, [prod_color(v) for v in month_summary.values],
                  "Average Productivity % by Month", ytitle="Productivity %", height=360))

    st.markdown("---")
    st.subheader("🤖 AI Performance Summary Report")
    prod_by_emp = df.groupby("Employee Name")[COL_PRODUCTIVITY].mean()
    prod_by_group = df.groupby("Group")[COL_PRODUCTIVITY].mean()
    best_emp, worst_emp = str(prod_by_emp.idxmax()), str(prod_by_emp.idxmin())
    best_prod, worst_prod = round(prod_by_emp.max() * 100, 1), round(prod_by_emp.min() * 100, 1)
    best_group, worst_group = str(prod_by_group.idxmax()), str(prod_by_group.idxmin())
    worst_deg = str(df.groupby("Deg")[COL_PRODUCTIVITY].mean().idxmin())
    idle_by_emp = df.groupby("Employee Name")["Total Idle Hr"].sum()
    top_idle_emp, top_idle_hr = str(idle_by_emp.idxmax()), round(idle_by_emp.max(), 1)

    if not ctx.sp.empty:
        top_row = ctx.sp.sort_values(COL_IDLE, ascending=False).iloc[0]
        idle_loc_html = (f"<li>Longest same-place stay: <b style='color:#EF4444;'>{top_row['Employee Name']}</b> — "
                         f"{fmt_hm(top_row[COL_IDLE])} at <i>{str(top_row[COL_ADDRESS])[:80]}...</i></li>")
    else:
        idle_loc_html = "<li>No idle location data available for the current filters</li>"

    avg = ctx.avg_prod
    if avg >= TARGET_PCT:
        prod_c, prod_status = "10B981", "Team is performing at a good level."
    elif avg >= LOW_PCT:
        prod_c, prod_status = "F59E0B", "Performance is average — improvement needed."
    else:
        prod_c, prod_status = "EF4444", "Performance is low — urgent action required."

    st.markdown(f"""
<div style="background:linear-gradient(135deg,#0F172A,#1E293B);border-radius:16px;padding:24px;color:white;
box-shadow:0 8px 24px rgba(0,0,0,.3);margin-bottom:16px;">
<h3 style="color:#10B981;margin-top:0;border:none;background:none;">📋 Monthly Field Force AI Report — {ctx.month_label}</h3>
<p style="color:#94A3B8;font-size:13px;">Period: {f.start} to {f.end}</p>
<hr style="border-color:#334155;margin:16px 0;">
<h4 style="color:#818CF8;">📊 Overall Performance</h4>
<ul style="color:#CBD5E1;line-height:2.2;">
<li>Active Employees: <b style="color:#10B981;">{ctx.total_emp}</b></li>
<li>Avg Productivity: <b style="color:#{prod_c};">{avg}%</b> — {prod_status}</li>
<li>Best Area: <b style="color:#10B981;">{best_group}</b> | Needs Attention: <b style="color:#F59E0B;">{worst_group}</b></li>
<li>Designation needing improvement: <b style="color:#F59E0B;">{worst_deg}</b></li>
</ul>
<h4 style="color:#22D3EE;">🏆 Employee Highlights</h4>
<ul style="color:#CBD5E1;line-height:2.2;">
<li>Best: <b style="color:#10B981;">{best_emp}</b> ({best_prod}%)</li>
<li>Lowest: <b style="color:#EF4444;">{worst_emp}</b> ({worst_prod}%)</li>
<li>Highest Idle: <b style="color:#F59E0B;">{top_idle_emp}</b> ({top_idle_hr} hrs total)</li>
</ul>
<h4 style="color:#F472B6;">📍 Idle Location Alert</h4>
<ul style="color:#CBD5E1;line-height:2.2;">{idle_loc_html}</ul>
<h4 style="color:#FBBF24;">💡 AI Improvement Suggestions</h4>
<ul style="color:#CBD5E1;line-height:2.2;">
<li>Monitor employees with repeated same-location idle patterns</li>
<li>{best_group} area outperforms {worst_group} — share best practices</li>
<li>Set location-based daily check-in targets to reduce idle time</li>
<li>Use GPS route data to optimize travel routes per employee</li>
<li>Weekly performance review recommended for bottom performers</li>
<li>Set minimum TP visit targets per day to improve productivity</li>
</ul></div>
""", unsafe_allow_html=True)

    if avg >= TARGET_PCT:
        st.success(f"Good Performance ({avg}%) — target of {TARGET_PCT}% reached. Sustain the momentum!")
    elif avg >= LOW_PCT:
        st.warning(f"Average Performance ({avg}%) — idle reduction & route optimization needed.")
    else:
        st.error(f"Low Performance ({avg}%) — immediate field activity review required.")

    st.markdown("---")
    st.subheader("📋 Detailed Data Table")
    display_df = df.copy()
    display_df["Date"] = display_df["Date"].dt.date
    display_df[COL_PRODUCTIVITY] = (display_df[COL_PRODUCTIVITY] * 100).round(1)
    st.dataframe(display_df.rename(columns={COL_PRODUCTIVITY: "Productivity %"}), width="stretch")
    st.download_button("⬇ Download Filtered Report (CSV)", df.to_csv(index=False).encode("utf-8"),
                       "field_force_report.csv", "text/csv")


# ===============================================================
# HERO + FOOTER
# ===============================================================
def render_hero(month_label, period_label):
    st.markdown(
        '<div class="hero"><div>'
        f'<div class="hero-title">📍 {APP_TITLE}</div>'
        f'<div class="hero-sub">ACI Premio Plastics • Digital Command Center • {month_label} ({period_label})</div>'
        '<div class="hero-steps"><span>① Pick filters on the left</span>'
        '<span>② Click a tab above</span>'
        '<span>③ Click any bar / line to see the numbers 🖱</span></div></div>'
        f'<img src="data:image/png;base64,{ACI_B64}" style="width:84px;height:84px;border-radius:50%;'
        'border:4px solid #10B981;box-shadow:0 8px 20px rgba(0,0,0,.4);"></div>',
        unsafe_allow_html=True)


def render_footer():
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div style="text-align:center;padding:20px 0 10px 0;border-top:1px solid #E2E8F0;">'
        f'<img src="data:image/jpeg;base64,{PREMIO_B64}" style="height:38px;margin-right:16px;vertical-align:middle;border-radius:4px;">'
        f'<img src="data:image/png;base64,{ACI_B64}" style="height:38px;vertical-align:middle;border-radius:50%;">'
        f'<p style="font-size:12px;margin-top:10px;color:#94A3B8;">© 2026 ACI Premio Plastics Ltd. | {APP_TITLE}</p></div>',
        unsafe_allow_html=True)


# ===============================================================
# MAIN
# ===============================================================
def main():
    st.set_page_config(
        page_title=APP_TITLE, page_icon="📍", layout="wide", initial_sidebar_state="expanded",
        menu_items={"Get help": None, "Report a bug": None, "About": None})
    st.markdown(f"<style>{APP_CSS}</style>", unsafe_allow_html=True)

    # ---- load data ----
    try:
        excel_path = find_excel()
        summary_df, image_df, sp_df, gps_df = load_data(str(excel_path), excel_path.stat().st_mtime)
    except Exception as e:
        st.error(f"Excel Loading Error: {e}")
        st.info(f"Make sure '{EXCEL_FILE}' is uploaded in the same folder as this script.")
        st.stop()
    if summary_df.empty:
        st.error("The 'Summarize_Data' sheet has no rows.")
        st.stop()

    # ---- sidebar filters ----
    f = render_sidebar(summary_df)
    if f.start > f.end:
        st.warning("Start Date is after End Date — please correct the dates in the sidebar.")
        st.stop()

    df = apply_filters(summary_df, f)
    if df.empty:
        st.warning("No records match the selected filters. Please widen the date / day / filter selection.")
        st.stop()

    ctx = Ctx(
        f=f, df=df,
        gps=filter_hours(apply_filters(gps_df, f), f),
        sp=filter_hours(apply_filters(sp_df, f), f),
        photos=Photos(image_df),
        month_label=f"{summary_df['Date'].max():%B %Y}",
        total_emp=df["Employee Name"].nunique(),
        total_working=round(df["Total Working Hr"].sum(), 1),
        total_travel=round(df["Total Travel Hr"].sum(), 1),
        total_idle=round(df["Total Idle Hr"].sum(), 1),
        avg_prod=round(df[COL_PRODUCTIVITY].mean() * 100, 1),
    )

    period_label = f"{summary_df['Date'].min():%d %b} – {summary_df['Date'].max():%d %b %Y}"
    render_hero(ctx.month_label, period_label)

    # ---- tabs ----
    tab_names = ["🏠 Overview", "🧑‍💼 Employee 360°", "📊 Charts", "📍 Same Place",
                 "🗺 Maps", "🏆 Rankings", "📋 Reports"]
    renderers = [render_overview, render_employee360, render_charts, render_same_place,
                 render_maps, render_rankings, render_reports]
    for tab, render in zip(st.tabs(tab_names), renderers):
        with tab:
            render(ctx)

    render_footer()


# ===============================================================
# LOGOS (Base64, safe for Streamlit Cloud). Kept at the bottom so
# the code above stays easy to read. Do not edit these two lines.
# ===============================================================
PREMIO_B64 = "/9j/4AAQSkZJRgABAQEAYABgAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAEAAAAAAAD/7QCcUGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAIAcAmcAFGFjV2hRNEhabUxPQzd6ajF5SDdrHAIoAGJGQk1EMDEwMDBhYzAwMzAwMDA0ZTE0MDAwMDdiMWQwMDAwNTEyMTAwMDBkYTIyMDAwMGE5MmIwMDAwYjQzNTAwMDA4MTQwMDAwMDgwNDMwMDAwNGU0NTAwMDBjODU1MDAwMP/iAhxJQ0NfUFJPRklMRQABAQAAAgxsY21zAhAAAG1udHJSR0IgWFlaIAfcAAEAGQADACkAOWFjc3BBUFBMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD21gABAAAAANMtbGNtcwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACmRlc2MAAAD8AAAAXmNwcnQAAAFcAAAAC3d0cHQAAAFoAAAAFGJrcHQAAAF8AAAAFHJYWVoAAAGQAAAAFGdYWVoAAAGkAAAAFGJYWVoAAAG4AAAAFHJUUkMAAAHMAAAAQGdUUkMAAAHMAAAAQGJUUkMAAAHMAAAAQGRlc2MAAAAAAAAAA2MyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHRleHQAAAAARkIAAFhZWiAAAAAAAAD21gABAAAAANMtWFlaIAAAAAAAAAMWAAADMwAAAqRYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9jdXJ2AAAAAAAAABoAAADLAckDYwWSCGsL9hA/FVEbNCHxKZAyGDuSRgVRd13ta3B6BYmxmnysab9908PpMP///9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgCUgOEAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMRAD8A/fyiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACigtt603zFB6igB2cUZzUburMBnPtUFxerbNjeq4GTk9KFd7DsyeSbY2PanRSeYPxrHj8baXcP5a6rprS7toRbpC2fTGc5rQtrhTJuLL6Dmq5ZLdFSpzj8SaLVFNjYeX1pQ6nvUkC0UZooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAozRUd2/lw7umD6ZoAkzRmvNfH37T3gn4XfFDw34L1zXbXTfE3jBsaTZOGJlbO0biOF3NkDOMngV6FZnBKtww/hznb+NAFiiiigAozQWAOO56CoLiXy5Pu8YyXY/KuO9AE+aM14n8R/8AgoD8HfhHqs1rrnxB8O29zGSrQwzG5dCOu4Rhgpz6msDwx/wVE+A3jG/ht7X4jaPbzSHK/bBLbBuccsVC/maAPouisjw/410vxXo8eoaXqmn6lYSDK3FtcpLEw/3lOK1x0oAKM0N9015j8RP2sfhz8J9futJ8T+MtD0TUrOAXM1rd3HlyCM8hgCBuzjtQB6dmisfwl4p0/wAa+H7HVtKu4b7T9QiWe2uYvmSaM8gr7Ed62KACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooJxTRIp70AOoqN7hYx8xxR9qjaPcrKV9c8UAE+0rzVO+njEO52+XJDY+nrXl/7Uf7Y/wAPv2VvB76l4y8RWem71LQWgxJdXWOuyPqfTPQZFfkv+2h/wWy+IH7Q8V7ofgkTeCPC83yCWN92o3cfTJb/AJZg9cCu/C4GpWfkfccK+Hub5/UisNTtC+spKy+T6s/RT9sv/gq38Mf2ObOaymvP+Em8WrGdmhaWd8qk9POfkRr/AL3XBx0r8mP2rf8Agpx8XP2t9TnXVNck8N6ArEwaFolxJawouRjfKo3Stj3x19a8Adi9xNI7SSSTOXkkkYvJKx6szEkkmjpX0+EymnQtJ6s/q/hDwayTJUquITrVf5paq/p5FxNc1BW8xdT1KKXjdN9tmVmPX5SrllPuRX1B+x9/wWB+KP7K+o21jqtxN8QPCKnbNYancu19bpwP3Fww44HCkEHnJFfKeP50V2YjDUqsbNH2OdcH5TmlB4bHUYu+zgrNfM/oR/ZG/wCCifw2/bD0VX8J6wsOpxYN5omoAQ6ha56/Ln5gPVcrjoa97juldNyOrr/eBGDX8velazfeHdXt9Q0u8vNN1KzPmW93aztBNAwOeHU5A9hwe9foD+xd/wAF2vEnw+ktdF+K1ude0XKw/wBtWsYW9tvTzI+kgH94c8dCa+dx2TSguamfzHxp4G43A3r5M3OH8u7Xl5n7HQvujHPbmng5rz74HftI+Dv2gPBtrrXhHXtP1qxugMNbyZaNu6up5Vh3BArthfAlRuwWGQK8GdOUNJqx+E4jD1cPUdKvFxktGmrP8S3RVcTse9Tb19anfYx23HUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAIzBBzxVfU7qK3s5JJJY40hUyOXYKoUckk+gAqedtkRYdRzgd/b8a+S/8AgrH+0xH+zl+ybqVnZXDLr3jBm0PTgrHIjfmaUY5yqbh9TQB+YH7cv7UF5+0N+17rfjjS76WG10W8jt/Dzq+Gt4bdyIJM/wB5pcN9Hr9rP2TvjtZ/tI/AHwv40tWCHWbINdxAf6i6Q7J4zzxtlVwB/dwa/O21/wCCbk0H/BIFNaTT93jm1n/4TUkoPMe1CFHtyfa1BfH94DvW3/wQR/aUVZvEXwr1G5PmMBr+k7mPykgLcIB/3y+Pc1TA/UAHNFQ2rbi/X5Tg/wA6mqQKWpzSRyYVVwQMEnAPJzk4+XA5HqeK/Ir/AIKJ/wDBRPxV+1P8XJPhp8M9SvP+EXa8GkQpp7NFceJ7ssVO6VeVgDcApycGv0q/bj8e3Hwv/ZK+I3iC2maG40nw7dzQsv3hJswhH0OPzr8pP+CLHw8tPHX7b+jzX0bTf8I9pVxqCo4DASkhQ59wzE/jQB9Y/ss/8ELvAvhnwhY3nxQN74u8S3SLPc2cV08Om2+esalVEjEcAkyHcQTxnFdJ8bv+CHHwh8b+GLlfB2n3vgXXFRjbXNndzT2rv286OUszA9PlPHWvtq1DCBd33up+tSUAfEP/AASV/YW8bfsg6f8AEJvGtpDb6pqF/BZaWbO8E9nPZQqXMirklBJKz8MA4ULnBr7chURwqq52qABnrinUUAFfjL/wXe06O1/bLMxCs03hyAjd64kFfs1X41f8F8W/4zAt/wDsW4D/AOPSVpR+ID9Rf2PtObSv2Vfhnbsm0x+GNO4UdMwRmvUK4D9mFT/wzb8Ox3XwzpwPt/o0Vd/Wf2mAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUVEl2sgyN2PX1oFcc0gK/h6VCJ1z1zt6+1QXWsWtrHJ500cMcaF5GkYIqKBySSemO9fCv7cX/BbfwN8B5Lrw54FaDxz4qhDpKbWXFhYuOgkkH3iD/Cv51tRw8qsuWKPXybIMwzXELDYCm5N9kfafxG+JegfDHwxdax4h1rTtD02zXdLd3cyxxpzjqep5AwMkk471+ZH7bP/AAX5lupb7w38F7BZVXMDeK9TgKorDgtbQcNu64Z1x0ODXwN+0b+1X47/AGrfEH2/xtrU2oxK5eDT42MdjaegWEHbkf3jya882qqqo/h74/lX02FyKNP3qjuf1FwR4F4LCv63nn76a2p6qK9X38jW8d+O9b+KPiqfWvEmsX3iDV7s77i91FzNJI3qOwHsoUe1ZbSmRFyzMx5bcozntgjtTc80v417caairU1ZH9BYXBU6NONOnFRSVkkrJLsIFxTsUHrTS2DWsYux0yi73HCikzzRurKW5FmLTQP3mf4cYx6+9Lu5oLUtVq2Gx0vwc+Nni/8AZ58YR+IPBOvah4f1eN1ZntpdsNyoPKzKchwe67R/vV+o/wCxR/wXp8P+PPsfh74s2cHhbW2xEmrWwP8AZd4x/vLy0LHqc/KCeGIr8khTbiHzom4Xc3XKhlOOnBrjr4GlX+JH5/xf4a5PxBDmrQUKn8y3P6gPC/i3S/FukwX2mX1rfWVwm+KeCQSRSg/3WHB/A1qCdf736V/On+yt+3v8SP2QNXt28M61Jd6Crf6Rod8xks5h/sDOYT/uYr9bP2K/+Cvvw4/apW10e+uP+ET8WPgNpmpyBRMc8GGTOH7eh9q+axmT1KWsNT+UOMvCjN8jftYr2lL+ZK6PsZHDjinVUtNQidRhj8wyDjg/iOKlW8Vj/F78dK8uzPy4mopolBoEmX20gHUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFI77ELHtzQBFdzbYW2gs68qv94joPx6V+Rn7at7rH/BS7/gpXD8NfDt8lvo/hHzNHguTGJ4Y2T57y7KgjPzBYxj/nmeuRX6Ef8ABQP9py3/AGU/2XPE3igXEMGsSRDT9GWRwjSXs/yRkZPOzdvOOyV8gf8ABvx8Ed+jeMviVelbmbUXTRrG4ZgzSkfPdOG65aQA/iaAKv8Aw4w+IdtaLGPj0yx7Xj8kaXOsTggIEI8/BTBIxjndXyZfaJ4g/wCCZ/7bFlb3Fwbi+8CahBKtxGvlx6pZOMM23k4dWcdTgrX7w3VnLNBtRlj9M/Nt44I+hwfwr83P+C+n7P8AGNG8IfEy2jUNaP8A8I/qYOMNC5LROzexBH/AqAP0W8BeJ7Txr4V0/WNPZZbHVbWK7t5F5DI43DH55/GtmvhH/ghX+1VD8UP2ftQ8A6hfLLrPgOVfssbv+9fT5izRAA8kRkMmQOm31r7simWZcr0/nQB5j+2f8O5vix+yp8RfDlujvcat4eu4YVXkvIInZVHuWAH41+TP/BGb4mWvw/8A23NCj1Bktf8AhJdOuNK/enZi4wCE57l0Ix2Ir9sLyOSQfu9vQnn+92/x/CvyV/4Kaf8ABNLxT8EvirffE/4YafeXGh3V1/ac0GnJ5l1oN3uDGZV6tGzAsQOm40AfrZHKqQL1HHAI5/KhrhU67snoMcmvzl/Zg/4LveF5/Dltp/xS0nVtJ12zTyZL7T4DPFdODgs0JIdDjrwecmuu+OP/AAXa+FPg/StvhS31jxRq0g/0aOSE2trvxwS7HJ57AUAfdguVP97/AL5NSZr5B/4JXfHn4rftJeD/ABt4s+JNvJZ2eqaxAPDdsLMQR21t5O5zGTy6ZIwx5IFfXduc28f+6KAHnpX40/8ABfGNm/bAhbHynw1AB/31JX7LN901+Mf/AAXZ1L7X+2Q0JeMfZfDluu0nlcmQgt6fU1pCyA/WD9mP5P2c/h/n+Lw3pxH42sVd5Xl/7HWtR61+yv8ADSZW3+Z4X05gw5BBgTHPTI2nNeoVmAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQA15Npr5o/bG/wCCmnw1/Yxs2h8Q30mpeIZBm10PTYzNet/voP8AVj3bFfSMkm9uhrlNU+C3hXW9TmvLzwz4dubq5/108unRPLJ9WK5P4mtqMop3kdmBlho14yxkHKC3Sdm/mfhN+2b/AMFRviV+2Ld3lhdSt4X8GzMQNC0uSRBcp2NzP9+RvULsU+hr5vimW3iaOGNo4mIKosZUD19yfcmv6Zo/gL4JVgP+ER8M8+umQ/8AxNSt+z94Ib/mUfDP4aZB/wDE17VPOKdNWjA/esj8asvyiisPl+X8qXXm1++x/Mp5px/q5P8Avg0CRmP+rk/74Nf00P8As+eCR/zKXhn8dNh/+JpjfALwOqAnwj4cXJxj+zIc5/75rePEFt0e9H6R1NbYJ/8Agf8AwD+Zp5GiXc0cu3vhCT/n6ULL5jqqxzbm5AMZU4+h5/rX9AX7RPxa+CX7P08el3XhrQdc8UXB22uhaXpkM97M/YYVCV/LPtXnmn/s+fFD9o6RrhvC/gv4Q+F7rBijXTIbjVGQ/wB/K/K2PXBz2oefQfxRYf8AExlNyu8G/wDwK5+ILv5e7dHIoU4JKEfz/rUS3schwrK3sp3fyr9/PAf/AASa+E/heX7Rq+lSeKr1zvkm1AriRvXYoAA9q9W8OfsifDbw5arFZ+BfClssfQLp0Z/mDSWfU/5GaVPpF0Le7hH96P5sfNznakzfSJjj8cYojukY7S2xh/eG2v6Yn/Zu8DyLtbwl4ZZT1B02Hn/x2uT8c/sE/CHx7Z/Z9U+HvhW4jblttgkbfgygEfnVrP49Ik4f6RVFytVwjt6o/nKLKOpVc+tPZCF549iMGv2o+MH/AAQJ+DfxB89/Dra14NuCMr9juTNCTjjKSZ4+hr4Z/am/4Il/Fn9n9bi/0E2vj7R4QXY2K+TeRoB3hYnd/wABP4V3UM3o1NGfe8P+MmRZpNUG/Y1HtzXt9+34nx4pyKWlmt5rC+ms7q3mtL23cxzW0yFJomHBDKwBGPcAemabnj19/Wu9Wa5k9D9Zo1I1EnBqV+q1X3iU5tsgVWXMYbd1O5COQQQQdw7MCCDzTcUrciiK5viNKkVW92ulKPax9gfsZ/8ABaD4l/swXVrpPiiSb4h+DY0EQgvJdmqWCjp5M4wJFA/gkAJx98da/WL9lL9u/wCHv7YnhwX3hHXIbq4UAzaZOfLvrRscq0bcn6jIr+dxU+X5vmq94Z8V6t4E8SQa1oep32j6paMHhu7Odo5kI6ZPcex4xXlYzKadRNw0PxPjLwXy7NVPE5fajWf/AIC/kf0+Ws25lzuRscg87T6EjjNXEP7zrX5I/sG/8F277SdY0/wx8ZEjubO6lSCDxLApjaMnjdOg4IzgFhjGcnpX6uaBr0GtW8NxbMtxbzxrLFNG26OVGAYMrdCCD1FfM4rBVMP8Wx/J/FHCOZZDilhsxjurxktpJen66mtRTVfcelOrkPmwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApsyeZEy/Lz/eGRTqKAOW+IXwc8NfFZtPbxJ4f0XXjpNx9qsl1C3FxHbygFRIFYY3YJHSrvg34f6V8PrBbPRNN0/R7FXZ/s1lAkMTM3JbaoGGJrcooAKwfHXw50b4l6IdM8QaPpOuaa0iym11C2S5hLLypKOCDg1vUUAcb4c+APg3wX4mbWtF8J+G9H1cwG2+2WWnxW0xiO393ujUHbxwDnHbqa6+KPy89Pb2H1p9FABUNzafaHVuVaPlSGxn61NRQB5X8TP2KfhX8Y743nib4e+D9VviSTdvpyR3LE+sqgN7dTmqfw5/YH+Dvwl1OO98O/DXwbpd5G29Z49OVpVPqHbJB969gooAptpf7rYreWNhQFTjZznI4796uDpRRQANytcL4o/Zt8C+NdWudT1jwX4R1jVruEQS3V9pUM8sqDOFZyuSoyeK7qigDM8MeFLPwdodnpmm2trYafp6CG1treIRwwRgYCqo6Ac1p0UUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFNd9pHFADqKhe5ZW+7x7tQbtUXLU7MCaioBegngceuaVbvcOn61Mmoq7Bak1FRfal981IjblB9aYC0UUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUFsUA5FABRRRQAm0eg/KjYPQflS0UAJtHoPypWOBQTgVWlum8pui7gSMnGAOpNACX+ox2VrNNJJHFFAu+R5G2qgHJJPsK+SfH/7TPjD9qvxzf8AgX4IhLHTLdymt+M7hD5UQPVLc924wCP061D8Z/iHr37dvxbvPhj4Hvp9L8EaFIB4n16DOLrHJt4mHXkAHHPWvpj4Q/CDQfgr4K0/w74bsY9O0vTYwkUSjDSerserMTnrQBwv7Nv7Eng39nW1a4tIZtb8SXWZLzXtTPn31xIepVzzGPZT+NexRWKwD5Sfx5//AFn3OTT44fLP3ievWn0AMSIoPvMfc04DHfNLQTigAoK57Um4UoOaAK8zBpGUjj1HFU7m089vvPt/u4B3fiQSPwIrQeAO27NH2cY60RVndMlX2f4aWPmH9uL/AIJkfDn9tTw+1xqOnx6F4wjVls/EWnIIruMjosp6Sx/7Lg47FTzX4w/tffsWeNv2KfiAdF8VWXnafcHOnaxbgtZ6hH0BDEnY/TKMSfc9a/o6+xLv3bm9OOK4P9oD9m7wr+0b8M9Q8L+KtOj1LSb9SGRwN0THo6H+FgeQRXqYHM6lKVpbH6nwB4n5jw9VjQnNzw99YvV/I/mnRtxx0pwORXtf7eX7EuvfsL/GL+xNQaTUNB1bzZ9A1MRnbfRp1gY9BcJnJHRh0rxLOBX2FOSnHmiz+3MnzfC5nhKeMwcuaE9vL17fMdSNyOaWhuRVcp6d7aoaeEblsSDayk/I4/ukeh6H2Nfst/wQH/aYvvit+ztqfg/WLqS8v/AtwILN3bLCyk+aNCe4Q5UH0FfjQ43Iy/3lK59M1+mf/Bt7Fv8AF/xMOcL5FpgD+EZbivPzqnB4e/VH4z445fh6vDs6nLrTacZdfe6H6yQtuapKjRNj49s1JXxJ/EoUUVFNMySBRQBKxwKYJGI/hr5K/wCCmf8AwUvk/wCCf134PWPwq3iaTxUblQi3Qh8nyQpzz67q+Xk/4OT5tg2/C3g8j/ibIf6V3Uctr1Yc8FofZ5P4f57meFWMwVHmg20nddD9V95/2aN5/wBmvypP/Byfcf8ARLf/ACqp/hR/xEoXH/RLf/Kqn+Faf2Rie34np/8AEJ+KOmFf3r/M/Vbef9mjef8AZr8qf+IlC4/6Jb/5VU/wqS2/4OT3Sb/SfhbIIvWPVUz/ACo/snE9g/4hPxR/0Cv71/mfqkpYjkAUAtnoK+Dfgf8A8F+PhP8AFHUbez8R2ut+A7i6fZHcX8PnWRPQDzU/wFfaXg/x1YeOdEt9U0nULHUtPuk3w3FrOJYpRz91hwenTqK5a2FqUtZo+RzXh/Mstly46jKn5taffsdFRVeG7Z9u7b8xqxXOeLGSYUUjHAqreXLYwoY8/wAIyf8A9VBRaY4pN5/2a/N/9pP/AIL3Sfs+ftBeL/Asfw9XVT4U1A2DXY1RU88hFbcFI4+9XFD/AIOT7g/80t/8qyf4V6EcrxDV0j7zC+GnEWIowxFHDuUZpNNNap/M/VXef9mjef8AZr8qT/wcoXA/5pb/AOVVP8KP+IlC4/6Jb/5VU/wqv7JxP8p1f8Qm4p/6BX96/wAz9Vt5/wBmjLH0r8qv+IlGbH/JLG/DVU/wrS8Pf8HJGnT3Krqnwv1VEduWs9SikZB67WAz+BoeU4lK7j+JnU8K+KILmeEl96/zP1D3469acDmvl39l7/grJ8Hf2qtQh0rSdfbQ/EkgBGka0n2O6kz/AM8wx2yfVT+FfTK3u2NeF6c4bpiuCpTlTlyyR8XmWWYvL6ro42nKEl0asWaKjWYsu7jGM1JUHCFFN3c153+1H8c2/Z0/Z/8AFfjj+z/7U/4RnT2vvsgk8sz7f4Q3b61UYuTSXU1w9GdaoqVNXk2kl5vY9GJwKbvP9386/KiP/g5MmaNW/wCFWtkjOP7WX29velH/AAcn3H/RLT/4Nk/wru/svEdj7+n4U8UTSlHCvXXdf5n6rbz/ALNG8/7NflT/AMRKFx/0S3/yqp/hR/xEo3H/AES0/wDg1T/Cq/snE/ylf8Qn4p/6BX96/wAz9Vt5/wBmgP64r8qf+IlC4/6Jb/5VU/woP/BydMx+b4Vs3PUasn+FJ5ViV9ka8JeKm7LCP71/mfqtvX1p1fIf/BNb/gpo3/BQW/8AF0beFP8AhGP+EXa1UE3gm8/zS+fpjaMV9cwP5kQb19a4q1KVKbhPc+IzTK8Vl2JlhMZHlnHdD6QuAetLXnP7S/x4tf2a/gv4l8cahZXN9YeGrY3c0EH+skRWUNt7ZAbI+lRFNuyOKnTlUkoQV22kvV6HowOaK4X4RfHTRvjj8O9L8VeF9S0/VtB1a3M9tdQvuVgMZyByCpO1geQQa7SO6LKudoYjpnOKJJxlyvcVSnOE5U5xacXZprsTUVD5zY7damHSkSndXCiiigAoqO4do48qMtmmvOynouMUPRXB6Imoqje6l5EDNmNQvO522qPqa8b+OH7cfhH4KfGTwD4Aurj7X4q+IWqR6bY2MDAyQK6MxuJR/DEu36nPFXTpym7I2oYerWk404ttJt+iPcqKp6fctIrcscgMMjHBz/gf0q0jbhUGIM+1vagvz/DVW+uHSbauxVwMszYr84v2gv8Agv6/wJ+OnizwYnw7GoJ4Z1N9PFy2pCIz7f4tuDiujD4WpXly01qe5kPDuPzirKjl8OaUVd6pafM/ScOO5pwOa/NX9n//AIL9v8c/jd4X8IN8OGsP+Em1FNPW6GpLIsJPfAFfpBDds0a/d6ZyDkfnRXwtSi0piz3h3H5PXjh8wp8kpK69L23LVFIh3IKU9K5zxAoqA3Rzj5c0y5vmt4SzbFC5JLHAA+tAdbFo9KaWYDoKpy6sEhfzCsewbnJOPL+v4c5r5p/aA/4K5/BP9na6ks9R8VR65qkbFTZaIv26QEcYYqcKR7nsa0p0Zz+FHoZdlWMx9T2WCpSqS7RTZ9Q+Yc9qVWye1fmT4m/4OSfDtvctHo/w1164XnZJd3sUBYf7mM/rWbov/ByfavOq6h8Lb7Znn7Pqke4D2yOT7V2LLMQ1ex9hT8LeKJw51hJfhf7rn6ksT2qKYZdS3GO3rXxN8Gv+C8vwS+KFzBa6xca14JvZTjGrWhFvn/rsvyjHqa+tPBPxI0f4neHbfWNB1bTdY0u4GY7u0uFmgfPcOvHY1yzw9SHxI+Xzbh/NMsdsdQlTf95NCfFDxVfeCvBGsatpujzeINS06ylurbToZBHJfOqkrErEYBYgDJ9a+MfFn/BXX4jeCPFfh7Q9Y/Zr8XabrHiqc2+k2txrNoHvXVQzKuMjIBHXmvukQrfMpb5SuAMj0IP64r4w/wCCidqkH7cH7KSrhv8AipbxckAn/UA5+vAFdGF9nKVpq+h0cM4fCVMR7CrSU3aTu21qotpaehpeJP8Agpp4y+G37PHi7x940+Bfibwqvhi4s7e30+51OBpNTM8wibYyggbMg89favev2V/2j9P/AGpv2fvDPjzTbVrGHXrbzJbNpN72UysY5YCcDLJIrLnAzjOBXiP/AAWlKWH7BetNiMxya3pJcMOAPt0J5+uP1ryP4W/Gj/hgiy/aW8B38hhh8IlvGnhZGOBNb6hDzHGB/Cl2SMD1NbKhTq0vdjZ3/A9j+x8HjsrWKw1PlqSqNJJt3ilBPR/3pJ+l+h7V8AP+CmOl/Hz9tzxf8JdL0C8jj8KxSyf2w90jQ3bxuqOiJjoCTzu/h6V9ZRnKL3461+V//BNz4Pz/AAR/4KD+GdD1MN/bN98M49X1Jj9+W6uLgyyMT6gvj6Cv1QiO6JcdMCscdGmpJU+x5nFGBw+FxFOGH2cE/XzXkOooorhPmgooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAKl1IwuMbV28Atk/LnvjGMfjXwF+1d/wXrsv2K/jJeeCPHXwZ8dafewl3s7tLm2a21aHgrNbsWG9SpG4cFWyMHGT+gklssj7iPmxjOe1eI/t0/sNeA/29PhHceDfGVgPOiH2jStThQLdaRcYIWWJ/Y9U6EUAfEY/4Onfh6R/ySjx368XdoR/6FR/xFO/D3/ok/jz/AMCrT/4qvy1/bX/Yi8ZfsDfGW48H+NLVpYZT5mj6vApFrrkOcbkPTzc53R/eBBPQivHVUhcFgxHBIPB+lG4H7Wf8RTvw9/6JP48/8CrT/wCKo/4infh7/wBEn8ef+BVp/wDFV+KmKMVXKFz9q3/4Omvh7KjK3wn8fYYYOLq0/wDiqpv/AMF6G/bm16z+Evwr+G/i3TfFPjiddPjv7q7h2WMOA00h8vOFEefmzgN1z0r8YD8vPX2zj9a/aX/g2l/Ymh8G/CTWPjVrVkG1DxhNJpfhtnXmHTYXKyzL/d82UOOOdsXo1JqwXP0L/Zt+Auk/s7fDDS/C+krGUtfnurkJte+uCMySsevJPHtmvRUt1jfcuR699350iWaRurfMSvqalpAFQ3ly1sq7VVix6FtucDP0/PFLePJHDujxuU5ORkY7+/vxXw9/wVh/4LHaH+wJoX/CLaHBZ+Ivipqts01rp0kn+i6TESQlzdkcjdwVj6kkduoB9ZfFn4/+EvgR4Rk17xl4i0bwzo0Klmu9RuVgRunABOST22g9PpXw78aP+Dln4DeANSms/DuneMPHVxDwsun2SwWrn0EkrLn6gGvxP+P37Rnjr9qrx/P4o+IPiXUvEuqXBygun/cWw7CKEfJGg6AYzxXG4wm3+HJO3+HnrxTsB+x6/wDB1F4ZOoMv/Cl/FX2cD/oNWvmfkM/zr1f4M/8AByx8A/HtzFbeI7Pxd4EkkODPqNj59pGT/ekhLFR7lce9fguE2rj+HOcdqd1+noeR+R4pAf1hfCz9oDwl8cfCceueDvEei+JdKkYKLrTrlZ4wT/Cdp+U9B82K7C2uTOe2Of0r+UH9nf8AaS8dfsn/ABCh8TfDvxNqPhvVYSA62zFre4XOTHPCfkkjP0yK/ef/AIJSf8Ff9B/4KC+G5NA1aOx8NfFDRoDJd6Wr5ttUi5AntSeWGR86DlTntigD7doIzTYWLRKScnHJFOoA8M/ba/ZJ0H9sj4Ear4N1qECac/atLvtvzadew8wSr6Ec59RnNfz5fEv4d6x8IPiLrXhTXrb7NrPh68azu0xgSEdHXP8AC33lPORxX9ORgVW+76Hr3FflP/wcA/sp2+n6jofxY0u12q5XRtcMa/eB5t5G+h+TPaveyXHOnV9kz9z8EuM6mX5j/Y9ef7qttfo/+CfmaKD0pu4Mfl/GjNfUWR/ZcJXBupr9Nf8Ag26/5Gv4l/8AXvZ/zavzKbqa/TT/AINuv+Rr+JX/AF72f82rz83/AIHyPynxo/5Jev8A9ufmfrJ/y1/CnU3/AJa/hTq+JP4XCobkYcN36VNTJRQHqflR/wAHKMOLv4Q8t8z6mTyRnCRcV+X6oIxtVOAMjHRRX6ff8HKUhW9+EP8Avan/AOgRV+YKfvDtJO2QBGx6HGa+2yaVqGp/cngnNx4Rpzl73vS02673GmZQflBZexwef0pyupHRv1/wr9iv+CcX/BN34L/HL9h74b+LPFHgXS9U1/XNHS5vLuQsGmkLuMkA46AD6AV7kP8AgkL+zsF5+G2k/wDj/wDjXPUzyEJuKW2h4OP8fMrwuJqYd0J3hJxe26dna7PwD3L6N+v+FIs2W2ptLN0XeDj61+/3/Dob9nX/AKJtpP8A4/8A41DqP/BHr9nXUrGW3b4d6fCswwzQySI/4ENkVKz+m/iRzQ+kRlfNrh5/h+jR+AghZHLLu3yMTknaZD3+bJ/KvfP2D/8AgoL42/YV8eRtpc0+s+Dbtw2reHJJj5M+SP3lsCf3cqjuDtJ6g19A/wDBRj/gjLcfs66Jf+M/hn9s1DwvYsZ9T0mY+bcWIP8Ay1iIyXj9uor4EW4WWBZFYtGwLx7G+8p4wP7vNehCrRxcLJH6Ng80yHjXLJQVpwktU/ii+j8n8z+lz4D/ABz8PftC/DXQ/F3hW+i1PQtagFxBMo2tgjowPKsDkFTyMV3iTh2xivx5/wCDf79qa98K/GTVPhXfXDSaX4gSXVtP3/diuVI8wL6blycDjiv2EDbVX7p9wK+Px2F9hU5T+KOOOGZ5Bm1XAT1S1T7p7X/IfO+xPxqo65kD905H8qsSEuMVCvzSbT0IJrj7nya303sfzv8A/BShm/4b6+LCs0j415sbmJ2/uo+AOleJblAOR93kjOOO/Ne1/wDBSRi37ffxaY/9B5v/AEVHXI/speE7Hx9+1D8ONB1a3S80nWvEtjZXtu/3Z4pJgjKfwY/nX6DRqKNBSeySP9DuH8dTwfDNDFTi7QpKTs+ijd6HAhlkG5WUr2IyQf0pwXPv+f8AhX77xf8ABIT9nZnkZvhro67nzxvx/On/APDoT9nUf8030of99/415UuIIJ2itD8tf0hsrT/gT/8AJf8AM/AQjH94fQ/40oLFWGPMAG4h13BR7ZGPyNfv2P8AgkJ+zv2+HOk/i0g/9mrkfix/wRJ+Afj3RJrex8JyaDdbDGl3pty6Sx9wdrHawye/PalTz6lJ2kmaUfpCZTKajOjNJ+UT8LpBvEePkaF1lhdfvW7jkNGTna3uOa/Sz/gkv/wVv1az8Raf8MPipq02qWt8RbaBrtzLvmhkztS1uHJy2VAKyHnsxZstXyJ+3T+w14k/YW+JkOk6xIuqeH9TJk0rWIkKpcIv34mH8MqjHHcc14ZLG3lglsSKVKshxt2ncMH2bn61316VPE0uaK1PuM6yXKONMmVVWl7RPlnZc0bdHbtsf1JW825Bg5VQARjp759KuV8rf8Env2o7v9qf9kPQr7UZhLrmgg6LqfzbnZ4wFDn/AHl2n86+o4p2eLceua+JrU3Cbg9z+Fcyy6pgcXVwVX4qcuVkh+9Xz3/wU+fH7AnxV/7AEv8AKvoT+IV89f8ABT//AJME+K3/AGL8v/oNVQ/iR9Tq4b/5GlD/ABw/9KR/PRbS5tUJ9OcLnp/+qnq+9sBc7huG07lUfXv+AqKzQPZoG5zgH8Qc196f8EPP2UPh/wDtTeIfiFD488O2niCLSUtfsaz5/wBH3M4IXBHXaK+7rYj6tS52tD/QHibibDZFlKzGtCTULJpO+r23PhJvkHzD/P40nmKfX8/8M1+/z/8ABIP9nZD/AMk30r85P/iqT/h0N+zt/wBE30n/AMif/FV5P+sEex+Q/wDExGV/9A8//JT8At6j+9/n64prSqM4LDAz2Nfv/wD8Ohv2df8Aom+k/wDkT/4qlX/gkF+zrIDn4b6P6fMX/q1NcQRvsNfSIyq+uHn+C/FHxr/wbdndrPxa2sS2/TcHaCQP3p6V+sFmqiNioI3HdyeteYfs+fscfDj9lSbVm8A+G7Pw+dcMTXvkE/vjHu29T7n68V6nHGsQwvA9K+dx2IVeu6sep/NvGmfUs4zermFCLjGb0T3HV84/8FW4c/sA/FRl6toz7ufRlx+VfR1fOf8AwVY/d/8ABP34rEfxaLID/wB9LU4W/to+p5vDumaYZr/n5D/0pH46/wDBPH/goR4i/YW8eDa0+qeBdamWTV9IJ5Uqdpu4F6JIvQgfLIOqk81+7Xwg+OHh745fD/S/E3hnUbTVtH1SISRTRHnpypHZh0IOMEEV/NDG2xU2/KBjIHRuMc/5619Cf8E+/wDgob4l/Ya8dq0bXGqeC9YlH9q6Uz5brtE0HZZRjODgN3r6jMMv9p78LH9VeK3hOs1oyzXKVy1Vq4rZ9/mf0FJJ5hx+Iq4OK4j4P/GXw/8AHP4eaZ4q8LapaatourQpLb3ELbsg9Qw7N/snkV15uCON36V8jOMlKzP4/q05wqOE4tSWjXW6LFFQSO4VcN/9ekM5/vfhSsTra6JZ/ufjVO9vfJi3FdwzjJ6KOeT7cfrS3VwwtmY7QBySeij1/CvlP/gpJ/wUr0P9iLwA0No9vq/jvVY/+JTpG7DKTkCaf/nnGCQfm68Vrh6Mq0+SJ3ZblOKzDFLBYOLc52St09Q/4KSf8FLdA/Yp+Hiw2qw6t471eMnS9KDAbODiaYfwovXnqcCvyd/Y9+IuvfGT/gpV8MfEfinUpta1zWvFsM93cS5w5aOUhUGSUjXAAUHHB4rxz4m/EfXPi98Q9U8SeJtUuNX1/VZS97cStnLjqoHRVGcKBxgV6P8A8E8D5f7evwb24Uf8JTAcAdT5cvP1r6unl8aVBt7n9f5T4b4ThvhjFVp/vMQ6Urt9PT0P6LLcFD97cWILZ7npVtV2iooIF+9zyBU1fHPc/jKyWiKd+mZc54YbSMDB+tfzl/t7uU/be+Ki7mwviOYcOwzz7Gv6Nb0/6Rj1Wv5yP2+Rj9uL4re3iOb+dfRcP29q/Q/fvo+t/wBr4hX+wvzHfsEx5/bg+Fv3v+Rhh/jPGSR61/RqIw7Felfzl/sEjd+3B8L/AG8QQfzNf0aoOd394c1jnmlSL8i/pEc0s3oc0rr2b/8ASiyg2oKU9KRDlRSt0rwz+fvIo4XbuPVea8+/aN/aV8K/syfDHUPFXjDU4NN02yQlgTuknfGRHEv8bt0AHeup8f8AjbTfhr4L1XXNWulsdM0mzku7qdj/AKpFBJI9wAeO9fgH/wAFA/249b/bk+Mdxq0zPb+EdJkaPQNN3/ukjzt8917yt94Z5FepluDlind7H33h7wHV4lzBUpNqjD4pbadrnXft0f8ABVn4gftg6peaXpd5feDvh+zMsWl2U5S5vlB4e5lB3c8EKhUDpg18rwQrbqyqAqt1wi5Y56k8FifViac8ysMysq/xbuM57fiemOpzgc19kfsRf8EZ/H37T+k6f4l8UzN4D8I3oElsZ4/+JlqKH+7E3EanjlvmxzX1EvY4SOq0P6/9tw7wZlenLCKdkrLmfo+vzPjnzmhO3LKW6L82D+hH601ZDIflOQ3TaGwfyGfxr9yPhh/wQ7+APgOzhW/8PX2v3I5M2pXzOzH12qduK3vFP/BGD9njxFYeSPAtva4UhWtbmSFxnng561wyz2hGVkvwPhJfSGyRVHT9jOS7+6vy1PwakIKmPbHjHzJ94Ee4Oc/iDXon7Nv7VfxA/ZN8VR6r4C8QXmlqpXz9OJMmn3KDJIeEnYByeUCtzweK+6f2r/8Ag37l8P2Laj8JdbuJm+9/YurSKfMHpFN13f7/AB0r84vF/hXU/h94lu9D16xuND1zTJDDeafeqY7iFwSAxHcHnDD5SOld0K1DFQs9mfeZPxLkHFuFlTT9o2vgnbmXp3sfuf8A8E8P+CovhT9tzwy1pMv9g+NtNiDX2kSSbw/HMkDdZEP4leldf+0v+ytqnx/+P/wZ8ZWWsafptv8ADfVp9RubaWFpW1ISRhAsbAjbjnJIP6c/gV8PvHWufDPxvpfiDw3qE2l69pVwJrG7R9uyT0b/AGT0OeOa/er/AIJz/tw6X+2/8ALPXoo47HxJprLp2v2O75ra8AHK+scnUHp1r53McveHftILQ/mnxM8OanDOIWZZXd0Z3v8A3W9LPtfob37fv7L1/wDtg/s56j4DsNUtNFn1G+s7gXc8RlQJDOsrDAI5wuAa81/bc/4Jqp+1d8d/h14yt9etdFt/Dmy01y2a3aRtfskmSZYeGwMOmfmBGGbjnI+toYw0Xscj8KcdMimXJ3ZbjOen0ry41nHRH5Vgc6xWEcHh5W5b/wDkySl+CPn+3/ZCvLP/AIKDzfGVdYsI9JHhVdBi0lbU/aEZZN3mGTdt29goXPHWvoiMYjX6d6hXS4VUrt+Vuoz1qxUSlc5MRjK2I5HWd+WKivJLZBRRRUHOFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAVHNZx3DAyIrbTuGegPTP1qSigDx/9rn9kLwX+2Z8KL/wR420yO80+8+a2uE+W70y4xhLqCU8rIvfByRwcivwJ8b/8E0n8H+ONc0iPxnHcw6Tfz2kcz6aUaVY3KhioYAE4zjAr+lCVF85TgZz/AI1+K/xl+f4zeMCeSdau857/AL56qIXPkP8A4d7zf9DZD/4AN/8AFUf8O95h/wAzZCfb7A3/AMVX0qBkUH5RnHTmqIbufPHhz/gmTe+P/EWm+H7PxZGt5rt3DYQMliS2Z22g8sR8oyeRX9FnwR+F2n/BL4U+GvB+lW8NrpvhfTrbToEiXClUjCk+vLAknuSa/Mv/AIJs+CF8aftjeGWlUyQ6PFNqZGM4ZEATP0L/AK1+sQjVTnHbH4VMhxHVHdSeVDu3bQpBPv7VJUN6nm2zL/e4qSjyn9sn9ob/AIZv+AmteIoXQ6tCPsulRTJuS4u3UhAQOdo+8SCPu478/gH4y/Ym1z4jeNdW8SeJPiAda8Ra5dyXuo301mztc3LHLMDu+6M4UDAUDgCv00/4K3fFmTxJ8btJ8K28jfY/DdmLidFb5TcTc8j1CgD8a+TwuB+tUlcVz5ev/wDgnrqkcRa18TafNIf+e1u0Zb8ctn8a4fxf+yD468JwNL/Zseowoc77OYSMR7pgH8q+2dtAOCOW+Xpz0qhcx+b1/aSadftb3EM1vNH96Jx5bj8GqNW3ew7ev4+9ff8A8Qfg/wCHvilpz22saXHdFvuywqEuIj/eD/0r5X+O37LGrfCWJtSs/wDicaEp/eTLnzrQf9NB/F9anlKPKT82Pmb5QQCDgrn0Pb8K2Ph78Rte+D3jrR/F3hvUbrTPEnh27F3pt3bvteGRccEnPDDIIOQw6g1io6su5WDKeVI7j/PalK5/nUgf07/8E4/22rH9ur9lPw348t/Jt9RuV+w6zZREH7Bfx8Spj+4fvAnsa+gImLxKxGCwBI9K/Cz/AINmf2gpvBf7Vviz4ZXU+NH+IGkHU7SLP3L6xI3gD1kt3JOOoiFfuohyi/SgCCZRj/eryn9rz4E2/wC0H+zd4y8KzL82qaVPFASMlLhQWjYf8CAr10ordqrtbp5bcYyOtXSk4S5lubYOtUoYiOIpu0otNfJ3P5cBZyWjNHcQm3niYpLGTzG43qV/Bk/8eppGK9Y/bv8AAMfww/bP+KGgwBVtrPXZriGMLgBLhxMoHsocYrycciv0HDyvTUn1R/pNkeLeMwGGxD+1BSfzBupr9Nf+Dbr/AJGv4l/9e9n/ADavzKbqa/TT/g26/wCRr+JX/XvZ/wA2rhzf+B8j878aP+SXr/8Abn5n6yf8tfwp1N/5a/hTq+JP4XCmy9Pyp1Ml/qKAPyk/4OVD/p/wh+uqf+gRV+YqDbIv+8v9K/Tv/g5TH+nfCH66p/6BFX5hQNumT/eX+Yr7bJ4N4dH9xeCsXLg+il3l+Z+/n/BIcf8AGtX4Pn10BP8A0Y9fTaRgoOv518w/8EjJNn/BNr4Of3f7CQH/AL7evpRSyqBu/WvkcVZVpp9z+M+IKiebYtf9PZ/i2WzEMd/zquQwH+1n1pu9s/e/8epZZlQj3rllyvQ8fnivefQp+INMtr3R7iGaFbiO4RonVhnKsCrD8QSK/m3/AGofhtD8Jf2lPHHhm12ix0jXbm2i7bU8x3wMds7R+Ff0OfHL446D+z98MNa8UeJL2Cx03R7ZriVnfG9gPkRR/EXOBgetfzk/Fbx9cfFz4q+JvFdwnlzeIdSm1NlJ+ZVml+VPqBhj9K+kyGEuax/SH0dcLiVi8TXirU1C73s30S8z03/gmjrNxoP7e3wnktpmhMmtJC+MfMjK4ZefUGv6H9ObJaPPTJx7ZNfz7f8ABKTwDcfEP/goF8NbWAM39l3j6jcgLnakUZySf95hX9B1vD5czfLt56+ox/jWOfSi66S7Hl/SCqQfEEI2972av97JHXAz61Xj/wBev0NWpR8tVY/9ev0NeGup+Cw+L5M/nd/4KR/8n8fFr/sOv/6KjrC/YgYj9tH4R/8AY3ab/wClKVu/8FIx/wAZ8/Fj/sPN/wCio6w/2IV3ftofCX/sb9N/9KEr7yP+669kf35R04JT7Ydv742P6QYBkv8A71SpGCO/X1qnGWDMM9+eakLMP4h/31Xwbtfc/gL2kXt+TJpECtTXgjmHzqrf7wzRHJ8vzc+9RzS7CTRG19BScVrI+Rv+C0Hwx03xl+wB46vbqH/SPDoh1i1lPzMkkcqKQuem4Eg+vevwsdtz/MuzbwFH8Pt+Ffsh/wAF4P2sdJ+H37NN18Oob2E+IvHUkcEluh3PDZRsJJHYD7pOABnrX42nr/jX2eRKcKDUlvsf2J9H7DYqnktV10+Scrwv1729T9Qv+DbvXrkw/FXSfMb7LHd2V7GmBhHZSr84zyB3r9UF+Rgv92vzK/4Nzfh5Ppnwy+IfiudJI4db1OO0tWZcCZIExuX1+ZiP+A1+moyZT+FfPZnb6y7H8++K1SnPirGOj8PMn+BNnJFfPf8AwU+5/YF+K3/YAl/lX0IeGFfPf/BT3/kwb4q/9gCX+VcdD+JH1Pl+G/8AkaUP8cP/AEpH88tkf9ET/PQHFfpx/wAG3HzeLPil/wBcrI/+Py1+ZVnxap91cnGW6DPFfZ//AAR2/bo8C/sP6v4yuPHE2rRLr1vaC1FpZm48zbuJJx9Sa+zzSEqlDlgtT+2fFbK8TjuFqlHBxc5Xjotz9yymZMc9PWl8pff86+Hv+H/XwGz/AMfXiz/wTPS/8P8Ar4Dn/l68Wf8Agmevkv7PxPWDP47/ANQ+If8AoDqf+As+4PKX3/OgRqP/AK9fD/8Aw/5+A/8Az9eLP/BM9Nl/4L8/Acj/AI/PFSnP8WkSCk8DX/lE+BeIF/zCT/8AAWfcRiUnOKdXz9+x9/wUN8A/txXGvR+BrrUpZPDrQrdi8tGt8ebu24B6/cPNe+20vmpu7Z4PqMA1z1KcoScJbo+cxmDxGEquhiYOMlumSV85/wDBVv8A5R+fFT/sDSfzWvoyvnP/AIKt/wDKPz4qf9gaT+a1rhP40fU9Dh3/AJGuG/6+Q/8ASkfz4L/q1/Cl27i3+0CGH8L5GPmHQ8evTtSJ/q1/Cuz8D/s+eKPiR8MPFHizQLBtW0/we0X9rQwkvPDE4LGdFH3lUAgqOtffScUveP8ARvFY6jhMPCtWny+9ZX2d3azPUv8Agnp/wUC8RfsLePAVebUPA2tTpJrOkFyfLc9LiA9FnUDDBcAjrk1+7Xwa+OWgfHvwPpniXwvqlvq2i6rGHikj55PVT/dYd1PIr+aCST7Q29csxBHpkHk/nX0b/wAE+P8AgoN4i/Yc8dRSK13qXgXUplGr6ZGd+wk5M0Oeky8/L0c4HevJx2WxqRU6Z+G+LXhXSzWlPNMsjy1ovVL7S76fh+J/QE5byW9VbimTzeXCzH76gfhXIfB3446B8ffhrp3irwpqlrrGj6tCJLe4gcNGT/ECeodejKehFeD/APBSX/gpRoX7EngI2cKpqnj7Voz/AGVpIYb416faJ8fdjHJA6sQB3r5elQqSrezSP5NweTYvFZh9QowftZPltbbzZD/wUl/4KW6H+xP8Pms4fJ1jx1qcZOm6X02t2mkA5SMdeeTivw8+I3xG174teOdQ8SeKNTm1bXtWlaS9nkO7ceyjjHlgHCjHb1pPiR8S/EHxh+IWpeLPFWqTax4n1Z/Ou7lySV5+RFU8RxjgKo70zwH4E1z4neMNN8PeGtNutb17U5RDa2UCb2lcnkHsFClmJPA219fgsFDDw5pn9ucA8A4LhLL/AK7ire1t70pbIyOibe2B29BgV7F/wTz/AOT9/g3/ANjTD/6KlrzP4geCrv4a+Ptb8OX7xy3vh+/uNNnkibMUkkEjRvsPU4ZSCT3r0z/gnlz+318Gv+xph/8ARUtdda3sJWPrOKq0K3D2Kq03dSpSd+9+qP6NIPu/h/jUlR25yn4f41JX56f5zlO9/wCPlf8Adr+cr9vr/k+L4rf9jHN/Ov6Nb3/j5X/dr+cr9vr/AJPi+K3/AGMc386+h4f/AIsvQ/ffo+/8jjEf4F+aH/sD/wDJ8Hww/wCxgg/ma/o1X7q/Sv5yv2B/+T4Phh/2MEH8zX9Gq/dX6Vln3xx9EafSE/5G1D/A/wD0onj+4KVjhT9KSP7gpWOFNeGfz8fmX/wcBftT3Hhj4a6L8LdOndZPFjm/1V422vHaR42x+uHYjOMHg84r8m2JaPYu1Ru3AKAozjHp/Ovpf/gsJ8TZviZ/wUH8Xh23Q+G1h0q2wcqQiZfP1Lg188+DvCk3jvxrpGh2xkWfWb2GxjKfeDSOEGPfmvt8vhGjheZbn92eGWV0ck4ThXSXvR55931sj79/4Irf8E4LH4xagnxY8cWEV9oVhcGHQdNnXdFdzx/euZF/iCn7oOVyAcZANfr1a26xMp2ru45x6DA/TiuU+Bfwo034KfDTQvC+lRCGw8P2kenxIoCrhEAJx6k5YnuTXbRRruXjp0r5XGYqdao3I/jzjPibEZ9mlTG13dJ2jHpFencj8hVOVX3wCcfl0o+ZsD+H2OKs7BSeWoPSuCzfxHy2i2SK01mksTKyrIrDBRvmDD3FfIH/AAVU/wCCcuj/ALYvwybW9Jt7XT/iNoNu50y8WNQ2oIo3fZJj/Eh2/KTkochSu5s/ZDDap+lUbyBJDuaISY6n2yP/ANf4V0YetOlNSi9j0MnzXEZXioY3CScZRd9PxXzWh/LvcW1xp99c2d3DJa3dq7W91A//ACykHDLj2I719If8Erf2mrn9nD9sHQZJLh49D8XSLomqJuCRIH4ilI6Aq2ORg+/Wt/8A4LQ/A21+DX7beqXGm2v2PT/GFlHrCxqu1FnYlZduPVhk+5r5M+3yae8VxEzJJDItwu3goY/mGP8AgQB/CvuI1FiMKnNbo/vL21HinhiNWov40dV/eS3P6irO7EkCFW3bgCDnqKvwf6oV5p+yv4zPxP8A2dvBfiRmZ5NX0e2uZGbqWKDOfxFemQn93XwdSPLNo/gGvh3RrVKb+zJr7mOoooqTMKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAI5f9av1H9a/FX4x/8AJZPF3/Yau/8A0c1ftVL/AK1fqP61+Kvxj/5LJ4u/7DV3/wCjmqokyOZHSg9KB0o7/jVEn1V/wSAhjb9pXW5G2tMmhNsIHYzJk/oPyr9Kq/Nf/gkEMftMaz/2ApP/AEclfpRUyKiFNkG5KdTZTiM5GR3qSj8ff22tXfWv2uviBI5b93qZhXJzwqgAfhXl9ewft+eGJPCn7X3jaORNqXt2l7EfVZI1/wAK8fqokS3CiiiqEI6LIPmUNg5GR0plxALiN1dVmjkwskT8rIvcH1qSjFBXMfJH7WH7OC/Dm7k8TaPGx8P3Uu25iRdzafI3RgO6e3QZrxBTx/XPX3/Hr+Nfo3rGj2mv6Rd2N8nnWd5E0E0eP4W7/hXwZ8X/AIZ3Hwk+IN9okqzNBB++tWI+WeOTABB9un4Gp5Sj2b/gkD4puPBv/BUH4K3tuWWS711rCXaxX93NbS27fmHAPriv6aoxiNccjHWv5sf+CJnw+l+I3/BUz4Rwqm620e6vNXuHHRRb2czoP++zD+df0nr90fSpAWqpl2K7c/K1WqqSf6o/T+lApXtofgh/wWXsP7K/4KRfECNPlEkOnTMAMbt1ouSfyH5Cvl/v+VfSv/BYbxDDr/8AwUg+IzQNuW0ksrCQ+jQ2sf8A8XXzUOP0r77ByvRhfsf6J+HvP/q7hPab+ygDdTX6a/8ABt1/yNfxL/697P8Am1fmU3U1+mn/AAbdD/iq/iV/172f82rnzf8AgfI+V8aP+SXr/wDbn5n6yf8ALX8KdTf+Wv4U6viT+Fwpspx+lOpsvSh6gflJ/wAHKf8Ax+/CH66p/wCgRV+YUbYcZYbeGIxgjGe/4j8q/T3/AIOVDi9+EP8Avan/AOgRV+YK8rX3GTStQR/cngnFS4RoxfeX5n3n+yN/wXEm/ZY/Z38J/D//AIV7Nq6+F7BbH7UNQSNZSpY5ClSRjOPwr0r/AIiSZ1/5pXKfc6un/wATX5jp92nA/X86qpluHk3KS3Z04rwd4ZxFaWIqU7yk7v3nu/JH6bf8RJc//RKpP/Buv/xNV9T/AODkHUZ7ORbX4XxxzuMQtNq67Qf9rCV+aOT/ALX51DuKnripjlOHesUmYLwV4Vur0n97/U9i/a6/b2+In7aWvxSeLNQt7fRbKTzLTQ7Astnbyf32bOZZP9tsr6KK8dJUAljtXJZmIzj1P5VLoml3ninVIdP0ixvtY1KYgRWdlB588uenAr9Gf+Cdn/BErWPEPiLS/GHxksxp+m2xS6sfDTn/AEh2BDK94R91QRwnfoa0qYihhlra56WdcScP8JZd7CjaLW0UtXbzPVP+CCf7F0/w68F3XxY8R2Mlpq3iaH7No1vMp822ss8yHPOZSN2e64HSv0mhOUrJ0DSIdHsrWzt4IoYbdVREiTEcSoMKo+grY6V8Xiqzq1HM/iTibiKvnmZVcyxDu5v7l0Q2b7tVY/8AXr9DVqb7tVY/9ev0Nc66nhQ+P5M/nd/4KRnH7fPxZ/7Dzf8AoqOvOPgt8QpPg/8AGfwj4vS3+3f8Itq9tqn2TzBH9q8mRZNgYg4JCHp716P/AMFIxn9vj4s/9h5v/RUdeJ9T9OR9a/QsPFSoxUtrI/0S4awscVw7h8NJXU6Ki+mjR+m2n/8AByDeW1svnfC2RpHAfP8AayjGe33e1WP+Ikm4b/mlc3/g4T/4mvzBU/5BqROR3/OuWplmEcrpHysvBfhdu7pfdKTX4aH6df8AESbKowfhTJn31hP/AIiuT+KX/BxP4y8S6O1v4U8Dabol5JGSLq+vTcLE3T5UUAMe/NfnfJ96mk7xtblfQ9qcMpw615TowvgrwtTqKtGk2ut3dfc/1Og+KXxR8R/GXx5f+JvFmrXmt6/qTbri7uCdxHZFBJ2IoOAo4H41F8NfhzrXxb8f6T4X8O2rXuta9cJb2sSRlsHozN/dRV5JH1rf+AX7M/jr9pvxeukeBfDl1r1wrKLiYMILOwU9XuJm+SNQOR/G3QCv2W/4Jw/8EtNF/Yl8PNq+r3MHiD4hasgW91VYSqWaY4gtlblVHQ5+ZsUsZjoYanyxOXjvjrKeFsveAw1va2/dwW8ezfke1fsd/s26f+yp+z74Z8F6fhl0ezVJ5OplnJ3SvnvlieTzivXQmTmq1qu2JV/ujHTHb07Vcr4upUc3dn8Q4zFVcTXniKzvKTu/mNP3q+fP+CoA/wCMBPit/wBi/N/6DX0GfvV8+f8ABUD/AJMF+Kv/AGAJv5VeH/iR9Tu4d/5GdD/HD/0pH88tgxFpHyf4f5GnrEqDaqqq8cAYFMsiwtk2feGSBjPr2r7D/wCCTH7AvhH9uvV/F9l4nu9Ytbfw7HBJbNYXAQMztJuzxmvvqtanTo88+h/oZxBxDhshy2WYYn4LRv1v5o+QNn+yv5UhT/ZH5V+zB/4N5/g731jxm30vVH9KP+Ief4N/9Bbxt/4Hr/hXlLPMItU2fl3/ABH7hv8All/4C/8AM/GfYv8Ad/SlAZFbayrxX7L/APEPT8GV+9q3jQf72oL/AIVDP/wb2fBlWG3WfGA+uoL/AIVpHPsM3o/wBeP3DaadpL/t1/5nk/8AwbcsRqfxkXd201sD123Ffq5ZKEhwPU9O1fOP7Ev/AATk8E/sKza43hS+1i6HiZomuzf3QkCiHJGOP9qvo+0BCNnOdx6ivmMwrRrV3Uj1P5b47zrC5vnlbMcJfkm7ry2Ja+c/+Crf/KPz4qf9gaT+a19GV85/8FW/+UfnxU/7A0n81rLCfxo+p5fDv/I0w3/XyH/pSP584x+6X8P5V+m//Bubb288vxVhmWORZBZqVddwIw2Rg9vavzIj/wBUv4fyr9N/+DcHm/8Aip/vWX8mr63NpWw7fof2f40Ta4PqxW3NF+ej0szkv+CsX/BJmb4PXd98Tvhnpgl8K3B8zW9BtIyf7Hc/fuLdVGTAOP3Y6HJr88RJz5i/vC6YYg/eHULn0J7dK/qEv7KPUIpI5I0kjkBDo67lYHqCO+a/IH/gq5/wSem+ED6p8T/htYSS+GJGa51zR4I9z6WxbLXMP96LGS6/wLk9q4MrzC1qVY/NPCPxcnCdPJs2qav4JPqn9l+fmfN/7FH/AAUB8b/sOX2rQ+HZItU0PWo2aTTbqQrB52P3VwndGDcEfdbHNeU/Ez4l698ZvHWoeJvE2oTatrusStPc3TkqOvyqufuqo42jA71gHax3N3ABYDJxWz4A8D698UPHGleHfD+l3mra3rVytrbW0Y5mY57/AMIABJb+FQT2r3FTpwk6q2P36jk+SYDEVM3hBQq2vKT2XmmN+H3gbWfin4z0vw54fsbrVtU1K4EVtbW67gXPUk/wqBks3UDJr9uP+CZP/BNbQ/2KvCDarq0VnqvxA1aLGoX5TcbNGKk28RI4TgZI+8VGak/4Jlf8EytB/Yn8HxX+qraax8RtUhxqWoLFiOyHX7LCD91VB5PV+W6cV9cTxhIMYGQPT3FfM5lmUqslCD0ufyj4neKVXOqjwGWzaoXSk7/G0912R/Nz+2DtX9r74s7BtV/GGqMQBgH/AEl/8T+Zre/4J5HH7fHwb/7GmH/0VNWF+1xz+198V/8Asb9U/wDSlq3v+Cen/J/fwY/7GmD/ANFy19E0lhXbsf0liP8Aki5L/pw190W/0P6M4Pu/h/jUlRw/fapK+BP4AKd7/wAfK/7tfzlft9f8nxfFb/sY5v51/Rre/wDHyv8Au1/OV+31/wAnxfFb/sY5v519Dw//ABZeh++/R9/5HGI/wL80P/YH/wCT4Phh/wBjBB/M1/Rqv3V+lfzlfsD/APJ8Hww/7GCD+Zr+jVfur9Kyz744+iNPpCf8jah/gf8A6UTx/cFEvMbfSiP7gpzfdNeGfz8fzg/t7iQftv8AxU8z5nPiScYP+6n+A/Kqn7DMVvN+2d8K2utrWv8Awk1iG3DILBwf58/WvR/+CvfwwuPht/wUB8cM0Hl2/iAw6xbf7kkaq/8A48lfPfgjxVN8PPHGh+Ibb/XaFfwahGPVonDj+VfdUfeweh/fuUxWN4Lpqg7uVG2ne2x/TtAq71Kqv948dvWrcJyRXF/B/wCJ2n/Fn4d6D4m0yTztP8QWcd7by+qSIGArr4W/fL75r4ip7s2mfwJWiqdWpSlpKLs/XsWqKKKgkRztWqVzNHN8hC8g1df7h+lZ2oEJ83cZI9sc5prcmbXK7+n3n5Hf8HGaW4+Nvw82kfa/7KnOPVBMvP4EmvzlmRZbabbyDHkHv9z/ABr67/4LY/HG3+MP7bV5Y2cy3Fl4Nso9I3j/AJ75LyD82r5Ga1kvJIoIg/2q4byI41H+sZxtx+or7rBrlwqT6I/vjwwwssFwnh3iNHyyevRbn9D3/BNMyP8AsJfDMS7t39hQdT2xXvNt/qR+NeZ/sn+B2+G37NPgfQCrJJpOjW0DqeqsIxn+denRDEYr4itLmqNn8N5xVVTHVZx2cpP8R1FFFZnmhRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQBHL/rV+o/rX4q/GP/ksni7/ALDV3/6Oav2ql/1q/Uf1r8VfjK234yeL88f8Tq76/wDXZ6qJMjmR0o7/AI0L0o7/AI1RJ9Wf8Egv+TmNZ/7AUn/o5K/SivzX/wCCQX/JzGs/9gKT/wBHJX6UVMiohUd1zA38PuByvvUlFSUfAf8AwV6+CdxDqXh/4gW8bfZ2Q6XqBxnY33omPoCA3J74r4nByM9uma/bH4tfDfSfi/8AD7UvDOuWwutN1iE20qd1BH3gezL1B9RX5CftEfs/ax+zF8R5/DetIzRrl9Pv8ERalag4jkHbeBw4HfrVJi5TiaKAcpuGSv8AeHSgNuHHT19aogKKKKACvCP27Ph//bnw+tfEEcf7zRZfKuCOS1vJx/PNe67xu2/xBS5HcKOp+nvXVfCP9lXUv20Nbn8F2fmQ6VeW7Lquohcx6ZbnBLk9CxHCA9TnFBob3/Bsd+y9ci68dfGrVIJIbeS0/wCEX0GWRfmlQSCe9nU+isIUU9PlIr9jlOVzXBfAb4D+Hf2cfgv4d8C+F7D7L4f8O2K2NumPnKgDLsR95nYZJ7k13aBvLXdjdjnHTNZgO3VVu5Y7K1kkklWNIlLMxb7oAzmnAbFP5183/wDBTj9o2H9m/wDZD8XaoJlGpX1k2m6cC4VpLicFRj1wpJGPStMPTdSpynZleDnjMXTwsN5NL72fhz+1N8Tj8Yf2l/HPiiPd5Ot6xcTwt3Kliif+Q42H41wfX9Kbb/uofL+YbQylmO5nZW5b6NuJB74p4Ga/QIxUYRiuh/pRlGEWFwNLDbckVFedhG6mv01/4Nuj/wAVX8Sv+vez/m1fmU3U1+mv/Bt0P+Kr+JX/AF72f82rhzf+B8j8z8aP+SXr/wDbn5n6x/8ALX8KdTf+Wv4U6viT+FwpsnT8qdTZfu/jQB+Un/Byof8ATfhD/van/wCgRV+YUO5m+VthXAVu6sCrD+Vfp5/wcq/8hH4Q/XU//RcVfmPH9+P/AHlr7TKf93sf3F4LScOD6Ml3l+dj9fv+Cbv/AATR+Bvxu/Yl+HfirxT8PNJ1TWtc0hJ7m5kkmDSvuZcna4HQAcDtXvSf8Ecv2ayo3fCbQWPc+fcc/wDkSj/gkR/yjV+Dv/YAT/0Y9fTkf3BXzWKxVb2slzOyZ/JefcQZnHNcUo4iaSqTWkpLZ+TPmP8A4c4/s1f9Ek0D/v8A3H/xyiw/4JF/s46ZP5lr8KNBWX/aaeT+clfT1Fc31ir/ADM8r/WLNf8AoJqf+By/zPOvh3+zn4L+DaeX4T8H+HfDqsRubT7CO3c/VlGTXeCDH4GrNFZSblqzya1arWblWnKTfWTuQqMyL161NRRSI1GzfdqrH/r1+hq1LytVY/8AXj6GhdQh8fyZ/O7/AMFI/wDk/j4s/wDYeb/0VHXH/steEtO8eftN/DzRdYt0vNH1fxFY2V7bMWxcxyTKpjO1gcHPP0rr/wDgpJ/yfx8WP+w83/oqOsT9iL/k9X4R/wDY36b/AOlCV97CVsNfyR/f1CU1wVCUHZqje630jc/aKx/4JAfs1y267vhRogwNo/0m6HA46eZ7dqtR/wDBHb9muRf+SUaGef8An4uT/wC1K+lbXo31qePofrXxP1ir/Mz+GlxFmltMTU/8Dl/mfBv7X3/BFj4S+JPg1r8Hw78E2PhvxZDaNLpl1ayzbnkUbjCSzEbZMbTX4v39nc6fqFxa3kLw3lrK9vNGybWSRGKMpHYgg1/UNqSLJL/F0AyP4cHIr8Zf+C6P7FMnwi+Nq/FLRbfy/DvjJ1h1YRKVWz1HA2SAYwqSquCT1k3dzXuZNj5uXspu9z9x8FvELEwzB5TmFVyVT4XJt2fbU8N/4Jq/tkzfsZ/tO6Vq11JJ/wAIrrbJpuvQq52sjnEc+PWNj+Rr9+tC1mHVrO3uoJkmt7iJZI5IzmN1IBBB9wQfoRX8wcsTlHWTaTsZmVf4u4wezD0r9kP+CFf7ay/GD4LXHw28QXqv4g8BxgWckkgZ77TWxsY/3mibMLEdAi5xmtc8wWntYnr+PnBSioZ/hI+9tN73XT0P0LQ5RfpS01Gwi57jv3p1fLn8vDT96vnz/gqB/wAmCfFb/sX5v/Qa+gz96vnz/gqB/wAmC/FX/sATfyrfD/xI+p63Dv8AyM6F/wCeH/pSP55LH/jzj+q/yNfpz/wbfHPi74pL1/dWn/oUlfmNaEm0TbkZOMj0Of6Zr74/4IcftSfD/wDZi1/4hTeO/FWj+E49US1SzN6xQ3AVpM4bvjI4HqK+uzSMpUPd+4/tbxfwdbFcLyhRi5S92ySu/wAD9nkQZ+7Tii/3f1r5zX/grB+zs3X4v+EP/Aoil/4ev/s6j/mr/g//AMCq+T+rV/5X9x/Fv+rua/8APip/4BL/ACPovavp+tMkUgja2B6V87/8PYf2df8Aor3g3/wKNL/w9e/Z1YcfF7wf/wCBVTLC1raxf3Evh/Nkr/V6n/gEv8j6JibIx3qW3Py15j8Cf2tfhv8AtOHUl8B+MtF8Vto2z7aLGbf9m3g7dx7EgEjPXB9DXplmoSHC/dBwOKwlBwlys8qpSq0ZeyrxcZdU1Zomr5z/AOCrf/KPz4qf9gaT+a19GV85/wDBVv8A5R+fFT/sDSfzWujCfxo+p63Dv/I1w3/XyH/pSP584/8AVL+Ffpv/AMG4H/H98Vf96z/rX5kR/wCqX8K/Tb/g3C/5CPxU/wB6z/ka+szdf7K36H9meNX/ACSFX1X5o/VLH7w1DqNlHc2UwaMyq0bZTZu3jByMd8+nerUXEhq2y7l/DFfGyl710fwxrp6L7+5+P3/BSn/gjPq2i+N/+Ew+DehyX2k67eA3/h+EcadcykZlgB4FsxOWXqjZIr68/wCCZH/BNHR/2IvCq6tqnl6p8QtYtwmpam0efs4Iz5EJ7IP/AB4jJ5Ar7BEG3mj/AIC35V2VMwqypeyZ9xmHiBneNymGT4mq3Tjo31kuzZUtYNrQsGbAzwR1qW8/49m+g/mKn/A1Xu/+PZvw/nXHHdHxMfdtFbJqx/Nr+1r/AMne/Fb/ALG/Vf8A0peui/4J3/8AJ/3wY/7GiD/0VLXO/tbf8nffFf8A7G/Vf/Sl66L/AIJ3nH7fvwY/7GmAf+Q5a+8l/ur9D/QCt/yRcv8Ary//AE3I/ots+n4VPUNmMCpq+BP8/wDoU73/AI+V/wB2v5yv2+v+T4vit/2Mc386/o1vf+Plf92v5yv2+v8Ak+L4rf8AYxzfzr6Hh/8Aiy9D99+j7/yOMR/gX5of+wP/AMnwfDD/ALGCD+Zr+jVfur9K/nK/YH/5Pg+GH/YwQfzNf0ar91fpWWffHH0Rp9IT/kbUP8D/APSieP7gpW+6fpSR/cFOb7prwz+fj8v/APg4I/ZjuPEfgnwz8VNLt/tEnhx20/V2RB8tvJ/q5f8AgL5X6Gvylkj3R7cYLZAB7/5Ff01fEX4d6b8Uvh/q3hzV7Nb7S9ZtZLK5tpANskb5zz6jPXsa/AH9vH9iHxB+wx8aZvD+pJLceGdWklufDupnJW7gX5vLZ+gmj6FM5K84xX1GT4yPJ7GZ/VfgXxxQlhXkeJdp814Nv4l28rH1x/wRN/4KLWvga4t/g/40u4rOxuJ3fw7eXB+WIt87WkjdPlJO09+lfrNaXH2l92R8rFc5zuBGRiv5dWODxv3LhxsOGXnhuORg4ORz6c19wfsTf8Ft/GX7PFhY+H/HOn3Hj3wzalYYbgTqmpWKj+ISSEefGOPlfDjHGeBRmmU8376mcnit4PYmri5ZvlEL82s4efdd35H7bRkF+MelOlOR+NfIPgD/AILZfs7+NbNXk8eJoFxyWt9ZsJ7ORPxZQjY9QTWt4g/4LC/s6aJZGVvip4dumXJVLXzZnY4PACA/rxXgfU6/8p/PdThnN4T9m8NO/wDhf+R9RSTpGCWZVUA7iTwPc18m/wDBUT/goJpn7GPwia2tLq11DxxrUbxaTYoRuQnA+0P6KgbPPU9OlfMn7Uf/AAcC2n9kTab8JfDsmoXLqUXWNcjEFvDngukWfMdh1AYBfWvzS+I3xD1z4ueM7/xF4o1i71rW9RkMk11dybpGHYAfdUDsqcCvVwOUVXJSkfr/AIfeDePxeKji85hyUoe9y2u5dlYzdT1G41jVrq8uriS7vL64eeeWTl5pXbczH/ePIFfQX/BLL9mK4/ah/bH8NWbwSTeH/Dsq61q8mf3YSEny0z6tLgY9j6GvB/B/g3V/iJ4x0zw7oOl3esa5q032exsYR81zIwxjI5Udy3RAMnFfvN/wTK/YKtf2HPgNa6XczRal4t1pku/EN+E4aYoo8mMj/lnGFC+7bm/ir1s0xUcPS9lHc/Y/FjjPCZLlE8Bh/dqVo8sVfWK6to+k7K3FpAoXPU8Y+77Vdtn8yEGo7cEJ82c96mi+5XxZ/EEW3eTdx1FFFAwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAifmRf96vxU+N0TL8aPGUbrtZNdvAQe3796/arOybLcCvxd/aCGPj/wCPF9PEN8n4+e5AqokyOTHT+VHf8aReFFL3/GqJPqz/AIJBf8nMaz/2ApP/AEclfpRX5r/8Egv+TmNZ/wCwFJ/6OSv0oqZFRCiiipKGylsDbnrXB/Hf9n3w7+0R4Gm0PxJYfaIW+a3mTC3Fi4BAeN+x56dCK76igD8rf2jv+Cbvj/4M3lxe6TazeLtBTLpPZR4vFHrJGP4sdSOuK+fL2SSwvGhvVktrleGS4/dyfQq2CMV+59xE0jqVx8uc8c/gc8frXL+MPgl4V8fEnW/DWi6szDBe4t0aT/vrbn8zVcwrH4r+av8AeX8xTrBDq9wLe0WS7uXbasUKGR2PoAoJr9dl/YX+FKtuXwHoA5zjyT1+m7Fdd4N+Cvhn4eH/AIkfhvR9LY/x21qiH8TyR9QaOYXKfnL+zn/wTU8dfGW5huPEUMnhDw6zh3mukxeSqP8AnjH/AAt/tnp6V+iHwQ+Bfh/4CeEI9D8N2P2OzjPmPK3zzXb95JH7t7dK7WMNs+ZPm6cn/wCuakXhemKOYoZEVhiVf7oxThIpqMjLVG06xg5bHc+3v9PejlCTS06kdzeqgJVuOP16V+Kf/BcP9riH45/HmPwLo9002g+CHJvWibKXF6w5Hv5YAX65r7x/4Kxf8FCLH9jX4KHTtHvbWT4geKInt9GtuCLVT9+6cdkT/axk9M1+GN1dT6jeTXVzPNcXV07zTyStueVyepNfRZLgnb2kz+ivA3giriK7z7Ew/dx0in1fcaOaMYNKKG5FfRcyP64jKy5WNPLGv00/4NuTnxX8S/8Ar3s/5tX5lrwTnjrzX6Yf8G3p/wCKt+Jqr95YLTPp1fvXn5x/A+R+T+NEf+MWrt9ofmfrOGDTde1PqODDncPTFSV8SfwqFNl+7TqhmbE2O2KFqB+Uv/BylIJdQ+Ee35vL/tQtjt+7ir8yEOHT/eXH6H+Vfur/AMFN/wDgmjd/8FAbjwbNY+Ko/DLeF3uS5azFx9pWZVUjqMY2ivlU/wDBt/rcYx/wtWFlK7cDRl+XJHOd390Gvqsvx+HpUeST1P6j8MfEvI8r4fhl2Nm1JXenm7n2b/wSHYN/wTY+D6/xLoKAj0PmPX07H9wV5P8AsZfAGT9l79mbwd4BuNRXVpvC9iLR7wQiIT4ZjnaCcY3Y98Zr1iHPljNfN4iSlVco7Nn83ZxiIV8wxFen8Mpya9G7jqKKKxPNCiiigAooooAa5wtVWkVJ13HrkD64Jq1NylZ9whlG3bv+YEg9xkZ7jtQEb3uj+eH/AIKRuG/b2+LB9debHH/TKOsP9iFt37avwj/7G/TP/SmMfzIr9GP2lv8AggvrHx9+P/izxpF8SbfSofE199tWzbSw5tiUVSu7dz0qj8C/+Df7VvhD8bPCXiqb4nWt/b+G9Wt9Te1XS9jTmFw6qG3ccivro5lQ+r+zb1sf11DxU4djwx/ZrrWqey5LNPdxsfpvaSq2/B/iqzGMD8ao6RZG2t9oXp3xjPX3NaFfIvc/kVetyC4Qs/3a81/ai/Z70n9pn4N694N1y286x1e2MYbHzQv2cHsVIUgj3r1GozHmQtt+bHX1rSnUdOXNE2w1edCtHEUnaUXdNH8yXxp+DmtfAH4p654P16FodT0G7NrJ8uFkBPyOvqrLgg9K0f2ZP2gNU/Zg+OWg+PNJkmWTRpiJ0BIW5t3CrLGw/iG0EjPev1+/4KL/APBIex/bl+I2m+LNP8QL4T160t2s76T7IJo9Qi6oMbvvIe59a+dv+IbfXAnzfFi1b1H9jgbvw3V9ZRzXDToWqvU/rjKfFzhzMMjjg84qWlJKMk02lZWbR+nPwk+KujfF74b6P4h0W6W60zW7aOeFlbcQGXP5jofQiuv89c/er5h/4J2/sZ+KP2KPhdfeEda8XR+MNIW78/SwLEQPp6P9+LJblN2H45wCB6V9KpC32gt/Cfevla0YqraGx/KGa0aFPGVYYWfNBP3X1a8yyWGa+ff+Cnp3/sC/FbHP/FPzfyr3/a2zOPevNv2qvgXcftE/s8eLvA9vqA0mXxJp7WaXpj8wQFv4iuecelFCSjUi30ZeSVqVHG0q1Z2ipRb8kmrn81+nvizjYDdwpHHbGP6inFVByzbj246V+mEP/BtnrkcSj/ha9v8AIMADSBzwP9r2p3/ENxrR/wCarW3/AIKB/wDFV9f/AGth11P7aj4ycJumnLEa2S+F9Ovqfmbhf7w/74pQinoR/wB81+mP/ENtrX/RVrX/AMFI/wDiqVP+DbbWgf8Akq9r/wCCkf8AxVP+2KH835Erxm4V/wCgn8GfmYUA9P8AvmlIVYz8w5yPu98V+mT/APBtvrR/5qva/wDgpH/xVNk/4NutaRf+Sr2/HPGjq3/s1H9r4Z6SZcfGrhVST9vfXaz/AFJv+DbpCNX+MHpnSx9TsnH86/WK1bfFn3NfHf8AwTB/4JpX/wCwJd+MJrzxaniT/hKDbbdtkLfyDDvIOMnP3v1r7CtAwRt3r/Svl8wqxqVeaOx/I3iHm2FzLP6+Nwck4SehNXzj/wAFWJlf/gn98VlDDK6NIW9uVr6Oryf9sD4BTftNfAXxV4Hj1L+x/wDhJLQ2n2vy/MEIJByV79K58PNRqxcu585k+Ihh8fQr1ZKMYzi232TTP5vonBiX6A1+mv8AwbhzKNS+Kgz1Nmw9x8wzVUf8G22tN974tW5bvjSB/wDFV9Uf8E0P+CZl3+wJJ4tkvvFVv4ofxGYPL22XkfZxHnr83OSa+kzLH0alBwgz+l/E7xM4fzfhueAwFbnqStZWtrfU+uk+Z8+1XKpxhkHTvVwdK+UTurs/lbye9kFFFFMY2T7tVbvm3b8P51ZuDiOq1xAJo2UruXIYD1xz6j0o5rNBHV27WZ/Nr+1yfL/bA+K4bg/8Jfqhx9bl63/+CerD/hvj4M8/d8V25/8AIctfdvxo/wCDfXVfit8avGHiiH4nQ2MfifW7zWFtzpIb7OJ5Wk2Z3c7cir/7Ov8AwQK1X4GfHrwj4ym+J0OoReGNTj1FrVdM8szbAw25yfWvsXmOH9i4X6H9e1vFPht8MPL3WtUdPltZ7uLX6n6ZWl3HLJIFcHaeas1nabD5ESrtxg9T1PJNaNfHadD+QI7b3Kd9xcKfav5y/wBvkZ/bj+K3X/kYpm6ds1/RtfK00nlmMtG2ASGA/wA9B+dfmZ+0h/wQS1r43fHvxd4yh+I8Omr4m1GS+FqdN83yA5zt3FhnFexlOKhQqOU+p+veDvFGX5JmlWvmUmoSjbTyPz8/YMbyf23Phg7cL/wkFuSfTJOK/o2Q5RT7V+ZPwB/4IEat8F/jl4V8YXHxHtdRh8N36XzW66V5bXBXou4NX6a2a4to+Cvy9PSpzjFUa1SPsn0H4wcVYDPcwpYnL3eEY2d/N3LMf3BTjyKROFFLXkn5EQNAojPy7uDx615h+03+yz4V/as+FWoeEfGWl/2lY3S5in4EtrKB8s0TdUZe2O45r1amXC74GHPI7VVOUoz5kzbC4ithqsa+Gk4zjs10PwA/bg/4JYfET9jrXJ7yGzu/FPgpXZrfV7SIyNDHnhZ41yVccDcOMV8xxXEdxHuSVfLB++pyAfTPr7V/URqeiw6pZTRXEEbpcDEiN8yMPdenSvlf9oX/AII1fBH9oC9l1CXw7ceG9YmJ33+iT/ZXbPdkA2Nn6A19Hhs6UY8tU/o7hHx+lhaSwueRcmvtx1fq0z8I2Yv827ft9BzQh2txuXHUkA1+pXib/g2w0+RpBovxa1SCDnZFeaNFMR9XDqT+VZ2kf8G21zDc7r74veWmR/x76BH8wx05kOD716Mc2wrVz9Mh44cLVKbqTrWfnDX8EfmQR5kvyjd0713X7Pn7NHjb9qXxhFofgrQLrWJmkHnXGDHa2w5G55ei45PGemK/WT4Lf8EBfhD8PLuC48RTa344lhO8pf3Xk2rn/rkgGfoSBX2V8M/hF4f+FnhyPSfD2h6b4f02FQEt7GFYFHXIKrn25yc5NcWKzyPK4Utz4riT6QVCKlHKablJqym9EvO2/ofNv/BOf/glb4a/Ym8PNqt7NH4g+IGpReXeayIwBaKeTHbqfuKO7dWOTX1rZweQm1VAUHjjrwOvqfepYohGTkenHp9KmiwF+Wvm6tadR802fzJm2aYvMsVLF42bnKW7YwgsakQYWlorI88KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigCCc+W+5vu5BH9a/K3/gpB8Drj4PftF6pqixyLofjSY6haXDD5Dcklpo8joQT0PWv1XdA/WuH+N/wK8O/tAeBLzw74jskurK4OY3X5ZbWTtJG3Zh+tAH4yA/uw20hT/479aUHp9a+kf2h/8AgmL44+DCXGoeH5B4y8P26l2khXbqEAH/AD0ix8w91LE+gr5tG4SsjJJHJGfnR1Kun+8p5X/gQFUmSz6t/wCCQX/JzGs/9gKT/wBHJX6UV+a//BIMf8ZL6weGH9gueDnGZk4r9KKJBEGbaKjkuFjQsf4eTTphmP0qrdMEh3MflHWp62Kd+m/YmN/GMfey2cDFO+0rXxP4A/4LM/D+b9obxH8OfH27wPq2i6rLp1rfSyebp2o4ICEyD/VscH7wCjnmvsDQfFFjrthDc2V1b3UNwu6OSKRZFcdsMCQcjng1tUw1SCu0ejj8nx2AUZY2lKKkk1ppZ7O5tRzLIxC/w0M4WqlnN88nzJt9Qen1qfzlkPysrFeuDnFY69TzFdrR381sSBwaUtikjbio5ZQsnvigpJvQkMqg0yW7WIfxE+gFU5J905+Vh2GASf0/rWL4t8c6T4X0OfUdX1Ky0vT7VS811cTCOONR1JY8f1q6cHJ2SJipSqKklq9rK/4G/JqEabchhuPcYx9a+ZP+Chf/AAUY8KfsT+CJ/tUyap4svoidL0WGQefcP2eQdUiHfPWvmb9uf/gvLovhywvPDfwVaDxFqigxy+IZ49+m2QJxmNSQZpPTIVV6gt0r8uPHXjfW/if4vvvEHiLVL7Wdd1GTzLi9upjJLIc5wTj7q5woAAAwOa9vAZROo71dj948O/BnHZlUWPzT93QXfRv/ACNP41fGTxN+0J8SdS8X+LNSfVNc1QuXdh8kEZ6QxD+CP29q5kninA0N0r6iFox5Ef2BgcHRwlCGHw0VGEdktg3Ufe49ajDZcLtbcRnB4z6fUntjNbnw8+HniD4s+LrfQvCuj6h4g1m4YLHbWMJlYH1Y/dQDnliBxTlJRV5MeLxdHC03VxM1BLXV9jFjmSA+ZLtWMBjuf7uQOh9K/aP/AIIN/ssal8Ev2cLzxPrlnLZat47lS6jgmXbJFaxjEWR2LD5se4ryj9hD/ghFDout2Pir4zXEF/qEEkd1B4ZtObS2I5V7iXA85sj7qggEDkiv040DTo9LMcMflpFDHsjRBhQBwMenGBgelfNZtmUKkfZ0z+TvGLxOwebUllOVe9FW55dHy7WNKOMo3+zipKKK+dP55Co5oy/SpKKAeqsVRbuoxinC1arFFO5EqcWkuxXFsyvwOKnXhaWikWFFN84A85HOOaElD0AOopok+tOzQAUUUUANlXctRRwN5h3AbamZtopPM9jQS43dyNoN69vxFNNsw9PwGKm8z2NKrbqChkCMgO6pKKKACiiigCnNZMZmYc7venLauuP5GrVFG4+aXVlM2kv96pvJ2/w/r0qbNG6haERpxXQQ5K1G8TMuKlooHYrraMppwtuamoqeUd2R+R7mk+z/AFqRm2ikV9xp8qAb5HuajmtWdlxyO9Ts22kD5p2QOzVmVzaMhG38asQrsShpQrYp1BMYxWyCoXtmkkY9AamozQVr0IBbEehpPs5HapjMoP3qcDkULQXKr3IzG1SDpRRQMKKKKAGygsnFRiJhn0qaila4WK/2XJo+ysT1qQ3Cg96kVtwoirDTsQLbsp9efWp6KKYhkqk9KiktGkOasU0vg96CZRUlZkKWpxzTzEQvSn+Z7GgTKT1/CgIxS2FQYQUtAOaKCgoblaKKAIzGxNNlgJTgfrU1Jv8AmxQHoQRW7A/N/OnSW24cetSq26hm2ihaaIPUia3ITpzRbxMhbd+FPMwB708HIp3J5VvYa6H+GiMELzTqKQ7a3CiiigYUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABSbRS0UAQmzHnbxjgYB/i/OvHf2g/2FfAv7RkU02rabDY6wwxHqlivk3H/A8cP+Ne0UHkUAfHv7F/7D3iP9lb9pLUr+4vbfW/Dt1pLwW9+g8uRXMqkI69zx1HFfYVQxWnlvnd+AGM+uamoAbJ9yqtyuYGDcg9RVqT7tV7kZhb6ZoW6D/gH84P7bm1v2v/idHtzHJrtyrL/Cfn/u9Kh+Af7ZHxQ/ZfvY5PBfjDVtPtVfe2nTSfaLKT/Z8t/ur7Lipf22T/xmF8Sv9rX7k/8AjxrzAjd/+uvv404zppTXQ/0TyXLcLj8iw1LGQU4+zirNX6H6L/Bz/g4i8U6BHDb+N/Amn6wp4lm0a4NrI3uEbIz+NfQ/g/8A4OCPgxqdvGNV0/xjos0gywlsRME+rKf6V+MAAAxg8+9M8tc/wrzk4jUZ+pxmuSWTUJvmsfFZp4H8NYqTlSjKnfpG9l6H7kn/AILvfs+rCrR63rUxY/cXTH3D8K4zx1/wcJ/B/TIJF0nQfGGuToSPL+xi3Rv+Bsa/G0y7vp2Axx+f9MU0qrdN3r2yT9cZqf7Dof0zycJ4A5DSlerOo0foR8YP+Dhzx34mjktvBXg/SfDsbZVbi+lN1MgPoBhQRXxj8cf2mviB+0lqjXPjnxVq3iAZ3R280pW2gPosS4X8wTXDKW9f0pScg5rto5fRpfCj9FyHw/yHKGp4Ogub+Zq7f3iN90Kvyw8fu/4RjpxSE7qDTZJ1hHzMq+54B/Hpn24PpmunmS02Pr/aRi/dtbzdl92w/NOhgkvb6G1t4prq6uGCQwQxmSaZicBVQct+GRXv37Hn/BNT4mftjajHNp2mt4Z8L/el1vV42iTb6xRcGX81HvX65/sT/wDBKn4Y/sc6XFeWFm3iLxUF/fa7qkYkut3fygSRCv8AspxXl4rNKdJWWrPybjXxdyvI/wBxRaq1v5VsvmfnX+xd/wAERPHXx5e31T4gTTeCfDjBZPsv3tSvQSMMF/5ZKPXrX6ufs4/se+Bv2VfCkGi+CNB0/SYkUJJchN11depkkPzNk846V6YkflOvzBtvGAuM/wBc/p7VaVv3q/L/APWr5rGZjUryuz+TeLOO82z+r7TGzcYdIrT77bgtptCD5ML22/y9Kkjg2S7vl6dqkorgPjgooooAKKKKACiiigAqNrkJ97jmpKrFBNIwb8KTYHlfxr/bU+F/7PXiyz0Xxz410LwzqOoQG6t7e9n8uSWHcV3gY6ZGK5iL/gqX+z7H/wA1W8K+n+vb/wCJr82/+DiO1SH9sHwh8ud3hgqDgEqpuzkcg9znjHpXoHhT/g3S03xP4Z0vUm+KmpW/9pWkV1sGhWzeXvUNjO7nGeteysvpRoxq1ZWufrWF4KyGjlWFzHNMXKDrRbSjG+qevT0P0F+GX7evwe+MusLpvhj4ieF9Y1CQ4S2hvAJX/wB1T1/CvVobrzDlRkdRyOfWvxT/AGzf+CJPiL9k74a3Hjbwv44l8UW/h0fab2H7Iun39pEMZlgkRz90nofTPtX2X/wRE/bK1r9pr4Caronia+bUfEPga5WzN44+e/tHQGKRz/fzvUnnJjzxnAwxGCioe1pSujyeIuEMvpZZ/a+SYh1qUXaXMrNN7H3TFceYfu1LVW0JwrY61arzUfnUZcy5l1GTvsTPvVK71ZLNHd/ljjTzHcnAVe5P0HP4Vavhm3I/vcVi+LLTz/BmoLv2t9jlUkdxsanHWSiOn71VQls+vzOD+Ef7bnwv+PXi6bQfBvjLRfEOr2qGW4tbSYtJAgOCzDHQHj616rb3G4HKlW7jPSvxQ/4N97GNf23/ABFldzf8I9dLuz2N0M1+1Fmux35+98/4kn+mB+FduYYVUJ8i7I+p4y4fp5NmKwVOfOuWLv8A4kn+Fy4j76dTIelPriZ8sFFFFABTfM56U6oHl2ls/dyQPU+vFAGN4l8b6d4RtFuNTu4LGFpo4A88gjXfIdqDJ7s2B+NaSXO8cf3iuM9cf41+TX/Bf/8Aa3bVPFei/Cfw/qUkZ00jWtaeGRg0dwp3WseQOMNh8ZOSAvGd1fbv/BMb9qyP9sD9lLw/r9xOkniTTQNL19BjKXsIG58cYEisrgc43EZOMn0KmCcKSqs+wzPgvG4LI6Ge1F7lZtW6r19T6Ujk8wU6o4vkRT/eqSvPPjwpsj+WuadTZl3pSewFeS/Hlhtvfk56f54H41j+KfibovgHQptV13ULPR9NtkLy3V3MsUMQ/wBpj0PtWlPBkqPl+X5st91T6n6dfqBX4r/8FVfi74y/a/8A+ChDfBvSZ2j0vRb6HSLDT5XMcFxdPGryzzc/MFDcEg59B1rvwmD9tUabskj6fg/hmWdYv2E6nJGMXKTtskfptN/wVL/Z/iuXhk+KvhFZIuGAu9wH4gVDN/wVP/Z+Xb/xdTwpt5Bb7Q2B6fw96+J9K/4Nv4dR06KS7+K98twVBfydDgEZJ/ubiflHbBNWk/4NtNNSf5fi1qbtwDnQ7Ue+OuK6vqmE/nPq5ZHwVBOM8xlzbW5Hv9x+lHwp+L/hv43+ELXxD4T1iz17Q7wssN5atuikKkqwBIHQ5FdZFJ5gPGOa8Q/YV/ZHh/Yq/Z/0rwBBrk2ux6XNPKLuW2SBnMsjSEbUOBjdjqfw6D2yzbdGfrXlVPZqfLB3PzXFU6UK844eXPBPR7EtV5SfNarFQSL+8+p9Kg59b2Ob+I/xJ0T4R+EdQ8QeJtQttG0PS1Mt3fXT7IYEwPmJ9MkCvJk/4Ko/s+7F/wCLqeFenX7Qx/8AZa6j9rn9nGH9rH4DeKfAcurS6QviSz+xNeRwiVrcbgchGIDcj1FfAV7/AMG32m2VrJM3xY1NvLDNt/sG1yce+6vQw1GhUj78rM+w4ZwHD+Jg/wC1cTKm7pLljff5H2vJ/wAFVf2eY/8AmrPhHPp9pOfyxV7wz/wUt+BfjLxHp2kaX8S/DN7qWrXEdraW8c7b55HOEVRt6lsAe5r8Q/2HP2NrT9sb9pWb4eHXJvD8cNreXP2+G2jkkP2dwu0ocfeznO7jHQ19+/Bj/g3u034VfFrwv4s/4WhqF7J4Y1ODUVtv7Gt41uPKkDiMsDlc4xkc12YvL8LRdvadD7PijgvhTKG6EsfL2vLzJOG99tkfpct0r9OnGD9alrP0uIRRqC2RwB+FaFeHeN/ddz8bTvqtugjNtWoprvyk3bd3IB5qSXlKrT52rg/xA/hRunYfWx5b8Zv20Phn8APGOm6D408YaR4b1TVoBc2lveS7ZJkLbMgAH+Lj8a9O0jVY9RsluIGWaCVQ8cinKup5BH86/HD/AIOCwsf7aPgWMKNqaAiJ/eUfa8Dnr1APX/Gv1w+EheH4e6HGx8wtYQEt0/5ZrXbisL7OhGoup9dn3DlLL8qwWYKpeVdNtdmtDrI5fMJ9qdUNrJvZvlxU1cR8jZrSQ15drYqtLfbZMbW5PU8D86lnfD15B+2h+1Dov7InwC8QeOtaja5TS4Vjs7INiTUbyUiO3hTrjdIRubB2qGbBxVRpylJJGlGjVq1Y0aUXKUmkkvP/ACO/8afFTQfhtoMmpeItV0/Q7GIZae9uFhjHtk9/avL9G/4KJ/A7xD4ibS7P4neEZtSZsC3N8I3J9PmwP1r8fvBPw7+N3/BYr4wX2oX2qCPSbOcLLd3kzDRtDRsbYUiU5klA65HJO7IzgfQXjL/g3D1SDwg39j/FKK71YR71gv8ARfKt2YdFz5jMik/UgdjXsf2fh4x/eVNT9YrcB5DljWEznMOXEdoq8U/7z6Pufq/pvjCz1W0hls5YruGZA8ckUgZJAfRhwa01m3Jux+FfnR/wR1/ZK+PX7PHiLxCvxA1B9I8I2ObOz0GSb7WL2UMR9phYkGKIY+UYO4YPydB+iCy7Y1XGa8vEU1TnaLuj83zvL6ODxssPhqyqxWzWxapsj+Whb0pwpG+6axPLfkQyXm1CcEcZrC8X/EfRfh9o02qa5qljpGn24JlubudYYk7YLMQM57da2b1Glt3VcKxGM5xivxS/4Kr/ABa8Zftgf8FBf+FO6PfeToOjahb6FYafLK0Nvd38saSPcTFSCxVXIAyBlQec4rrweFdeTXY+o4R4aedY90J1OSnCLlKVtktWz9Mx/wAFS/2f7a4eNvip4TO04bbclgD9QMUlz/wVP/Z9ZP8Akq3hNepy1yVHAPqtfEek/wDBt3BeWEMlx8WLlbpkxKsehwNGCOu0sScfVjV63/4NtbOKT/kr2qQt1BTQbYk/+PCuuWFwy0U9T6z+w+B4T/e5jOy0doa3+4/SL4TfGTwz8cPBsHiLwjrNlr+g3TFYb60ffDIVOGAPseK6y2uPPzgcL3z1rxf9iP8AZRt/2MP2fdL8AW+s3HiK302WWX7ZNAtu0hkcsRsUkDGexr2TTiMuF24HYdq82tFRlaJ+Z42NKGJlChK9O/ut6Nr0LVFFFZnOFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUANk+7Ve5OIW+mKsSfdqvcjMTfShboP80fzf8A7bXH7YXxL/2deuR/49XmANen/ttuF/bG+Jm7dg+ILgcDkkvgD8SQKtfBn9hz4nftCfCnXvGHg3w5Jrel+H7g2siRy7Jb6RRmRbYEYl2dDjHP5V+gQqRp0oymz/RDJc3weX5JhKmOqKEXCKTbtqeTMM0gBHpU2pabdaLqtxY3tvcWOoWp/f2dxEYprfHVXB6H6Ag4qEMsihlO5W5B6j8+9bRkpK8T6rD4qlXjz4eXMvJi80hHNKB81J1dRy2c/KvLHFBtGp9p3CnLE0i/KsjtjO0DnHqPUe9XvBnhDVviH4kt9G0HS9Q1jVr1tlvZ2cJmldj0yB0HqSRX6OfsYf8ABA2+1uex13403htLIlZIvDWnSbnZhzm4uMHCkf8ALNAvuc5rGvjaVFas+L4q48ynIabni6q5ukN5P0/4J8F/s/fs2eOP2pPFVvpXgfQbjVWkkCTXp+WytR3Ly9OP7o5r9Vf2Kv8Aghh4O+Dq2+vfEhofHHiRMTJaumNNsyBx8h/1jDszdK+1/hZ8FPCfwY8J2ei+FtB0/Q9Ks0CRW1tAI0AAwMjufc5NdWsGwKMs23vn/Ir5nHZzKvpDQ/lLjbxgzPOpewwl6VHst36ve/oVNM0K20i0jgtYIYLeEBY40jCrGuOg9K0vswX7vy08DjpS14123dn5BL3neWrIxbAHrTtnzU6ikAUUUUAFFFFABRRRQAUUUUAFVyMzE1Yqu7bZTQTLp6n4xf8ABxOAv7YHg0Fvu+GSW9h9qJ/pX1D8Pv8AgvB8DfDvgvRLG9k8XC50+wht5gmjMyhkRQcHPqK+Xv8Ag4lG79sHwf8A7Xhgg/8AgWR/WvpzwR/wQN+A/ivwRoupTL4u+0ahZwXc23XJFUs6AkAY6ZPSvoJxoPDQ9vsf0RjP7AfDGWrOudPllyuFu+t/wseNft9/8FufA/xj/Z81jwh8P7LVru48R2zWlxfajbiKKzjYZcgAkliAcDpXp3/Bvd+zhqXwv+CniXxpqlvcWp8cXMK6fHMpVpLOCNwspU/d3NIxx9K9Z+F3/BD34B/DTxRb6x/wj+qa5e2jrJD/AGzqcl6kbL0IVsA/jmvrHStDtdHso7W1hW3t49qxxxgKqAdAABgCuXEY6jGh7DD7M+EzzifKKWTPI8kpyjGUlKUpbuxetpOQuOlWKgii8th9anryVsj800vZbDJxujrG8W2+fC+pEMw/0SQf+OGtmf7n41keMW8vwlqJAz/okn/oBrWj8a9TTD/xU1vdfmfjP/wb+MqftweIif8AoAXWP/AoV+0lsdu1j0ZR/WvxE/4IX+P9H+HH7Y3iTUte1Cx0mxOiXUKz3c6xIX+0AgfMRnp2r9dj+1f8NolXd458MYIwg/tGHsT/ALXpivVzSlKdW9nsvyP1TxYyvFzz9ulSk1yw1S0+FHpyORT0bcK5XwD8WvDvxLjuG8O67peuLZkC4+x3CTeVu6AlSQCfeumspGlh3Mu0tg49OB/WvGaadmflMqc4PlqJp9nuTUUUUiSGS4ZXwAK8/wD2iPjrpv7PPwb8QeNdXdVtPD1jNcsmcGRhkKo92YbQPeu7nl2zHJwucYxkk1+Tf/BwV+2Ada8Q6P8ABjR7qT7PAV1bxD5JyWYnFtbkDnriTHXLL24PXg8O6tVLofUcHcOzznNqWCXwN3m/5Yrdnz1+xl+zhrX/AAVG/as8Za34kmubfTpo57/VLiJ/nSeTItYgP9ltrfSP3rsv+CS3x21b9hr9t3V/hz41DafD4hvB4e1VekcGoQExwTgZ+64O3ceqsDXmf7NX7V/x+/Y48MzaX4D8PyaRZ3s7XVy174UmvZpnIwNzlgcIOFxjGT1rzn9oHx/8SPjJ8Sbj4g+ONFvdM124EIn1GHQZtJi8xf8AVSSNswWACjIx90fj9M6U6i9jJrlP6ixWU1Mwq4nKcRVpfU3FRpWmrqa2aXd/5H9JEV9vRfunkjjsR/8AWqb7QTt4+8a+a/8AgmV+1kv7XX7L2ja5dso13T8afq8f8Ud1FhCSMkjeuHH+9X0jD+8Zf9mvlKtGVKXJLc/kPM8vrYLGTwlZWcNH6lmmyHC06mzcpWJwt21K08YZOv3jjp+OPx6V+P8A/wAFfv2K/Hvwn/aom+N3gWz1S8sdQkhv7iTT4jPcaXexhV8woOXjYKMgd/av181DK2zMM/Kd2FGWOATxyP8AP51594k/aD8A6HrN1pWteLvDdreWr4mtpr+KN4m7AgsCDtKnrXZl9SdObnHU+s4O4gxmU5i6+EpOopR5ZJK94vdbM/J21/4L5fGrTNMhs7jw/wCF5bi3jKSSvZzxO7Lj7ykfKenA6V2XwO/4OGvE1j4oht/H/hPTH0eRlE8+mlo7i1U9XKN94DrxX6MTfHP4OyFnbxR4FYnlma+tlYnseM5r8+P+C8fxJ+EPjLwZ4Ph8O32g6r4yhvt8cumskgjtQP3oldeFU8cHqa9ajGnWlyTp2T6n6ZksspznFwy+plXslNu87S93z1SP1F+HfjjTfiP4K03XdIvI7/TdSgW4trlDlZkYBgQfocH3BrprcYjr5J/4IsWeqQ/8E4fASar9o3Mty9sZhhvIM7mPj0xwK+trY/J9OBXh1qahJpH4tnGCpYTH1qFKXMoykk+6T0ZJTWiy2c06isTgK726nd2PrWbrloV0i4yzN+7cc1sGPPrWfrp/4lNxn+4/8qqlpOKXcujFe0ivNP53PxM/4IbuH/4KY3WUU7tK1nrzj96tft2tnhvvcZ6Yr8Q/+CGn/KTGf/sFaz/6NWv3C7/jXpZx/HXoj9N8Wtc5hJ7ulT/Jh9mA244wc1JRRXln5f0sNl/1bfSqr/IB7j+lWpf9W30qrN0X8KL2TaFLW3r+h+Nv/Bwedv7bPgVf+oBGc/8Ab6xr9d/hbB9m8A6Kud2LC3PPvEpr8hv+DhH/AJPc8C/9gBB+V4f8a/X34bn/AIoPRf8Arwt//RS17WYyf1ejHpY/VuNFbhvKF/ck/wATfg6tUlRwfxVJXiI/KY7FPUJ2SYAD0r8w/wDg5H8TX0Pgv4WaLFLt0+/1C9vJ0HV5IYVCke4SSXHuc9q/UCaBZZuc9K+Mf+C037IWqftNfsw/bNAs473xD4JnbWbSLB8y5iCFZoU5+8V5x/FjA55r0cvlD2qUtD7Hw9x1DCcRYWti3aCcrt9Lqy+46z/gkt8NtI8Ef8E/vhy2m2sMY1jSxqdyVADPJPkvuP8AEVGF55+WvpiO2DRY+Urjtzn05r8jP+CSH/BV3Rfgd4Qs/hZ8Sb6TSdH0+Qto2q3UTIlokrEm1mGSU2sTgkfLnac4yf0X8Zft6/Cb4feEn1rUPH3hP7FHGXDR6hG7SDHRFUlmb/ZwDWmOwlR1rpXV9Dq4x4ZzOlm9RThKftJOUWldNSd0/mmevJBsTaM7c5571P5ee/vmvm79jn/gpr8PP23NX1nS/Ct1Pa6tpMjlbTUY/s819bjpdQrk74ieDg7lPUV9HxSHyx0zjmvPqU3CfLLc+LxmDxGGrywuKjyyW5bFI/3TS0j/AHTWZi9itc/6pvmC8YBPqeBX4/8A/BXv9ibx78I/2n2+N/gDT9QvrO+u4b64ksYTPPo+oQoiLKUGSyNsUZ9WOeK/YC7bZbyMSRx1AzXnviL9oTwD4e1WbSdW8WeH7O9t2MU9tNqUUbxsOocFwe46+tdmX4mrSm3TWj39D6jhPPMXleOVbC0vaK1pLvF7rqfkzp3/AAXw+N2iaXBay+GfD00kKBGllsZlkdvfoFJ9K7P4G/8ABw74ntvGNvb/ABC8I6a+kSP+/utLZo7q0TIyxR/vKOc49q/RX/hcvwbaV3fxJ8PwW43NdWn+PNfnv/wXa+I/wn8W+A/CNv4avvDuq+OVvlkjl0pklaKwCSCbzWQbAmSMg5Pp616tFUa0uWULX6n6tk9XJM5xqy95SqXtL+9eXuu3xbLY/Ur4deP9N+KfgzTfEGi38OpaRq9utzazxNuV0YZH49j710FjD5TvzkHHavkn/givp2qWf/BP3wV/aaTRBhM1or8bbfzW2YH90jkexr66ts+XXi4imqc3GOx+I5vgKeCx9bCU5cyhJpS7ruSUUUVgeeFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAMuG2xVSllaS1bAy3Tj0q5cf6qqTP5UUf3fmPQ/xc/zoW6JvaafTV/dY/Ajxr+zbr37Wn/BS7xp4L0OOS3fUvEN019dr8w0+1Vx5kp9+VUe7g9q/cf4DfBLQPgP8JtB8K6BZR2Ok6FYrb28aDGe7OT3Z2+Zj1JNebfslfsU6X+zp4/8e+K5Ghv/ABR491mfU7q8Cf6iBn/dwJ2AXHzYAycZzxXv0NrGi5+bPP8AEe/J/wA9q9XG45VrRjskffcbca1M5jh8HDSjRivS9tT5z/bJ/wCCafwz/bC0eSXXtL/s3XIxuh1iwCxXUTcYJPRl9Qc5r8oP2rv+CRHxe/Zt1a6u9O0e58ceG929NQ0yHdOFPaS3++D/ALvBwa/ee8TbNGypuYZAPdabJDnC+X97jKmowuZ1aLs3dG3CHilnmRT5MPP2kP5JbJf3WfzKRfC3xTc30lrD4T8VyXanb5CaTN5m702kcN7Gvq79kD/gin8Svj5eQah40t5/APhnaHf7SA2o3MZ7Ki8R56Zb3r9ul0iCB2KwxeY/8YQbh+PWlS2+zfdVeTuyAM56Z+vvXZiM+m17iPss68fs3xWH9jhKKo3+1fmaPJf2U/2Fvh1+yZ4TjsPCGg21rcMAJtRZd15cY7tIefy4r2ZNNhR9wXnp1p1gMW/TGSSeOpqavFqVpVHzS6n4pjMdiMbUeIxc3OctW31ZGbdSe9O8ladRWZzeQUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAVVc5aQ+lWqhZFDN7nn3oJlt59D4K/4KZf8ABKnXv28PjdoXijSfGGj6DDpemnS2trm0eeSRhN5m4FWAH419t/D7w5/wi/gzSdPZ/MksLOK1dgMBmRQM+2cZq/NbqWb93uXpjPB/DpV6OFVT7uM8/jXRUxEpQUL6I9bE51j8ThaWDxMrwp35V2vuJ9nU9eaBbKDUlFc+2x5a0VkNEIBHXinUUUB5DJ/ufjWV4qtnvdAurdTtN0hh3dkDAjJ5HArUuG2x1XkDv05X09aqGj5iqd1JSjo19x+Qp/4NxfGmoXEgf4keHceY8o36VMeHZjjh8HginJ/wbYeMI1P/ABcrwufc6RK3t3f2r9dY4FDNuX73POaeEVPuqozXpf2tidk9D9D/AOIr8S8qUquq/up/i9T5E/4JYf8ABNrWP+CfOm+MYdS8Tad4hk8UXFtJH9ktGt1hEYwchic59q+wbJt8G4cgs2Cfqajjh8wdB8vcVPAmxT9a8+tUlUm5z3Z8PmGZV8fiJYzEu85avS3lsh9FFFZnEZOsNMyyi3ZVkyVU4J2ZU8nHPXHSvzn8Af8ABFPxJqH7ZLfFb4jeN9L8URT6zLrN1ZRafJC08m9jDFl2xsRSq7cYIQV+j09iLudllVZFDEruGcfSpAAvy/3Rjr0ropYidP4D2Mpz7H5Z7VYSXLzxs2t7PoNhtoksFTbtXAG30rzv9qH4A6R+0v8AA3xB4J1pMWesWpt1cjIt5Bkh+PQ4P4V6ULbKY7de9Atio79MYyeR/nvWSqST5kefhsXWoVY16MmpRd0+z7nwz/wTF/4JseO/2CviLrV5fePdI1/w3r1nHbz2EdjJE7XCN8k4csVzs+U4HOB6V9yRfI6+5qNbfMmdtWAi7l+tXXxEq0ueS1OnNc2xWZYqWLxrTlLdrr6k1Nk+7+NOpHG4VjvoecUtXiaSLapP0HVuxH5E1+Yf7U//AAQm8TftI/tLeL/G1n470PT7fxFdi7jtbjSpHaAbFTBYMAfuj9a/USeHzR3B9jioWgVDu2rz1OOTXVhsU6KtA9rI+JMblFZ1sG+WXe1/wZ+Q8f8AwbY+MM/N8SPDHTH/ACCZf/i67X4Jf8G6ln4Z8Zw3fj7xlH4g0mN0ZtP0yxMAvNpztlZiTtz2FfqGqx/3UpQqK25VQN0yBXRUzOu1Zs+or+KXEdWlKjKs0pb2ik/vRneEPB1j4L8PWml6daxWNjZRLDBbRcRwIoACgemK2YBhaZbRLtOBt5qRF2ivPlLmd+p+fyqTm3Ke7HUUUVJIVm6khvbaSBjtaVWC1pVReAmZjt55Gc0czTugXOneFrpo/P3/AIJ+f8EjfEX7G/7V1x8QdS8ZaXrVpJa3lolpDYvC6m5cMDlmOcYr9Codznn1qJ7WNh93b2+U7f5VbUYFb4jEOq1KW56mb51iszr+3xTu0lFPyQUUUVgeWNmPyVXnX5PyxzViUZWomXdt6/LQKV90fB3/AAUp/wCCVXiD9uX4/wDh/wAXaf4u0/w/b6Jp4sJLebT5JmlIl83cCDjHUfjX2x4H0+bRfDdlZzyLIbW2igDBdu8qgUnHuRWq1lE//LNNw745/On+Vg565966JYhzgoS6HrY7OsXi8NSwtSV401aKfRPexJAc5qSo4c81JXPseUNZNxqve6bHOMtu47Z69x+R5q1Qy7hzR1uB8K/tjf8ABEv4c/tL+Lr7xFo9/deCvEl43mzyWkKyW15KeskkZ6se5FfOPh//AINvtasdVje++JWjx2KyFiLLSdl0yn08xsZ+tfrZPAqysefpuOPypsVqjjbtAXoApKj9OK7I5lVjonsfbZf4j8RYKj9Wp4huOyvFOy6K77Hz1+xJ/wAE2fhz+xPpLXHh+wur7XryDyr3WdQw1zKOuxeyLuycL+tfRy2y+/IqK0sFgmZtvXjrVquWdSU5c0tz5XMMwxGNrSxGLm5zlu3uwpshwh+lOpspxG30qTj2Kt/u+xPtJB46fXmvy/8A2p/+CDviv9oX9pLxb40tfHPh/TrfxZfyXogn0ySV4AwXgsGx1Rf1r9RJd3lgrioViAk3Mv1rqw+IlR+Hc9zI+IsflFb6xl7Sk9NUm18mfkQv/Btf4xTp8SvCw3cZXSJMj/x7j612fwL/AODdK38K+MLW88deMrfW9Ht5hLNp2lWLW63eP4WdmJ2nnIHWv1I2w/3T+tJ5EcjZVfmxgseuPrXRLNMRKPLJ6H1VbxT4lnCUZVrcys7QinZ+a1RleD/C1l4P8P2Ol6dZw2Gn2ECQ21vCNscEajaqAdsAVuQ8VGsO1v6VLF1NcMpXPzf2kpycpj6KKKzKCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAAjcKjezjkTay5XpUlFAEYtY1bIUA4xxThCoHSnUUBZWsRy2ccxG5c46Uot1Axin0UA9VZkf2ZC2dvNK1urnJFPoouKKS2EVdgwKWiigYUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFIUBPSlooAb5antTqKKACiiigAooooAbJEsy7WGRR5K4xinUUBurMbsX0o8pT2p1FACKgTpS0UUAFFFFACbBnpTfs6+nWn0UXYeQAYFB5oooAaIlXtR5YzTqKAWmwUUUUAGMGmmNW7U6igBvkr6UGFT2p1FLlQDVQJ0pwGKKKYBRRRQAU3YvpTqKAGiJQelOoooDbYKKKKAAjNNMamnUUANESijy19KdRQAioEHFLRRQAUUUUANMKsc460CFQOlOooDyAcUUUUAFDLuXB70UUAN8lSu3HFAhUDpTqKBcqvcb5K/3aBEq9BTqKLIdkN2CnBcUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB//9k="
ACI_B64 = "iVBORw0KGgoAAAANSUhEUgAAASwAAAErCAYAAABkeL7NAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAHExSURBVHhe7V0HfFRV9n40e+9lddW1r2XXsvbV1f/qqrurSFOSmQlViggiiEgTsdBEUEQQkJSp6ZWEEEIPEAgEEhJCgIRUSO89nP85991hY3hJZpKZyUxyv9/vM5jMvHffffd895xbzpUEBH4Hb7fLJG/3GyWD+h5Jr35Y8nF7StK7/0MyaEZKWvcJks59iaRTr5O0qtWSTuUt6VV78f8PIg90wAS8Zjx+xxe/8wv+/zr5WqrxklE9Fv/2L/leeE+6N5WByiIgINDLYRrYV9JpHmBCpFW5o0jMkoyaH1FI/JDbUESS8HeF+LNeChwFUuBIkPw5/UaA5OvxPxo11rHld+la5uvSPehedE/53lgGLAuViZUNy6hXjZCMqtekDVh2egYBAYEehBWTLpJ0H1wvaYc/IhlUQ1EEFqPhh0p6zX40/gz8dzH+bJKCR4MUNhbYTxINEhCzGBnUMvV2pvk+ZlFjIoZlaVk2+rtOVYrEstMz4LPQM+nch+FzPMm8MtOgS/jTCwgIODVMo65lhqtXD8IQbg6GXf7ITPxdLf6umQkBeTJBJErkKSFN+DtHCFJX+Tsxw3LTM9Cz0P/r8NkMmjr8mYvCHIaf/xY9RxRo9dNMsAUEBJwAsfP6YWj3BHoZ41GUfkGj3oH/LpICzKLEDZpEiXlLSCUxcGXSM9Gz0TO2FDOiXlWCn9klGT1+w3qZhnxWhJICAo7CmjGXSroRd6OBvoeGuBIN8iCKVD4z2FAMmchIyWhdwWNyBKkuKKwMGyOPlxnUBShahzCEXId/H4Qh5F1iYF9AwJYgg9Jqnkdx+hANLgQNrZB5E2SAFBLRTxIsJYMV/B+pjlrWGdWhDutSrw5HTpG06pdYhyAgINAJGNyfQ0OajyIVgz9Lz4/ZnA/vFIxS0DKax8OoLil8DkHvVK+qwL9tRn4r6T3+T5qrvpi/CQEBgQtAnpTR/THJoJqNYUucZFQXsfEYmh2j0EbJ8ARtR6pjqmsSMIO6HN/BPvYujG5/lXyGX8nfkoBAL4eP2/0YmoyVDBoK9ypZuEKD5fSzJw6SOzupzs2D9/QOdOpGSa+JkvTuE1DAHuRvTUCgF8F37gDsxf+JBvIbhibHz8/qUYgiQj3noTl0pFCcvC+j5gQKmBf+7V1pgwgZBXo6aIaPtrjQdhajppKJFAtBhCflEqR3FYzvzKiuQeGKl4we4yVv9z/xtysg0EPg7f4c9tYLsdFnslCDKMalXJfkCZvXexnUORgqfs9mGQUEXBbzXuon6VSvYwP3Qw/qrBR6fi2QshEIuibpndIso0FThJ5zIHvnvoMH8FYgIODkWDP4Kkmr+q+k0wSjONVIQdiYqVErNXbBnkN6x+RxUbioV2+XjKrBrC0ICDglVrxxkaRzH449bTQL92i8g0IHpcYt2HNpXttFbYDaArUJahsCAk4BaowG1QcYCmCvqm5mg7JiEF2Q2oA8QN/M2ga1kTX/FivpBboJq8f0x4b4GoZ94ZJJ03R+FbpS4xXsvaQ2QW3DqG5CEduGv3uXjW8KCDgMJtWzkl6lR7GShUrs5RPsiOb1XHpaiKr2l/Qe/+CtSUDATmDrqDxWYK9ZcH6hp1LjFBRsiyRcNKto8ihBr2u1WMclYHt4uV+HHtVH2NgyWLoSMesn2FWaZxVpHZfefbrk43YLb20CAl2ATvUmNrCd5/eWiXVUgrYitSVqU9QJmjT7JP2IQbzVCQhYCU/1PZLRY5Xk6yFvoREzf4L2IgkXEy2PWvz3BslLbLIWsBTz5vWTtO4jUayOs0YkttAIOorU1qjNGWm7j3qiWL8l0D60wx/HEDCADabT4j+lRiUoaG+ahx+0qhDJ2/1vvHUKCHDQ0VBa90mSTp3P04goNyRBQUeR2iANRejURZLB/VNxfJmADJZAD70qc6+m1HgEBbuL1CapbRrUgdhW/8pbrUCvhH6EG/Zk8liV8KoEnZXUNmXPP0vSe4zgrVeg14AO16QZQJNHA199LCjo/AxAT8vk0cTarli31UtAp9GYNLvE3j9BlyS1WXmXRYLkpfo7b9UCPRJ02INJc4a512IBqKArk0TLqCnC9jxJmi314S1coEdAN/4aiY5yN4rlCoI9iHS6ki+GiXQU/9phN/HWLuDS0A5/BF9uNF9JrPziBQVdldQJU1Zbg2aP5OP2FG/1Ai4Jnfu/Jb0mnW8yVX7hgoI9gfIwR5akUw3mrV/AZWAa2FfyVk9Esapgg+tCrAR7A2kW0aCpk7zdpzMbEHABsNNq1Esw/GsQC0EFex1pPyIb+lD9KA7BcHbQ+iqD2od5VWLTsmBvpXnpg0FjlPQjb+XWIeBUMKjvQW4UewEFBZHmvYgG9T7JoLqPW4mAU4D2WOnUCeLEGkHBFiRboGjDqDkqcsg7C9jJNZoMtr5KDK4LCl5IedlDLrMVgW6EVv2KZNJksTVWSi9KUFBQpny4a5ak07zDrUfAodB5DJaMHnli5bqgoIWk8NDPo0ryUam5FQk4BN5qDwz/aoRYCQpaSTm3Vi2zIQEHQOuuQde2klW80gsRFBRsn2Q7ZENkSwJ2hB5dWV90aYVYCQp2jbJoVTGbErADKO42aGqFWAkK2oiyLdVKOhEe2ha0hsSgKWXpNJQqXlBQsHOk7Wu+IyolnWY4tzaBLkHv8X+SQZ0vPCtBQTtRnrwqZbYm0AUYNc+jZ5UjZgMFBe1McgiMaGticWknoVc/LOlVyWJRqKCgg0i2plNnSSaPR7gVClgEb/e70LPaKxLvCQo6mBTNGDQJ7LxOAQvg7XYZuqXB4qAIQcFuINkcOQo6VZRkGnQDt0oBRawe0x9d0lU8l49yhQoKCtqXJFrkMOhVnswmBdqATjWbLV0Qh0UICnYvTegwUHioUy3k1inwO+jcVZJR0yiJTKG9hCoyBk73dsg/Q59XvI6g3Ui2aNQ0SVrNh9xKBRh0qhewQRaItVY9mCQ6WjeQfN4HyXsY/vwA/3/470Spj14DfQ0a/EmfN/8ef9Ln6PP0PaL5e63vIWh7yucilEpeYrmDDB+329H9PMTGrZQqTNC1SeLiOYgJzRXGUfCn4E/gyeiv4Z2dq2DBQV/wPh4L4afjIeBUHBwsOgXpFWdgT0Ea/v9u9vu1qZtgboIBhu5eA09HL4D7Qz+FG/zG4XVRsDa8h9cdKsTL3qTlDkb1UZFqOWTMpdiLBkghYkawR5EEhDwpryFwvd94GBi7FOYeCYKI0wegsLoUuoKKuirYl58Ky1Kj4cN4T3hh41xsO+gF4L3YPXXoxdkjfCTPjoTXHiRRR+9S8b7OwJYzh75jevFJPHr3uXzdh3JFdRfNBudosrEahfK4Csnw0LD74b8fQyH5KSUK4gtPwLmmJi43tgcJYNSZFPg+KQSeivoSLjKieFH4yOrTBp4XD1fHHdCCX1a8XfhKzLcouOgpKt3fWUg2Suce6lSLufX2MhhUb2MlVDvdIDs2zmsx5HggdBrcF/Ip3Bsy1e6k+9yNoZJkoLpwRdHCMqOHcxF6Cf9Fb2pj1gEoqa/mkuI4lDXUQAIK5Ef7PDF0nCaLFnlGXalTEj1sE/qsBH4X22Psvt9A+m2g8v2difIgfK3k4zaEW3EvgXwk1wmn3HaDXsLdQVNgD4YdeZWFkF1x1u7MqyiAU+X58GrMQtk7UCqXs5K8KhSGFzd9Bb6Z8dwEux9nqopgIXp4f42YJYsW1WtnPFguWGtPxfEr2x6aPb/K43FK93c2yvt6M5EPc2vu4Vj3nyuwAcgr2ZUqxBnoMwzmHPLjzclx8DweKw8iu4KXRcaPZb3O90NYnBQGhbUV/CmcC9nVxbAmLQbuQk+WCRcbL1J4nrYoBOtCsvEs9SZpg/pqbtU9GDrVHKcct2pJbNRXm8bA7rPHeZNyDPYUnYJLDLRYjwaOFcrlLKTyeQ2Dv0UvwDpK46V3buSgxzUt3hsuonZHoaLScylRCNaFZONZbFHpPG7VPRRG9UsYClbytR3OTa+h8Na25VDfWM+blf1R1VALb8Uskme8lMrkDNTS+ik3mLLPCwprynjJXQemrP1wY+Ak5kUrPl9rCsFSJtmwUVMleane5Nbdw2BS3yYZVPvZUUNKFeBsRKO8GBvq+pM7eLNyDOYlh4HkSYLlhGGhzwfQz+ABy45GAJw7x0vseqBZxRsDP2bvWPE5W1IIVttktqw5IulG3M2tvAdBr14qb2pWeHBnJTaiTw8aebNyDPZiGHoTLYx0tsF3DKMu0WtgXfpWXlLXxrLUSHkNlNKztqQQrPZJayh16nXcynsI9G5vYThY41KbmrH3vcIwAsJzDvNm5Rg0NTfB05u/4YPvCuXqDmqHwwCDBtanb+OldH3oM+MsEwkhWO2TbNqgaZB81IO4tbs4TCNuQK8qwWVCQTO9hsAzm+ZDdZ3j1xN9mRgIfXyGM0NRLJsjiWJFC0HXpMfy0vUMeGfsEoJlK8pLHY5LXqo7uNW7KOa91E/Sqn5w+lnBC0hT9sNg/AEdb1KOxdGSLOjvO8ayMRZ7UucG/TAU/PqQPy9Zz4EQLBuSbDuYljq4r5JMA/ty63dBUEJ7OmHW1wVmBVsSheJS02jYnXeUNynHorqpAf6ycV73jmORd4eiPWr3GofOlLbGmaoSSC46BRGZ8RCasRf2nzkGmeX50NDUyD/ROQjBsjFpFbxJUynpPFx01tB38FX4one5ZBYG9Cruj5gF9c322wPXLs6dg/mH/Lp3eQOK1Z8j50F+Tdc2K1uLszVlsDFrP3y0dwP8GUPyu4KnwC0BE1mmh8uNI+FGv/FwR9DH8EDYdHh96zJYkRwKB4pOQkV9Db+CZRCCZQeyYR/VQUk39nquAi4EvWYqyx7qUqEgJxrroiPBvDl1D7afTYNLaQc/GYxSGe1Jn+Fwg+9Y2JGfwktjf5ysKoSvDgfAY6Ez2LgZ8y5pZTot8qR/U3jM2Ppv8sr1NzZ/C7+kx8LpqiJ+xfYhBMsOJFsnm9e5zeYq4CLwdv8zPkCeSybk07pDH8MIiC/K4M2pe1BSUw4vRn2JhulgL4sMFMXgq6RQXhL7Q3diOzwU9hlInoNlEaJw1NJ1aPRZVuahbLHvo+EzYGFqJNQ3tR/GCsGyE2lBqUFdgO/vSa4GTg4aaKfk9a6akA8b/b9iFkKllSGGPTAj0Vc2YqVy2ov4/C9FzYeqhjpeCvuhrL4aRuz7DQZQMj6bCDONu6FwoQf2evTXEF94kt/pQgjBsiPZ2iyVrxQ7rx9XBSeGXvOW5DuihiWxV3oYpyY2eBSIBcmO8y7aw5acw3ClEV8+hUGK5bUx6T7GkbA93/6TDWdqSuH/tiyWBdnWeydJZFB4r/OfAL+d3Mnv+HsIwbIjyfZNHg1Yd+9xVXBSrGEZRLe7rHeFPfONfh9CnIM3PbeF2sY6eDxyLgvRFMtra6J3MnLPen53+6G2qQHe3PY9igp5jxaGfp0hhpd9UHC+PRIE0NzM7y5DCJadKW+OPiAFj7iCq4MTwtt9HE/ypfwQzk6vIfBG7BKn2ic3fb+WDyzb0bCJKNbX+0+E5NIcfmc7Aev244NGWaxIDJTKYkuyQfr3YV7i79MFCcGyM9kA/Aisf/VnXB2cDD5ut+CLPe5yK9rNxIbZV+sOK45G8WbkHNiSe5gZjH2NG6+PYj05Qc/vaj9EZSfAVZRZ1VFhLhHv1Qfr8Mdj0bwUQrAcQvKytKoTkmHYH7hKOBG07rNYAV3Vu8JGfbXfOMgoz+fNyDlQXF8F97GpfjsuIkUP5NbASXC0NJvf1T4owWd5LHKePDCuVA57Ep9xAAplRF4SK4sQLAfwf4dXOFkeeJPbQ5LBI9ulzxXEBv3WlqXQZEU4WNfUCDsKjkNUzkE4UJLJf2tjYHE+O4Bhod1mC9G7wmef6IBtSN9S6hxaQ0Ueo2JZ7EwUyicjZkFdYx34no4XguUIynmzciSD6kGuFk4ArXqRU6c8toRotL8dty51ysHiDOijReNb+xY8HjkHGpoa+F9si4jcROhP97FHWKh1g76+YyDVzmNXZ6tL4I7Q6Y6bQGiLGPp+g8LpR4Il0ss4hqQNBvUyrhbdDMOYeySdutDpTr+xhhgO3hUwGU6Un+FNyDJ8lRgghzfe78Mt/uNhn53SBZ9BY5cPU7BDKIXXVO1eDc12nmiYlxQK/ZhY2UF0rSGG1veEfQafJugsE08hWF2nnFaqHHXiaa4a3QiXzMbQitjTqslordg7SFPz94Z/wTwzFuJ4DYVFqZv4X22PCfu97BAWqtBzc4fw7EP8LvZBdlUxPBQ20z6CazXxXWFbJa/SoplXIVhdJ2kDnZClU3lLwHWjW2AaeT8WIo8JllJBXYHYIPv7DIf1J6xLhRyXnwKXm/AlmGe7UPQ0u1ZBo53CwpDTB+AyI8262HCRJQrIi1FfQXldJb+LfWDK3Ne9Y1cX0IrwWgiWbejPvKwyyeDxHFePbgCN/rv62BWGBX8M/gQyyvJ487EMI/Z5Mq/qfC+Nocb9QVMhs7KQf8K2KK4pgzuDp8oeXetn6AzRCPtgKDvXzpu8G9BrHbRtuZN4V52gECzbMZRyvHXXWJaP20OSUZMr+bnw2BURDWnorp9507EMxdWl8HgEhoMtjZC8B/QiwnPsF14N3vkz3tNGgoVlvcV/AiTZeZM3ZU+42H/c/zxRV6MQLNuRtIJmDL26Y8ZQp/nKZbfgmKl1h774cyOGW9bAhJ+/iDbstg4r0FubuteTf8r2CM+Ml0NCS8OZ9ohlfXTTfDh37vdbVmyNtcdi2PYYm5S5OygEy7aUNWMRVxEHQV7VnuES5wu2R+z1b8BwsKi+ijedjkED82Pi1spbS1pfD72W+9Dzqu4gtUlncbq6GG4InGwDb0X2BifaUVzNeGXr97/3RF2NQrBsS9IMnfqMtEFzM1cTB0CnmsbSHrvqqnYz0ZA+3vubVVP6uTWlcEdbooG/u9w0GhILT/BP2xYklh/tWa8sltaQwlc0RC/0fuyJM5WFcF/oNBTHbl571RUKwbItSTNkR2cWVxM7Q/fB9ZJWtY9NUyoVyFVIRoteRpCVY076EztRmGibDB9s/9013aEPhmzfHrFfehrvjDhWblb+1ve3lFROvQb220lYzfDN2AcX6T3wfjZOH+NICsGyPeWw8JC0fsQNXFXsCJ1qsGT0aHSpMwaV6D0MHg/7HLItTKlrxpuxGOK0N1PnORhej13Czhe0B7LQa3ksbAYrv+L9LSGK6sV+4yATQ0x74osjgZatJndmCsGyPeVMDk3Y6Y/gqmInrJh0Eb48X3akj1JBXIkoLNZmJ8iqOAs3BkyUPRylaxJRzO4KmgJpxVn8W7bHoF2ruiYEPsPhbhTrM7UV/Iq2R3NzM1uXJgSrY/Q6wSIyL8s9XDINuoSrix2g0zzBTnB29bEr9DCuNo6CyJxE3mQsw8KUjdCvo1k6/BstRP21jUyXtsCK1GgsBxlSO+Vojyiqb8YshKqGWn5F26OwuhRe2OjA5IP2ohAs+5A0xKipkwweL3J1sQN07ktcNt9VS6IRPRwxC2qtMNim5kYMB3lKX6VrtiR6FVP2e9stEWAuhoXXBX7cvqfXHn2Ggdv25dDYaJ9V+YTdxRlwse+HrHNQLIOrUAiW/ShnJf2Jq4uNQYPtelWGS6eQMRMN/aN4H95cLMOewhNwg6+FCyDRg3k4bCbk2fFMv392NJbWHr2GwKfxXnbNrBp9NpULahcmB5yBQrDsR3mJQyZbJmVz0ACZQd3k8uEgNsB+hhGw90wqby6W4dukEMuXE1Ajx5Btx9lj/Nu2h+HEDhQEFE80JsUytEcUrLmHjPxK9oEPlY9tXVK4vytRCJb9KIeFzShaE7nK2Ah0Xr5OHcqO71G6sSsRvZIXohZAVaPlx1iV1dfAUzQeY7FHgyLi/T4sPhzIr2B7HK88C1f6jUdj6kTIhYI1+6CBX8k++I6l3nHx8SuiECz7Uj4SLELyHTyAq40NQHlsdKoil983SPQcBEusPMbrQNEpkCgPOTVepWsqEcOhv0V/TQlD7QJKFqiivYUoPor3b48OEKyPaDdAZ0NWZ6IQLPuS8ujpVKXIZ7na2AB69+k9wrvSurFz/nYVWHeM17R4OrmGjM+K8Es7HK72nwCnrMwCYQ1+St8qh13WhoV2Fizan6ghMRWCZRF6tWARSVu06rlcbbqINYOvwovtZmfmK93Mleg5GF7d/A1U1lfzptIxqhtq4b7wz60Pb7Ch09KDFSn2S+p3oOAE3OI3gXlzimVoi3YWrMrGengkcr715XJGCsGyP2VtOSDpPriGq04XoBv5bM8YbJfzPy1MCuHNxDKEZ+7nifM6seEYBZIWedoLDU318M/YJXgfK8NCBwjWQ5FfCsGyEL1esJi2aBrx5zNcdboAnXohW8pAR/Yo3cxVqKWDQifA0WIrTrc5dw7G7lnHhEfxmh3Reyj8Bb2zvErrtv9Yg8VJYdCHzRZaMb7mAMF6WAiWxej1gkXaImvMCq46nQSN3OtVB106BbKZGNK9HLOQNxHLcLqiEB4M/azzs106N7hMr7FrUr9jJdlwhR8t0LTCA0TBmnPQfssahGBZh14vWESmMapUydvtMq4+nYBO9RperNKlT8QhYoMjg155dCNvIpYh8PR+DCPR6Oj7Ste1hJ6DYHaiL7+iffDq5m+tEwf0/Kbvt27hrDUQgmUdhGAh2WyhugpDwze4+nQClBmQ0si4ejiIns5FptGQZsWpzo3NTeC+cxXzRhSvaSl93ofn0XhLrRjotxarUqNQhKyYxfQZBsO3L4d6Ox2YUYGCdR9btyYEyxIIwUKSxsgn63TylGjaRa1TbXL5QyaI6FEM274C6qww0MyqQriMFmZ2dS+c1h36GUdBYqn9sjckl+XA1aaxlpcVBes/GB7X2GnzM52I/X9bFgvBshBCsDhJa3Tq7dK6/1zBVcgK6FVP4kVc+4BUIoVzKFjrTu7izcMyrE/bLAtAV8JBIn1f+wH8lma/7J5V6L29tXkhe07FMrQmludPYTPhrL3Sy5w7BxPj0AjFOiyLIASLU9aaCkk34gWuQlbARzW+R3hXaDR/CpoMx0pzefOwAGhwL2xZhAJgo60lWIa3t/7AL24ffJUUbPlsJgpxf/Qes2rK+LdtjzkHDLarv+6kECzHki1QV03mKmQh5g4egC/J1CNSyaARv7NzpVWZCY4WZcKtAR/ZLqRBwbot8GMorCnhd7A9dpw5BtdYmk1C52b3FMm/Hou23ONzZgrBcixptlCrCpRWvHERVyMLQCda6NWZrh8OurHV5iuPbeZNwzLMTw6HPsw76GI4aCaW42K9BxhOWBeWWgPaW/hC9FeWTRJQmIqGaM9DKCJyD2NZSLBsVIfdRSFYjqWcer1A8lTfxtXIAuhUr7Mvu/rqdvSQbg6aAvlWLNysrKti23e6PDv4O6LReg4BD8pBZUfMPuSPz2xJyhn8O3pin+zX8m/aHgklp+FKXxtMWnQ3hWA5lqQ5lPNdp3qTq5EF0KsXMO/K1ZczYCjmvnsNbxaWIb7oFFxpQLcUvSLFa3aW6G28FDUfKlAQ7YXEwpNwiYlcagvCQvQgn4v5jn/T9iipKYd/RKHH5+phoRAsx5I0R07sZ8WR9nrVFpcfvyIvA72NjVnWner8ZSJ6KfZIPIeexlWmsbDzbBq/k+1R19wIz5BIWDL25jMM/hQyFYXFTgPv5wDG7l3PFs4q3t9VKATL8WTrsdR7uBp1AF/VHfilDJffjoNidXvwJ+y0ZEtRVFcBNwVMAmndv7EBDbQ917wO81AQ7YmFhwMt82rQC6NUOyFWHtNvDVakbbF+n6OzUQiW40nZG3SqHMln5P1cldqBQTUQP1zr8gPu2LNP3reB5WayFGX11Whk0bD8WBRyk+2ZGgGxeUlWzVhai035SShGNIbVkUjQuNpgmJbox79pexwvyYbbA7EDcOX1WEKwHE8TjWOp6/HfblyV2oFO9Q1bf+XK41fYyPrp1GA8Hc+bRO9BQU05PLtxnmVroLyGwL9iFkJ1XQ3/tu3xj63fu/Y4lhCs7mHoGPq5lKtSG5j3Uj/8UKDLZxdFA3li4ywori3nTaJ3YWqCwbKxI50bXGscw1JA2wvBmfH8Xk4UFrKydDSTyikEq3tIGkTnSPjObSfXu8n9LvzQUZcev8LGJXkNg/Hx3rw59D5E5BxiB8VaNFvoNRiWpViXxcIaVDTUwKMRXzhPWEjLLGja3EjH1VkgWkKwuoeUH0unPiYZxtzD1UkBOtULGArWSkYXHr/CBnmJYQRstvJU556EWhSJxyLnWBaK+XwAT2IIWWHHU6B/PbYZ+jjDkV8+w+BK//HwfWoU3BE6wzJBF4LVPZSPAGuUfDSvcnVSgF6ldvn9g2iA94XNgGorjvHqiZiw34fVRYdeBA/VYuyYZLC4tgKei5rPBvkvuL+jiOJ9pXEkhGUnoDjXwR3YRoRgOTnlsHAUV6dWAKRe9YNrr79C4/R+H2YftG/CPFfAnjMp0JeOJrNkpbn3MHh3+wr+TftgY3YiXIWeL1uJr1QGu5GGCIbCbUFTYCPN0CLy6irhttDpQrCcnUyLND9Ls2f3kUWqJTaoL8YXs9Glj6LHxjWANvUWneRNofeiqL4KbgqeaplRoid2vd842F9o33r76ICOdSgden22Iu1W8BoCz0R+CXtaTCzk1JYJwXIFyuNY25XTJgepr8aGdErydWHBQk/hlaiv7Lr9xVXQfK4ZpuzdwEIhxbpqTc8haDjroLmpiV/B9iipr4aXKbEfO/LfzqJF4bDP+zA6bjWcqf59lgwhWC5C0iJaQBqudPyX0f0xVLMGvlvaNek5COYdCeLNwHIkF5+GvfkpsO/MMbtz75lUdq/86lJ+d/shKOeQbLhocIr11ZIYOl6JIeQ2LJs9kVNZCC+yrBIoWpaUy1pSCIxe1S0Bk+DHtBhoar5QgIVguQhpAalO1Sjp3f/GVaoFfNSD8I+ue/4gNr5rjKNhW95R3gwsA+Vavzt8JlyBxnqFabRjiEY15ZD9x9lyqorhMZoNY2GYQp21Jnpjb6MH1GSnXO9mnK4sgOdItGgQ3paiheXv6+MG72xbDiklbaekFoLlIiQt0qma8d8KK9716rnnUzsofdnZ6TUU/rb5G6iut27V9racw9Bfj14lrROi3tkR9HwPXo1eYNcV5maMjvdknqdinbUmGafPcFhuxzxZZpxGT+vtnStlMaW676xw0TgVfd97GPwlYhboTuzscOuTECwXoREpR3wLuEq1gE69jqV1cMUtOdTYsfF9beWpzoTZiQHMSDttMJ0hlvVGvwlw0JpDXTsJPRrapXrshEgolcrSmmj8t2I4leKAslU11sPS5HC42bzfkMbbSCCUytWSTFjx8xj6XYT/f2/457AiJRKyLQyzhWC5EOVF7IFcpThoSYNOvUsKGnXhF1yB2PAu9vsQ0kqzeROwDLQx+p+xi+XQROm6diOKIxqnLnMPL4n9UFZbgQaNYSEZuGJZFIj18d+t30O1nUNDM05XnoXPEvTwKB1aS8JKxkt7IanMJCpEJmhDmHdKHcw9gZPhP9uXQ2DGHqhstG7RqxAsFyLTJNVeKXZeP1msCOHjr8FfJrM8NEpfcnZiY35u0zfQ0Gx5ZgbCseIsuDtwCn7fRrnbrSEK1uQ9663KJtFZDKLzFS0dx2IkQX0fZqD36UjkVxWz040+O6CFf8d8B89EzIFrAj6G24OmwsuR82DYtmXwxUEjrD6xg727zkIIlgtRPhE6RfIedyNXK4SP20OSVnXaNddgoXF5DYVVqZv467cc69E4+nZXviY0lrvQo6D1UvaGEQ2ceS7WPCd+foDeA7yPb+dXcSyqsF7OYoh3tDwfjlecheKaMqi30e4FIVguRKZJqmzJT/0oVyuEQf1PfCnlLrkGCw3rOr9xcMjag0rPnYNpGIZ0W0ZMLDetRD9ox0wJZuShsd9AiQlJtJTK0hbRc70N63Zztv227XQHhGC5EOW8fJWSzqNFjnfaQ0h/dMUlDSg4qp0/Q1NTI3/9lqGktgJeoLxR1ozt2JJoBGQMPyZH8BLZD00YKo9gYWEnNiD7DIPb/cZDTA8SraKGahQsGi8TguX0lDdB47vSjORqhdB5fML27bjakgZsSH29P4BlnThV+WBpNgygFCzWhEm2pvcwluDOEdBn7un80WX4vVv9J8BGO6ZTdiR8M/fCAPQcLTpoRAhW95JWLQSjnRrUn3G1Qug030jhY5W/4MxE7+iuoMlwqiyPv3rL4ZO+TfauyNNRurYjiEJwb+g0KGy1dcQeOFmWD/cFTcVn7uQEA5b1BtMY2JAWy6/oekivOAvT9vvAxWxTuIUdlRCs7qesTYu4WiH0qvWul2WUZrKGwmud9FCG7VjZeeO1FTEkuRK9vMisg7xU9kNzcxO4xa3u2pgd1tclOjVMOqCDqgb7L3q1FRoaG+DnY1vgARq3oi1BlooVUQhW91NOM7OBqxVCp/J1uTxYvCHpT1p/mjLlyrqDGm93jV+1pOdgmHnYMcsHfju+FfqRR9mVMJjGfdDbej1mISQU2n/CoCuoxfe8KScRXt/8Lb5rudyKz9QehWB1P4OYNgVKpoF9JZa6QaeKdbk1WGg4twROhsJO5G3fmXsELjPh81o7a2YPomANRC+RvAB7g0LnmwM+6rpQkxGjd3uj/0RYkrIRKuuq+R2cBOfOQWj2QRhMB2CQQJNQdfZgXCFY3U9ai6VV7ZU2qK+mPFhXs/9xtTzuaDBjdv8KzXRip5WYfsgP+jCj7cbxKzNReG9AETls5Sr9zuKNWDrFxkbPTiE1isHzUV+CKXMfNNnx+DJLQJ6z78nd8K+tS6A/bUfyGiILjlLZLaUQrO4naZNOvU9aRWlm9KNuwv85yg4vVPqwU5J6zffB//Q+/sotB60sf3Pr0u5bf6VEr6EQkp3AS2hf/Ja2RRYa8jyUymItmQdDe/o84J+xSyEy9zA739FhQJE8UpYDP6VEwWMbZ7O9hUyobOU9C8HqfsprsfIkg+aPkqQbcTf+T6FLHZyK3tHDoTMgrxM5pU6V5cO9QZ90PSyyJdFL+Szeh5fQvsipLoHrMZSzaA2SNSSBYGNEGvjn5u9gVfo2yK8s4ne1PbIqzsLqE9vRy/4FrvAfL9+b3mlnQ7+2yAXLK9N+Z1yO3rceBWug8v0FZcHSoUaRVrH/aNUFLiVY2IO+i7154tk0iMs/ajETC9JhzuFA6EteQVdDBVsSxePesM8hJucw7DuTqlh2W5CSB8ZkJ8JTEXPxnnaaIaV6pQWq6MU9GjodhsX9CpuyDkJhXYViMj1LQNlTyxtrIR09qeXJG+HfO36Eh0Km4T1QdMlTZt6UjTzG1mTtxAO9oPXgfSwG1qVGX0Av/P2CI8EwfN8GUMd7KdAbBu5eheH4IngDw9XWvC1oMn8GhfsLcg9LUygZ1PdIks+Ih9DVqmbZ/ZQ+7JRUwSWmUXCV7xg2eG4przCNgX50EIKtwiEb81LjKLjcNFqx7LYgXZs4gOpA4f42JRMu9HqQdOzazejVDtm+An5KjoCNuYmwo/AE7C85zZL5FaDXd7aqmDGtLA/2FmfC1oI0CMUweWFiINvJ8AcUv6vxfbN3R0eG0bUd2On0QfZD9tVdSPq97GGih0de3gWksT4sM4krZQZpTebtutiibUeSFo8GjmySvFRPoof1wdMueTQ9NVxqJNbSmTyr37GTz9MZ2ssbUSTei8I0MkoyXiZiaLzGkXC53wR4LGwGvLBxDjy3cTY8j7wLhe0itgq9hTCZjZ6u0V3vz+yVt0dWr21R4ZqCltG82t0TtUrycXuK5ZxxxcR9gi5KNGAycLOItSSFqiSq5wVA6fuCvYqkTaRRpFVCsAQFBZ2a/xOsv0osLhSCJSgo6KwkbaLkDDrVm5KkUz8tBEtQUNCpydaJqj6UJJN6CFtJKgRLUFDQWXlesAzqca61yl1QULDXkTTKqB4rsf8IwRIUFHRm/k+wVOOFYAkKCjo1zwuWXv0fKUAMugsKCjoxSbBMNIYlljUICgo6O2UPa5RYOCooKOjcJG2idVh6938IwRIUFHRukjaRRnkOf0IIlqCgoHPTLFhsL6GrZmsQFBTsHSRt+l+2hpH3S0ZNqWvlwxIUFOw1pAOefT2aJK3mcUnyHn2XpFOddamMo4KCVpPnpWJ5rdqh+XOK1xDsFsraVCh5UsZR+o+r5XQXFFQiS6RnThb4ge3I8nMJEes2kjadz+nukqfmCAoiSaAo6R+dksNOdHZjx9Bf6zcObgucBH+OnAfvbf8JPt27AX46GgFrj0VDePZBiCtMV+T2s6ngnb4Vfk6OgI/2rIeno7+F24I/gSuMaBskXOzYMMrYqlAWQfvRj9KaqzMlg8cfJHbWl061z+XOJRTsZaSQjWcppRTLdMqMTg33B0+FV6IXwHtxv8Ccw0EQcGInJBWdgsrGWqhvboJzXTgrsaG5GUrrqmBbdiJMTjDifb6GAXrs7W1x3qGg5WTnEqJG6ehcwg3qi/F/NrFpQ6UPCwp2C1GgKBRrmToZReLO4E9Bs2MlfJ8SDqbT+yG9NBsaG+u5xNgXdJ+w7EMwJm4N9Ddir+/ggzB6LelUep1qp7RuxBUSg05jZEsblD4sKOhIkgdF4Z3XULjcNAruCpkK/9q6DFYkhUFcwXHIqCqCpqZGLiHdBPS8NuYkwvNRX6GYomixgz0UnkXQNpS1yU+aPbuPLFgG9RopdMyFHxQUtDdpMJsMnrwVDLVu8hsPr8d8C+MOaCECPahKR54ibSXKayvhk/1aWWSFaNmPISRYGk9ZrAha9XwpfOyFHxQUtAfNg+V0jh8a+uUBE2HY1h/g57QtEH82HRocFOLZBOcA5hwOEKJlT8ratIirFULrPknyGwmSUSweFbQjaaaNDhNFwbo7aAr8E0M904kdcBLDvDpXEqnWOHcOPk8wsDBWjGnZmGyVO3pYOtUMrlYIg+oD/EWTWO0uaFOeD/do6cFguDPgY3DftQqWH9sMeZWF3Np7Bqrra+Dd2KXy+JtYeGo7nneiVGquVggfzauSVlXK1ju0/oKgoLU0j0mhR3Wx7xj4b8xi+DUtBpJKTnPz7pnIRk/x5U0LZE9LqV4EracvrcFSlSNf52qF0Ksfxl+clvyFYAl2gSzkGwyXGEbAXyLnwMxDfpCMIlXTWMdNuucjuSwXbgn+lE0gKNaRoHVki0ZVOZJ2+ONcrRC0vkGnOsLWOyh9SVBQkbSYE70pNoA+BO4NnAKj9qyFgMx41xo4tzF2F52Cu0Omy6IltvR0jSxxnypZCh9/DVcrDlqYJRaPClpENEImUsPgYvSmnoqaD17pW+E4ehcCMlLKz8BrLDwUq+K7RFmTdkvAdeo8dOo1zP0SebEE2yQJlbxe6p6gj8Ft92rYnncUKl0g5CuvrYB0FJFdxRkQezYVYvKTYcPxWFiRHA4/JtM+w80QnX8EthakwaHSbCiqLuPf7Dxyq4phyNZl8ngeLXtQrFPBNklaJO9x9uEq1QI69UzJ6CHnnlH6smDvJQ2isy0yH8Dj4V/AN8mhkFaaxc3SOVFSUwabcxLhu9QomLTPE16P/hr+RGNLpjHsOdjAOAlJSzJvyA2u9p8IL0XOg8n7vcBwKg5q6mv4Va1HFX53xkET9DWgbZHYixlEy0laJM8SzuUq1QJa1dsoWo1iaYPgedL+PTTsS9HYXt78LehP7Yac6hJuis6FuqYGyKwsgF9SouDN7Svg0bAZcDWJEwnThve4l0P7EVF82bhSG2RLMWiDNX7e8z3op1PDExtnw5r07dB8rpnfzXr4Z8bDY6Ez5PIIb8sykljRcisf9SCuUi1gcrsfX1i9ZBJ5sXo9yaA8B8ENfh/CkJ0rYXP2Iahv7ub9e20gqTQHPNO3wtuxS+Fyv/FYdncUG+490XN0ddCbBI6JjArG710PNQ21/M7WI6eyENz3rINLUATlWcQulq2nk2kRapLR/TGuUi1AqRv0qgy27kHpy4I9nGg8FCqhoV+FIdHUeC/YeSaVm5rzYWtuEny0bwP8IXCybPxsmw8JlJ0GuOnaKFyUJ6u2oWtjduFZCfDcpvlymSnUFoPyypTXYOVIplHXcpVqAd/BA7DRBoq8WL2Q3BhvCvgIpu7XQkp5Pjct50JxfTX4ndoDL29ZDNeaRsvbfEhklZ7JHmTe1hBYhx5dV1FYXwVr07bAwxhusjFC8uKEcP2etC5Up4qSTIMu4SrVClrV11KQEKxeQwqXMHy6VK+BEXvWwsmyPG5OzoXi2kpYf2IHPLZxDho3hXw83FN6JnvTewi8HP0tlNkoiwRlo6CtSi9GfQl9Mew0Z07tcijbEyivwfqBq5MCdO7D7ZoXi14C9SLUU9FLYcQeknpJM6kxWkI+a8VIjdd8PXZ9etnihbdJqn8WjgyHpyPnQUTWAW4+zoX6xjr48VgMvBRlDp/wvXe3F0J1Z/CAw2U5vJS2QUVdFQRk7gO37cuhny+fyWTPjD97q+fFtKjlHsLW0Lv/TTJoyro08E6VSwLCFhaim0s9Bs3SkPvOjWQA9uiXGUfCFcZRcLP/RLgVQ5FbAibCHwI/hvvDPoMHwjvmncFT4Gb8Dn33Rv/xcDle63LKvU1iRS+ZxjUojS7dn8rBXjw2tt4sZCTonoOhD9bRkygCq1OjoRANxdmQV10Cy5LD4MEICpewPTnTADVr3x/AoVL77ItsaGpkIflPRzfCE+h13YptnGxGDn+dQLAdRbbESl2L9vwCVycFGIb9AT+Q2P44FjYcqjTmJbUSJjQGyrN9nd94+HPINHgRe++3YxeDR9waGBW/AaYcDoAvk8NBn74DQjP2QhT27MlFGZBWkgUpxZlwuvwMNJ1r4q+ufeRXlbDvHC/JhsSCExCRGY/cByuxR552JBgmJhhg1N618J+tS+EFLMc9NDiLosbGCqiczEsj7wyfgz1PDxUybmD0fi43job3ti2DtRhe1Tc539aZTHz/S49GwP3YdmQPGulsBsrqc7jdBKs10rF9z08KhWE7f4Inwz/HMqAhs/bLO+GeGlWwPO7qo5LJ/S6uTgqIndcPP2xqN/sor5j+5CEFTYY3Nn8H43avgYVHgsDr1C7wz0mEuIJ0OFWayxbvNTU18KrvHpxrboLimlI4WnwaIvKPgvb0PliCDWDsrl/gqcgvWTYB8jjYi2fhJhoJ/duVXz41YNaRyM9ye/BUmH3QCNvOpLDcTc4GGjubvl8Hj5JQmQ3RWesfPb5+plFwpNzx25ByKwogPO8I/JgSCaN2/gyPoAfaH8tyvv1SvbE2THVIgk+Chk4FORfnRU3hmWxGvD67DyeViRHLYe6AqHzkNNBPFvEoXQfJsoyqA6V5L/Xj6tQGdJo5bBP0BSveqVLc4a1tP0IkJf+vOAunUQiqaF1Kc+cX1HUXSMhK6qtYjvAtKLILEwPhxdglcF/Ip3AleiKsUtkAKL7w39WDk5I1WmwYFDpgI7kbO5N/ojcVnLEH8rDjcDqgbqaU5sCsAwa4BcN6uREjnb2jQKN7EoWisKacP0j3oBlDx8K6SjhZWQBbcw/DT0lh8NE+T3gJ2/BfImfDPSGfsHDyatNo6E9eGYWV1IHRMImZFGp3hdTWWl+PCSeJpAquQju6EqOam/wnwF3BU+Du4E9QZGfCXzbOguc3fwP3UU58PelMq3duPpper/qaq1I7MKgG4s1qLzhYlQwCG9XX6LL3ZBRWl6IgJ8BnhwPhPQwnr/WbIL8Is+dFPUfLeukW4rsw92DUi5KxoyHdgWGv+66fYWnqJsjC8MpZseXMMfhwzzq4ynec3NM6s0fVmp7vweforTqjp9oSuZWFcKTwJMRkHQTtyV3wVWoUzEoOhckJWqQPTNzvBe9i+35n6xIk/bSO/0VhVO1exa8nX3PGkUD4Btvel+gBrjoeC9FZCRB1+gAcOJsGp8vzWZtsQkfBjJkYlUmeaFutPT/abaNT1Uta1VCuSu3AV3UHfilDcRwLFXU4xtLNTZaNM7k6mjGcjS86BevTYuDtLUvhUt+xskiYXW6HiRcJFBcnujeJJ77kSzAcuD9sJnwW7wNGDHWTSpx3fx+d8RedfRje370Grvb9ENsSea/4TIrP66REb/sqwwjYSqF1DwAN11AaIDq+zFo2dHHDO917EEYALIppXc+04ZkWjOo0D3BV6gA6Vay8BqLVhdBg7g76BI456Xode6IWK/gE9hIrU6LgpZhF8MfASdiAUbCwx6W1OUxMbCVgbEKDixO53cj+Bg+4A0OnJ9CdHrxrFaxP3QyppdkY1jrvqTKEWmzc0blJ8NqWxXAZGrss9uRRKTy3sxM97f9u/Z6FYwJdQ0pZLtxFE2FKbYE5S6otXI0sgFY9ly1tuCDVDPb0GJ+G5iTy2/ZeHENvZnX6NvjkgBb+u2UR3E8J22hHPomXefDePNBIL0WRLT7DvkMDkvh9jPvvwI7h9ehvYNTu1TA3KRiMGXvhaOEpqEcBcAXUN9SB14mdcp5zajv0bMyjcpHQrzXRu+qP72VXvvNuV3IlmLLi5TbRuj2Q5sjDUQu4GlkAWvtAX1Y6RQcNav5BXzZoKiCjtr4Gva8zsLXgOARn7YefMYafuV8LH+76BYZv/wkei5oPD0bOg4cjvzzP+yPnwr/RmMfhZ6bs/Q0WYzy/7vgWiM5Pgp2FJ7AHyoOK2kp+B9dBcW0FrEXv77lNX8kCRS6/U4z7dYVoVNgBfZag508p0BWcO3cOpu3zwraB3nbrujZrzu9yuHcEb/cb8UsnLxh4J6Jn8AwaX003L1dwCZxrhqbmZqhCr6iyFSuQDZQBoQspS5wJR8vz4JfUTfBI6GfQh2ZWKfRzeaHixDb/AD4XLWgV6DrIHh4Om8nq9YK6ljUnU9qguZmrkQUwDeyLvcpviimTdW7QzziSHS4gILD3bDpMOaCD24OmyD0mNcKeIlRmbnhPeFc2xMHCk9CPhk+U2gmNnetUJmnu4AFcjSyE1l3DF2/9nmy9jxssPBLMb68Mmr6sbKiDvNpyyK4uhj1nUiHqdAL8nLwRfjgSyjhqrye8s2sNDNz96+/4LuMamIOh549JYbDsSAj8cjQSorMPsdxMx8pyIbemFM7UVUJpQw000iCok08z9yTQKTg785LhvZ2r4Gb/CfLEgFJv2RNIRoUe49pTu/nTC3QVcw6asL3Ia7UuqG/aP+ijGs9VyApoPR6RdOocxbMK0d1/fssiFou2hV+PRsHLUQvgqsCP4VLTKBhgGAH9aJEYawA0C4akRk6D022RQguaiaPP4vdopox4CXp4l/mOgRtDpsFfouaDavuP8AmK308pkRCUkwCbzx6DlJIsKKut4KURsAWyqopBn7EH/rX5OxjA2sGwnitUZrL26g6rT+3htSDQFZBmPLf5G3lyqXVdy+FgIUZ3T3IVsgKUH0urjlT0srCR3uo/EQ4XnuLFuBALD9OisMFyg2aCI4uOTFRWpq4d0Pw5RvN3qQFxwSNBo+ubZ9moEoiGkWzK9B+bFoB69yqYkugH69O3Q8KZNKhCj0zAOqQUn4bPEv3h8YhZvDPBuqb3wN5Tq7bRE+k5BL5IDOC1IdAVkGaQdih2dPKR9Jvazn/VEfTuX7IkWkrbdLyGwoJ2Vr0fL8+HiyinNgnL777rALZex4Qi1he9u5v9JsBDYdPh77FLWLhJW4zSKs5AmRNmLOhuFGC4TSG4+46f4XbKGEDjU0rT0E5DLBfrwLCM5vduzhLCdwLIIqv03Q6IxnUfevOnKs7y2nFyoBdTh2F7rSVsqHPoGZKkGfLsYKt2RMsZ5HNRF3H16QQCPP6KFyhWDAu9BsM7W7/HilF+WDoU4N0tS3gjb/VdhxMr57yIUYNGUuPGBtzf90N4KXI+fLrPEzZk7IK4ghPQ0MUUuC4LjPB3nz0OS5KC4RHypsjAzevEnE2omNeN75TKRp48GsGV+C6fCPscXt70FQzcvhzc4lbD4B0/wivoaT9AG6tpeIE88M4IF3qW72JHd6obtjzRQtVK7FRPlubCidIc2HzmGITmHQXfrP2gxTbrjfwWI5qZB3TYERvh4z2/weDY72FQ7FL2sz2+hzY6fNsKmJ1gYPs6FyWF4DV3gzHrAGzKT2XedZmNlteQVpBmKK5ul8PBSnyvr3H16QRotlCnjuNng/2e6DldiR7U/uJMXpwLsSZ9m6ymlHS/9fe7m+YwkzxAck95OEmpcf61+VsYGe8N+hM74UxVMX+anouC6lL4MSUK3t22DK6hbTPUoKhOmGE7mVBRR0NhKZavj3EUvIhi9NUhX/gtIw62oSHnVRRAVSuPubquGtJL88CQGc8yi1zCjvoiEVa4fpvEesCO7mEUvu+PbrSZEZtRXl/NUi3R3j/v49th+pEQGBvvBWoU3be3LGai+1Dwp/Bg8FQ5cjGgTdI7ok3HtNuC3hlvw/JuAuyUzZ1NR6QO3Pxdug5dE6/d3zga/hQ4Gf4eNZ+J/5REfwg/HY8RSeeefX/RKbiSctYpRV3y6vaDbCiqS9BrpiuHhUh80B/aCQtpMeX9WMGsgbX+rlMSGyUZBL00zyFwkd4D/hg0Gf6BvdDa1E2QWJoFJT0hfDzXDOmVBRCcsRfe2f4j/DH4E3xuFCe2vw+fX7FuuplkWGhQ15jGwovRX8OKpDA4WpZr9fs419wMkblH4K6wzzr3rGTcPsPhTyhcHrt/BcOJ7ZBclgPpVUVQVFsB1fU17CxCNmuNNP9/KZbzFHZ+SfjZTRhq/4J2Mz3eB15AL4jyulN2BUpmeVXL7Ar0PlgWBPzJRIgLDBk8GxO2Y2dC1zZ35uYQG+2YkmTeEzIV/oNe2cbTB+CMFRNblIKKiWLre7EDU1FjdOqFXHW6AL3qL3jROsVV7/gwf4v8EhrbSbqn3rOWGf8F33V64gszhxyM+OIMIzDcWACLk8MgPOcwlLvSSnS0n8PYg687uQPGx/0K1/jTXsgWz9eZMMneNIekaLAPYMc3Ot4TYnOTmeB2FTr0FJhI6zvx3OZyUZtA0etj8ICrA6fAs2gL/41ZxDhxz3qYtPc3eAf//Z/NC9kE0C3B09hnmVAS6ft0HSaC+JOGLOj3TIyc8X2YRYzKTOUdDo9HzAHPEzugsUUWBiXUYDj4FH6WPXPr68ra0izpRj7LVacL2OR2GRZ0Mx8Q+z2xYimR3652joMKRhW+mL0kfNDW33clmkNIalzY8w1Al/z5yDkwat8GFjpmVZxB4XaeVevN2LvTZMKm0wkw8YAO3tyyCO4M+Fju4ajHZoZhxx66K6S2QvWMPx8Pnwm/HIuGdBtvuPfLPiAbXZdCXmoTvF0wAcLrsUF/pDnEYmJElA2cfZZ9h+7blXs7AekZyPvCehwWt4aNW7cFOpatH0VpJMatrxPAtGWr5Dvhcq46XYTefYIUpLC8gSocX4hm73perAtBifKeiJonv7gLvu/CpIpnjXIoipcHyy//eNSXMC3eGwIy9sCBktNwurII6rpw8KaloFQdtG0koTQbInMOwdeH/OCVLYvh9qCPWa57VvfmJSbObCS8TikUfyr6G/A/FQfFnRwv6QhLkkMw3MIwR6kcgtYRbeChjXPaPRn7A/TqZQ1QaH+UrM+g+oSrjQ2gVz+KaprD4szWN0OFfSjsMyhoZ3B6bWq03MO4eo+iSHwmMjSze2/+iSL2YMhnbCbmi4NGWHZ8M3hm7oXteUdZnu7i6lKrxKyhsQEKq0rYTNH+s8fBB0Oa5ce3wJeJfqDa8RP8NXwWSGwwme7foizMs3Xyeqf6o54ahWoQPoshYy/r6OyJWQf0coejVB5BK4htC9uZHsPCtkDJBO8PpfTXCk4LrUDQqc5IPm5/5WpjA1Cud9rfo3QEGLq4/bHAv6Rv48W7ELkoZjQGwQyo9fd7HEkcSMQoTMDnJUOk9UAUhtEBHb4fwn3Bn8ATEbPghU3z4c1tS+Et5L+3LwNNvBeMiNfCyP1aGL7PC97asQLe2kp//x5ejP4KRekLuD/4U3nhHe0aIIMzp6YlgSLDd6VQg0IkrB/aBfHOthWwOS8Zmu0sVIQG9EjVKIyySCqUS9Byok0/gG3ybDtpuFemxUBfVtcK7ZK8K53av+Pc7dZCr3aTDOqmtlLOvBLzHZTT7Egb+OygQQ5LWn/3dzQbW0sqfc4Vic/ScqyDehsSGpp9YaSZIPo995Do32y2kv+dPkvfoe+SEJI4uWr9sNBvMFyK4erb25ZDTG4S1DlAqMyoRc+WBsbZ2JJS+QQtJ7bLeYfb3gVQ0VAD/9j8ndx+W3+XtIQ0Raf24CpjQ6wZfJWkVR9TDAvREPtgj7+zII0X80LsLkhH72Ks3EiY0XLDZKELJzPYVmz5d/N32OCl2ZtoVRaXJokaZ48Sa04SbB76Ddn+I4Rkd08iSNos77HzZ2xLQrC6ROw4r/cbD/sKT/KavRCx+UeZNrB33/r7spacknRjr+cqY2No3b+UF3gpeFnY+38Yt44X80LQmMTQuDXQHz93NQrXDUGTWSK7wdhwJ8WthTkJBvglJQp+Php5nqtTN8E3hwJgwu5fsYGtgkeiFrAjq67B719M6XZprQrzPNAISMR6opH3BDKPagj0RyF+PXYpbMk7iqLhOI9KCbOwvSn2+oIWEm0N62/w7tXthvGjdq1m2qB4Dfko+sVcXewAGnzXq8oVT4ZG8bjGfyJklOfzol6IpJLT4Ju+HdJKc6Ac3fKqpgaW4I4tsmsP+Hfa5V2Nn69orIWTZblsH+DMQ/7w1o6V8FTELLjahN4bNUCiKww09wriO2CdiTu8jGGB/uRuqHdg6NceVqZvkUNtxXLbmtQWkWbvmUhtlDrZ9kheScvvmK/DqHQfBxLLd7FxJGzLS+Y1eiEyyvPgWr9xsmPR+vsUDho11ZJO8wRXFztg/5j+WFmeygetYiViCDftkC8vruNQWlMGMTlJ8F1qFLy/9Qe4At3U8yEke/FO8IJ7E8nQWP0Pg2c3zgHPjDg452QHN2w/mwbXmkbL7UPpGTpFbGfkTdI1zw95KJP2sF4V8DFcETBJkVfSmjkWRSh//380t3MaIuECp1g2GxO9prfQW25vKcO0BD33rhTKxM4dVJukFZMu4upiJ+jc35FMmnpFLwsL90DYDMiqLORFdjxo53kyemD69G3w4uaFbFaOVRodDiGEy/4kjwo7rofDZ8JvaVugoJsPHG0LTSig79JGXJrBVXoOS2kWZzZpMgwG6D3gOr8P4fbgKfBE9DfgvvMX+CrBCL+mRkPw6XiILzrJmFiWA0fLz0AyRiTKPAMJGJGYP0/f9UyLZQf+Ton7jc2q/nnTV3BnyFS2pedKtqUHvRYSL/MkDZXLHpMz+Mz9dGo2WdIWsqoK4YHQ6bLttf4+aYfRoxHtcTBXFTsiZMyleNMdigtJSRCwohanRPJidy+aMIxMKsyAmdhgno74Qn6ZZFDUyFqXXbBr5Es46Lh5yhzgCil74s8ehztpixK1CYuNGj/HRIqEAcUOw90HQj6Ft7cuhdH7tdhR7mSbmKsb6+y684EGURrw+vXNjZBTWQQ7c46Az/FtMAUjnGFxq+HV6K/h0dDPoJ8BPRlKQEAiRuLRMtTsrJB5DYaB235oM1ML4buj4bJoKjkJlHpdp9on6T6w02B7a2hVH6JKKi9xwEp5NGI2FDlZgy3CsHHtiR3wyqZvsJxYbvbyqPdpVX5B60gGgHV5JYbhMxIMkFVRwGvcNbAZDf0Ofzoqn3skzKDN4RWFd0QK8Sj04uEXGuHjYTNhNoY8JvR8TpTm8qs5D+iwUzppOTT3CKw+uRO+SgyAd7YsgSvpWUlEzGErzcSzZTIWduJYJ1egF7kx5xC/04U4g7Z2L4olq6/W3yfNYNGZZhJXEwfAd/Dl+NDHFfNk8cpY185C0u5EaX01RGQfglfx5V1CYwSsFxDCZTWpzrCnpXEg97i1kOqERmspUkuyQY3PcHfwJ3C1aYw8A43tmH5eZhwJN/qPZ9kUXoxZBN+h4e8+ewyNspR/23VQ3VALxysLYF9hOkth7r5rNdyPkcdNNObLsnWQF0YRSDv2gPbyJh0m287kCdm+Wdgv+D5phlZ1gp3M5VBQ2hml06GJqKwPb5zDKshpgT61b+Z+eCv2e6xYLDN7UZ10kXsTqY6wrvrgz/9gw92Zf5RXqOuDwthtOYfBdGIXrEjdBL74MzRjLxwsSIcKJz9hu7Ogo+YOF5yEH1KiQLV3LTwbMQv6U64t5nHyKMRsF/hvyjaxGeuoLdSgzT+yca6yd0WUw8EZXEUcCMOwP2AMnKaY3A/dy374cCvTtvDHcF7Q4GtYdiL8ffO3WMnkJuNLstQ97k2kRkt1g3X0atQCCM06SKkweS0K9BTQftVNeUfhx9RIeCd2KVxEi71JtEiANgyEQdtXsEwgbWHl8Vi0fQqnFWyIvCujJlfyUt3HVcTB0Ku/VDykgug1FJ6MmgfFLjD4SijDnkGbvh0eDJ8px/Vt9RC9kVQX2OPSuMTP2Am1twVLoOeAjnM7iSGk6cROeGfbcngodDrsKWp7VXtRXSU8GYneFQ3yK7Uj+ZCJxRJw/XA4dJoHsCC58ur3VoUjhcU4dpULeFktUVJbAd8eDpQXvPX28S16h16D4Tb/iTD9kD+crXa9cRsB26ESw+L21l19nxKJnT1FKApDK7JG5Eqmkfdz9egm6NTLWGEUUyi/D38ImWZVGlVnQXxxJry/YyX0ZeNbbbyEnkoKAbyHskHnj/esh8MlWbxWBASUkV9TDrdSqm0abG/dnkgbSCNIK7oduhF3S3pVbpszhihaXx8O4o/lWmhoboKgU3vhyY2zZTdX6WX0KNL7kicf/i9mkbwwsKNtU3bCuW7eayhgHcjG2VCKUsfOTsRRFUgG9T1cNboZOtU0xRTKRDTyGwI/glM2TnHrSBSihzgLw8Tr6Fh2Ei5aw6L0rK5MeiZscE9GzIYNp3azyQhHgFK9xOWnsqPgv0vZCAuSgmDC3vWg2bkKxqN3tyApmC1CXHNyF0TnHGFri9iqSQGnAdk22XibHTobu7LFARO2gpfqDizQMcWxLNZrD4Nx+zbwx3NdHCrNhpG7VsNlJr56uKcIF3pV15vGwpIjoZBr73EqFJvcqiLQpW8Ht91r4Bn0Xtk6IOqZaR3Qb+/iz8Hy+CH9/G2gvH1G6w5XmcbAQyHT4O/RC8B9zzrwSouFYxiu0kb67vIEBQDG4rtg70tp5TxpglGTKumG383Vwkmgd58uBYxoVlz9joZ9LTa2uHYOq3AV0JTuvrNpMGjrMriExrcoyZ6rjm/xQfU7Qz6FiKwD/Antg/rGetiC73983Dq4LWgy9GPDBcPkhk7C39FSEvo7fY6WVtB30Bvsi7+nPXR3hc2AMbt/Bf3JnbCz8ARkVxYJAXMQYvOS4RJat6XUeZMW0FpNrfvnXCWcCLQviA5CbGsxKTayV7csglon27XfaTQ3Q8jp/Shc3+ML4cbHXporiBeWkZYqoNH/Z8sSSETP0W5A4QjM3AcDty7F+2IDpjGy8wLVlbrC754XMQxF+NILOmD0r2FfgHrnSvj8sD+sObEDEs6kQV43bcjPR28yp6IAcjkppO3sQaTOBrJlsmnZu1J4R6QFOlWy5ON2C1cJJ4NRM5ypqpKXRT2qjxv8dnwrf9yeAdoAGp2dCKodK+Fq3zGyQTrzrCIZOJbx2oBJsDQ53G67EeqxMUdmJcBr2KAvJqFinmgHXpQtSPcg8aJwkjYn43u41X8CPBA6DZ7f/A28ju/ps/16WJkUBiEZe2Bz1kFIQK8sv6asQ+7OT4HN2Qch6NQe+Dk5AhYnBoJ6z3p4F727f+1cBc/HfAPPR3+FXMD4zKYv4cHQ6XA/hrFm3ofe7KPhM89/xsw3dvzIrqPZ+xssORwMa1M2wcbM/XCwCMtWWwaF9VXsbD9nAtky2bRiWzfrgE49iquDE2KD+mIs4EZ5+X2rByCiId+FLzCruoQ/cs8BJRikLRwj4z3hz7Tx0xzudNmTsAXx/rRUwXMQXGMcDaP3rGX75+yF7VgPgzFk7k/3Zh5VN65lY14YeWDmUBJ/0o4G/H0/FNL+Bg+WiO5y7Gw64kWGEezz9D35uvhcZu+OZsjo+q3J7od/a0lz21D8HF6ProsiQOliLjaOwHuPhTtCZ8Cb0Qth2j4vWJsWAyG5iXCw+DSUdVP6HrJhsmVWbqV6Jw3QqqOZJjg1jOqXJF+PYnkqs/WDoOHgyxm/z5M/ds9EJrr+a9K3snCxD50LaG6o1uyK7zK5SPF7X4GN/sPda2BrOxkiu4o89EI+PaBDIxstG6HDntUaYr2QR0BlM5PqyRLSIuKW32OeRUsq3c8a8uu0LN/5+5uFkcQOvUcUyJvRS/5H1HwYs3ctfH00nHmAje0cZmpLkA2zd6z03JSNwdejXDJ6vM5VwYlBy+617ivZVCadmd/6YTAkuQx7qojctjdQ9hTUogt/qCQLVqdEwt+iv4YbafU8NUY2C2aeZVR44Z0mXouuyXrtoXAJegN3h82A+Yn+kESLP+00EN3Y3MxOv34g/HP53uy5lMonaDNSO+LhPZtFRTG73vdDeCRiFoyIWwfhmfGQSxkl7PDOw3ISmQ23+Z7lDc6/SdBte3CsBG1u1KuPK26MJmIv8feor6Cqh+6AVwLNLh5DF355ShS4YUj2UtQ8uNqXtv9gj2lOrsa8MGwErXvxC9iy98Wel77LM15eYRqLdfsljIn3hsCMvXbPMnCiPB9UdKqveUeATQVY0GJSu2ACRu1I9uRvD5wCE+PWgl/WfiixUehYUVcFz1A2BvauFcrBTsJR5eG/H+Zq4CLQuo/ku7MvfCg9Ghs+8MzDgbwaeh8qaitg+5ljsD5jD3x3JBg+2LYcG8I8uCpoMvQ1Yr21FKbW5KJwMeVpCv0c3t2yBGZiKLbm5E7YgmFBpYNmodaf3M3WRTlv+NdbydsOCRi9GxSwFyPnwXTytIsy+dvrHMhmZbFSeN+0BYcEy6D+lKuAC8E06BJJywfgWz8YET0DOq5rdw/Kp9QVUC764tpKOFldDGkVZ1gIF356P/ie3AX+p+LOk3bO7y04DsfK8yC9sgDd/jKWg8iRyKgshFG7V7N1UG32tILOQ+r8SLjwXd3kNwFGYsh4tDSHv03LQbZKNsu8eqX7yAPtW6V1I67gKuBi8NE8g2pbrHj4KhEr8VkMX/LrXG9zdK8EhrU6FNC7Q8mr6ubZP6spe6W9miRcFC6icF3s+yFMTzDAKQvXpxWg1/7spq9k4VO6NkVTfh5Fko/bi9z6XRR69Xy+PF/5QbECpmI4I+DcOFF+BtzifoW+FGrQEgGld9ndNIdB52fTUFSJ9G9uqMzgzv/OPGZIwtvLBI2e23Mw3Bs2A3afbfvEdjOmH2jnyC4KBWkvsVPtF+wsyD3UquLa3ByNvfQVhhEQmX2QV42AM4FmALXoVd0T8qls6M7kVTGPAUWHLRIdRAbDZsruCpoCj0bMhKei5jG+snUZ/HfXGngNfz4VOYf97tGIL9jnbvYbDwNoxouejWbb6Dg4MubeIGDeQ+HKgImQ3EFiArLNyw0evF4UrkMr2g3qBMk04gZu9S4OchON6nLltVlIrLg7Qz+FtErXOnGlp+MENuThcWv+JwzdbsR4f/KGaAyFBMZnOPw5ZDoM3/EjfHrYD1anb4dduUmQgeVu7wgqAv39VGkeHDhzDPSn9sD0pBD4+IAPvBXzHdxKh5iyhaB0D7wXiXRPm1TA93mdaRRE5STyGlFGZlUx/JEtEG0jFCSbNqpr0Mbf4tbeQ0DplGlQrp3QcOj2n9jxRALdD//T++F+aqgUQnW7sXLBxJDkIlpbFvo5fHnQCJF5RyCDpZux3Vqj6vpqlrAwKOcgzD9oghc3LWCr3ClbBFvzROUgAVcsp4sQBbg/vtNfUqP4UyuDbNFt58+ycCuGgkh5v+BibuU9CL6Dr8JYN7rN0BArsA/2akuObuTVJdAdOFFZCBP2rMMwCTuWtsYsHEkSCDSYOzF0Ue/6BWKyE6HEgev3qhrrIKPyLKxMjoTXty6FOwPp4FUqEwm5K006tCDW58S9Had7Iltkz9pWhyWvANguBamv5lbew+Dj8VdU46I2Zw3Rxe9nHAWBHbipAvaB7uROuIetq8IetVuNEUWSxku8hsFdwZ/CnCPBkFx8mpeye3G0KBOWH4tmJ8rIW5B42OgqIaPnIPh77GI420Ha8si8ZLRFfD7ad6l0HZoVpBUABlefFewIOtX4theUItHtph3t6Z1YJyLQOaRgCDRs12oMEyj0Iq9K4b04iiRU6Nld6z8evkr0gzQK+ZwQtQ11sKcgHeYmGNi5BUywKGR05qUe6BU+HjETsjtYynCyLBcebm9jM9kuOR061SfcqnswVo/pjz2opxSE6q1UGUSs2Le3LIaq+o6PCKPcQskl2bCz6BTsKDrJftqK2wpPQmpZHpypKoJcfMnEHGRRdRk0NTZAc1Mj+2mvPXr2RmFNOfyIbv+dgR+zOu9eY0OxRAOhDAUeGPrZM5uErVGCbXBFyiZ4LvpruETvgXU52Pk8LuyI7g6c3GESzdqmehi4bZn8DErXIdI+Yb3aJMU6eyYGW4GlVFaltrkKnnp6bLxj4714NbaN5KIMeDz0c+hjGAl9TWOhn+lDJP3sOiXjGPhj0FR4JmIWPBk2k/GvYZ/DS5Ffsoyjg7f9wH4O37kSPj1ogOmHjIxfJIfBd8c2w9JjW2Dj6QMQm50IO9HFpuRtxMbm7k1iWI/ewYq0LfACnSHHBpK7eRqfhBLL8RTWs39WAi+l66EO6zUMy//+tuXsZGzmoTiDx6X9AK5G+4jOPsRL2jYmJujkcrc1sUBj0Ab1CXb4TK+CTvUaupaliqftEOlFY6X9iIbfHurR+FfjZ8yNXp72thHpWjQ+QQO/LUkvlDwS6oXoJ60D2kC5xzlpQzMlNtO6w5XG0XAtit8NfuPg3uCpjH/dOBueYeuDlrLEbxPjfeDHpFDQH98OkacTILE4A7KqiyG3thwK6qqgtrGuy4dC1DU1QA56ij+mRMLjkfOwfBR6YX1196A6luFi9Eom7PeBfMou0APQgF53TO5hGBj7vbx+qTu9V2y/A9Br9Ty+jZeubfyUFsPabJtl9WVDOZWSQfU2t+JeBr3bR5Kvh3IeeCL2DNcaRzEvpSPMoGOGaPzDKWZuUAQoJCAy8cNyEc8LHgoFNWISPpqFIfFgAqmCvti4BhhHwMWmUXBd0BS4N3w2vLNlCYzYsRJmJ+jhh+Rw8D29FwKzE2DTmWMQjTxSnAmn0XPLLMtnpH/TGqqYs8dZorcVR8PRG1wO1wdM5EJFPWg3hyzUg3sNhT+FfMr2SPZEUFJH8mre2fo9XGoWrrY8F3uQd7izkkJ4idpGDJbzJtonSB200rVo1pg5F+7TufX2QvgOHiDp1Gva3bqDxnV30GRIpXxOHWAa7SRn4uAMomUpqQGTwLUQufNeIokc9/CYuHGho9S/NDtFszjIPwRMgsdCP4NHQqYz0r8fDJ6GYe0YWaBoFThdgwm6Aw2mLVI58Hn+sWUxpJbn87fXg4HCFZV9EF7Z/C2+B3yvrMOw83vg7YdsoiOQbZGNyR63wrVovVUotjWtyuj8GUTtDS/36/DlbWtzPIuIRkph1MmKs7yKlUE92mckWmTkLiVanSEJHD4jEzcSNu69tST32pS/301Er7kfivLsA3oodXCWie5GSX0NeJ3YCXeHfY7vh4YN8B0p1VFXycWKbIFsoj1kVBbAX9hBwdgRKl2LyJJxavZJAe43cqvt5aBkX3r1KeVzDYlodJ6D4YUti9iu8Y4wau967n53c9gj+HtiD36FaTT8mhbD31TvREltBcxN9Ic/sIN5bTyjSNfCtj/KggmrM3UV8Cqd+sRspY2OLQDDQIM6R/J6/0lurQIMevc3MCwsanMQnkQLw6DXtv0AlR1s36luqAHVnl/5i+jpnpYLkIwB391tQVMgko6/F2BILMoA1e7VWEcaORzrqjdMbR3bvHvcGmYD7YEmYQbvWiXbCBuWULieeZ+gXj2IW6nA7+CtnsgT2CtXIGU5xIb/8T4vaOhgaQAlxBuDvYw8oI1UvJ6g3UnhCb6zf0Z/AwfQQO2K5maWdru8roql8m1J+l1NvfOdEN2AZfbL2MNysbNxyc62VTYsMJy1+frGOn51ZdARbJPjvfHz7azQJxukWUEf1TRunQKK0KoXMTe0rUF46kWwoqegaNEhDx1hhksOxPcAkrfAV81PQiMqrinjb8Q2oPVscWfSWIrplWmbYdp+Hxi7ew38Z/NCeGPTN/Cv6G9/R/rdwC1LYMKe9fBFgh5WndgKvlkJcARFtJEW/3YzCjFMnHZAD5fSRArztqwIE/mYFbX1jtCAYjUtXov3eF+2JaXrke3RmLJOvZwt9BZoB2zmUOXNVsLTVKpShVJFY4V/eciXv4a2Qf3pdDYQTz0QDURz8aIGYWZbLrFgJ0hCRYP+H8BfwmdCAB2B30Wnhg7wKK6rhIiMeJiUYIC3YpfAI6HT4Qa/8fJ7ZWvf+Jo4moFj96efLYm/o7/RZ2jWlNbO6T3gj4EfwyvRC+D9Pb/Cr8diILuiAJrONfM7Ox4x+SlMYOXnaGOJQUvSZ7AOqI1bUs0rksOhL32nrQ6cZgTJ9nSqAMl37gBulQLtwuud67DiQqRgWlXbhmih8NDet1Up7afIMONzfKE3mcbAdcjLDCPxJeMLM9MTG/Fv78oNnxqzufGzht9K5JTKIiiT6grr8raASbDwcBAUdMGrOoehUnJpDvgcj4W3ty6DW4Mnw6XsvdE9aJkGGjS9l06/ExRW7q2zd42kHPU3BUyAN2IWgTZ9G6R3MCttL1Rj6DoH6+9G3w/lsrXVoWo/gMv1GpiVFMy/2T5+SdnEFumyelO6HolVCDkKqi2Sr+Zmbo0CFmHDsD+iSxovL3doQ7SwsV2KP1fji7AEe/NTYWt2IoRgT03bZhZhj0r8PCkEZhwywmfIEXFrYMi25fDe1u/hb+Gz4OnwL7ARfwz96UguWgVMYkYu+/menEg9FolaC2FjA6i9wHOjZ6WEelgnf0BvZdJ+LRwq6nxmhaq6KvA5tRs0O3+Ga/0myHXNvAjuFXR1YLo90rOwe+E7xfd7R/BUmInh4x4MPbsDtO/vleivZdFq7RFhGa82jALdiZ380+2DbORSOoqNnq/ldVqSZul1qkOS/n0XO6LLWeClehA9rIQ2c2gR0ViYaHWQjMwa0KZmSl6WX1UM+ZVFcLg0G3YXZcAWbEBReYfB68Q2WHIkBKbu2wDDt62AxyK/hKuDpsBVvmPhEuNI+ShzamDmEIR+klH3FA+NjU+hUXsNxWdVw70Ynn2fFApHS3N5DVqPXKzrr7FOn944B+uKdwytjdShpGfEd7bhPbgJPZ2Ru1bDoYKTvLSOA22s/hRF8yLyLtnYFpYL29Td6MXSqUqWYHUqihU9E7XBC56TU7axk5LW4xFufQKdgk79NDae7LZFC8UBjeci/Pl9UhibrnUYzp2DpuYmqGisg6L6asiqKoRdecngm74D5iUGwLu7V8PLm7+BR8I+g5vJWyBDpLzhZIxk8EzAXMULw3KSB8lC5qHwp6DJ8Ob25eB3Ko4ZVWdxsqoIvsbw59bgT7BO8PrWDjg7giRc+Mx0xNXsAzo4UeH4lDfROYfhQfT2pd8GwiMRX8Cu/BT+l/ZBQyYXsWfA9tZWpEKelUGdjw7Ca9zqBLoErfoVyeiR3fbCUiQZk/f7bPaw3pGiZQHyKwthT/4xWHViB8w9EgiaHT/DY2EzsZF4yN4XC0OQzKNwAgGjXpzGd6hOWdnkEOnGwEkwevcvGEZHw/Eupn+hDcKex7fCA6GfMTFg93FqD5QLNno396G4rsKyU2flSJyqLIDPsSM8asE2NbKBrxL94TLWoWPdKj4Tkm2LQ9siGxOwIUj9jZoclkNaqeKJbPZQXvLQ4ODGZBWamyEbPYs9xafA6/gWUO9aBXeGTYdL6fQWmm6mMJINKjtAwEic6D7UqOm+NHuG/74YxfR6//HwRNR8to0mPDsBjtLpKjZYxkTZQ/8v9nu8D/eolMrlzGSdzHAYtnMlnO4Gb6tDoOdPNsCyhpDIKj0DkWyJbEp4VnaCj+ZV5rq262nJxjds588WbeNxBtB+r6qmOojOOgjj4n3gtZiF8GjodDkFL0tdQ2EkHwsjb4caIZE8EuYNtUf8jPnzJEokhCQS5pQ4eM1LDKNYcrcXor6EtzHMm5JghNBTeyALPcPaLqa1aQ0thpD34L3Y8zi1R9UBqXNEz/Dx0BkQ7ERH1FGbH7dnvfyu2xsDNIeBZFMCdoQcHma1K1pkCJ5D4GXsxTPRk3E5oIBll5+BsJzDsOx4DCw4EgQj0Qv79+aF8LeIOXAdGvylAR/hs2JISV6KWZBakxqtcSRcGfAxXEYZHcJnwb82fQPDtv0An8R7w5LUCPjpxDb0nhLhaFEG1NS3v52jK6isq4JPEgz4bvD9sPGUVu/MJYkdAj7LJcZRsDg5jHnO3Yn82nJ4mfYGUmfUXmcgh4FZIgx0FKhX0Gty2w8PsTFhD/jXyLmQbEHM7+w4hyEubTs5W1MGxyrOQjKGIgmlWZBQnAEJJW0Q/5ZYmg0p+Pmj+Pm86lK2TaXegh0CtsT+wlPwf2x6Hj07ElKl9+XKpNk37Bxm7td1eB6ivZBamgNPoYfM6pjavlI5icxm0HaEZ+VgMNFS57WbloZ6QHyB9wZ/AlvRWxFwMM4BrD4eC1dTVgJnnP2zJSn88n4f3t/+I8uV70gcOJsGj4TPwLZOnlUHYmUQYtV90KpfkkyaBCkARYtW6Sq9JCIay21+48AnfTt/xQL2Bk0q0L49JlI9JgTsgPSsXkPg3dilUOAg0aKFo3f6T8Q2PkS5TESyDbIRk+a0CAO7Gz5uD2GvsV/eUtDGOhMi9n4XaVXw6SGTY9dq9ULE5iXDoxvnMuNlg9NK76Onkg9FDIxdhuG3/Q58rWtqhC8OB8EAGhOkyRilshDJJuQEfPslrw+e5VYj0K2gBIAG9Sa2uLStLA9EGj9BIxq840fIc8XBeCcHGdHipDC4hvbAMSNqJzzpySRPC736cXFr7TJeSG138O5V8qRKe2OCJFZsyES1E/9fbLdxKtBx2Xq1H3tBlFNL6QUSWQ84GP4WMRu2oicgYBukl+XBv3es5EaEVKr73kStG/RBr352goHN+toK8QXp8BzzXjsYEzSniNGrAiX9+7dyKxFwKlCCfL16qeQ7olnya0e0qOfHHvBKv3GwJCUSmrr5nEBXR8jpA/AwnX5MWS968sC6tUTvpw+GxHS0fVdx7lwz/JgWA9een8Box3ults8G2FW/SD7Dr+TWIeCUmPdSP8lH9Qn2MFXtrtUi8sV1o3av7vD4boELUVFfBVPRg+jLVuj34hCwPWIb628aAxtzj/Basx5FNeUwkc4pYIuiO5jAoCPkTR71KGjzmC0IuAi83YdhDJ/L3OL2BuP5zM6jG+dAdJ7IN24pDhaegre3LJJ7+564tsqW3PAePBAyHbI6sY1n29k0eCpqPmuj7a5cp5lAdioztnmt+0huBQIuBd0HT+PL3MdmENsbjGch4jC4wjgaFh4JhlI7rvh2ddAC1jXpsXBDwMcdj6P0dvLO8AbfD2HSAT2U1VbwWuwYtU0NsOhICBu2YJ1Ce94rdcjUxvXqI5KP24u89Qu4JPQjb5V0ahMLD9s83IKTPAXv9+HVmEUQhz2bwO+RU10Co+PWyuMnFE4r1aGgLFTYAfbBf7+/YyXsPHOM16BlOFycCe9sW451jO2xowkMatP+2LZ16iDJS3Ufb/UCLg0ajNep5kgGVU2723mIbBZxCNyEPdv85FCo7eAEkt6CmLxkeNi8tqq90KQ3kwkVhchu8I9NCyD4dDxLBmkpKMPIitRouDNwErC8Y+0NrBPltlyHYrVQ8na7jLd2gR4DOl+NsirSuFa7ISKSBjfRi3hj83cQV3iCN6neh+rGeliaHAHX0LH3JFZiYF2BWCc06YDe+ePhX8DKtFirFycfKcmCwdvJq0KPijJqKN6Hk62vwhDQpMnCTngob90CPRLe7n/GniuYzaa0u/QBST0c9phXobdFydAq7bhy2RlBa6so3YxFoUlvJHlU1LGhUD2KQrUcvSNrt+TU1tfCd0mhcEsAba+xYEyQQkCWGkazRdJqHuetWqBHg9xnnWo29oxV8iyiQsNoSer1sDE9GjkPgk7vd+7kgDZCYOY++DPLBtpBqpLeSKoPFhoPh/tRqH5GocquLuU1Zzmic4/AM5TFgnlVFuy3lBeC1iC/lnQfXMNbs0CvgdHjTRSuA2yvVXur44nc26KBVLddq9jAaE9EaV0FTD1ogD5ibdWFZEKFAq51h1c2LYBFKZFQVlfFa85ypJTmwNR4L/nYMku8Khq+YEfeqZLws+/w1ivQK7Hu/VvRKH9BF7uBu9rKjcZMloZ5KNyOLvysI0E2P9G4O7G38CS8Ev0N9vYoVGJgXSaJidkDMnjA0K3fg+lUHDR0Yo9gRX01LDwaAXcF0pIQ8lw7qGPy/GnoggRLp14r+aru4K1WoFcDkFrVQGwkKcztJm+LdsErNSIzqQF7DYOHIr6AX9JioKqhljdL1wMdWro6bQvcGDBJDnU6mp3qDTR7UyhWt6HAjIpbB7vPpHYqQR+1DW1GHPw5Yvb/xE/pnmZS26M2SF6VXpMuaTUaafbsPnJjFRAwwwt7MIP6J+zRGiwa26JG7Y2NTzsc3or5jo37dKbn7U6criqGcXvWyR6VJeMoPZYo0iTUtBaPhEqvgdeiv4LFR0Ihs7zzB01E5hzCtrFQrl8KsTsK/8wr1o3qeuxEPTEMFGurBDqAXv0fydcjji3II5dcqWG1JBvfktcnvRm7FGK7sG/MkYjNS4JHqdfvzSvWWciHQk1EsfpD0GSYsd8HQnMSobYLXjO1gcE7foJLjTQWaGH9+uFnaVjC5JGI/z+It0YBAQuwdvC1kkHzFXpbhRYNyhOZxzUUrjaNgve2L4ft+UehxgkXntLaqm8OB2E5x8pCq/QsPZrYwZBA0WlE2MncEfQxDNn+E/if3A3Z1SW8lqxHY1MjvvMUGITvntqAxYts2aA62z5WLBk9Fkn6YTfxViggYCW83Z/DRhWKLnqTRYPyRLbFh2bYNPAeelwBGXuxMTtHltO00hz477ZlaLBYRhpPUSp/TyN5wNSZsPeC3g6+m/uDp4J69ypYlhoFJ7twxD7hHApVOIZ+6p0r5fZB757upVSWlqTP/s+Lj5R0I/7OW52AQBcQMekibFCDJIMqkTUuJlwKDbA1yUiwl+2jdYdXN3/LBrbL662fBrcVjJn72KnFFm37cGnSWBQfMyKBQvHob/CAu0NnwJS9v4FfVjwc76JIEeh0Zf+MPfBGzHcwgE5YZhMWFoR+RGpHJFYGdbKkU3tIa8ZcylubgICNsH7EDdgrzkDXPYPt46IxB6XG2Jpm4UKReCDiC1icFAop5fm82dsfZXWV8Ml+Hzn/d0fbPlyVVMcU5pEYew6GS42j4IGQqfDilkXw9eFAiD+bJh+sa4Psn8cqzsCPRyPh4Y2z4WKzUFkS+hF9sc2w7CEe2ei1zxXhn4D9ofO4F3vFn7DRlZxfBqHUOFvTbFTe78P1gZNg0t71sIUG6O14AGdqaTb8PWoBSOv/i4YlT9OzcIUMjMrjMp4WhXXm0A7LTs/A6hIFmATDMAIeD/2chWUzjwRCYGY8nLFxUsZteUdhRoIebqaTq+neREuFitoIy1elLsH/X4t8lLcmAQEHwcftKcmk+QXFq5IJV0fpa86TQhY0OPQGrjaOhpcxXFyWEgk5lQXs2Hpbgo6b988+AD8khcConb/AAxvnwJX+4+ASmr0iAaCxLJolJKMn4zcLGgkDlVOx/HYkE9EWgnS+bLI4DMDQ7nLTSLg9ZBq8FbMEPkfP0S9zL2wrOA7ZFQX8qW2H4tpyWJe2Ff4RuxiupUkKOlWZ6szSuqE2ISfWq0Masa08zVuPgEA3weDxIjbGAAwVq5i7b6nHRSTD5CHFTQGT0EP4BSJyDqF42f5Un2b05MoaaiGnuhh25x+FdSmbYFqCEf5v+3L4++av4fHwmXAHluEaMkwSMy8UsA0D5dk08s6ITNTIuzCLGxcXi4ifZQsn+TVofMl8XXYfJIpBX/0IuMVvPNwX8ik8hQL7GoZ17+9ZD18nBkJYxl70GrOguL7abse1FdSUQmTuYZi85ze4M2gK9GV1IY+FKb5DJZpn/gzYJihXlY/6n9LqMf15ixEQcALo3P+NjdWALGdjXNS7WjKrSGShDhoEm2H6AP4aNhOmHTRAWPZBhy1GpW1GhwtOQEzOYVhzchcsSN0EXyYFwdQDWhTSn8F9x0p25PxzEXPg7xvnwZ+Cp8HFKCwX+02ASzogfYY+S9+h7z4bMRve3PwdaHb9DCN3r4bZh/1gflIwLDseC/qMfbA3PwVOleVBjYN2EFAdU13PSPSFFyjvV4t3IXubCu+sNWkixpxNQa9uwnfvy9qEgIDTwjSwLwsVjerV6KWUstkgdoKJhcJFZKEaGguGH5cZR8FTUXPhw3gv2JyVACU0y2jjsNEioHdG6ZHLaiuhoLoEiqvL4GTFWThQmm0V6Tv03bNVJXK6Hhq/64bHoXsWY13uO5MK47Bun4qcy+qaDdqTUNE7UHo3bZFEikI/vYrGqLT4/deERyXgWtCr/oKNdxGGiMeYaFm6jqslWU+P4QiGUgPw/+8KmQYTdq8FQ0YcxBdlQFOjOMXaUtB6qcMlWWDK3Auf7N0Af8SQ8xIDekQUpvLlD4rvoC3Su6R3yoYBNOn4vn+RfEc/Ifb9Cbg2DOp7sDFPxZ/72LQ2NXL6aa3X1TJs9BoC12CINWjrUphxOAD8M/ZCnh0GnV0dFOZGZx2CWUcCYOi2ZXAry5SA4mQWKBrgt8abahn20VilQXMIfz9L0mke4G9bQKCHwNv9Rknn8W/JoKIB+mIWLtLsYkfpmtsiGRwNWGPoeLFhJNwb/Ak8H/MdfJFggsjT+yELQ7cyOgHIjssmnAYY4pXVV7Nnjs1JZHXwfOwSeDBsGlxlTu9MdcW8KCvDPSK9I3pX5Ckb2dKESOQgDP9v4W9XQKAHw1fzOPbO3zCvi3rtUDQqawbpf0c0QDYDRzNvNFA8nM1sXer3ITwT9RVMxBDy59QoCEZDPohhZFEX9ss5C4qqSyGp+DQE5ybCuuMx8Ok+T3zW+eyZ+1KdmD1RXh+dEil6F/ROzDO/Rs1+7GyWS0aPp0TYJ9A7sX7QDWgc76JB/IyhyVkWKrJtG8jOel7nw0e+jICMlgaTte5wa8BHbBZs+M6VMC5BB4tToiAEw8mDZ9OhpNa6fOSOAJUpuegUbMyMhx+ORcOHCVpWdnqGP9ICTnpOz/fkEM+8zIJ+Z+2guZlU57SDgeqffmrVBfhu1qJIDRar0gUEWsLkfhd6XsPYuh2DOot5XiHoecl7zpQNzBq2HAPjW1jIwC/DcPIm//FwT8hU+HPEF/Di1qWg2v0rfJVghF+ORkLgqTjYlpcMx8rzGE9WFcHpmlLIRE8tC3+WYyhGqVhoCYIS6W/0GfosfYe+m1FdfP56dG26x/rUzbDwkD+M2rMeXoxdyspCZaIDGq6gGTwmTlRuPpNnzZKD9kh1S3Vs9nKNmnz8fTR6ZWoM4+9iCR4FBATaQcDo2yWt6kMULwMaUDqbMqcMlNTrkxfQqdBRgUzE0CshITOHlGxhKHkr+DsFj+UK9NJuDpkGNwZPZSvO/2/TAhi4ZQm8s2WxIulv9Bn6LH3nFvx5fRB6R+fL0MIjZOXAe19QFl4O83e6Qqo7sydFdcqWI2Ad61RGFK/PMOx7kL8FAQEBq0EGZFB9gMa0Ao0qCXv+JjamQjNV52erFAyzSySBIiExk0SLkwlZC5rFpT3SZ1p/r+U1z98H2ZnxpvZIdUN1RF4Um6HFf1Md6lVp+LdVKGAaIVICAvaA7oPrJf3IJ5knoFdFoSGeQiOvZdtAKKQho+zs2FdPItUB1QUNmJMXRXVkxLrSq2ORCySt5nk2aysgIOBA0EkqNL2u03yD4mXCfyczYyVDpSl4Gjwmj8LEwyBbhZLOQPKc6JmI9IwU4pEHRc9Ov9Opj+LPIBSrb/Dzg8SpMwICzgbd8LvZRluThsa/fkJDjUWDLUHjrmdGTkZNYzZEs5iRcbcWA2ejWZSozCTEtCaKnoUE2KBuwGcsx+fdh17nesk0Ypzk4/ZPSTfibmneS/14zQgICDg96JRrX83NklH9Ch/EX4XGvQkNPRF/ZiMrmNGHYThJpNCSiRmKAvPMkOTJ2HyMrBXN40xMlPDeJEpUFnO5qIxUVr06S9JrjuC/t+HPXyWt+yQM9V5Hgbpd8hl+pQRifZSAQM/DuhFXoPE/KmnV/5K0mpH47/koGuSRYVipQq9MfRj/nctEgsIs8mjoJ4kJkbwbEpfWNIeeRPq30mfou+brtLy2LEh0z8PIXZJB5cvLNJ+VkZXV4xFJN14c0y4gIMBBXpmcDvqP6LncL3mpnpR0o55GARmKXg16aarxKCKLUNTWIdcwalWrWTimJ89HfVAm/Rt/R38zf46+w75L10CPz0elxn8/y+5B96J70r3JWxIQ+B0k6f8Bkv8og9Zb8k0AAAAASUVORK5CYII="

if __name__ == "__main__":
    main()
