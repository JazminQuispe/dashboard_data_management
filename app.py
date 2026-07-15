import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os
import base64
import calendar
from datetime import datetime


def _clean_html(s: str) -> str:
    """Collapse a multi-line HTML f-string into a single line.
    Prevents Streamlit's markdown parser from misreading indented /
    blank-line-separated HTML as a code block instead of rendering it."""
    return " ".join(line.strip() for line in s.strip().splitlines())


# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Data Management Dashboard | Statkraft",
    page_icon="🔵",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ----------------------------------------------------------------------------
# STATKRAFT BRAND COLORS
# ----------------------------------------------------------------------------
C_BLACK = "#2C0F00"
C_OCEAN_BLUE = "#20BAFF"
C_SKY_BLUE = "#ADDFFD"
C_MIST = "#DFF3FF"
C_WHITE = "#FFFFFF"
C_WARM_SOIL = "#3A0E00"
C_RED_SOIL = "#914600"
C_DESERT_SUN = "#FFBA8A"
C_SUNSET = "#FFE8D7"
C_GREY_BG = "#FAFAF8"
C_GREY_BORDER = "#E7E5DF"
C_TEXT_MUTED = "#6B6459"

STATUS_COLORS = {
    "On Track": "#1E8E5A",
    "At Risk": "#C77700",
    "Delayed": "#C0392B",
    "Pending": "#8A8580",
}
STATUS_BG = {
    "On Track": "#E6F4EC",
    "At Risk": "#FCEFDD",
    "Delayed": "#FBE7E4",
    "Pending": "#EFEEEA",
}

TREND_LINE_COLORS = [C_OCEAN_BLUE, C_WARM_SOIL, C_DESERT_SUN, C_RED_SOIL]

SCHEDULE_STATUS_COLORS = {
    "Not Started": "#B9B4AA",
    "No Deviation": "#1E8E5A",
    "Moderate Deviation": "#C77700",
    "Critical Deviation": "#8B1E1E",
    "Stopped": "#ADDFFD",
}
SCHEDULE_STATUS_LABELS = {
    "Not Started": "Not Started",
    "No Deviation": "No Deviation",
    "Moderate Deviation": "Moderate Deviation (≤15%)",
    "Critical Deviation": "Critical Deviation (>15%)",
    "Stopped": "Stopped",
}

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

BASE_DIR = os.path.dirname(__file__)
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
DATA_FILE = os.path.join(BASE_DIR, "data_source.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "logo_stk.png")
CHEVES_BG_PATH = os.path.join(ASSETS_DIR, "cheves_bg.jpg")
CYCLE_DIAGRAM_PATH = os.path.join(ASSETS_DIR, "actualizacion.jpg")

# ----------------------------------------------------------------------------
# GLOBAL CSS
# ----------------------------------------------------------------------------
st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"] {{
            font-family: 'Inter', 'Segoe UI', sans-serif;
        }}
        .stApp {{ background-color: {C_WHITE}; }}
        #MainMenu, footer, header {{ visibility: hidden; }}

        h1, h2, h3, h4 {{ color: {C_BLACK}; font-weight: 700; letter-spacing: -0.01em; }}
        h1 {{ font-size: 26px; margin-bottom: 2px; }}
        h4 {{ font-size: 15px; text-transform: uppercase; letter-spacing: 0.04em; color: {C_TEXT_MUTED}; margin: 22px 0 10px 0; }}

        button[data-baseweb="tab"] {{
            font-size: 14px; font-weight: 600; color: {C_TEXT_MUTED};
            padding: 10px 16px;
        }}
        button[data-baseweb="tab"][aria-selected="true"] {{
            color: {C_OCEAN_BLUE};
        }}
        div[data-baseweb="tab-highlight"] {{ background-color: {C_OCEAN_BLUE}; height: 3px; }}
        div[data-baseweb="tab-border"] {{ background-color: {C_GREY_BORDER}; }}
        div[data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {C_GREY_BORDER}; }}

        .kpi-card {{
            background-color: {C_WHITE};
            border: 1px solid {C_GREY_BORDER};
            border-left: 4px solid var(--accent, {C_OCEAN_BLUE});
            border-radius: 8px;
            padding: 14px 16px;
            box-shadow: 0 1px 3px rgba(44,15,0,0.06);
            height: 100%;
            min-height: 104px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        .kpi-label {{
            font-size: 11.5px; color: {C_TEXT_MUTED}; margin-bottom: 6px;
            text-transform: uppercase; letter-spacing: 0.04em; font-weight: 600;
        }}
        .kpi-value {{ font-size: 26px; font-weight: 800; color: {C_BLACK}; line-height: 1.1; }}
        .kpi-sub {{ font-size: 11.5px; color: {C_TEXT_MUTED}; margin-top: 4px; }}

        .badge {{
            display: inline-block; padding: 3px 10px; border-radius: 6px;
            font-size: 11.5px; font-weight: 700; white-space: nowrap;
        }}
        .source-tag {{
            font-size: 10.5px; padding: 2px 8px; border-radius: 5px; font-weight: 600;
            background-color: {C_MIST}; color: {C_BLACK}; border: 1px solid {C_SKY_BLUE};
        }}

        .pbi-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        .pbi-table th {{
            text-align: left; background-color: {C_BLACK}; color: white;
            padding: 10px 14px; font-weight: 600; font-size: 10.5px;
            text-transform: uppercase; letter-spacing: 0.04em;
        }}
        .pbi-table td {{ padding: 11px 14px; border-bottom: 1px solid {C_GREY_BORDER}; vertical-align: middle; }}
        .pbi-table tr:nth-child(even) td {{ background-color: {C_GREY_BG}; }}
        .pbi-table tr:hover td {{ background-color: {C_MIST}; }}
        .pbi-desc {{ color: {C_TEXT_MUTED}; font-size: 12px; }}

        .progress-track {{ background-color: {C_GREY_BORDER}; border-radius: 6px; height: 8px; width: 100%; }}
        .progress-fill {{ background-color: {C_OCEAN_BLUE}; border-radius: 6px; height: 8px; }}

        .checklist-row {{
            display: flex; align-items: center; gap: 12px;
            padding: 10px 14px; border: 1px solid {C_GREY_BORDER}; border-radius: 8px;
            background-color: {C_WHITE}; margin-bottom: 6px; font-size: 13px;
        }}
        .checklist-icon {{ font-size: 16px; width: 20px; text-align: center; }}
        .checklist-label {{ flex: 2; font-weight: 600; color: {C_BLACK}; }}
        .checklist-target {{ flex: 1.4; color: {C_TEXT_MUTED}; font-size: 12px; }}

        .plant-card {{
            background-color: {C_GREY_BG}; border: 1px solid {C_GREY_BORDER};
            border-radius: 10px; padding: 14px; height: 100%;
        }}
        div.st-key-logo_box img {{ max-width: 400px !important; width: 100% !important; }}
        [data-testid="stMetricValue"] {{ color: {C_BLACK}; }}
        div.block-container {{ padding-top: 1.4rem; }}

        div.st-key-portfolio_hero {{
            border-radius: 16px;
            padding: 40px 24px;
            margin-bottom: 24px;
            text-align: center;
            background-size: cover;
            background-position: center;
        }}
        div.st-key-portfolio_hero div[data-testid="stMarkdownContainer"] {{
            text-align: center;
        }}
        div.st-key-portfolio_hero h1 {{
            color: white !important; font-size: 34px !important; margin-bottom: 6px;
            text-align: center;
        }}
        div.st-key-portfolio_hero p {{ color: white !important; text-align: center; }}
        .hero-eyebrow {{
            display: inline-block; color: white; font-size: 11.5px; font-weight: 700;
            letter-spacing: 0.12em; text-transform: uppercase; background: rgba(255,255,255,0.18);
            padding: 4px 12px; border-radius: 20px; margin-bottom: 10px;
        }}

        /* Project overview cards (Portfolio tab) — real bordered containers */
        div[data-testid="stVerticalBlockBorderWrapper"].st-key-projcard {{
            border-radius: 8px !important; box-shadow: 0 1px 3px rgba(44,15,0,0.06);
        }}

        /* Chevron / arrow roadmap for Data Governance */
        .chevron-wrap {{ display: flex; width: 100%; }}
        .chevron-step {{
            flex: 1; position: relative; color: white; font-weight: 700; font-size: 12px;
            text-align: center; padding: 16px 14px; margin-left: -16px;
            clip-path: polygon(0% 0%, 88% 0%, 100% 50%, 88% 100%, 0% 100%, 12% 50%);
        }}
        .chevron-step:first-child {{ margin-left: 0; clip-path: polygon(0% 0%, 88% 0%, 100% 50%, 88% 100%, 0% 100%); }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# DATA LOADING — single source of truth: data_source.xlsx
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Loading data_source.xlsx…")
def load_workbook(path, _mtime):
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")


if not os.path.exists(DATA_FILE):
    st.error(
        f"No se encontró **data_source.xlsx** junto a app.py (ruta esperada: `{DATA_FILE}`). "
        "Coloca el archivo en la misma carpeta que app.py y recarga la página."
    )
    st.stop()

sheets = load_workbook(DATA_FILE, os.path.getmtime(DATA_FILE))

df_portfolio = sheets["Portfolio"]
df_weekly = sheets["Weekly_Data"]
df_phases = sheets["Phases"]
df_milestones = sheets["Milestones"]
df_gov_domains_w = sheets["Governance_Domains"]
df_gov_stage_info = sheets["Governance_StageInfo"].sort_values("StageOrder")
df_gov_datasets = sheets["Governance_Datasets"]
df_gov_ownership = sheets["Governance_Ownership"]
df_gov_quality = sheets["Governance_QualityRules"]
df_gov_pipeline = sheets["Governance_Pipeline"]
df_gov_monitoring = sheets["Governance_Monitoring"]
df_exec_cycle_config = sheets["Executive_Cycle_Config"].sort_values("Order")
df_exec_cycle_status = sheets["Executive_Cycle_Status"]
df_exec_cycle_history = sheets["Executive_Cycle_History"]

GOV_STAGES = df_gov_stage_info["StageName"].tolist()

_weekly_idx = df_weekly.set_index(["Module", "Item", "Week"]).sort_index()
_gov_idx = df_gov_domains_w.set_index(["Domain", "Week"]).sort_index()

MODULE_ITEMS = {
    "Portfolio": df_portfolio["Name"].tolist(),
    "Cheves": ["Anomaly Detection", "CWS", "UH Prediction"],
    "Rollout": ["Cahua", "Yaupi", "Malpaso"],
    "AIAgents": [
        "Autonomous Anomaly Detection Agent", "CWS Health Monitoring and Forecasting Agent",
        "HPU Health Monitoring and Forecasting Agent", "Vibration Monitoring Agent",
    ],
    "SOKnowledge": [
        "O&M Manuals Agent", "Root Cause Analysis Agent", "Cross-Functional SO Meetings Agent",
    ],
}


def week_label(week):
    m = (week - 1) // 4
    wim = (week - 1) % 4 + 1
    return f"{MONTH_ABBR[m]}-W{wim}"


def get_row(module, item, week):
    try:
        return _weekly_idx.loc[(module, item, week)]
    except KeyError:
        return None


def get_series(module, item, upto_week):
    d = df_weekly[(df_weekly["Module"] == module) & (df_weekly["Item"] == item) & (df_weekly["Week"] <= upto_week)]
    d = d.sort_values("Week")
    labels = [week_label(w) for w in d["Week"]]
    return labels, d["Planned%"].tolist(), d["Progress%"].tolist()


def get_module_avg_at(module, week, col="Progress%"):
    items = MODULE_ITEMS[module]
    vals = []
    for it in items:
        r = get_row(module, it, week)
        if r is not None:
            vals.append(r[col])
    return round(np.mean(vals)) if vals else 0


def get_gov_row(domain, week):
    try:
        return _gov_idx.loc[(domain, week)]
    except KeyError:
        return None


def get_gov_series(domain, upto_week):
    d = df_gov_domains_w[(df_gov_domains_w["Domain"] == domain) & (df_gov_domains_w["Week"] <= upto_week)]
    d = d.sort_values("Week")
    labels = [week_label(w) for w in d["Week"]]
    return labels, d["Planned%"].tolist() if "Planned%" in d.columns else None, d["PlanningProgress%"].tolist()


def get_phases_for(module, item):
    d = df_phases[(df_phases["Module"] == module) & (df_phases["Item"] == item)].sort_values("PhaseOrder")
    return dict(zip(d["PhaseName"], d["Value%"]))


def get_milestones_for(module, item, upto_week):
    d = df_milestones[(df_milestones["Module"] == module) & (df_milestones["Item"] == item)].sort_values("MilestoneOrder")
    out = []
    for _, row in d.iterrows():
        wc = row["WeekCompleted"]
        done = pd.notna(wc) and wc <= upto_week
        out.append((row["Milestone"], done))
    return out


def get_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


# ----------------------------------------------------------------------------
# UI HELPERS
# ----------------------------------------------------------------------------
def badge(text):
    color = STATUS_COLORS.get(text, "#8A8580")
    bg = STATUS_BG.get(text, "#EFEEEA")
    return f'<span class="badge" style="background-color:{bg}; color:{color};">{text}</span>'


def kpi_card(label, value, accent=C_OCEAN_BLUE, sub=None, big=False):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    value_style = "font-size:26px;" if big else "font-size:20px;"
    st.markdown(
        _clean_html(f"""
        <div class="kpi-card" style="--accent:{accent};">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" style="{value_style}">{value}</div>
            {sub_html}
        </div>
        """),
        unsafe_allow_html=True,
    )


def status_dot(color, size=12):
    return f'<span style="width:{size}px;height:{size}px;border-radius:50%;background:{color};display:inline-block;flex-shrink:0;"></span>'


def kpi_dot_card(label, value, color, sub=None):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    st.markdown(
        _clean_html(f"""
        <div class="kpi-card" style="--accent:{color};">
            <div style="display:flex; align-items:center; gap:9px;">
                {status_dot(color, 15)}
                <div class="kpi-value" style="font-size:22px;">{value}</div>
            </div>
            <div class="kpi-label" style="margin-top:6px;">{label}</div>
            {sub_html}
        </div>
        """),
        unsafe_allow_html=True,
    )


def sparkline_chart(labels, values, color=C_OCEAN_BLUE, height=54, key=None):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=labels, y=values, mode="lines+markers",
        line=dict(color=color, width=2.5), marker=dict(size=3, color=color),
        hovertemplate="%{x}: %{y}%<extra></extra>",
    ))
    fig.update_layout(
        height=height, margin=dict(l=0, r=0, t=2, b=2),
        xaxis=dict(visible=False), yaxis=dict(visible=False),
        plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=key)


def project_overview_card(name, description, status, source, trend_labels, trend_values, key=None):
    color = STATUS_COLORS.get(status, "#8A8580")
    with st.container(border=True, key=f"projcard_{key}"):
        st.markdown(
            _clean_html(f"""
            <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px;">
                <span style="font-weight:700; font-size:14px; color:{C_BLACK};">{name}</span>
                <span title="{status}">{status_dot(color, 13)}</span>
            </div>
            <div class="pbi-desc" style="margin:4px 0 2px 0; min-height:32px;">{description}</div>
            """),
            unsafe_allow_html=True,
        )
        if trend_values:
            sparkline_chart(trend_labels, trend_values, color=C_OCEAN_BLUE, height=52, key=f"spark_{key}")
            last_val, last_label = trend_values[-1], trend_labels[-1]
        else:
            st.caption("No trend data yet.")
            last_val, last_label = "--", ""
        st.markdown(
            _clean_html(f"""
            <div style="display:flex; justify-content:space-between; align-items:center; margin-top:-4px;">
                <span class="source-tag">{source}</span>
                <span style="font-size:11px; color:{C_TEXT_MUTED}; font-weight:600;">{last_val}% as of {last_label}</span>
            </div>
            """),
            unsafe_allow_html=True,
        )


def line_chart(labels, actual_series: dict, planned_series: dict = None, y_suffix="%", height=280,
               single_label=("Actual", "Planned")):
    fig = go.Figure()
    multi = len(actual_series) > 1
    for i, (name, values) in enumerate(actual_series.items()):
        color = TREND_LINE_COLORS[i % len(TREND_LINE_COLORS)]
        actual_name = name if multi else single_label[0]
        fig.add_trace(go.Scatter(
            x=labels, y=values, name=actual_name, mode="lines+markers",
            line=dict(color=color, width=3), marker=dict(size=6, color=color),
            legendgroup=name,
        ))
        if planned_series and name in planned_series:
            planned_name = f"{name} (plan)" if multi else single_label[1]
            fig.add_trace(go.Scatter(
                x=labels, y=planned_series[name], name=planned_name, mode="lines",
                line=dict(color=color, width=2, dash="dot"), opacity=0.55,
                legendgroup=name,
            ))
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        yaxis=dict(gridcolor="#EDEDED", ticksuffix=y_suffix if y_suffix else ""),
        xaxis=dict(showgrid=False),
        font=dict(color=C_BLACK, family="Inter"),
    )
    st.plotly_chart(fig, use_container_width=True)


def cycle_compliance_chart(months, planned, actual, compliance, height=280):
    colors = [STATUS_COLORS["On Track"] if c == "On time" else STATUS_COLORS["Delayed"] for c in compliance]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=months, y=planned, name="Planned (day 14)", mode="lines",
        line=dict(color=C_TEXT_MUTED, width=2, dash="dash"),
    ))
    fig.add_trace(go.Scatter(
        x=months, y=actual, name="Actual", mode="lines+markers",
        line=dict(color=C_OCEAN_BLUE, width=3),
        marker=dict(size=11, color=colors, line=dict(width=1, color="white")),
    ))
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        yaxis=dict(title="Cycle day", gridcolor="#EDEDED"),
        xaxis=dict(showgrid=False),
        font=dict(color=C_BLACK, family="Inter"),
    )
    st.plotly_chart(fig, use_container_width=True)


def phase_bars(title, phase_values: dict, compact=False):
    if title:
        st.markdown(f"**{title}**")
    for phase, pct in phase_values.items():
        fs = "11px" if compact else "12px"
        st.markdown(
            _clean_html(f"""
            <div style="margin-bottom:{'5' if compact else '7'}px;">
                <div style="display:flex; justify-content:space-between; font-size:{fs}; color:{C_TEXT_MUTED};">
                    <span>{phase}</span><span style="font-weight:600; color:{C_BLACK};">{int(pct)}%</span>
                </div>
                <div class="progress-track"><div class="progress-fill" style="width:{int(pct)}%;"></div></div>
            </div>
            """),
            unsafe_allow_html=True,
        )


def checklist_item(label, target, done):
    icon = "✅" if done else "⬜"
    status = badge("On Track" if done else "Pending")
    st.markdown(
        _clean_html(f"""
        <div class="checklist-row">
            <span class="checklist-icon">{icon}</span>
            <span class="checklist-label">{label}</span>
            <span class="checklist-target">{target}</span>
            <span>{status}</span>
        </div>
        """),
        unsafe_allow_html=True,
    )


def project_like_card(name, phases, milestones, status=None, compact=True):
    st.markdown('<div class="plant-card">', unsafe_allow_html=True)
    status_html = badge(status) if status else ""
    st.markdown(
        _clean_html(f"""
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
            <span style="font-weight:700; font-size:13.5px; color:{C_BLACK};">{name}</span>
            {status_html}
        </div>
        """),
        unsafe_allow_html=True,
    )
    phase_bars(None, phases, compact=compact)
    with st.expander("Milestones"):
        for milestone, done in milestones:
            icon = "✅" if done else "⬜"
            st.markdown(f"{icon} {milestone}")
    st.markdown("</div>", unsafe_allow_html=True)


def render_gov_stage_detail(domain, stage):
    if stage == "Inventory & Prioritization":
        df = df_gov_datasets[df_gov_datasets["Domain"] == domain].drop(columns="Domain")
        st.markdown("**Critical Datasets Inventory**")
        st.caption("No datasets cataloged yet for this domain.") if df.empty else st.dataframe(df, use_container_width=True, hide_index=True)

    elif stage == "Ownership Definition":
        df = df_gov_ownership[df_gov_ownership["Domain"] == domain].drop(columns="Domain")
        st.markdown("**Ownership Matrix**")
        st.caption("No ownership defined yet for this domain.") if df.empty else st.dataframe(df, use_container_width=True, hide_index=True)

    elif stage == "Data Quality Rules":
        df = df_gov_quality[df_gov_quality["Domain"] == domain].drop(columns="Domain")
        st.markdown("**Quality Rules**")
        st.caption("No quality rules defined yet for this domain.") if df.empty else st.dataframe(df, use_container_width=True, hide_index=True)

    elif stage == "Pipeline & Dashboard Implementation":
        row = df_gov_pipeline[df_gov_pipeline["Domain"] == domain]
        if row.empty:
            st.caption("No pipeline info yet for this domain.")
        else:
            info = row.iloc[0]
            st.markdown(f"**Pipeline status:** {badge(info['PipelineStatus'])}", unsafe_allow_html=True)
            st.markdown(f"**Dashboard status:** {badge(info['DashboardStatus'])}", unsafe_allow_html=True)
            st.caption(info["Notes"])

    elif stage == "Monitoring & Continuous Improvement":
        row = df_gov_monitoring[df_gov_monitoring["Domain"] == domain]
        if row.empty:
            st.caption("No monitoring info yet for this domain.")
        else:
            info = row.iloc[0]
            st.markdown(
                f"**Datasets OK:** {int(info['DatasetsOK%'])}% · **Issues:** {int(info['Issues'])} · "
                f"Backlog items: {int(info['Backlog'])}"
            )


# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
h1, h2, h3 = st.columns([0.20, 0.60, 0.20])
with h1:
    if os.path.exists(LOGO_PATH):
        with st.container(key="logo_box"):
            st.image(LOGO_PATH, use_container_width=True)
with h2:
    st.markdown("<h1>Data Management Dashboard</h1>", unsafe_allow_html=True)
    st.markdown(
        f"<span style='color:{C_TEXT_MUTED}; font-size:13px;'>Asset Management · Statkraft Perú · "
        f"Refresh cadence: weekly</span>",
        unsafe_allow_html=True,
    )
with h3:
    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.caption(f"Source updated: {datetime.fromtimestamp(os.path.getmtime(DATA_FILE)).strftime('%b %d, %Y %H:%M')}")

# ----------------------------------------------------------------------------
# GLOBAL MONTH + WEEK FILTER — drives every tab, always cumulative from Week 1
# ----------------------------------------------------------------------------
f1, f2, f3 = st.columns([0.16, 0.14, 0.70])
with f1:
    sel_month = st.selectbox("📅 Month filter (applies to every tab)", options=list(range(1, 13)),
                              format_func=lambda m: MONTH_ABBR[m - 1], index=5)
with f2:
    sel_week_in_month = st.selectbox("Week", options=[1, 2, 3, 4], format_func=lambda w: f"W{w}", index=3)
selected_week = (sel_month - 1) * 4 + sel_week_in_month
selected_week_label = week_label(selected_week)
with f3:
    st.caption(f"Showing data accumulated from **Jan-W1** through **{selected_week_label}** in every tab below. "
               f"Weeks beyond Jun-W4 are still placeholders at 0 — fill them in weekly in Weekly_Data.")

st.write("")

PAGE_NAMES = [
    "Portfolio Overview", "Cheves Value Pack", "General Scaling",
    "AI Models Integration Agent", "SO Knowledge Integration Agent",
    "Data Governance", "Executive Dashboard",
]
tabs = st.tabs(PAGE_NAMES)

# ----------------------------------------------------------------------------
# PAGE: PORTFOLIO OVERVIEW
# ----------------------------------------------------------------------------
with tabs[0]:
    if os.path.exists(CHEVES_BG_PATH):
        hero_b64 = get_base64(CHEVES_BG_PATH)
        st.markdown(
            f"""
            <style>
            div.st-key-portfolio_hero {{
                background-image: linear-gradient(135deg, rgba(44,15,0,0.80), rgba(32,186,255,0.45)),
                    url('data:image/jpeg;base64,{hero_b64}');
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )
        with st.container(key="portfolio_hero"):
            st.markdown(
                _clean_html("""
                <span class="hero-eyebrow">General Overview - Data Management</span>
                <h1>Portfolio Overview</h1>
                <p style="font-size:14px; max-width:640px; margin:0 auto;">
                    Asset Management · Statkraft Perú — Kevin Carrión & Jazmin Quispe.
                </p>
                """),
                unsafe_allow_html=True,
            )
    else:
        st.markdown("<h1 style='text-align:center;'>Portfolio Overview</h1>", unsafe_allow_html=True)

    st.caption(f"📅 Viewing as of: **{selected_week_label}**")

    prev_week = max(1, selected_week - 6)
    active_6w = 0
    sched_counts = {k: 0 for k in SCHEDULE_STATUS_COLORS}
    progress_vals = []
    for name in MODULE_ITEMS["Portfolio"]:
        cur = get_row("Portfolio", name, selected_week)
        prev = get_row("Portfolio", name, prev_week)
        cur_prog = cur["Progress%"] if cur is not None else 0
        prev_prog = prev["Progress%"] if prev is not None else 0
        if (selected_week <= 6 and cur_prog > 0) or (cur_prog - prev_prog) > 0:
            active_6w += 1
        if cur is not None and cur["ScheduleStatus"] in sched_counts:
            sched_counts[cur["ScheduleStatus"]] += 1
        if cur is not None:
            progress_vals.append(cur_prog)
    avg_progress = round(np.mean(progress_vals)) if progress_vals else 0

    cfg = df_exec_cycle_config
    status_map_exec = dict(zip(df_exec_cycle_status["Milestone"], df_exec_cycle_status["Done"]))
    next_ms = None
    for _, r in cfg.iterrows():
        if status_map_exec.get(r["Milestone"], "N") != "Y":
            next_ms = r
            break
    if next_ms is None:
        next_ms = cfg.iloc[-1]

    kpi_cols = st.columns(8)
    with kpi_cols[0]:
        kpi_card("Active Projects (Last 6 Weeks)", active_6w, accent=C_OCEAN_BLUE, big=True)
    for i, cat in enumerate(["Not Started", "No Deviation", "Moderate Deviation", "Critical Deviation", "Stopped"]):
        with kpi_cols[i + 1]:
            kpi_dot_card(SCHEDULE_STATUS_LABELS[cat], sched_counts[cat], SCHEDULE_STATUS_COLORS[cat])
    with kpi_cols[6]:
        kpi_card("Portfolio Avg. Progress", f"{avg_progress}%", accent=C_OCEAN_BLUE, big=True)
    with kpi_cols[7]:
        kpi_card("Next Milestone", next_ms["Milestone"], accent=C_WARM_SOIL,
                  sub=f"Planned day: {int(next_ms['PlannedDayOffset'])}")

    st.markdown("#### Projects")
    proj_cols = st.columns(3)
    for i, (_, p) in enumerate(df_portfolio.iterrows()):
        labels, _, actuals = get_series("Portfolio", p["Name"], selected_week)
        with proj_cols[i % 3]:
            cur = get_row("Portfolio", p["Name"], selected_week)
            status = cur["Status"] if cur is not None and cur["Status"] else "Pending"
            project_overview_card(p["Name"], p["Description"], status, p["Source"], labels, actuals,
                                   key=p["ProjectID"])
            st.write("")

    st.markdown(f"#### Portfolio Progress — Planned vs Actual (Jan-W1 – {selected_week_label})")
    agg_rows = []
    for w in range(1, selected_week + 1):
        planned_vals, actual_vals = [], []
        for name in MODULE_ITEMS["Portfolio"]:
            r = get_row("Portfolio", name, w)
            if r is not None:
                planned_vals.append(r["Planned%"])
                actual_vals.append(r["Progress%"])
        agg_rows.append((week_label(w), np.mean(planned_vals) if planned_vals else 0,
                          np.mean(actual_vals) if actual_vals else 0))
    agg_labels = [r[0] for r in agg_rows]
    agg_planned = [round(r[1], 1) for r in agg_rows]
    agg_actual = [round(r[2], 1) for r in agg_rows]
    line_chart(agg_labels, {"Portfolio": agg_actual}, {"Portfolio": agg_planned}, single_label=("Actual", "Planned"))

# ----------------------------------------------------------------------------
# PAGE: CHEVES VALUE PACK
# ----------------------------------------------------------------------------
with tabs[1]:
    prow = df_portfolio[df_portfolio["Name"] == "Cheves Value Pack"].iloc[0]
    st.caption(f"Type {prow['Type']} · Development / Deployment")
    st.caption(f"📅 Viewing as of: **{selected_week_label}**")

    CHEVES_ITEMS = MODULE_ITEMS["Cheves"]
    overall = get_module_avg_at("Cheves", selected_week, "Progress%")
    planned_avg = get_module_avg_at("Cheves", selected_week, "Planned%")
    deviation = planned_avg - overall
    open_tasks = sum(get_row("Cheves", it, selected_week)["OpenTickets"] if get_row("Cheves", it, selected_week) is not None else 0 for it in CHEVES_ITEMS)
    closed_week = sum(get_row("Cheves", it, selected_week)["ClosedThisWeek"] if get_row("Cheves", it, selected_week) is not None else 0 for it in CHEVES_ITEMS)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Overall Progress", f"{overall}%", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Schedule Deviation", f"{deviation:+d} pts", accent=STATUS_COLORS["At Risk"], sub="Planned − Actual")
    with c3:
        kpi_card("Open Tasks", int(open_tasks), accent=C_WARM_SOIL)
    with c4:
        kpi_card("Closed This Week", int(closed_week), accent=STATUS_COLORS["On Track"])

    st.markdown("#### Progress by Sub-project, Phase, Status & Milestones")
    cols = st.columns(3)
    for col, item in zip(cols, CHEVES_ITEMS):
        with col:
            r = get_row("Cheves", item, selected_week)
            status = r["Status"] if r is not None and r["Status"] else "Pending"
            project_like_card(item, get_phases_for("Cheves", item),
                               get_milestones_for("Cheves", item, selected_week), status=status, compact=True)

    st.markdown(f"#### Progress Trend — Planned vs Actual (Jan-W1 – {selected_week_label})")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in CHEVES_ITEMS:
        labels, planned, actual = get_series("Cheves", item, selected_week)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: GENERAL SCALING
# ----------------------------------------------------------------------------
with tabs[2]:
    prow = df_portfolio[df_portfolio["Name"] == "General Scaling"].iloc[0]
    st.caption(f"Type {prow['Type']} · Fleet-wide expansion: Cahua · Yaupi · Malpaso")
    st.caption(f"📅 Viewing as of: **{selected_week_label}**")

    ROLLOUT_ITEMS = MODULE_ITEMS["Rollout"]
    overall = get_module_avg_at("Rollout", selected_week, "Progress%")
    live_plants = sum(1 for it in ROLLOUT_ITEMS
                       if (get_row("Rollout", it, selected_week) is not None
                           and get_row("Rollout", it, selected_week)["Progress%"] >= 60))

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Overall Scaling Progress", f"{overall}%", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Plants Near Go-live", live_plants, accent=STATUS_COLORS["On Track"])
    with c3:
        kpi_card("Total Plants", len(ROLLOUT_ITEMS), accent=C_WARM_SOIL)

    st.markdown("#### Progress by Plant, Phase, Status & Milestones")
    cols = st.columns(3)
    for col, item in zip(cols, ROLLOUT_ITEMS):
        with col:
            r = get_row("Rollout", item, selected_week)
            status = r["Status"] if r is not None and r["Status"] else "Pending"
            project_like_card(item, get_phases_for("Rollout", item),
                               get_milestones_for("Rollout", item, selected_week), status=status, compact=True)

    st.markdown(f"#### Progress Trend by Plant — Planned vs Actual (Jan-W1 – {selected_week_label})")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in ROLLOUT_ITEMS:
        labels, planned, actual = get_series("Rollout", item, selected_week)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: AI MODELS INTEGRATION AGENT
# ----------------------------------------------------------------------------
with tabs[3]:
    prow = df_portfolio[df_portfolio["Name"] == "AI Models Integration Agent"].iloc[0]
    st.caption(f"O&M Agents · AI Models Integration Agent · Analytical & AI Models")
    st.caption(f"📅 Viewing as of: **{selected_week_label}**")

    AI_ITEMS = MODULE_ITEMS["AIAgents"]
    overall = get_module_avg_at("AIAgents", selected_week, "Progress%")
    deployment_vals = df_phases[(df_phases["Module"] == "AIAgents") & (df_phases["PhaseName"] == "Deployment")]
    in_prod = int((deployment_vals["Value%"] >= 60).sum())
    closed_week = sum((get_row("AIAgents", it, selected_week)["ClosedThisWeek"] if get_row("AIAgents", it, selected_week) is not None else 0) for it in AI_ITEMS)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Overall Progress", f"{overall}%", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Sub-agents in Production", in_prod, accent=STATUS_COLORS["On Track"])
    with c3:
        kpi_card("Total Sub-agents", len(AI_ITEMS), accent=C_WARM_SOIL)
    with c4:
        kpi_card("Closed This Week", int(closed_week), accent=C_OCEAN_BLUE, sub="Tickets across sub-agents")

    st.markdown("#### Sub-agents: Progress, Phases, Status & Milestones")
    for row_start in range(0, len(AI_ITEMS), 2):
        row_items = AI_ITEMS[row_start:row_start + 2]
        cols = st.columns(2)
        for col, item in zip(cols, row_items):
            with col:
                r = get_row("AIAgents", item, selected_week)
                status = r["Status"] if r is not None and r["Status"] else "Pending"
                project_like_card(item, get_phases_for("AIAgents", item),
                                   get_milestones_for("AIAgents", item, selected_week), status=status, compact=True)

    st.markdown(f"#### Progress Trend by Sub-agent — Planned vs Actual (Jan-W1 – {selected_week_label})")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in AI_ITEMS:
        labels, planned, actual = get_series("AIAgents", item, selected_week)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: SO KNOWLEDGE INTEGRATION AGENT
# ----------------------------------------------------------------------------
with tabs[4]:
    prow = df_portfolio[df_portfolio["Name"] == "SO Knowledge Integration Agent"].iloc[0]
    st.caption(f"O&M Agents · SO Knowledge Integration Agent · Documents & Knowledge Management")
    st.caption(f"📅 Viewing as of: **{selected_week_label}**")

    SO_ITEMS = MODULE_ITEMS["SOKnowledge"]
    overall = get_module_avg_at("SOKnowledge", selected_week, "Progress%")
    deployment_vals = df_phases[(df_phases["Module"] == "SOKnowledge") & (df_phases["PhaseName"] == "Deployment")]
    in_prod = int((deployment_vals["Value%"] >= 60).sum())
    closed_week = sum((get_row("SOKnowledge", it, selected_week)["ClosedThisWeek"] if get_row("SOKnowledge", it, selected_week) is not None else 0) for it in SO_ITEMS)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Overall Progress", f"{overall}%", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Sub-agents in Production", in_prod, accent=STATUS_COLORS["On Track"])
    with c3:
        kpi_card("Total Sub-agents", len(SO_ITEMS), accent=C_WARM_SOIL)
    with c4:
        kpi_card("Closed This Week", int(closed_week), accent=C_OCEAN_BLUE, sub="Tickets across sub-agents")

    st.markdown("#### Sub-agents: Progress, Phases, Status & Milestones")
    cols = st.columns(3)
    for col, item in zip(cols, SO_ITEMS):
        with col:
            r = get_row("SOKnowledge", item, selected_week)
            status = r["Status"] if r is not None and r["Status"] else "Pending"
            project_like_card(item, get_phases_for("SOKnowledge", item),
                               get_milestones_for("SOKnowledge", item, selected_week), status=status, compact=True)

    st.markdown(f"#### Progress Trend by Sub-agent — Planned vs Actual (Jan-W1 – {selected_week_label})")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in SO_ITEMS:
        labels, planned, actual = get_series("SOKnowledge", item, selected_week)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: DATA GOVERNANCE
# ----------------------------------------------------------------------------
with tabs[5]:
    prow = df_portfolio[df_portfolio["Name"] == "Data Governance MVP"].iloc[0]
    domains = df_gov_domains_w["Domain"].unique().tolist()
    st.caption(f"Type {prow['Type']} · {len(domains)} domains × 5-stage maturity framework")
    st.caption(f"📅 Viewing as of: **{selected_week_label}**")

    if "gov_selected_stage" not in st.session_state:
        st.session_state["gov_selected_stage"] = GOV_STAGES[0]

    stage_vals, planning_vals = [], []
    for d in domains:
        r = get_gov_row(d, selected_week)
        if r is not None:
            stage_vals.append(r["Stage(0-4)"])
            planning_vals.append(r["PlanningProgress%"])
    avg_stage_idx = int(round(np.mean(stage_vals))) if stage_vals else 0
    avg_planning = round(np.mean(planning_vals)) if planning_vals else 0

    g1, g2 = st.columns([0.72, 0.28])
    with g1:
        st.markdown("#### MVP Maturity Roadmap")
        st.markdown('<div class="chevron-wrap">' + "".join(
            f'<div class="chevron-step" style="background-color:{STATUS_COLORS["On Track"] if i < avg_stage_idx else (C_OCEAN_BLUE if i == avg_stage_idx else "#D9D6CF")};">{stage}</div>'
            for i, stage in enumerate(GOV_STAGES)
        ) + '</div>', unsafe_allow_html=True)
        st.write("")
        stage_cols = st.columns(len(GOV_STAGES))
        for i, (col, stage) in enumerate(zip(stage_cols, GOV_STAGES)):
            with col:
                if st.button(stage, key=f"gov_stage_btn_{i}", use_container_width=True):
                    st.session_state["gov_selected_stage"] = stage
        with st.expander("ℹ️ Stage guide"):
            for _, r in df_gov_stage_info.iterrows():
                st.markdown(f"**{int(r['StageOrder'])}. {r['StageName']}** — {r['Description']}")
    with g2:
        st.markdown("#### Overall Planning Progress")
        st.markdown(
            _clean_html(f"""
            <div class="kpi-card" style="--accent:{C_OCEAN_BLUE}; min-height:180px; align-items:center;">
                <div class="kpi-label" style="text-align:center;">All domains average</div>
                <div class="kpi-value" style="font-size:44px; text-align:center;">{avg_planning}%</div>
            </div>
            """),
            unsafe_allow_html=True,
        )

    selected_stage = st.session_state["gov_selected_stage"]
    selected_idx = GOV_STAGES.index(selected_stage)

    st.markdown("#### Domains — Cataloging, Ownership & Planning Progress")
    cols = st.columns(len(domains))
    for col, domain in zip(cols, domains):
        r = get_gov_row(domain, selected_week)
        stage_idx = int(r["Stage(0-4)"]) if r is not None else 0
        cataloged = int(r["Cataloged%"]) if r is not None else 0
        owner_pct = int(r["OwnerAssigned%"]) if r is not None else 0
        planning_pct = int(r["PlanningProgress%"]) if r is not None else 0
        with col:
            st.markdown(
                _clean_html(f"""
                <div class="plant-card">
                    <div style="font-weight:700; margin-bottom:4px;">{domain}</div>
                    <div style="font-size:12px; color:{C_TEXT_MUTED}; margin-bottom:10px;">
                        Stage {stage_idx + 1}: {GOV_STAGES[stage_idx]}
                    </div>
                    <div style="font-size:12px;">Cataloged: <b>{cataloged}%</b></div>
                    <div style="font-size:12px; margin-bottom:8px;">Owner assigned: <b>{owner_pct}%</b></div>
                    <div style="display:flex; justify-content:space-between; font-size:11.5px; color:{C_TEXT_MUTED};">
                        <span>Planning progress</span><span style="font-weight:600; color:{C_BLACK};">{planning_pct}%</span>
                    </div>
                    <div class="progress-track"><div class="progress-fill" style="width:{planning_pct}%;"></div></div>
                </div>
                """),
                unsafe_allow_html=True,
            )
            with st.expander(f"Supporting detail — {selected_stage}"):
                if selected_idx > stage_idx:
                    st.caption(f"{domain} hasn't reached this stage yet.")
                else:
                    render_gov_stage_detail(domain, selected_stage)

# ----------------------------------------------------------------------------
# PAGE: EXECUTIVE DASHBOARD
# ----------------------------------------------------------------------------
with tabs[6]:
    prow = df_portfolio[df_portfolio["Name"] == "Executive Dashboard"].iloc[0]
    st.caption(f"Type {prow['Type']} · Recurring monthly delivery to Management")
    st.caption(f"📅 Viewing as of: **{selected_week_label}**")

    status_map = dict(zip(df_exec_cycle_status["Milestone"], df_exec_cycle_status["Done"]))
    data_updated_done = status_map.get("Data Updated", "N") == "Y"
    sustentacion_done = status_map.get("Monthly Review", "N") == "Y"
    fap_done = status_map.get("FAP Presentation", "N") == "Y"

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Data Updated", "Yes" if data_updated_done else "Pending",
                  accent=STATUS_COLORS["On Track"] if data_updated_done else STATUS_COLORS["Pending"])
    with c2:
        kpi_card("Monthly Review", "Completed" if sustentacion_done else "Pending",
                  accent=STATUS_COLORS["On Track"] if sustentacion_done else STATUS_COLORS["Pending"])
    with c3:
        kpi_card("FAP Presentation", "Done" if fap_done else "Pending",
                  accent=STATUS_COLORS["On Track"] if fap_done else STATUS_COLORS["Pending"])

    left, right = st.columns([0.62, 0.38])
    with left:
        st.markdown("#### Monthly Cycle Checklist")
        actual_map = dict(zip(df_exec_cycle_status["Milestone"], df_exec_cycle_status["ActualDayOffset"]))
        for _, r in df_exec_cycle_config.iterrows():
            ms = r["Milestone"]
            done = status_map.get(ms, "N") == "Y"
            actual_day = actual_map.get(ms)
            target_label = f"Planned day {int(r['PlannedDayOffset'])}"
            if done and pd.notna(actual_day):
                target_label += f" · actual day {int(actual_day)}"
            checklist_item(ms, target_label, done)
    with right:
        st.markdown("#### Timeline")
        if os.path.exists(CYCLE_DIAGRAM_PATH):
            st.image(CYCLE_DIAGRAM_PATH, use_container_width=True)
            st.caption("End of month → 14 days → Data Updated → 1 day → Dashboard Update "
                       "→ Monthly Review → 15 days → FAP (4th week)")
        else:
            st.caption("Upload assets/actualizacion.jpg to display the reference diagram.")

    st.markdown(f"#### Data Update Compliance — Planned vs Actual Day (through {MONTH_ABBR[sel_month - 1]})")
    hist = df_exec_cycle_history.copy()
    hist["MonthNum"] = pd.to_datetime(hist["Month"], format="%b-%y").dt.month
    hist = hist[hist["MonthNum"] <= sel_month]
    hist = hist[hist["PlannedDay"] > 0]
    if hist.empty:
        st.caption("No cycle history available on or before the selected month.")
    else:
        hist["Compliance"] = np.where(hist["ActualDay"] <= hist["PlannedDay"], "On time", "Delayed")
        cycle_compliance_chart(hist["Month"].tolist(), hist["PlannedDay"].tolist(), hist["ActualDay"].tolist(),
                                hist["Compliance"].tolist())
