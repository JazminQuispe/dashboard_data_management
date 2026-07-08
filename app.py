import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os
import base64
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

        .stage-pill {{
            border-radius: 8px; padding: 10px 8px; text-align: center;
            font-size: 12px; min-height: 58px; color: white;
        }}
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
        div.st-key-portfolio_hero h1 {{
            color: white !important; font-size: 34px !important; margin-bottom: 6px;
        }}
        div.st-key-portfolio_hero p {{ color: white !important; }}
        .hero-eyebrow {{
            display: inline-block; color: white; font-size: 11.5px; font-weight: 700;
            letter-spacing: 0.12em; text-transform: uppercase; background: rgba(255,255,255,0.18);
            padding: 4px 12px; border-radius: 20px; margin-bottom: 10px;
        }}
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
df_portfolio_trend = sheets["Portfolio_Trend"].sort_values("Week")
df_phases = sheets["Phases"]
df_milestones = sheets["Milestones"]
df_item_status = sheets["Item_Status"]
df_trends = sheets["Trends"]
df_gov_domains = sheets["Governance_Domains"]
df_gov_domains = df_gov_domains[df_gov_domains["Domain"].notna() & (df_gov_domains["Stage(0-4)"].notna())].copy()
df_gov_stage_info = sheets["Governance_StageInfo"].sort_values("StageOrder")
df_gov_datasets = sheets["Governance_Datasets"]
df_gov_ownership = sheets["Governance_Ownership"]
df_gov_quality = sheets["Governance_QualityRules"]
df_gov_pipeline = sheets["Governance_Pipeline"]
df_gov_monitoring = sheets["Governance_Monitoring"]
df_exec_kpis = sheets["Executive_KPIs"]
df_exec_cycle_config = sheets["Executive_Cycle_Config"].sort_values("Order")
df_exec_cycle_status = sheets["Executive_Cycle_Status"]
df_exec_cycle_history = sheets["Executive_Cycle_History"]
df_extra_kpis = sheets["Extra_KPIs"]

GOV_STAGES = df_gov_stage_info["StageName"].tolist()


# ----------------------------------------------------------------------------
# DATA ACCESS HELPERS
# ----------------------------------------------------------------------------
def get_items_for_module(module):
    seen = []
    for it in df_phases[df_phases["Module"] == module]["Item"]:
        if it not in seen:
            seen.append(it)
    return seen


def get_phases_for(module, item):
    d = df_phases[(df_phases["Module"] == module) & (df_phases["Item"] == item)].sort_values("PhaseOrder")
    return dict(zip(d["PhaseName"], d["Value%"]))


def get_milestones_for(module, item):
    d = df_milestones[(df_milestones["Module"] == module) & (df_milestones["Item"] == item)].sort_values("MilestoneOrder")
    return [(row["Milestone"], str(row["Done"]).strip().upper() == "Y") for _, row in d.iterrows()]


def get_status_for(module, item, default="Pending"):
    row = df_item_status[(df_item_status["Module"] == module) & (df_item_status["Item"] == item)]
    return row.iloc[0]["Status"] if not row.empty else default


def get_trend_for(module, item):
    d = df_trends[(df_trends["Module"] == module) & (df_trends["Item"] == item)].sort_values("Week")
    return d["WeekLabel"].tolist(), d["Planned%"].tolist(), d["Actual%"].tolist()


def get_extra_kpi(module, kpi_name, default="--"):
    row = df_extra_kpis[(df_extra_kpis["Module"] == module) & (df_extra_kpis["KPI_Name"] == kpi_name)]
    return row.iloc[0]["KPI_Value"] if not row.empty else default


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
    value_style = "font-size:30px;" if big else ""
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


def gauge_chart(value, title, color=C_OCEAN_BLUE, height=220):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={"suffix": "%", "font": {"size": 30, "color": C_BLACK}},
        title={"text": title, "font": {"size": 13, "color": C_TEXT_MUTED}},
        gauge={
            "axis": {"range": [0, 100], "tickcolor": C_TEXT_MUTED, "tickfont": {"size": 9}},
            "bar": {"color": color, "thickness": 0.3},
            "bgcolor": C_GREY_BG,
            "borderwidth": 0,
            "steps": [
                {"range": [0, 40], "color": "#FBE7E4"},
                {"range": [40, 75], "color": "#FCEFDD"},
                {"range": [75, 100], "color": "#E6F4EC"},
            ],
        },
    ))
    fig.update_layout(height=height, margin=dict(l=20, r=20, t=40, b=10),
                       paper_bgcolor="white", font=dict(family="Inter"))
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


def project_table_html(df):
    rows = ""
    for _, p in df.iterrows():
        prog = int(p["Progress%"])
        rows += f"""
        <tr>
            <td><b>{p['Name']}</b></td>
            <td class="pbi-desc">{p['Description']}</td>
            <td style="min-width:140px;">
                <div class="progress-track"><div class="progress-fill" style="width:{prog}%;"></div></div>
                <div style="font-size:11px; color:{C_TEXT_MUTED}; margin-top:2px;">{prog}%</div>
            </td>
            <td>{badge(p['Status'])}</td>
            <td><span class="source-tag">{p['Source']}</span></td>
        </tr>
        """
    return _clean_html(f"""
    <table class="pbi-table">
        <thead><tr><th>Project</th><th>Description</th><th>Progress</th><th>Status</th><th>Source</th></tr></thead>
        <tbody>{rows}</tbody>
    </table>
    """)


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
        st.caption("No data owners assigned yet for this domain.") if df.empty else st.dataframe(df, use_container_width=True, hide_index=True)

    elif stage == "Data Quality Rules":
        df = df_gov_quality[df_gov_quality["Domain"] == domain].drop(columns="Domain")
        st.markdown("**Data Quality Rules**")
        st.caption("No quality rules defined yet for this domain.") if df.empty else st.dataframe(df, use_container_width=True, hide_index=True)

    elif stage == "Pipeline & Dashboard Implementation":
        row = df_gov_pipeline[df_gov_pipeline["Domain"] == domain]
        st.markdown("**Pipeline & Dashboard Implementation**")
        if row.empty:
            st.caption("No pipeline info yet for this domain.")
        else:
            info = row.iloc[0]
            st.markdown(
                f"Pipeline: {badge(info['PipelineStatus'])} &nbsp;&nbsp; Dashboard: {badge(info['DashboardStatus'])}",
                unsafe_allow_html=True,
            )
            st.caption(info["Notes"])

    elif stage == "Monitoring & Continuous Improvement":
        row = df_gov_monitoring[df_gov_monitoring["Domain"] == domain]
        st.markdown("**Monitoring Summary**")
        if row.empty:
            st.caption("No monitoring info yet for this domain.")
        else:
            info = row.iloc[0]
            st.caption(
                f"Datasets OK: {int(info['DatasetsOK%'])}% · Recurring issues: {int(info['Issues'])} · "
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

st.write("")

PAGE_NAMES = [
    "Portfolio Overview", "Cheves Value Pack", "General Scaling",
    "AI Models Integration Agent", "SO Knowledge Integration Agent",
    "Data Governance", "Executive Dashboard",
]
tabs = st.tabs(PAGE_NAMES)

# ----------------------------------------------------------------------------
# PAGE: PORTFOLIO OVERVIEW — hero banner, 6 KPI cards, weekly progress trend
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
                <span class="hero-eyebrow">General overview · not just another tab</span>
                <h1>Portfolio Overview</h1>
                <p style="font-size:14px; max-width:640px; margin:0 auto;">
                    Asset Management · Statkraft Perú — consolidated view across the six Data
                    Management workstreams, used as an input for management reporting.
                </p>
                """),
                unsafe_allow_html=True,
            )
    else:
        st.markdown("<h1 style='text-align:center;'>Portfolio Overview</h1>", unsafe_allow_html=True)

    active = len(df_portfolio)
    on_track = int((df_portfolio["Status"] == "On Track").sum())
    at_risk = int((df_portfolio["Status"] == "At Risk").sum())
    delayed = int((df_portfolio["Status"] == "Delayed").sum())
    avg_progress = round(df_portfolio["Progress%"].mean())

    cfg = df_exec_cycle_config
    status_map_exec = dict(zip(df_exec_cycle_status["Milestone"], df_exec_cycle_status["Done"]))
    next_ms = None
    for _, r in cfg.iterrows():
        if status_map_exec.get(r["Milestone"], "N") != "Y":
            next_ms = r
            break
    if next_ms is None:
        next_ms = cfg.iloc[-1]

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        kpi_card("Active Projects", active, accent=C_OCEAN_BLUE, big=True)
    with c2:
        kpi_card("On Track", on_track, accent=STATUS_COLORS["On Track"], big=True)
    with c3:
        kpi_card("At Risk", at_risk, accent=STATUS_COLORS["At Risk"], big=True)
    with c4:
        kpi_card("Delayed", delayed, accent=STATUS_COLORS["Delayed"], big=True)
    with c5:
        kpi_card("Portfolio Avg. Progress", f"{avg_progress}%", accent=C_OCEAN_BLUE, big=True)
    with c6:
        kpi_card("Next Cycle Milestone", next_ms["Milestone"], accent=C_WARM_SOIL,
                  sub=f"Planned day: {int(next_ms['PlannedDayOffset'])}")

    st.markdown("#### Projects")
    st.markdown(project_table_html(df_portfolio), unsafe_allow_html=True)

    st.markdown("#### Portfolio Progress — Planned vs Actual (last 8 weeks)")
    st.caption("Replaces the previous man-hours chart: tracks the portfolio-wide average progress against the S-curve plan.")
    weeks = df_portfolio_trend["WeekLabel"].tolist()
    planned = df_portfolio_trend["Planned%"].tolist()
    actual = df_portfolio_trend["Actual%"].tolist()
    line_chart(weeks, {"Portfolio": actual}, {"Portfolio": planned}, single_label=("Actual", "Planned"))

# ----------------------------------------------------------------------------
# PAGE: CHEVES VALUE PACK
# ----------------------------------------------------------------------------
with tabs[1]:
    prow = df_portfolio[df_portfolio["Name"] == "Cheves Value Pack"].iloc[0]
    st.caption(f"Type {prow['Type']} · Development / Deployment · Source: {prow['Source']}")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Overall Progress", f"{int(prow['Progress%'])}%", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Schedule Deviation", get_extra_kpi("Cheves", "Schedule Deviation"),
                  accent=STATUS_COLORS["At Risk"], sub="vs. baseline plan")
    with c3:
        kpi_card("Open tasks", get_extra_kpi("Cheves", "Open tasks"), accent=C_WARM_SOIL)
    with c4:
        kpi_card("Closed This Week", get_extra_kpi("Cheves", "Closed This Week"), accent=STATUS_COLORS["On Track"])

    CHEVES_ITEMS = get_items_for_module("Cheves")
    st.markdown("#### Progress by Sub-project, Phase, Status & Milestones")
    cols = st.columns(3)
    for col, item in zip(cols, CHEVES_ITEMS):
        with col:
            project_like_card(item, get_phases_for("Cheves", item), get_milestones_for("Cheves", item),
                               status=get_status_for("Cheves", item), compact=True)

    st.markdown("#### Progress Trend — Planned (S-curve) vs Actual (last 8 weeks)")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in CHEVES_ITEMS:
        labels, planned, actual = get_trend_for("Cheves", item)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: GENERAL SCALING
# ----------------------------------------------------------------------------
with tabs[2]:
    prow = df_portfolio[df_portfolio["Name"] == "General Scaling"].iloc[0]
    st.caption(f"Type {prow['Type']} · Fleet-wide expansion: Cahua · Yaupi · Malpaso · Source: {prow['Source']}")

    ROLLOUT_ITEMS = get_items_for_module("Rollout")
    latest_actuals = {}
    for item in ROLLOUT_ITEMS:
        _, _, actual = get_trend_for("Rollout", item)
        latest_actuals[item] = actual[-1]
    overall = round(np.mean(list(latest_actuals.values())))
    live_plants = sum(1 for v in latest_actuals.values() if v >= 60)

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
            project_like_card(item, get_phases_for("Rollout", item), get_milestones_for("Rollout", item),
                               status=get_status_for("Rollout", item), compact=True)

    st.markdown("#### Progress Trend by Plant — Planned (S-curve) vs Actual (last 8 weeks)")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in ROLLOUT_ITEMS:
        labels, planned, actual = get_trend_for("Rollout", item)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: AI MODELS INTEGRATION AGENT
# ----------------------------------------------------------------------------
with tabs[3]:
    prow = df_portfolio[df_portfolio["Name"] == "AI Models Integration Agent"].iloc[0]
    st.caption(f"O&M Agents · AI Models Integration Agent · Analytical & AI Models · Source: {prow['Source']}")

    AI_ITEMS = get_items_for_module("AIAgents")
    latest_actuals = {}
    for item in AI_ITEMS:
        _, _, actual = get_trend_for("AIAgents", item)
        latest_actuals[item] = actual[-1]
    overall = round(np.mean(list(latest_actuals.values())))
    deployment_vals = df_phases[(df_phases["Module"] == "AIAgents") & (df_phases["PhaseName"] == "Deployment")]
    in_prod = int((deployment_vals["Value%"] >= 60).sum())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Overall Progress", f"{overall}%", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Sub-agents in Production", in_prod, accent=STATUS_COLORS["On Track"])
    with c3:
        kpi_card("Total Sub-agents", len(AI_ITEMS), accent=C_WARM_SOIL)
    with c4:
        kpi_card("Adoption Rate", get_extra_kpi("AIAgents", "Adoption Rate"), accent=C_OCEAN_BLUE, sub="Across active plants")

    st.markdown("#### Sub-agents: Progress, Phases, Status & Milestones")
    for row_start in range(0, len(AI_ITEMS), 2):
        row_items = AI_ITEMS[row_start:row_start + 2]
        cols = st.columns(2)
        for col, item in zip(cols, row_items):
            with col:
                project_like_card(item, get_phases_for("AIAgents", item), get_milestones_for("AIAgents", item),
                                   status=get_status_for("AIAgents", item), compact=True)

    st.markdown("#### Progress Trend by Sub-agent — Planned (S-curve) vs Actual (last 8 weeks)")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in AI_ITEMS:
        labels, planned, actual = get_trend_for("AIAgents", item)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: SO KNOWLEDGE INTEGRATION AGENT
# ----------------------------------------------------------------------------
with tabs[4]:
    prow = df_portfolio[df_portfolio["Name"] == "SO Knowledge Integration Agent"].iloc[0]
    st.caption(f"O&M Agents · SO Knowledge Integration Agent · Documents & Knowledge Management · Source: {prow['Source']}")

    SO_ITEMS = get_items_for_module("SOKnowledge")
    latest_actuals = {}
    for item in SO_ITEMS:
        _, _, actual = get_trend_for("SOKnowledge", item)
        latest_actuals[item] = actual[-1]
    overall = round(np.mean(list(latest_actuals.values())))
    deployment_vals = df_phases[(df_phases["Module"] == "SOKnowledge") & (df_phases["PhaseName"] == "Deployment")]
    in_prod = int((deployment_vals["Value%"] >= 60).sum())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Overall Progress", f"{overall}%", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Sub-agents in Production", in_prod, accent=STATUS_COLORS["On Track"])
    with c3:
        kpi_card("Total Sub-agents", len(SO_ITEMS), accent=C_WARM_SOIL)
    with c4:
        kpi_card("Documents Indexed", get_extra_kpi("SOKnowledge", "Documents Indexed"), accent=C_OCEAN_BLUE, sub="ACRs · O&M · Procedures")

    st.markdown("#### Sub-agents: Progress, Phases, Status & Milestones")
    cols = st.columns(3)
    for col, item in zip(cols, SO_ITEMS):
        with col:
            project_like_card(item, get_phases_for("SOKnowledge", item), get_milestones_for("SOKnowledge", item),
                               status=get_status_for("SOKnowledge", item), compact=True)

    st.markdown("#### Progress Trend by Sub-agent — Planned (S-curve) vs Actual (last 8 weeks)")
    labels, actual_dict, planned_dict = None, {}, {}
    for item in SO_ITEMS:
        labels, planned, actual = get_trend_for("SOKnowledge", item)
        actual_dict[item] = actual
        planned_dict[item] = planned
    line_chart(labels, actual_dict, planned_dict)

# ----------------------------------------------------------------------------
# PAGE: DATA GOVERNANCE
# ----------------------------------------------------------------------------
with tabs[5]:
    prow = df_portfolio[df_portfolio["Name"] == "Data Governance MVP"].iloc[0]
    st.caption(f"Type {prow['Type']} · {len(df_gov_domains)} domains × 5-stage maturity framework · Source: {prow['Source']}")

    if "gov_selected_stage" not in st.session_state:
        st.session_state["gov_selected_stage"] = GOV_STAGES[0]

    avg_stage_idx = int(round(df_gov_domains["Stage(0-4)"].mean()))
    avg_planning = round(df_gov_domains["PlanningProgress%"].mean())

    g1, g2 = st.columns([0.72, 0.28])
    with g1:
        st.markdown("#### MVP Maturity Roadmap (overall)")
        stage_cols = st.columns(len(GOV_STAGES))
        for i, (col, stage) in enumerate(zip(stage_cols, GOV_STAGES)):
            if i < avg_stage_idx:
                bg = STATUS_COLORS["On Track"]
            elif i == avg_stage_idx:
                bg = C_OCEAN_BLUE
            else:
                bg = "#D9D6CF"
            is_selected = st.session_state["gov_selected_stage"] == stage
            with col:
                with st.container(key=f"gov_stage_box_{i}"):
                    if st.button(stage, key=f"gov_stage_btn_{i}", use_container_width=True):
                        st.session_state["gov_selected_stage"] = stage
                st.markdown(
                    _clean_html(f"""
                    <style>
                    div.st-key-gov_stage_box_{i} button {{
                        background-color:{bg} !important;
                        color:white !important;
                        border:{'3px solid ' + C_BLACK if is_selected else '3px solid transparent'} !important;
                        border-radius:8px !important;
                        font-weight:700 !important;
                        min-height:58px !important;
                        white-space:normal !important;
                    }}
                    div.st-key-gov_stage_box_{i} button p,
                    div.st-key-gov_stage_box_{i} button div {{ color:white !important; }}
                    div.st-key-gov_stage_box_{i} button:hover {{ filter:brightness(0.93); }}
                    </style>
                    """),
                    unsafe_allow_html=True,
                )
        with st.expander("ℹ️ ¿Qué significa cada etapa? / Stage guide"):
            for _, r in df_gov_stage_info.iterrows():
                st.markdown(f"**{int(r['StageOrder'])}. {r['StageName']}** — {r['Description']}")
    with g2:
        st.markdown("#### Overall Planning Progress")
        gauge_chart(avg_planning, "All domains average", color=C_OCEAN_BLUE, height=200)

    selected_stage = st.session_state["gov_selected_stage"]
    selected_idx = GOV_STAGES.index(selected_stage)

    st.markdown("#### Domains — Cataloging, Ownership & Planning Progress")
    cols = st.columns(len(df_gov_domains))
    for col, (_, d) in zip(cols, df_gov_domains.iterrows()):
        domain = d["Domain"]
        stage_idx = int(d["Stage(0-4)"])
        planning_pct = int(d["PlanningProgress%"])
        with col:
            st.markdown(
                _clean_html(f"""
                <div class="plant-card">
                    <div style="font-weight:700; margin-bottom:4px;">{domain}</div>
                    <div style="font-size:12px; color:{C_TEXT_MUTED}; margin-bottom:10px;">
                        Stage {stage_idx + 1}: {GOV_STAGES[stage_idx]}
                    </div>
                    <div style="font-size:12px;">Cataloged: <b>{int(d['Cataloged%'])}%</b></div>
                    <div style="font-size:12px; margin-bottom:8px;">Owner assigned: <b>{int(d['OwnerAssigned%'])}%</b></div>
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
    st.caption(f"Type {prow['Type']} · Recurring monthly delivery to Management · Source: {prow['Source']}")

    kpis_completed = int((df_exec_kpis["Actual"] >= df_exec_kpis["Target"]).sum())
    status_map = dict(zip(df_exec_cycle_status["Milestone"], df_exec_cycle_status["Done"]))
    data_updated_done = status_map.get("Data Updated", "N") == "Y"
    sustentacion_done = status_map.get("Monthly Review", "N") == "Y"
    fap_done = status_map.get("FAP Presentation", "N") == "Y"

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("KPIs Completed", f"{kpis_completed}/{len(df_exec_kpis)}", accent=C_OCEAN_BLUE)
    with c2:
        kpi_card("Data Updated (Teams)", "Yes" if data_updated_done else "Pending",
                  accent=STATUS_COLORS["On Track"] if data_updated_done else STATUS_COLORS["Pending"])
    with c3:
        kpi_card("Monthly Review (Sustentación)", "Completed" if sustentacion_done else "Pending",
                  accent=STATUS_COLORS["On Track"] if sustentacion_done else STATUS_COLORS["Pending"])
    with c4:
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
        st.markdown("#### Reference Cycle")
        if os.path.exists(CYCLE_DIAGRAM_PATH):
            st.image(CYCLE_DIAGRAM_PATH, use_container_width=True)
            st.caption("Fin de mes → 14 días → Datos Actualizados → 1 día → Actualización Dashboard "
                       "→ Monthly Review → 15 días → FAP (4ª semana)")
        else:
            st.caption("Sube assets/actualizacion.jpg para mostrar el diagrama de referencia.")

    st.markdown("#### Data Update Compliance — Planned vs Actual Day (last 6 months)")
    st.caption("Tracks the day teams actually finished updating their data each month against the day-14 target — "
               "since some teams update the same day as, or right after, the sustentación.")
    hist = df_exec_cycle_history.copy()
    hist["Compliance"] = np.where(hist["ActualDay"] <= hist["PlannedDay"], "On time", "Delayed")
    cycle_compliance_chart(hist["Month"].tolist(), hist["PlannedDay"].tolist(), hist["ActualDay"].tolist(),
                            hist["Compliance"].tolist())

    st.markdown("#### Management KPIs (10)")
    cols = st.columns(5)
    for i, (_, k) in enumerate(df_exec_kpis.iterrows()):
        completed = k["Actual"] >= k["Target"]
        with cols[i % 5]:
            kpi_card(k["Name"], f"{k['Actual']}", sub=f"Target: {k['Target']}",
                      accent=STATUS_COLORS["On Track"] if completed else STATUS_COLORS["Pending"])
