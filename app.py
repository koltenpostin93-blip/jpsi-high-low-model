# ============================================================
# JPSI Futures High / Low Forecast Model
# John Stewart & Partners
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from pathlib import Path
from datetime import datetime
import plotly.graph_objects as go

# ─── PAGE CONFIG ────────────────────────────────────────────
st.set_page_config(
    page_title="JPSI High/Low Futures Model",
    page_icon="🌽",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── COLORS ─────────────────────────────────────────────────
DM_BG       = "#0d1210"
DM_SURFACE  = "#141c18"
DM_SURFACE2 = "#1a2620"
DM_BORDER   = "#253328"
DM_TEXT     = "#e8ede9"
DM_MUTED    = "#7a9485"
JSA_GREEN   = "#5e7164"
JSA_LT      = "#8db89a"
COL_HIGH    = "#8db89a"
COL_LOW     = "#6fa8c4"
COL_GOLD    = "#c4b456"
COL_PURPLE  = "#9b89c4"
COL_RED     = "#e07070"

# ─── CSS ────────────────────────────────────────────────────
st.markdown(f"""
<style>
html,body,[data-testid="stAppViewContainer"]{{background:{DM_BG};color:{DM_TEXT};}}
[data-testid="stSidebar"]{{background:{DM_SURFACE};border-right:1px solid {DM_BORDER};}}
[data-testid="stSidebar"] *{{color:{DM_TEXT} !important;}}
#MainMenu,footer,header{{visibility:hidden;}}
.stDeployButton{{display:none;}}

.hl-row-label{{
  font-size:0.72rem;text-transform:uppercase;letter-spacing:0.08em;
  color:{DM_MUTED};margin-bottom:6px;padding-top:4px;
}}
.hl-card{{
  background:{DM_SURFACE};border:1px solid {DM_BORDER};border-radius:12px;
  padding:16px 18px 14px;text-align:center;height:100%;
}}
.hl-card-high{{border-top:4px solid {COL_HIGH};}}
.hl-card-low{{border-top:4px solid {COL_LOW};}}
.hl-contract{{color:{DM_MUTED};font-size:0.65rem;text-transform:uppercase;
  letter-spacing:0.10em;margin-bottom:2px;}}
.hl-label{{color:{DM_MUTED};font-size:0.60rem;text-transform:uppercase;
  letter-spacing:0.08em;margin-bottom:8px;}}
.hl-price-high{{color:{COL_HIGH};font-size:2rem;font-weight:700;line-height:1;}}
.hl-price-low{{color:{COL_LOW};font-size:2rem;font-weight:700;line-height:1;}}
.hl-sub{{color:{DM_MUTED};font-size:0.76rem;margin-top:5px;}}

.sec-hdr{{
  font-size:1.0rem;font-weight:600;color:{DM_TEXT};
  border-left:4px solid {JSA_LT};padding-left:10px;
  margin:1.4rem 0 0.7rem 0;
}}
.sec-div{{
  border-bottom:1px solid {DM_BORDER};margin:6px 0 12px 0;
  color:{DM_MUTED};font-size:0.66rem;text-transform:uppercase;letter-spacing:0.10em;
  padding-bottom:4px;
}}
.ratio-badge{{
  background:{DM_SURFACE2};border:1px solid {DM_BORDER};border-radius:6px;
  padding:3px 10px;display:inline-block;font-size:0.82rem;
  color:{DM_TEXT};margin-top:4px;
}}
.tbl-header{{
  color:{DM_MUTED};font-size:0.68rem;text-transform:uppercase;
  letter-spacing:0.07em;margin:0.8rem 0 0.3rem;
}}
.indicated-box{{
  background:{DM_SURFACE2};border:1px solid {DM_BORDER};border-radius:8px;
  padding:10px 14px;margin-bottom:10px;font-size:0.88rem;
}}

.stTabs [data-baseweb="tab-list"]{{
  background:{DM_SURFACE};border-radius:10px;
  padding:6px 8px;gap:6px;border:1px solid {DM_BORDER};
}}
.stTabs [data-baseweb="tab"]{{
  color:{DM_MUTED};font-size:0.95rem;font-weight:600;
  padding:9px 22px;border-radius:7px;
}}
.stTabs [data-baseweb="tab"]:hover{{background:{DM_SURFACE2};color:{DM_TEXT};}}
.stTabs [aria-selected="true"]{{color:#fff !important;background:{JSA_GREEN} !important;}}
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"]{{display:none !important;}}

div[data-testid="stDataFrame"]{{
  background:{DM_SURFACE};border-radius:8px;border:1px solid {DM_BORDER};
}}
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING ───────────────────────────────────────────
DATA_PATH = Path(__file__).parent / "data" / "High Low Model.xlsx"


def _parse_year(v):
    if v is None:
        return None
    try:
        return int(str(v).strip())
    except Exception:
        return None


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


@st.cache_data
def load_data():
    wb = openpyxl.load_workbook(DATA_PATH)
    return {
        "cz_low":  _load_cz(wb, "CZ  Low"),
        "cz_high": _load_cz(wb, "CZ High"),
        "cn_low":  _load_cn(wb, "CN Low"),
        "cn_high": _load_cn(wb, "CN High"),
        "sx_low":  _load_sx(wb, "SX Low"),
        "sx_high": _load_sx(wb, "SX High"),
        "sn_low":  _load_sn(wb, "SN Low",  is_low=True),
        "sn_high": _load_sn(wb, "SN High", is_low=False),
    }


def _load_cz(wb, name):
    """CZ Low/High: col mapping (0-indexed tuple)
       [1]=Year [2]=Usage [4]=Prod [8]=Jan1 [10]=Date [12]=Price [13]=Est flag
    """
    rows = list(wb[name].iter_rows(values_only=True))
    out = []
    for row in rows:
        yr = _parse_year(row[1])
        if not (yr and 1985 <= yr <= 2030):
            continue
        usage, prod, jan1, price = row[2], row[4], row[8], row[12]
        if not all(_is_num(v) for v in [usage, prod, jan1, price]):
            continue
        out.append(dict(
            year=yr,
            usage=usage,
            production=prod,
            ratio=prod / usage,
            jan1=jan1,
            date=row[10] if isinstance(row[10], datetime) else None,
            price=price,
            price_pct=price / jan1,
            is_est=(row[13] == "Est"),
        ))
    df = pd.DataFrame(out)
    if not df.empty:
        df["section"] = df["ratio"].apply(
            lambda r: "Prod ≥ Use" if r >= 1 else "Prod < Use"
        )
    return df


def _load_cn(wb, name):
    """CN Low/High: col mapping
       [1]=Year [4]=C/O% [5]=Price [6]=Est [7]=Date [9]=Jan1
       Stop before crop-scare sub-section (row 49, i=48).
    """
    rows = list(wb[name].iter_rows(values_only=True))
    out = []
    for i, row in enumerate(rows):
        if i >= 48:
            break
        yr = _parse_year(row[1])
        if not (yr and 1985 <= yr <= 2030):
            continue
        co_pct, price, jan1 = row[4], row[5], row[9]
        if not all(_is_num(v) for v in [co_pct, price, jan1]):
            continue
        out.append(dict(
            year=yr,
            carryout_pct=co_pct,
            jan1=jan1,
            date=row[7] if isinstance(row[7], datetime) else None,
            price=price,
            price_pct=price / jan1,
            is_est=(row[6] == "Est"),
        ))
    return pd.DataFrame(out)


def _load_sx(wb, name):
    """SX Low/High: col mapping
       [0]=Year [1]=World C/O% [3]=Jan1 [5]=Date [7]=Price [8]=Est
       Two sections detected via row text.
    """
    rows = list(wb[name].iter_rows(values_only=True))
    out = []
    section = "Prod ≥ Prev Yr Use"
    for i, row in enumerate(rows):
        if i < 8:
            continue
        rstr = " ".join(str(v) for v in row if v is not None)
        if "Less Than" in rstr:
            section = "Prod < Prev Yr Use"
            continue
        if "Equal to or Greater" in rstr:
            section = "Prod ≥ Prev Yr Use"
            continue
        yr = _parse_year(row[0])
        if not (yr and 1985 <= yr <= 2030):
            continue
        co_pct, jan1, price = row[1], row[3], row[7]
        if not all(_is_num(v) for v in [co_pct, jan1, price]):
            continue
        out.append(dict(
            year=yr,
            carryout_pct=co_pct,
            jan1=jan1,
            date=row[5] if isinstance(row[5], datetime) else None,
            price=price,
            price_pct=price / jan1,
            is_est=(row[8] == "Est"),
            section=section,
        ))
    return pd.DataFrame(out)


def _load_sn(wb, name, is_low):
    """SN Low/High: col mapping
       Low:  [1]=Year [6]=C/O% [7]=Price [8]=Est [9]=Date [11]=Jan1
       High: [0]=Year [5]=C/O% [7]=Price [8]=Est [9]=Date [11]=Jan1
       Three sections detected via row text.
    """
    rows = list(wb[name].iter_rows(values_only=True))
    out = []
    section = "No Crop Scare"
    yr_idx  = 1 if is_low else 0
    co_idx  = 6 if is_low else 5

    for i, row in enumerate(rows):
        if i < 9:
            continue
        rstr = " ".join(str(v) for v in row if v is not None)
        rl = rstr.lower()
        if "without" in rl and "crop scare" in rl:
            section = "No Crop Scare"
            continue
        if "south america" in rl:
            section = "SA Crop Problem"
            continue
        if "with" in rl and "crop scare" in rl and "without" not in rl:
            section = "Crop Scare"
            continue

        yr = _parse_year(row[yr_idx])
        if not (yr and 1985 <= yr <= 2030):
            continue
        co_pct, price, jan1 = row[co_idx], row[7], row[11]
        if not all(_is_num(v) for v in [co_pct, price, jan1]):
            continue
        out.append(dict(
            year=yr,
            carryout_pct=co_pct,
            jan1=jan1,
            date=row[9] if isinstance(row[9], datetime) else None,
            price=price,
            price_pct=price / jan1,
            is_est=(row[8] == "Est"),
            section=section,
        ))
    return pd.DataFrame(out)


# ─── COMPUTATION HELPERS ────────────────────────────────────

def rank_df(df, ratio_col, current_ratio, current_jan1):
    """Filter est rows, rank by proximity to current ratio, add indicated price."""
    d = df[~df["is_est"]].copy()
    d["distance"] = (d[ratio_col] - current_ratio).abs()
    d = d.sort_values("distance").reset_index(drop=True)
    d.index = d.index + 1
    d["indicated"] = (d["price_pct"] * current_jan1).round(2)
    return d


def headline(df, ratio_col, current_ratio, current_jan1):
    """Return (median_indicated_price, median_pct, top5_indicated_price)."""
    d = rank_df(df, ratio_col, current_ratio, current_jan1)
    if d.empty:
        return None, None, None
    med_pct  = d["price_pct"].median()
    top5_pct = d.head(5)["price_pct"].median()
    return round(med_pct * current_jan1, 2), med_pct, round(top5_pct * current_jan1, 2)


def fmt_pct(p):
    return f"{p*100:.1f}%" if p is not None else "—"


def fmt_price(p):
    """Display price in $/bu (inputs are in ¢/bu, divide by 100)."""
    return f"${p/100:.2f}/bu" if p is not None else "—"


# ─── SECTION EMOJI MAP ──────────────────────────────────────

SECTION_EMOJI = {
    "Prod ≥ Use":          "🟢",
    "Prod < Use":          "🔴",
    "Prod ≥ Prev Yr Use":  "🟢",
    "Prod < Prev Yr Use":  "🔴",
    "No Crop Scare":       "🟢",
    "SA Crop Problem":     "🟡",
    "Crop Scare":          "🔴",
}


# ─── TABLE BUILDER ──────────────────────────────────────────

def build_display_table(df, ratio_col, ratio_label, current_ratio, current_jan1,
                         price_col_label, current_section=None):
    d = rank_df(df, ratio_col, current_ratio, current_jan1)
    if d.empty:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["Rank"]              = d.index
    out["Year"]              = d["year"].astype(int)

    if "section" in d.columns:
        out["Category"] = d["section"].apply(
            lambda s: f"{SECTION_EMOJI.get(s, '⚪')} {s}"
        )
        if current_section:
            out["Cur Yr"] = d["section"].apply(
                lambda s: "★" if s == current_section else ""
            )

    out[ratio_label]          = (d[ratio_col] * 100).round(2).astype(str) + "%"
    out["Dist. from Current"] = (d["distance"] * 100).round(3).astype(str) + "%"
    out["Jan 1 (¢/bu)"]      = d["jan1"].round(2)
    out[price_col_label]      = d["price"].round(2)
    out["% of Jan 1"]         = (d["price_pct"] * 100).round(1).astype(str) + "%"
    out["Indicated ($/bu)"]   = (d["indicated"] / 100).round(4)
    out["Date of H/L"]        = d["date"].apply(
        lambda x: x.strftime("%m/%d/%Y") if isinstance(x, datetime) else "—"
    )
    return out


# ─── SCATTER CHART ──────────────────────────────────────────

SECTION_COLORS_LOW = {
    "Prod ≥ Use":          COL_LOW,
    "Prod < Use":          COL_GOLD,
    "Prod ≥ Prev Yr Use":  COL_LOW,
    "Prod < Prev Yr Use":  COL_GOLD,
    "No Crop Scare":       COL_LOW,
    "SA Crop Problem":     COL_GOLD,
    "Crop Scare":          COL_RED,
    "All":                 COL_LOW,
}
SECTION_COLORS_HIGH = {
    "Prod ≥ Use":          COL_HIGH,
    "Prod < Use":          COL_PURPLE,
    "Prod ≥ Prev Yr Use":  COL_HIGH,
    "Prod < Prev Yr Use":  COL_PURPLE,
    "No Crop Scare":       COL_HIGH,
    "SA Crop Problem":     COL_PURPLE,
    "Crop Scare":          COL_RED,
    "All":                 COL_HIGH,
}


def make_scatter(df_low, df_high, ratio_col, ratio_label, current_ratio, title):
    fig = go.Figure()

    for df, lbl, cmap, sym in [
        (df_low,  "Low",  SECTION_COLORS_LOW,  "circle"),
        (df_high, "High", SECTION_COLORS_HIGH, "diamond"),
    ]:
        d = df[~df["is_est"]].copy()
        if "section" in d.columns:
            groups = [(s, d[d["section"] == s]) for s in d["section"].unique()]
        else:
            groups = [("All", d)]

        for sec, sd in groups:
            if sd.empty:
                continue
            color = cmap.get(sec, DM_MUTED)
            fig.add_trace(go.Scatter(
                x=sd[ratio_col] * 100,
                y=sd["price_pct"] * 100,
                mode="markers+text",
                text=sd["year"].astype(str),
                textposition="top center",
                textfont=dict(size=8, color=DM_MUTED),
                marker=dict(size=9, color=color, symbol=sym,
                            line=dict(width=1, color=DM_BORDER)),
                name=f"{lbl} — {sec}",
                hovertemplate=(
                    f"<b>%{{text}}</b> ({lbl})<br>"
                    f"{ratio_label}: %{{x:.2f}}%<br>"
                    "Price % of Jan 1: %{y:.1f}%<br>"
                    "<extra></extra>"
                ),
            ))

    fig.add_vline(
        x=current_ratio * 100,
        line_color=JSA_LT, line_dash="dash", line_width=2,
        annotation_text=f" Current: {current_ratio*100:.2f}%",
        annotation_position="top right",
        annotation_font=dict(color=JSA_LT, size=11),
    )

    fig.update_layout(
        title=dict(text=title, font=dict(color=DM_TEXT, size=12), x=0),
        plot_bgcolor=DM_SURFACE2,
        paper_bgcolor=DM_SURFACE2,
        font=dict(color=DM_TEXT, family="Arial"),
        xaxis=dict(title=ratio_label, gridcolor=DM_BORDER, color=DM_MUTED,
                   ticksuffix="%", zeroline=False),
        yaxis=dict(title="Price as % of Jan 1", gridcolor=DM_BORDER, color=DM_MUTED,
                   ticksuffix="%", zeroline=False),
        legend=dict(bgcolor=DM_SURFACE, bordercolor=DM_BORDER, borderwidth=1,
                    font=dict(color=DM_TEXT, size=10)),
        height=430,
        margin=dict(l=55, r=20, t=40, b=50),
    )
    return fig


# ─── HEADLINE TILE BUILDER ──────────────────────────────────

def headline_tile(contract, label, price, pct, top5, kind):
    cls        = "high" if kind == "high" else "low"
    price_html = fmt_price(price)
    pct_html   = f"Median: {fmt_pct(pct)} of Jan 1" if pct else "—"
    top5_html  = f"Top 5 Comps: {fmt_price(top5)}" if top5 else ""
    return f"""
<div class="hl-card hl-card-{cls}">
  <div class="hl-contract">{contract}</div>
  <div class="hl-label">{label}</div>
  <div class="hl-price-{cls}">{price_html}</div>
  <div class="hl-sub">{pct_html}</div>
  <div class="hl-sub">{top5_html}</div>
</div>"""


# ─── SIDEBAR ────────────────────────────────────────────────
LOGO_URL = "https://www.jpsi.com/wp-content/themes/gate39media/img/logo-white.png"

with st.sidebar:
    st.image(LOGO_URL, width=175)
    st.markdown(
        f'<div style="border-bottom:1px solid {DM_BORDER};margin:10px 0 16px;"></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="sec-div">Current Year Assumptions</div>', unsafe_allow_html=True)

    # ── CZ Dec Corn ──
    st.markdown(f'<div style="color:{DM_TEXT};font-weight:600;margin-bottom:4px;">🌽 Dec Corn (CZ25)</div>',
                unsafe_allow_html=True)
    cz_jan1 = st.number_input("Jan 1 Price (¢/bu)", value=458.50, step=0.25,
                               key="cz_jan1", format="%.2f")
    cz_prod = st.number_input("US Production (Mil Bu)", value=15979.0, step=50.0,
                               key="cz_prod", format="%.0f")
    cz_use  = st.number_input("US Usage (Mil Bu)",      value=16120.0, step=50.0,
                               key="cz_use",  format="%.0f")
    cz_ratio = cz_prod / cz_use if cz_use else 0.0
    st.markdown(
        f'<span class="ratio-badge">Prod/Use: {cz_ratio*100:.2f}% '
        f'({"Prod ≥ Use" if cz_ratio >= 1 else "Prod < Use"})</span>',
        unsafe_allow_html=True,
    )

    st.markdown(f'<div style="border-bottom:1px solid {DM_BORDER};margin:14px 0;"></div>',
                unsafe_allow_html=True)

    # ── CN Jul Corn ──
    st.markdown(f'<div style="color:{DM_TEXT};font-weight:600;margin-bottom:4px;">🌽 Jul Corn (CN26)</div>',
                unsafe_allow_html=True)
    cn_jan1   = st.number_input("Jan 1 Price (¢/bu)",  value=452.00, step=0.25,
                                 key="cn_jan1", format="%.2f")
    cn_co_raw = st.number_input("Carryout % of Use",   value=12.9,   step=0.1,
                                 key="cn_co", format="%.1f",
                                 help="Enter as a percentage, e.g. 12.9 for 12.9%")
    cn_co_pct = cn_co_raw / 100
    st.markdown(
        f'<span class="ratio-badge">C/O Ratio: {cn_co_pct*100:.2f}%</span>',
        unsafe_allow_html=True,
    )

    st.markdown(f'<div style="border-bottom:1px solid {DM_BORDER};margin:14px 0;"></div>',
                unsafe_allow_html=True)

    # ── SX Nov Soybeans ──
    st.markdown(f'<div style="color:{DM_TEXT};font-weight:600;margin-bottom:4px;">🫘 Nov Beans (SX25)</div>',
                unsafe_allow_html=True)
    sx_jan1   = st.number_input("Jan 1 Price (¢/bu)",     value=1062.75, step=0.25,
                                 key="sx_jan1", format="%.2f")
    sx_co_raw = st.number_input("World C/O % of Use",     value=28.4,    step=0.1,
                                 key="sx_co", format="%.1f",
                                 help="Enter as a percentage, e.g. 28.4 for 28.4%")
    sx_co_pct = sx_co_raw / 100
    sx_section = st.selectbox(
        "Current Year Category",
        ["Prod ≥ Prev Yr Use", "Prod < Prev Yr Use"],
        key="sx_section",
        help="Is world production this year ≥ or < previous year's use? Highlights (★) matching historical years.",
    )
    st.markdown(
        f'<span class="ratio-badge">World C/O: {sx_co_pct*100:.2f}%</span>',
        unsafe_allow_html=True,
    )

    st.markdown(f'<div style="border-bottom:1px solid {DM_BORDER};margin:14px 0;"></div>',
                unsafe_allow_html=True)

    # ── SN Jul Soybeans ──
    st.markdown(f'<div style="color:{DM_TEXT};font-weight:600;margin-bottom:4px;">🫘 Jul Beans (SN26)</div>',
                unsafe_allow_html=True)
    sn_jan1   = st.number_input("Jan 1 Price (¢/bu)",     value=1072.00, step=0.25,
                                 key="sn_jan1", format="%.2f")
    sn_co_raw = st.number_input("World C/O % of Use",     value=29.5,    step=0.1,
                                 key="sn_co", format="%.1f",
                                 help="Enter as a percentage, e.g. 29.5 for 29.5%")
    sn_co_pct = sn_co_raw / 100
    sn_section = st.selectbox(
        "Current Year Category",
        ["No Crop Scare", "SA Crop Problem", "Crop Scare"],
        key="sn_section",
        help="Classify this marketing year — used to highlight (★) matching historical years in the ranked tables.",
    )
    st.markdown(
        f'<span class="ratio-badge">World C/O: {sn_co_pct*100:.2f}%</span>',
        unsafe_allow_html=True,
    )

    st.markdown(f'<div style="border-bottom:1px solid {DM_BORDER};margin:16px 0 8px;"></div>',
                unsafe_allow_html=True)
    st.markdown(
        f'<div style="color:{DM_MUTED};font-size:0.68rem;text-align:center;">'
        f'Data: USDA NASS &nbsp;·&nbsp; Model by JPSI</div>',
        unsafe_allow_html=True,
    )


# ─── LOAD DATA ──────────────────────────────────────────────
try:
    D = load_data()
except FileNotFoundError:
    st.error(
        "📂 Data file not found. Make sure `data/High Low Model.xlsx` "
        "is in the same folder as `app.py`."
    )
    st.stop()
except Exception as e:
    st.error(f"Error loading data: {e}")
    st.stop()


# ─── CURRENT SECTION LABELS ─────────────────────────────────
# CZ: auto-determined from ratio
cz_section = "Prod ≥ Use" if cz_ratio >= 1 else "Prod < Use"

# ─── HEADLINES (all years — no section filtering) ────────────
cz_hl_low,  cz_lp,  cz_t5_low  = headline(D["cz_low"],  "ratio",        cz_ratio,  cz_jan1)
cz_hl_high, cz_hp,  cz_t5_high = headline(D["cz_high"], "ratio",        cz_ratio,  cz_jan1)
cn_hl_low,  cn_lp,  cn_t5_low  = headline(D["cn_low"],  "carryout_pct", cn_co_pct, cn_jan1)
cn_hl_high, cn_hp,  cn_t5_high = headline(D["cn_high"], "carryout_pct", cn_co_pct, cn_jan1)
sx_hl_low,  sx_lp,  sx_t5_low  = headline(D["sx_low"],  "carryout_pct", sx_co_pct, sx_jan1)
sx_hl_high, sx_hp,  sx_t5_high = headline(D["sx_high"], "carryout_pct", sx_co_pct, sx_jan1)
sn_hl_low,  sn_lp,  sn_t5_low  = headline(D["sn_low"],  "carryout_pct", sn_co_pct, sn_jan1)
sn_hl_high, sn_hp,  sn_t5_high = headline(D["sn_high"], "carryout_pct", sn_co_pct, sn_jan1)


# ─── PAGE HEADER ────────────────────────────────────────────
st.markdown(
    f"""
    <div style="display:flex;align-items:baseline;gap:14px;margin-bottom:20px;">
      <h1 style="margin:0;font-size:1.55rem;font-weight:700;color:{DM_TEXT};">
        Futures High / Low Forecast Model
      </h1>
      <span style="color:{DM_MUTED};font-size:0.85rem;">
        John Stewart &amp; Partners &nbsp;·&nbsp; Historical Analog Pricing
      </span>
    </div>
    """,
    unsafe_allow_html=True,
)

# ─── HEADLINE GRID ──────────────────────────────────────────
st.markdown('<div class="sec-hdr">📌 Indicated Price Range — Current Marketing Year</div>',
            unsafe_allow_html=True)

# Row 1: Lows
st.markdown('<div class="hl-row-label">📉 Indicated Lows (Median of All Historical Years × Current Jan 1 Price)</div>',
            unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)
c1.markdown(headline_tile("CZ — Dec Corn",  "Jan–Dec Indicated Low",  cz_hl_low,  cz_lp,  cz_t5_low,  "low"),  unsafe_allow_html=True)
c2.markdown(headline_tile("CN — Jul Corn",  "Jan–Jul Indicated Low",  cn_hl_low,  cn_lp,  cn_t5_low,  "low"),  unsafe_allow_html=True)
c3.markdown(headline_tile("SX — Nov Beans", "Jan–Nov Indicated Low",  sx_hl_low,  sx_lp,  sx_t5_low,  "low"),  unsafe_allow_html=True)
c4.markdown(headline_tile("SN — Jul Beans", "Jan–Jul Indicated Low",  sn_hl_low,  sn_lp,  sn_t5_low,  "low"),  unsafe_allow_html=True)

st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)

# Row 2: Highs
st.markdown('<div class="hl-row-label">📈 Indicated Highs (Median of All Historical Years × Current Jan 1 Price)</div>',
            unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)
c1.markdown(headline_tile("CZ — Dec Corn",  "Jan–Dec Indicated High", cz_hl_high, cz_hp, cz_t5_high, "high"), unsafe_allow_html=True)
c2.markdown(headline_tile("CN — Jul Corn",  "Jan–Jul Indicated High", cn_hl_high, cn_hp, cn_t5_high, "high"), unsafe_allow_html=True)
c3.markdown(headline_tile("SX — Nov Beans", "Jan–Nov Indicated High", sx_hl_high, sx_hp, sx_t5_high, "high"), unsafe_allow_html=True)
c4.markdown(headline_tile("SN — Jul Beans", "Jan–Jul Indicated High", sn_hl_high, sn_hp, sn_t5_high, "high"), unsafe_allow_html=True)

st.markdown('<div style="height:16px;"></div>', unsafe_allow_html=True)

# ─── CONTRACT TABS ──────────────────────────────────────────
tab_cz, tab_cn, tab_sx, tab_sn = st.tabs([
    "🌽  Dec Corn (CZ)",
    "🌽  Jul Corn (CN)",
    "🫘  Nov Soybeans (SX)",
    "🫘  Jul Soybeans (SN)",
])


# ══════════════════════════════════════════════════════════════
# CZ TAB
# ══════════════════════════════════════════════════════════════
with tab_cz:
    st.markdown(
        '<div class="sec-hdr">Dec Corn (CZ) — Production as % of Same Year\'s Use</div>',
        unsafe_allow_html=True,
    )
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Jan 1 Price",  f"${cz_jan1/100:.2f}/bu")
    m2.metric("US Production", f"{cz_prod:,.0f} Mil Bu")
    m3.metric("US Usage",      f"{cz_use:,.0f} Mil Bu")
    m4.metric("Prod/Use Ratio", f"{cz_ratio*100:.2f}%",
              delta="Prod ≥ Use (Surplus)" if cz_ratio >= 1 else "Prod < Use (Deficit)",
              delta_color="normal" if cz_ratio >= 1 else "inverse")

    st.markdown("#### Historical Scatter: Prod/Use % vs Price as % of Jan 1")
    st.plotly_chart(
        make_scatter(D["cz_low"], D["cz_high"], "ratio",
                     "Prod/Use %", cz_ratio,
                     "Dec Corn: Production/Use vs Jan–Dec High/Low as % of Jan 1"),
        use_container_width=True,
    )

    col_l, col_h = st.columns(2)
    with col_l:
        tbl = build_display_table(D["cz_low"], "ratio", "Prod/Use %",
                                   cz_ratio, cz_jan1, "Jan–Dec Low (¢/bu)",
                                   current_section=cz_section)
        st.markdown(
            f'<div class="indicated-box">📉 <b>Lows</b> ranked by similarity &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_LOW}">{fmt_price(cz_hl_low)}</b> '
            f'({fmt_pct(cz_lp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_LOW}">{fmt_price(cz_t5_low)}</b>'
            f'&nbsp;·&nbsp; ★ = {cz_section}</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)

    with col_h:
        tbl = build_display_table(D["cz_high"], "ratio", "Prod/Use %",
                                   cz_ratio, cz_jan1, "Jan–Dec High (¢/bu)",
                                   current_section=cz_section)
        st.markdown(
            f'<div class="indicated-box">📈 <b>Highs</b> ranked by similarity &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_HIGH}">{fmt_price(cz_hl_high)}</b> '
            f'({fmt_pct(cz_hp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_HIGH}">{fmt_price(cz_t5_high)}</b>'
            f'&nbsp;·&nbsp; ★ = {cz_section}</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)


# ══════════════════════════════════════════════════════════════
# CN TAB
# ══════════════════════════════════════════════════════════════
with tab_cn:
    st.markdown(
        '<div class="sec-hdr">Jul Corn (CN) — Carryout as % of Use</div>',
        unsafe_allow_html=True,
    )
    m1, m2 = st.columns(2)
    m1.metric("Jan 1 Price",      f"${cn_jan1/100:.2f}/bu")
    m2.metric("Carryout / Use",   f"{cn_co_pct*100:.2f}%")

    st.markdown("#### Historical Scatter: Carryout % vs Price as % of Jan 1")
    st.plotly_chart(
        make_scatter(D["cn_low"], D["cn_high"], "carryout_pct",
                     "Carryout % of Use", cn_co_pct,
                     "Jul Corn: Carryout/Use vs Jan–Jul High/Low as % of Jan 1"),
        use_container_width=True,
    )

    col_l, col_h = st.columns(2)
    with col_l:
        tbl = build_display_table(D["cn_low"], "carryout_pct", "Carryout %",
                                   cn_co_pct, cn_jan1, "Jan–Jul Low (¢/bu)")
        st.markdown(
            f'<div class="indicated-box">📉 <b>Lows</b> ranked by similarity &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_LOW}">{fmt_price(cn_hl_low)}</b> '
            f'({fmt_pct(cn_lp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_LOW}">{fmt_price(cn_t5_low)}</b></div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)

    with col_h:
        tbl = build_display_table(D["cn_high"], "carryout_pct", "Carryout %",
                                   cn_co_pct, cn_jan1, "Jan–Jul High (¢/bu)")
        st.markdown(
            f'<div class="indicated-box">📈 <b>Highs</b> ranked by similarity &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_HIGH}">{fmt_price(cn_hl_high)}</b> '
            f'({fmt_pct(cn_hp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_HIGH}">{fmt_price(cn_t5_high)}</b></div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)


# ══════════════════════════════════════════════════════════════
# SX TAB
# ══════════════════════════════════════════════════════════════
with tab_sx:
    st.markdown(
        '<div class="sec-hdr">Nov Soybeans (SX) — World Carryout as % of Same Year\'s Use</div>',
        unsafe_allow_html=True,
    )
    m1, m2 = st.columns(2)
    m1.metric("Jan 1 Price",      f"${sx_jan1/100:.2f}/bu")
    m2.metric("World C/O / Use",  f"{sx_co_pct*100:.2f}%")

    st.markdown("#### Historical Scatter: World C/O % vs Price as % of Jan 1")
    st.plotly_chart(
        make_scatter(D["sx_low"], D["sx_high"], "carryout_pct",
                     "World C/O % of Use", sx_co_pct,
                     "Nov Soybeans: World C/O/Use vs Jan–Nov High/Low as % of Jan 1"),
        use_container_width=True,
    )

    col_l, col_h = st.columns(2)
    with col_l:
        tbl = build_display_table(D["sx_low"], "carryout_pct", "World C/O %",
                                   sx_co_pct, sx_jan1, "Jan–Nov Low (¢/bu)",
                                   current_section=sx_section)
        st.markdown(
            f'<div class="indicated-box">📉 <b>Lows</b> ranked by similarity &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_LOW}">{fmt_price(sx_hl_low)}</b> '
            f'({fmt_pct(sx_lp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_LOW}">{fmt_price(sx_t5_low)}</b>'
            f'&nbsp;·&nbsp; ★ = {sx_section}</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)

    with col_h:
        tbl = build_display_table(D["sx_high"], "carryout_pct", "World C/O %",
                                   sx_co_pct, sx_jan1, "Jan–Nov High (¢/bu)",
                                   current_section=sx_section)
        st.markdown(
            f'<div class="indicated-box">📈 <b>Highs</b> ranked by similarity &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_HIGH}">{fmt_price(sx_hl_high)}</b> '
            f'({fmt_pct(sx_hp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_HIGH}">{fmt_price(sx_t5_high)}</b>'
            f'&nbsp;·&nbsp; ★ = {sx_section}</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)


# ══════════════════════════════════════════════════════════════
# SN TAB
# ══════════════════════════════════════════════════════════════
with tab_sn:
    st.markdown(
        f'<div class="sec-hdr">Jul Soybeans (SN) — World Carryout as % of Use '
        f'<span style="color:{DM_MUTED};font-size:0.75rem;font-weight:400;">'
        f'· Current year: {SECTION_EMOJI.get(sn_section,"⚪")} {sn_section} '
        f'(★ in tables below)</span></div>',
        unsafe_allow_html=True,
    )
    m1, m2, m3 = st.columns(3)
    m1.metric("Jan 1 Price",     f"${sn_jan1/100:.2f}/bu")
    m2.metric("World C/O / Use", f"{sn_co_pct*100:.2f}%")
    m3.metric("Current Category", sn_section)

    st.markdown("#### Historical Scatter: World C/O % vs Price as % of Jan 1 — All Three Categories")
    st.plotly_chart(
        make_scatter(D["sn_low"], D["sn_high"], "carryout_pct",
                     "World C/O % of Use", sn_co_pct,
                     "Jul Soybeans: World C/O/Use vs Jan–Jul High/Low as % of Jan 1"),
        use_container_width=True,
    )

    col_l, col_h = st.columns(2)
    with col_l:
        tbl = build_display_table(D["sn_low"], "carryout_pct", "World C/O %",
                                   sn_co_pct, sn_jan1, "Jan–Jul Low (¢/bu)",
                                   current_section=sn_section)
        st.markdown(
            f'<div class="indicated-box">📉 <b>Lows</b> ranked by similarity — all categories &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_LOW}">{fmt_price(sn_hl_low)}</b> '
            f'({fmt_pct(sn_lp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_LOW}">{fmt_price(sn_t5_low)}</b>'
            f'&nbsp;·&nbsp; ★ = {sn_section}</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)

    with col_h:
        tbl = build_display_table(D["sn_high"], "carryout_pct", "World C/O %",
                                   sn_co_pct, sn_jan1, "Jan–Jul High (¢/bu)",
                                   current_section=sn_section)
        st.markdown(
            f'<div class="indicated-box">📈 <b>Highs</b> ranked by similarity — all categories &nbsp;|&nbsp; '
            f'Median indicated: <b style="color:{COL_HIGH}">{fmt_price(sn_hl_high)}</b> '
            f'({fmt_pct(sn_hp)} of Jan 1) &nbsp;|&nbsp; '
            f'Top-5 comps: <b style="color:{COL_HIGH}">{fmt_price(sn_t5_high)}</b>'
            f'&nbsp;·&nbsp; ★ = {sn_section}</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=480)
