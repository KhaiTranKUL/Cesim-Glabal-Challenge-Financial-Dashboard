"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         CESIM GLOBAL CHALLENGE – FINANCIAL ANALYSIS DASHBOARD v3           ║
║  Upload one or more round .xls files → ratios, valuation, risk & trends    ║
╚══════════════════════════════════════════════════════════════════════════════╝

HOW TO RUN
  pip install streamlit plotly xlrd pandas openpyxl numpy
  streamlit run Cesim_dashboard_v3_1_0_fixed.py

NEW IN v3
  • Region-level analysis tab: US, Europe, China
    - Regional Income Statements (revenue, EBITDA, EBIT, net profit)
    - Regional Balance Sheets (assets, equity, liabilities)
    - Regional market shares (Combustion / Hybrid / Electric / Hydrogen)
    - Revenue contribution % charts  (who earns where)
    - Regional margin comparison across all companies
    - Heatmap-style colour-coded regional P&L tables

RETAINED FROM v2
  • Multi-round upload → historical trend charts
  • Valuation tab: P/E, P/B, P/Sales, EV/EBITDA, EV/Sales
  • Capital & Risk tab: WACC, Altman Z-Score
  • Growth & Trends tab: YoY growth, SGR, Residual Income Model
  • DDM (zero-growth + constant-growth) in Valuation tab
  • Assumption sidebar: Re, g, tax override, forecast DPS, payout ratio

FORMULA ASSUMPTIONS (see CFO Cheat Sheet tab for full details)
  WACC    : (E/V)×Re + (D/V)×Rd×(1−t)  — E = market cap, D = LT+ST debt
  Altman Z: Public-firm model (1968); X4 uses market cap / book liabilities
  SGR     : ROE_after_tax × (1 − payout ratio)
  DDM-0   : P = DPS / Re  (zero growth; N/A when DPS = 0)
  DDM-g   : P = DPS×(1+g) / (Re−g)  (Gordon Growth Model; requires Re > g)
  RIM     : IV = BVps + ((EPS×(1+g)) − Re×BVps) / (Re−g)  (single-stage Ohlson 1995)
  EV      : Market Cap + LT Debt + ST Debt − Cash
"""

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 0 — Imports & page config
# ═══════════════════════════════════════════════════════════════════════════
import re
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np

st.set_page_config(
    page_title="Cesim Financial Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 1 — CSS
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
  .stApp { background:#F8FAFF; }
  section[data-testid="stSidebar"] { background:#1E2A4A; }
  section[data-testid="stSidebar"] * { color:#E8EDF8 !important; }
  section[data-testid="stSidebar"] .stSelectbox label,
  section[data-testid="stSidebar"] .stFileUploader label { color:#E8EDF8 !important; }

  .kpi-card {
    background:white; border-radius:12px; padding:16px 20px;
    box-shadow:0 2px 10px rgba(0,0,0,.08);
    border-top:4px solid #3B6FE8; margin-bottom:12px;
  }
  .kpi-card-warn  { border-top-color:#F59E0B !important; }
  .kpi-card-ok    { border-top-color:#10B981 !important; }
  .kpi-card-bad   { border-top-color:#EF4444 !important; }
  .kpi-label { font-size:12px; color:#6B7A99; font-weight:600;
               text-transform:uppercase; letter-spacing:.5px; }
  .kpi-value { font-size:26px; font-weight:700; color:#1E2A4A; margin:4px 0 2px; }
  .kpi-sub   { font-size:12px; color:#9AA5B8; }

  /* Region flag pills */
  .pill-us  { background:#DBEAFE; color:#1D4ED8; padding:2px 9px;
              border-radius:20px; font-size:12px; font-weight:700; }
  .pill-cn  { background:#FEE2E2; color:#991B1B; padding:2px 9px;
              border-radius:20px; font-size:12px; font-weight:700; }
  .pill-eu  { background:#D1FAE5; color:#065F46; padding:2px 9px;
              border-radius:20px; font-size:12px; font-weight:700; }

  .badge-green { background:#D1FAE5; color:#065F46; padding:3px 10px;
                 border-radius:20px; font-size:12px; font-weight:700; }
  .badge-amber { background:#FEF3C7; color:#92400E; padding:3px 10px;
                 border-radius:20px; font-size:12px; font-weight:700; }
  .badge-red   { background:#FEE2E2; color:#991B1B; padding:3px 10px;
                 border-radius:20px; font-size:12px; font-weight:700; }

  .section-header {
    font-size:20px; font-weight:700; color:#1E2A4A;
    border-left:5px solid #3B6FE8; padding-left:12px;
    margin:24px 0 16px;
  }
  .assume-box {
    background:#FFFBEB; border-left:4px solid #F59E0B;
    border-radius:8px; padding:10px 14px; margin-bottom:12px;
    font-size:13px; color:#78350F;
  }
  .styled-table { width:100%; border-collapse:collapse; }
  .styled-table th { background:#1E2A4A; color:white; padding:8px 12px;
                     font-size:13px; text-align:left; }
  .styled-table td { padding:7px 12px; font-size:13px;
                     border-bottom:1px solid #E8EDF8; }
  .styled-table tr:hover td { background:#F0F4FF; }
  .my-company { background:#EFF6FF !important; font-weight:700; }
  h1,h2,h3 { color:#1E2A4A; }
  .stTabs [data-baseweb="tab"] { font-size:14px; font-weight:600; }
  .stTabs [aria-selected="true"] { color:#3B6FE8; border-bottom-color:#3B6FE8; }
  .cheat-box {
    background:#F0F4FF; border-radius:10px; padding:14px 18px;
    border-left:4px solid #3B6FE8; margin-bottom:12px;
  }
  .cheat-title   { font-weight:700; color:#1E2A4A; font-size:14px; }
  .cheat-formula { font-family:monospace; font-size:12px; color:#3B6FE8; }
  .cheat-body    { font-size:13px; color:#374151; margin-top:4px; }
  .assume-tag    { background:#E0E7FF; color:#3730A3; padding:2px 8px;
                   border-radius:10px; font-size:11px; font-weight:600; }
  .na-cell       { color:#9CA3AF; font-style:italic; font-size:12px; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 2 — Default assumptions
# ═══════════════════════════════════════════════════════════════════════════
DEFAULT_ASSUMPTIONS: dict = {
    "cost_of_equity":        0.09,
    "lt_growth":             0.03,
    "tax_rate_override":     0.00,
    "payout_ratio_override": 0.00,
    "forecast_dps":          0.00,
    "use_market_cap":        True,
}


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 3 — Excel row-index constants
# ═══════════════════════════════════════════════════════════════════════════
#
# Each constant maps to an absolute row index (0-based) in the Cesim XLS.
# Regions have separate Income Statements and Balance Sheets; row offsets
# were verified against results-r03.xls (1 098 rows × 6 cols).
#
# Key structural differences between regions:
#   USA    – full IS (incl. inhouse/contract mfg, warranty, internal loans)
#   China  – full IS but no warranty row; no LT debt on BS
#   Europe – distribution-only IS (no inhouse/contract mfg, no R&D, no
#            warranty); no LT debt, no inventory on BS

# ── Global (existing, unchanged) ──────────────────────────────────────────
_G = {                        # row: field
    "sales_rev":          6,
    "inhouse_mfg":        9,
    "feature_costs":      10,
    "contract_mfg":       11,
    "transport":          12,
    "rd_costs":           13,
    "promotion":          14,
    "warranty":           15,
    "administration":     16,
    "total_costs":        17,
    "ebitda":             19,
    "depreciation":       20,
    "ebit":               22,
    "net_fin_exp":        23,
    "profit_before_tax":  25,
    "income_taxes":       26,
    "net_profit":         28,
    # BS
    "fixed_assets":       34,
    "inventory":          35,
    "receivables":        36,
    "cash":               37,
    "total_assets":       38,
    "share_capital":      41,
    "retained_earnings":  44,
    "total_equity":       45,
    "lt_debt":            48,
    "st_debt":            49,
    "payables":           50,
    "total_liabilities":  51,
}

# ── USA regional ──────────────────────────────────────────────────────────
# IS header: row 55  |  BS header: row 87
_USA_IS = {
    "sales_rev":         62,   # "Sales revenue total"
    "inhouse_mfg":       65,
    "feature_costs":     66,
    "contract_mfg":      67,
    "transport":         68,
    "rd_costs":          69,
    "promotion":         70,
    "warranty":          71,
    "administration":    72,
    "total_costs":       74,
    "ebitda":            76,
    "depreciation":      77,
    "ebit":              79,
    "net_fin_exp":       80,
    "profit_before_tax": 82,
    "income_taxes":      83,
    "net_profit":        85,
}
_USA_BS = {
    "fixed_assets":      91,
    "inventory":         92,
    "receivables":       93,
    "cash":              94,
    "total_assets":      95,
    "share_capital":     98,
    "retained_earnings": 101,
    "total_equity":      102,
    "lt_debt":           105,
    "st_debt":           106,
    "payables":          108,
    "total_liabilities": 109,
}
_USA_MS_ROWS = {
    "combustion": 365, "hybrid": 366, "electric": 367,
    "hydrogen": 368,   "total":  369,
}

# ── China regional ────────────────────────────────────────────────────────
# IS header: row 143  |  BS header: row 173
# Note: no warranty row in China IS; no LT debt in China BS
_CHN_IS = {
    "sales_rev":         149,
    "inhouse_mfg":       152,
    "feature_costs":     153,
    "contract_mfg":      154,
    "transport":         155,
    "rd_costs":          156,
    "promotion":         157,
    "warranty":          None,   # absent in China IS
    "administration":    158,
    "total_costs":       160,
    "ebitda":            162,
    "depreciation":      163,
    "ebit":              165,
    "net_fin_exp":       166,
    "profit_before_tax": 168,
    "income_taxes":      169,
    "net_profit":        171,
}
_CHN_BS = {
    "fixed_assets":      177,
    "inventory":         178,
    "receivables":       179,
    "cash":              180,
    "total_assets":      181,
    "share_capital":     184,
    "retained_earnings": 186,
    "total_equity":      187,
    "lt_debt":           None,   # absent in China BS
    "st_debt":           190,
    "payables":          192,
    "total_liabilities": 193,
}
_CHN_MS_ROWS = {
    "combustion": 403, "hybrid": 404, "electric": 405,
    "hydrogen": 406,   "total":  407,
}

# ── Europe regional ───────────────────────────────────────────────────────
# IS header: row 224  |  BS header: row 249
# Europe is a distribution hub: no inhouse/contract mfg, no R&D, no warranty
_EU_IS = {
    "sales_rev":         228,
    "inhouse_mfg":       None,   # absent in Europe IS
    "feature_costs":     231,
    "contract_mfg":      None,   # absent in Europe IS
    "transport":         232,
    "rd_costs":          None,   # absent in Europe IS
    "promotion":         233,
    "warranty":          None,   # absent in Europe IS
    "administration":    234,
    "total_costs":       236,
    "ebitda":            238,
    "depreciation":      239,
    "ebit":              241,
    "net_fin_exp":       242,
    "profit_before_tax": 244,
    "income_taxes":      245,
    "net_profit":        247,
}
_EU_BS = {
    "fixed_assets":      253,
    "inventory":         None,   # absent in Europe BS
    "receivables":       254,
    "cash":              255,
    "total_assets":      256,
    "share_capital":     259,
    "retained_earnings": 261,
    "total_equity":      262,
    "lt_debt":           None,   # absent in Europe BS
    "st_debt":           265,
    "payables":          267,
    "total_liabilities": 268,
}
_EU_MS_ROWS = {
    "combustion": 441, "hybrid": 442, "electric": 443,
    "hydrogen": 444,   "total":  445,
}

REGIONS = {
    "USA":    {"is": _USA_IS, "bs": _USA_BS, "ms": _USA_MS_ROWS,
               "pill": "<span class='pill-us'>🇺🇸 USA</span>",
               "color": "#3B6FE8"},
    "China":  {"is": _CHN_IS, "bs": _CHN_BS, "ms": _CHN_MS_ROWS,
               "pill": "<span class='pill-cn'>🇨🇳 China</span>",
               "color": "#EF4444"},
    "Europe": {"is": _EU_IS,  "bs": _EU_BS,  "ms": _EU_MS_ROWS,
               "pill": "<span class='pill-eu'>🇪🇺 Europe</span>",
               "color": "#10B981"},
}
REGION_COLORS = {r: REGIONS[r]["color"] for r in REGIONS}


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 4 — Excel parsing
# ═══════════════════════════════════════════════════════════════════════════

def _get_val(df: pd.DataFrame, row: int, col: int) -> float:
    """Safe numeric extraction; returns 0.0 on missing/NaN."""
    try:
        v = df.iloc[row, col]
        return float(v) if pd.notna(v) else 0.0
    except Exception:
        return 0.0

def _get_str(df: pd.DataFrame, row: int, col: int) -> str:
    try:
        v = df.iloc[row, col]
        return str(v) if pd.notna(v) else ""
    except Exception:
        return ""

def _round_num(df: pd.DataFrame) -> int:
    title = _get_str(df, 0, 0)
    m = re.search(r'Round\s+(\d+)', title, re.IGNORECASE)
    return int(m.group(1)) if m else 1

# Aliases for backward compat
get_val = _get_val
get_str = _get_str
parse_round_title = _round_num

def _sdiv(a: float, b: float, default=None):
    """Safe division; returns `default` when b is zero or None."""
    try:
        if b and b != 0:
            return a / b
        return default
    except Exception:
        return default


def _parse_region_section(df: pd.DataFrame, col: int,
                           is_map: dict, bs_map: dict, ms_map: dict) -> dict:
    """
    Extract all IS, BS and market-share fields for one region / one company.

    Parameters
    ----------
    df       : full DataFrame for the round
    col      : company column index (1-based)
    is_map   : {field: row_index | None}  for the regional Income Statement
    bs_map   : {field: row_index | None}  for the regional Balance Sheet
    ms_map   : {field: row_index}          for regional market shares

    Returns
    -------
    dict of raw values and first-level derived metrics (margins, %)
    """
    def gv(row):
        if row is None:
            return 0.0
        return _get_val(df, row, col)

    # ── Income Statement ──────────────────────────────────────────────────
    sales_rev         = gv(is_map["sales_rev"])
    inhouse_mfg       = gv(is_map["inhouse_mfg"])
    feature_costs     = gv(is_map["feature_costs"])
    contract_mfg      = gv(is_map["contract_mfg"])
    transport         = gv(is_map["transport"])
    rd_costs          = gv(is_map["rd_costs"])
    promotion         = gv(is_map["promotion"])
    warranty          = gv(is_map["warranty"])
    administration    = gv(is_map["administration"])
    total_costs       = gv(is_map["total_costs"])
    ebitda            = gv(is_map["ebitda"])
    depreciation      = gv(is_map["depreciation"])
    ebit              = gv(is_map["ebit"])
    net_fin_exp       = gv(is_map["net_fin_exp"])
    profit_before_tax = gv(is_map["profit_before_tax"])
    income_taxes      = gv(is_map["income_taxes"])
    net_profit        = gv(is_map["net_profit"])

    # ── Balance Sheet ─────────────────────────────────────────────────────
    fixed_assets      = gv(bs_map["fixed_assets"])
    inventory         = gv(bs_map["inventory"])
    receivables       = gv(bs_map["receivables"])
    cash              = gv(bs_map["cash"])
    total_assets      = gv(bs_map["total_assets"])
    share_capital     = gv(bs_map["share_capital"])
    retained_earnings = gv(bs_map["retained_earnings"])
    total_equity      = gv(bs_map["total_equity"])
    lt_debt           = gv(bs_map["lt_debt"])
    st_debt           = gv(bs_map["st_debt"])
    payables          = gv(bs_map["payables"])
    total_liabilities = gv(bs_map["total_liabilities"])

    # ── Market shares ─────────────────────────────────────────────────────
    ms_combustion = gv(ms_map["combustion"])
    ms_hybrid     = gv(ms_map["hybrid"])
    ms_electric   = gv(ms_map["electric"])
    ms_hydrogen   = gv(ms_map["hydrogen"])
    ms_total      = gv(ms_map["total"])

    # ── Derived ───────────────────────────────────────────────────────────
    ebitda_margin = _sdiv(ebitda, sales_rev, 0.0) * 100
    ebit_margin   = _sdiv(ebit,   sales_rev, 0.0) * 100
    net_margin    = _sdiv(net_profit, sales_rev, 0.0) * 100
    cost_ratio    = _sdiv(total_costs, sales_rev, 0.0) * 100
    gross_profit  = sales_rev - total_costs  # before fin exp / tax

    return dict(
        # IS raw
        sales_rev=sales_rev, inhouse_mfg=inhouse_mfg,
        feature_costs=feature_costs, contract_mfg=contract_mfg,
        transport=transport, rd_costs=rd_costs,
        promotion=promotion, warranty=warranty,
        administration=administration, total_costs=total_costs,
        ebitda=ebitda, depreciation=depreciation, ebit=ebit,
        net_fin_exp=net_fin_exp, profit_before_tax=profit_before_tax,
        income_taxes=income_taxes, net_profit=net_profit,
        # BS raw
        fixed_assets=fixed_assets, inventory=inventory,
        receivables=receivables, cash=cash, total_assets=total_assets,
        share_capital=share_capital, retained_earnings=retained_earnings,
        total_equity=total_equity, lt_debt=lt_debt,
        st_debt=st_debt, payables=payables, total_liabilities=total_liabilities,
        # Market shares
        ms_combustion=ms_combustion, ms_hybrid=ms_hybrid,
        ms_electric=ms_electric, ms_hydrogen=ms_hydrogen, ms_total=ms_total,
        # Derived margins
        ebitda_margin=ebitda_margin, ebit_margin=ebit_margin,
        net_margin=net_margin, cost_ratio=cost_ratio,
        gross_profit=gross_profit,
    )


def parse_cesim_xls(uploaded_file) -> tuple[dict, int]:
    """
    Parse one Cesim round-results .xls/.xlsx file.

    Returns
    -------
    data      : {company_name: raw_data_dict}  — raw + derived + regional metrics
    round_num : int
    """
    df = pd.read_excel(uploaded_file, sheet_name=0, engine="xlrd", header=None)
    round_num = _round_num(df)

    companies = []
    for c in range(1, 7):
        nm = _get_str(df, 4, c)
        if nm and nm != "nan":
            companies.append((c, nm))

    data: dict = {}
    for col, name in companies:
        # ── A: Global Income Statement ────────────────────────────────────
        sales_rev         = _get_val(df, 6,  col)
        inhouse_mfg       = _get_val(df, 9,  col)
        feature_costs     = _get_val(df, 10, col)
        contract_mfg      = _get_val(df, 11, col)
        transport_tariffs = _get_val(df, 12, col)
        rd_costs          = _get_val(df, 13, col)
        promotion         = _get_val(df, 14, col)
        warranty          = _get_val(df, 15, col)
        administration    = _get_val(df, 16, col)
        total_costs       = _get_val(df, 17, col)
        ebitda            = _get_val(df, 19, col)
        depreciation      = _get_val(df, 20, col)
        ebit              = _get_val(df, 22, col)
        net_fin_exp       = _get_val(df, 23, col)
        profit_before_tax = _get_val(df, 25, col)
        income_taxes      = _get_val(df, 26, col)
        net_profit        = _get_val(df, 28, col)

        # ── B: Global Balance Sheet ───────────────────────────────────────
        fixed_assets      = _get_val(df, 34, col)
        inventory         = _get_val(df, 35, col)
        receivables       = _get_val(df, 36, col)
        cash              = _get_val(df, 37, col)
        total_assets      = _get_val(df, 38, col)
        share_capital     = _get_val(df, 41, col)
        retained_earnings = _get_val(df, 44, col)
        total_equity      = _get_val(df, 45, col)
        lt_debt           = _get_val(df, 48, col)
        st_debt           = _get_val(df, 49, col)
        payables          = _get_val(df, 50, col)
        total_liabilities = _get_val(df, 51, col)

        # ── C: Market / KPIs block ────────────────────────────────────────
        cum_earnings   = _get_val(df, 297, col)
        market_cap     = _get_val(df, 298, col)
        shares_out     = _get_val(df, 299, col)
        share_price    = _get_val(df, 300, col)
        avg_price      = _get_val(df, 301, col)
        div_yield      = _get_val(df, 302, col)
        pe_ratio       = _get_val(df, 303, col)
        tsr            = _get_val(df, 304, col)
        gross_margin_f = _get_val(df, 307, col)
        ebitda_margin_f= _get_val(df, 308, col)
        ebit_margin_f  = _get_val(df, 309, col)
        ros_f          = _get_val(df, 310, col)
        equity_ratio_f = _get_val(df, 311, col)
        nd_equity_f    = _get_val(df, 312, col)
        roce_f         = _get_val(df, 313, col)
        roe_f          = _get_val(df, 314, col)
        eps            = _get_val(df, 315, col)

        # ── D: Interest rates ─────────────────────────────────────────────
        int_usa_long  = _get_val(df, 318, col)
        int_usa_short = _get_val(df, 319, col)
        int_cn_short  = _get_val(df, 320, col)
        int_eu_short  = _get_val(df, 321, col)

        # ── E: Global market shares ───────────────────────────────────────
        ms_combustion = _get_val(df, 327, col)
        ms_hybrid     = _get_val(df, 328, col)
        ms_electric   = _get_val(df, 329, col)
        ms_hydrogen   = _get_val(df, 330, col)
        ms_total      = _get_val(df, 331, col)

        # ── F: HR ─────────────────────────────────────────────────────────
        wage_month    = _get_val(df, 451, col)
        train_budget  = _get_val(df, 452, col)
        hr_eff        = _get_val(df, 454, col)
        rdstaff_prev  = _get_val(df, 457, col)
        rdstaff_now   = _get_val(df, 460, col)
        vol_turnover  = _get_val(df, 462, col)
        hr_total_cost = _get_val(df, 473, col)

        # ── G: ESG ────────────────────────────────────────────────────────
        esg_e   = _get_val(df, 498, col)
        esg_s   = _get_val(df, 499, col)
        esg_g   = _get_val(df, 500, col)
        esg_rep = _get_val(df, 501, col)

        # ══ DERIVED METRICS (global) ══════════════════════════════════════
        current_assets = inventory + receivables + cash
        current_liab   = st_debt + payables
        lt_fin_res     = total_equity + lt_debt

        nwc          = lt_fin_res - fixed_assets
        net_cash     = cash - st_debt
        current_ratio= _sdiv(current_assets, current_liab, 0.0)
        quick_ratio  = _sdiv(current_assets - inventory, current_liab, 0.0)
        cos          = inhouse_mfg + feature_costs + contract_mfg + transport_tariffs
        stock_days   = _sdiv(inventory, cos, 0.0) * 365
        rec_days     = _sdiv(receivables, sales_rev, 0.0) * 365
        cred_purch   = feature_costs + contract_mfg + transport_tariffs + promotion + warranty
        pay_days     = _sdiv(payables, cred_purch, 0.0) * 365
        cash_cycle   = stock_days + rec_days - pay_days

        gross_sales_margin = _sdiv(ebitda, sales_rev, 0.0) * 100
        net_sales_margin   = _sdiv(ebit,   sales_rev, 0.0) * 100
        gross_roa    = _sdiv(ebitda, total_assets, 0.0) * 100
        net_roa      = _sdiv(ebit,   total_assets, 0.0) * 100
        roe_bt       = _sdiv(profit_before_tax, total_equity, 0.0) * 100
        roe_at       = _sdiv(net_profit,        total_equity, 0.0) * 100
        ce           = total_assets - current_liab
        roce         = _sdiv(ebit, ce, 0.0) * 100

        avg_cost_debt = _sdiv(net_fin_exp, lt_debt + st_debt, 0.0) * 100
        fin_lev_sum   = roe_bt - net_roa
        _r1           = _sdiv(profit_before_tax, ebit,         0.0)
        _r2           = _sdiv(total_assets,      total_equity, 0.0)
        fin_lev_mult  = net_roa / 100 * _r1 * _r2 * 100

        fin_independence = _sdiv(total_equity, total_assets, 0.0) * 100
        debt_ratio       = 100 - fin_independence
        fin_stability    = _sdiv(lt_fin_res, total_assets, 0.0) * 100
        net_debt         = lt_debt + st_debt - cash
        gearing          = _sdiv(net_debt, total_equity, 0.0) * 100

        def _pct(v): return _sdiv(v, sales_rev, 0.0) * 100

        # ── H: Regional data (NEW in v3) ─────────────────────────────────
        region_data: dict = {}
        for rname, rconf in REGIONS.items():
            region_data[rname] = _parse_region_section(
                df, col,
                rconf["is"], rconf["bs"], rconf["ms"],
            )

        # Revenue share per region (% of global)
        for rname in REGIONS:
            rrev = region_data[rname]["sales_rev"]
            region_data[rname]["rev_pct_global"] = (
                _sdiv(rrev, sales_rev, 0.0) * 100
            )

        data[name] = dict(
            name=name,
            # IS (raw global)
            sales_rev=sales_rev, inhouse_mfg=inhouse_mfg,
            feature_costs=feature_costs, contract_mfg=contract_mfg,
            transport_tariffs=transport_tariffs, rd_costs=rd_costs,
            promotion=promotion, warranty=warranty,
            administration=administration, total_costs=total_costs,
            ebitda=ebitda, depreciation=depreciation, ebit=ebit,
            net_fin_exp=net_fin_exp, profit_before_tax=profit_before_tax,
            income_taxes=income_taxes, net_profit=net_profit,
            # BS (raw global)
            fixed_assets=fixed_assets, inventory=inventory,
            receivables=receivables, cash=cash, total_assets=total_assets,
            share_capital=share_capital, retained_earnings=retained_earnings,
            total_equity=total_equity, lt_debt=lt_debt, st_debt=st_debt,
            payables=payables, total_liabilities=total_liabilities,
            # Derived — liquidity
            current_assets=current_assets, current_liab=current_liab,
            lt_fin_res=lt_fin_res, nwc=nwc, net_cash=net_cash,
            current_ratio=current_ratio, quick_ratio=quick_ratio,
            stock_days=stock_days, rec_days=rec_days,
            pay_days=pay_days, cash_cycle=cash_cycle,
            # Derived — profitability
            gross_sales_margin=gross_sales_margin, net_sales_margin=net_sales_margin,
            gross_roa=gross_roa, net_roa=net_roa,
            roe_bt=roe_bt, roe_at=roe_at, roce=roce,
            avg_cost_debt=avg_cost_debt,
            fin_lev_sum=fin_lev_sum, fin_lev_mult=fin_lev_mult,
            # Derived — solvency
            fin_independence=fin_independence, debt_ratio=debt_ratio,
            fin_stability=fin_stability, net_debt=net_debt, gearing=gearing,
            # Cost %
            inhouse_pct=_pct(inhouse_mfg), feature_pct=_pct(feature_costs),
            contract_pct=_pct(contract_mfg), transport_pct=_pct(transport_tariffs),
            rd_pct=_pct(rd_costs), promo_pct=_pct(promotion),
            warranty_pct=_pct(warranty), admin_pct=_pct(administration),
            total_cost_pct=_pct(total_costs),
            # Market (global KPI block)
            cum_earnings=cum_earnings, market_cap=market_cap,
            shares_out=shares_out, share_price=share_price,
            avg_price=avg_price, div_yield=div_yield,
            pe_ratio=pe_ratio, tsr=tsr,
            gross_margin_f=gross_margin_f, ebitda_margin_f=ebitda_margin_f,
            ebit_margin_f=ebit_margin_f, ros_f=ros_f,
            equity_ratio_f=equity_ratio_f, nd_equity_f=nd_equity_f,
            roce_f=roce_f, roe_f=roe_f, eps=eps,
            int_usa_long=int_usa_long, int_usa_short=int_usa_short,
            int_cn_short=int_cn_short, int_eu_short=int_eu_short,
            ms_combustion=ms_combustion, ms_hybrid=ms_hybrid,
            ms_electric=ms_electric, ms_hydrogen=ms_hydrogen, ms_total=ms_total,
            # HR
            wage_month=wage_month, train_budget=train_budget, hr_eff=hr_eff,
            rdstaff_prev=rdstaff_prev, rdstaff_now=rdstaff_now,
            vol_turnover=vol_turnover, hr_total_cost=hr_total_cost,
            # ESG
            esg_e=esg_e, esg_s=esg_s, esg_g=esg_g, esg_rep=esg_rep,
            # Regional data (NEW v3) — nested dict keyed by region name
            regions=region_data,
        )

    return data, round_num


def load_all_rounds(uploaded_files) -> dict[int, dict]:
    """Parse every uploaded file and return a round-indexed dict."""
    rounds_data: dict[int, dict] = {}
    for f in uploaded_files:
        try:
            rdata, rnum = parse_cesim_xls(f)
            rounds_data[rnum] = rdata
        except Exception as exc:
            st.sidebar.warning(f"⚠️ Could not parse **{f.name}**: {exc}")
    return rounds_data


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 5 — Advanced metric computation (v2, unchanged)
# ═══════════════════════════════════════════════════════════════════════════

def _yoy(curr: float, prev: float | None) -> float | None:
    if prev is None or abs(prev) < 1e-9:
        return None
    return (curr - prev) / abs(prev) * 100


def compute_advanced_metrics(d: dict, assumptions: dict,
                              prior_d: dict | None = None) -> dict:
    Re  = assumptions["cost_of_equity"]
    g   = assumptions["lt_growth"]
    adv: dict = {}

    ev = d["market_cap"] + d["net_debt"]
    adv["adv_ev"]        = ev
    adv["adv_ev_ebitda"] = _sdiv(ev, d["ebitda"])
    adv["adv_ev_sales"]  = _sdiv(ev, d["sales_rev"])

    bvps = _sdiv(d["total_equity"], d["shares_out"])
    adv["adv_bv_per_share"] = bvps
    adv["adv_pb"] = _sdiv(d["share_price"], bvps) if bvps and bvps > 0 else None
    adv["adv_ps"] = _sdiv(d["market_cap"], d["sales_rev"])
    adv["adv_pe"] = d["pe_ratio"] if d["pe_ratio"] else None

    if assumptions["tax_rate_override"] > 0:
        eff_tax = assumptions["tax_rate_override"]
    elif d["profit_before_tax"] > 0:
        eff_tax = min(_sdiv(d["income_taxes"], d["profit_before_tax"], 0.25), 0.50)
    else:
        eff_tax = 0.25
    adv["adv_eff_tax_rate"] = eff_tax * 100

    E_wacc = d["market_cap"] if assumptions["use_market_cap"] else d["total_equity"]
    D_wacc = d["lt_debt"] + d["st_debt"]
    V_wacc = E_wacc + D_wacc if (E_wacc + D_wacc) > 0 else 1
    Rd     = d["avg_cost_debt"] / 100
    wacc_decimal = (E_wacc / V_wacc) * Re + (D_wacc / V_wacc) * Rd * (1 - eff_tax)
    adv["adv_wacc"]               = wacc_decimal * 100
    adv["adv_wacc_equity_wt"]     = E_wacc / V_wacc * 100
    adv["adv_wacc_debt_wt"]       = D_wacc / V_wacc * 100
    adv["adv_wacc_after_tax_rd"]  = Rd * (1 - eff_tax) * 100
    adv["adv_roce_wacc_spread"]   = d["roce"] - wacc_decimal * 100

    ta = d["total_assets"]
    if ta > 0:
        X1 = d["nwc"] / ta
        X2 = d["retained_earnings"] / ta
        X3 = d["ebit"] / ta
        book_liab = d["lt_debt"] + d["st_debt"] + d["payables"]
        X4 = _sdiv(d["market_cap"], book_liab, 0.0)
        X5 = d["sales_rev"] / ta
        z  = 1.2*X1 + 1.4*X2 + 3.3*X3 + 0.6*X4 + 1.0*X5
        adv["adv_altman_z"]    = z
        adv["adv_altman_x1"]   = X1
        adv["adv_altman_x2"]   = X2
        adv["adv_altman_x3"]   = X3
        adv["adv_altman_x4"]   = X4
        adv["adv_altman_x5"]   = X5
        adv["adv_altman_zone"] = ("Safe","green") if z > 2.99 else \
                                 (("Grey Zone","amber") if z >= 1.81 else ("Distress","red"))
    else:
        for k in ["adv_altman_z","adv_altman_x1","adv_altman_x2",
                  "adv_altman_x3","adv_altman_x4","adv_altman_x5"]:
            adv[k] = None
        adv["adv_altman_zone"] = ("N/A","amber")

    actual_dps = d["div_yield"] / 100 * d["share_price"] if d["share_price"] > 0 else 0.0
    dps = assumptions["forecast_dps"] if assumptions["forecast_dps"] > 0 else actual_dps

    if assumptions["payout_ratio_override"] > 0:
        payout = min(float(assumptions["payout_ratio_override"]), 1.0)
    elif d["net_profit"] > 0 and actual_dps > 0:
        total_div = actual_dps * d["shares_out"]
        payout = min(_sdiv(total_div, d["net_profit"], 0.0), 1.0)
    else:
        payout = 0.0

    retention = 1.0 - payout
    sgr = d["roe_at"] / 100 * retention
    adv["adv_sgr"]             = sgr * 100
    adv["adv_payout_ratio"]    = payout * 100
    adv["adv_retention_ratio"] = retention * 100
    adv["adv_dps"]             = dps
    adv["adv_actual_dps"]      = actual_dps

    # DDM uses the dividend base (actual or overridden forecast) as D0.
    # Constant-growth DDM prices the next dividend D1 = D0 × (1 + g).
    adv["adv_ddm_zero"] = _sdiv(dps, Re) if dps > 0 else None
    if dps > 0 and Re > g:
        adv["adv_ddm_const"] = dps * (1 + g) / (Re - g)
    else:
        adv["adv_ddm_const"] = None
    g_sgr = min(sgr, Re - 0.005) if sgr < Re else Re - 0.005
    if dps > 0 and Re > g_sgr and g_sgr >= 0:
        adv["adv_ddm_sgr"] = dps * (1 + g_sgr) / (Re - g_sgr)
    else:
        adv["adv_ddm_sgr"] = None

    # RIM needs next-period EPS. We estimate EPS1 by growing current EPS by g.
    if bvps and bvps > 0 and Re > g:
        eps_next = d["eps"] * (1 + g)
        ri_ps = eps_next - Re * bvps
        adv["adv_rim_iv"]      = bvps + ri_ps / (Re - g)
        adv["adv_ri_per_share"]= ri_ps
        adv["adv_normal_eps"]  = Re * bvps
    else:
        adv["adv_rim_iv"]      = None
        adv["adv_ri_per_share"]= None
        adv["adv_normal_eps"]  = None

    if prior_d:
        adv["adv_rev_growth"]      = _yoy(d["sales_rev"],    prior_d["sales_rev"])
        adv["adv_earnings_growth"] = _yoy(d["net_profit"],   prior_d["net_profit"])
        adv["adv_ebitda_growth"]   = _yoy(d["ebitda"],       prior_d["ebitda"])
        adv["adv_eps_growth"]      = _yoy(d["eps"],          prior_d["eps"])
        adv["adv_equity_growth"]   = _yoy(d["total_equity"], prior_d["total_equity"])
        adv["adv_price_growth"]    = _yoy(d["share_price"],  prior_d["share_price"])
        adv["adv_ms_growth"]       = _yoy(d["ms_total"],     prior_d["ms_total"])
    else:
        for k in ["adv_rev_growth","adv_earnings_growth","adv_ebitda_growth",
                  "adv_eps_growth","adv_equity_growth","adv_price_growth","adv_ms_growth"]:
            adv[k] = None

    return adv


def enrich_all_companies(data: dict, assumptions: dict, prior_data: dict | None) -> None:
    """In-place: compute and merge advanced metrics for every company."""
    for name, d in data.items():
        prior_d = prior_data.get(name) if prior_data else None
        adv     = compute_advanced_metrics(d, assumptions, prior_d)
        d.update(adv)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 6 — Display helpers
# ═══════════════════════════════════════════════════════════════════════════

COLORS    = ["#3B6FE8","#F59E0B","#10B981","#EF4444","#8B5CF6"]
COLOR_MAP: dict = {}

def assign_colors(companies: list[str]) -> None:
    for i, c in enumerate(companies):
        COLOR_MAP[c] = COLORS[i % len(COLORS)]

# Formatters
def fmt_usd(v: float, decimals: int = 0) -> str:
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:,.{decimals}f}M"
    if abs(v) >= 1_000:
        return f"${v/1_000:,.{decimals}f}K"
    return f"${v:,.{decimals}f}"

def pct_str(v: float, d: int = 2) -> str:
    return f"{v:.{d}f}%"

def fmt_pct(v: float, d: int = 2) -> str:
    return f"{v:+.{d}f}%"

def fmt_x(v: float | None, d: int = 2) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return '<span class="na-cell">N/A</span>'
    return f"{v:.{d}f}×"

def fmt_opt(v: float | None, fn=None) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return '<span class="na-cell">N/A</span>'
    return fn(v) if fn else f"{v:.2f}"

def fmt_growth(v: float | None) -> str:
    if v is None:
        return '<span class="na-cell">N/A (need prior round)</span>'
    color = "#10B981" if v >= 0 else "#EF4444"
    arrow = "▲" if v >= 0 else "▼"
    return f'<span style="color:{color};font-weight:700">{arrow} {v:+.1f}%</span>'

def alert(val, thresholds, labels=("🟢 Good","🟡 Caution","🔴 Poor")):
    if val >= thresholds[0]:
        return f'<span class="badge-green">{labels[0]}</span>'
    if val >= thresholds[1]:
        return f'<span class="badge-amber">{labels[1]}</span>'
    return f'<span class="badge-red">{labels[2]}</span>'

def alert_low(val, t1, t2, labels=("🟢 Good","🟡 Caution","🔴 High")):
    if val <= t1:
        return f'<span class="badge-green">{labels[0]}</span>'
    if val <= t2:
        return f'<span class="badge-amber">{labels[1]}</span>'
    return f'<span class="badge-red">{labels[2]}</span>'

def zone_badge(zone_tuple: tuple) -> str:
    zone, color = zone_tuple
    cls = f"badge-{color}"
    return f'<span class="{cls}">{zone}</span>'

def rank_badge(r: int) -> str:
    return {1:"🥇 #1", 2:"🥈 #2", 3:"🥉 #3"}.get(r, f"#{r}")

def rank_companies(data: dict, field: str, higher_is_better: bool = True) -> dict:
    vals = [(n, d.get(field, 0) or 0) for n, d in data.items()]
    vals.sort(key=lambda x: x[1], reverse=higher_is_better)
    return {name: i+1 for i, (name, _) in enumerate(vals)}

def bar_chart(title, companies, field, data, fmt=None):
    vals  = [data[c].get(field, 0) or 0 for c in companies]
    clrs  = [COLOR_MAP[c] for c in companies]
    texts = [fmt(v) if fmt else f"{v:.2f}" for v in vals]
    fig = go.Figure(go.Bar(
        x=companies, y=vals, text=texts, textposition="outside",
        marker_color=clrs,
    ))
    fig.update_layout(
        title=title, height=320, margin=dict(t=40,b=20,l=20,r=20),
        plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(gridcolor="#E8EDF8"),
        xaxis=dict(tickangle=-20),
        font=dict(family="sans-serif", color="#1E2A4A"),
    )
    return fig

def bar_chart_optional(title, companies, field, data, fmt_fn=None):
    vals, clrs, texts, cnames = [], [], [], []
    for c in companies:
        v = data[c].get(field)
        if v is not None and not (isinstance(v, float) and np.isnan(v)):
            vals.append(v)
            clrs.append(COLOR_MAP[c])
            texts.append(fmt_fn(v) if fmt_fn else f"{v:.2f}")
            cnames.append(c)
    if not vals:
        return None
    fig = go.Figure(go.Bar(x=cnames, y=vals, text=texts,
                           textposition="outside", marker_color=clrs))
    fig.update_layout(
        title=title, height=320, margin=dict(t=40,b=20,l=20,r=20),
        plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(gridcolor="#E8EDF8"),
        xaxis=dict(tickangle=-20),
        font=dict(family="sans-serif", color="#1E2A4A"),
    )
    return fig

def radar_chart(title, categories, values_dict, companies):
    fig = go.Figure()
    for c in companies:
        vals = values_dict[c]
        fig.add_trace(go.Scatterpolar(
            r=vals + [vals[0]], theta=categories + [categories[0]],
            fill="toself", name=c, line_color=COLOR_MAP[c], opacity=0.6,
        ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, gridcolor="#E8EDF8")),
        title=title, height=380, showlegend=True,
        paper_bgcolor="white", font=dict(color="#1E2A4A"),
        margin=dict(t=50,b=20,l=20,r=20),
    )
    return fig

def _table_header(companies, my_company, extra_cols=None):
    h = ['<table class="styled-table"><thead><tr><th>Metric</th>']
    for c in companies:
        sty = "background:#2A3F7A;color:white;" if c == my_company else ""
        h.append(f'<th style="{sty}">{c}</th>')
    for col in (extra_cols or []):
        h.append(f"<th>{col}</th>")
    h.append("</tr></thead><tbody>")
    return "".join(h)

def _table_row(label, companies, data, field, fmt_fn, my_company,
               extra_cells=None, bold=False):
    lbl = f"<b>{label}</b>" if bold else label
    r = [f"<tr><td>{lbl}</td>"]
    for c in companies:
        cls = "my-company" if c == my_company else ""
        v   = data[c].get(field)
        r.append(f'<td class="{cls}">{fmt_fn(v)}</td>')
    for cell in (extra_cells or []):
        r.append(f"<td>{cell}</td>")
    r.append("</tr>")
    return "".join(r)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 7 — Sidebar
# ═══════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📊 Cesim Dashboard v3")
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "Upload Round Results (.xls)",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        help="Upload one or more round files for trend analysis",
    )

    with st.expander("⚙️ Valuation Assumptions", expanded=False):
        st.caption("Used for WACC, DDM, RIM, and Z-score calculations.")
        Re_pct = st.number_input(
            "Cost of Equity Re (%)",
            value=9.0, min_value=1.0, max_value=40.0, step=0.5,
        )
        g_pct = st.number_input(
            "Long-term Growth Rate g (%)",
            value=3.0, min_value=0.0, max_value=20.0, step=0.5,
        )
        tax_override_pct = st.number_input(
            "Tax Rate Override (0 = auto)",
            value=0.0, min_value=0.0, max_value=50.0, step=1.0,
        )
        payout_override = st.number_input(
            "Dividend Payout Ratio (0 = use actual)",
            value=0.0, min_value=0.0, max_value=1.0, step=0.05,
        )
        forecast_dps = st.number_input(
            "Forecast DPS for DDM (0 = actual)",
            value=0.0, min_value=0.0, step=0.10,
        )
        use_mktcap = st.checkbox(
            "Use Market Cap for WACC equity (recommended)",
            value=True,
        )

    assumptions: dict = {
        "cost_of_equity":        Re_pct / 100,
        "lt_growth":             g_pct  / 100,
        "tax_rate_override":     tax_override_pct / 100,
        "payout_ratio_override": payout_override,
        "forecast_dps":          forecast_dps,
        "use_market_cap":        use_mktcap,
    }

    rounds_data: dict = {}
    if uploaded_files:
        rounds_data = load_all_rounds(uploaded_files)

    if rounds_data:
        all_rounds   = sorted(rounds_data.keys())
        latest_round = max(all_rounds)
        prev_round   = latest_round - 1

        data      = rounds_data[latest_round]
        companies = list(data.keys())
        prior_data = rounds_data.get(prev_round)

        assign_colors(companies)
        enrich_all_companies(data, assumptions, prior_data)

        st.success(
            f"✅ Rounds loaded: {all_rounds}\n"
            f"Showing Round **{latest_round}** — {len(companies)} companies"
        )
        if len(all_rounds) < 2:
            st.info("ℹ️ Upload prior round for growth metrics.")

        my_company = st.selectbox("🏢 Your Company", companies)
        st.markdown("---")
        for c in companies:
            st.markdown(f"{'⭐' if c == my_company else '▪️'} {c}")
    else:
        data, companies, my_company, prior_data, latest_round = {}, [], None, None, 0

    st.markdown("---")
    st.caption("Cesim Financial Dashboard v3.0")


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 8 — Main area + tab navigation
# ═══════════════════════════════════════════════════════════════════════════

st.markdown("# 📊 Cesim Global Challenge — Financial Dashboard")

if not data:
    st.info("👈 Upload one or more Cesim round results (.xls) files from the sidebar.")
    st.markdown("""
### What's new in v3
| Feature | Detail |
|---------|--------|
| 🌐 Regional Analysis tab | Full IS, BS and market-share breakdown for USA / China / Europe |
| Revenue contribution % | Stacked bars showing each company's revenue split by region |
| Regional EBITDA / net margin | Side-by-side comparison across all companies and all 3 regions |
| Regional market shares | Combustion / Hybrid / Electric / Hydrogen per region |
| Regional P&L table | Colour-coded table per company with region rows |
| Regional BS snapshot | Assets, equity, liabilities by region |
""")
    st.stop()

tabs = st.tabs([
    "🏠 Overview",
    "📈 Profitability",
    "💧 Liquidity",
    "🏦 Solvency",
    "📉 Market",
    "💰 Cost Breakdown",
    "🌐 Regional Analysis",   # ← NEW in v3
    "🌍 ESG & HR",
    "💎 Valuation",
    "📐 Capital & Risk",
    "📊 Growth & Trends",
    "📄 CFO Cheat Sheet",
])


# ══════════════════════════ TAB 0 — OVERVIEW ════════════════════════════════
with tabs[0]:
    st.markdown(f"<div class='section-header'>Round {latest_round} — Company Snapshot</div>",
                unsafe_allow_html=True)
    my = data[my_company]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>Net Profit</div>
          <div class='kpi-value'>{fmt_usd(my['net_profit'])}</div>
          <div class='kpi-sub'>ROS: {my['net_sales_margin']:.1f}%</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>Share Price</div>
          <div class='kpi-value'>${my['share_price']:.2f}</div>
          <div class='kpi-sub'>Market Cap: {fmt_usd(my['market_cap'])}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>ROE / ROCE</div>
          <div class='kpi-value'>{my['roe_at']:.1f}%</div>
          <div class='kpi-sub'>ROCE: {my['roce']:.1f}% | WACC: {my.get('adv_wacc',0):.1f}%</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>Market Share</div>
          <div class='kpi-value'>{my['ms_total']:.1f}%</div>
          <div class='kpi-sub'>Z-score: {f"{my.get('adv_altman_z',0):.2f}" if my.get('adv_altman_z') else "N/A"}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(f"<div class='section-header'>🚦 Health Alerts — {my_company}</div>",
                unsafe_allow_html=True)

    alert_rows = [
        ("Net Working Capital",         my['nwc'] > 0,                fmt_usd(my['nwc']),          "Positive"),
        ("Net Cash Position",           my['net_cash'] > 0,            fmt_usd(my['net_cash']),      "Positive"),
        ("Current Ratio ≥ 1.0",         my['current_ratio'] >= 1.0,    f"{my['current_ratio']:.2f}", "≥ 1.0"),
        ("Quick Ratio ≥ 0.8",           my['quick_ratio'] >= 0.8,      f"{my['quick_ratio']:.2f}",   "≥ 0.8"),
        ("EBITDA Margin ≥ 10%",         my['gross_sales_margin'] >= 10,pct_str(my['gross_sales_margin']),"≥ 10%"),
        ("Net Margin > 0%",             my['net_sales_margin'] > 0,    pct_str(my['net_sales_margin']),"≥ 0%"),
        ("ROE (before tax) ≥ 15%",      my['roe_bt'] >= 15,            pct_str(my['roe_bt']),        "≥ 15%"),
        ("ROCE ≥ 15%",                  my['roce'] >= 15,              pct_str(my['roce']),          "≥ 15%"),
        ("ROCE > WACC",                 my.get('adv_roce_wacc_spread',0)>0,
                                        f"Spread: {my.get('adv_roce_wacc_spread',0):+.1f}%", "ROCE > WACC"),
        ("Financial Independence ≥ 30%",my['fin_independence'] >= 30,  pct_str(my['fin_independence']),"≥ 30%"),
        ("Financing Stability ≥ 50%",   my['fin_stability'] >= 50,     pct_str(my['fin_stability']), "≥ 50%"),
        ("Net Gearing ≤ 100%",          my['gearing'] <= 100,          pct_str(my['gearing']),       "≤ 100%"),
        ("Altman Z-Score ≥ 1.81",       (my.get('adv_altman_z') or 0) >= 1.81,
                                        f"{my.get('adv_altman_z',0):.2f}" if my.get('adv_altman_z') else "N/A",
                                        "≥ 1.81"),
    ]

    col_a, col_b = st.columns([2, 1])
    with col_a:
        ht = ['<table class="styled-table"><thead><tr>',
              '<th>Indicator</th><th>Your Value</th><th>Threshold</th><th>Status</th></tr></thead><tbody>']
        for label, ok, val_str, thresh in alert_rows:
            badge = f'<span class="badge-green">🟢 OK</span>' if ok \
                    else f'<span class="badge-red">🔴 Watch</span>'
            ht.append(f'<tr><td><b>{label}</b></td><td>{val_str}</td>'
                      f'<td>{thresh}</td><td>{badge}</td></tr>')
        ht.append("</tbody></table>")
        st.markdown("".join(ht), unsafe_allow_html=True)

    with col_b:
        green_cnt = sum(1 for _, ok, _, _ in alert_rows if ok)
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=green_cnt,
            title={"text": "Health Score"},
            gauge={
                "axis": {"range": [0, len(alert_rows)]},
                "bar": {"color": "#10B981" if green_cnt >= 10 else
                                  "#F59E0B" if green_cnt >= 7 else "#EF4444"},
                "steps": [
                    {"range": [0, 5],              "color": "#FEE2E2"},
                    {"range": [5, 9],              "color": "#FEF3C7"},
                    {"range": [9, len(alert_rows)],"color": "#D1FAE5"},
                ],
            },
        ))
        fig_gauge.update_layout(height=260, paper_bgcolor="white",
                                margin=dict(t=30,b=10,l=20,r=20))
        st.plotly_chart(fig_gauge, use_container_width=True)
        st.markdown(f"**{green_cnt} of {len(alert_rows)} indicators OK**")

    st.markdown("---")
    st.markdown("<div class='section-header'>📊 All-Company Snapshot</div>",
                unsafe_allow_html=True)
    rows_snap = []
    for c in companies:
        d = data[c]
        rows_snap.append({
            "Company": f"⭐ {c}" if c == my_company else c,
            "Sales": fmt_usd(d["sales_rev"]),
            "Net Profit": fmt_usd(d["net_profit"]),
            "Net Margin": pct_str(d["net_sales_margin"]),
            "ROE(BT)": pct_str(d["roe_bt"]),
            "ROCE": pct_str(d["roce"]),
            "WACC": pct_str(d.get("adv_wacc", 0)),
            "Spread": fmt_pct(d.get("adv_roce_wacc_spread", 0)),
            "Share $": f"${d['share_price']:.2f}",
            "Mkt Cap": fmt_usd(d["market_cap"]),
            "Z-Score": f"{d.get('adv_altman_z',0):.2f}" if d.get("adv_altman_z") else "N/A",
        })
    st.dataframe(pd.DataFrame(rows_snap).set_index("Company"), use_container_width=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>🕸️ Multi-Dimensional Comparison</div>",
                unsafe_allow_html=True)

    def _norm(field, hi=True):
        vals = {c: (data[c].get(field) or 0) for c in companies}
        mn, mx = min(vals.values()), max(vals.values())
        rng = mx - mn or 1
        return {c: (v - mn)/rng*10 if hi else (mx - v)/rng*10
                for c, v in vals.items()}

    cats     = ["Revenue", "Profitability", "Liquidity", "Solvency", "Market Share", "ESG"]
    _n_rev   = _norm("sales_rev")
    _n_ros   = _norm("net_sales_margin")
    _n_cr    = _norm("current_ratio")
    _n_fi    = _norm("fin_independence")
    _n_ms    = _norm("ms_total")
    _n_esg   = _norm("esg_rep")
    vals_dict = {c: [_n_rev[c],_n_ros[c],_n_cr[c],_n_fi[c],_n_ms[c],_n_esg[c]]
                 for c in companies}
    st.plotly_chart(
        radar_chart("Company Comparison (0–10 normalised)", cats, vals_dict, companies),
        use_container_width=True,
    )


# ══════════════════════════ TAB 1 — PROFITABILITY ════════════════════════════
with tabs[1]:
    st.markdown("<div class='section-header'>📈 Profitability Analysis</div>",
                unsafe_allow_html=True)
    my = data[my_company]
    rank_ros  = rank_companies(data, "net_sales_margin")
    rank_roa  = rank_companies(data, "net_roa")
    rank_roe  = rank_companies(data, "roe_bt")
    rank_roce = rank_companies(data, "roce")
    rank_ebit = rank_companies(data, "gross_sales_margin")

    c1,c2,c3,c4,c5 = st.columns(5)
    for col, (lbl, val, rk) in zip([c1,c2,c3,c4,c5],[
        ("EBITDA Margin",     my["gross_sales_margin"],   rank_ebit[my_company]),
        ("Net Margin (ROS)",  my["net_sales_margin"],     rank_ros[my_company]),
        ("Gross ROA",         my["gross_roa"],            rank_roa[my_company]),
        ("ROE (before tax)",  my["roe_bt"],               rank_roe[my_company]),
        ("ROCE",              my["roce"],                  rank_roce[my_company]),
    ]):
        with col:
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value'>{val:.1f}%</div>
              <div class='kpi-sub'>{alert(val,(15,8))} {rank_badge(rk)}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(bar_chart("Net Profit Margin %", companies, "net_sales_margin",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
        st.plotly_chart(bar_chart("ROE before Tax %", companies, "roe_bt",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
    with col_r:
        st.plotly_chart(bar_chart("EBITDA Margin %", companies, "gross_sales_margin",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
        st.plotly_chart(bar_chart("ROCE %", companies, "roce",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Detailed Profitability Table</div>",
                unsafe_allow_html=True)
    metrics = [
        ("Sales Revenue",           "sales_rev",          fmt_usd),
        ("EBITDA",                  "ebitda",             fmt_usd),
        ("EBIT",                    "ebit",               fmt_usd),
        ("Net Profit",              "net_profit",         fmt_usd),
        ("EBITDA Margin %",         "gross_sales_margin", lambda v: pct_str(v)),
        ("Net Margin (ROS) %",      "net_sales_margin",   lambda v: pct_str(v)),
        ("Gross ROA %",             "gross_roa",          lambda v: pct_str(v)),
        ("Net ROA %",               "net_roa",            lambda v: pct_str(v)),
        ("ROE (before tax) %",      "roe_bt",             lambda v: pct_str(v)),
        ("ROE (after tax) %",       "roe_at",             lambda v: pct_str(v)),
        ("ROCE %",                  "roce",               lambda v: pct_str(v)),
        ("WACC % (calculated)",     "adv_wacc",           lambda v: pct_str(v)),
        ("ROCE – WACC Spread",      "adv_roce_wacc_spread", lambda v: fmt_pct(v)),
        ("EPS ($)",                 "eps",                lambda v: f"${v:.2f}"),
    ]
    ht = _table_header(companies, my_company, ["🏆 Best"])
    for lbl, fld, fn in metrics:
        ranks = rank_companies(data, fld)
        best  = min(ranks, key=ranks.get)
        ht   += _table_row(lbl, companies, data, fld,
                           lambda v, f=fn: fmt_opt(v, f),
                           my_company, [f"<b>{best}</b>"])
    st.markdown(ht + "</tbody></table>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Financial Leverage</div>",
                unsafe_allow_html=True)
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(bar_chart("Leverage Sum (%)", companies, "fin_lev_sum",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
    with col_r:
        st.plotly_chart(bar_chart("Leverage Multiplier", companies, "fin_lev_mult",
                                   data, lambda v: f"{v:.2f}×"), use_container_width=True)


# ══════════════════════════ TAB 2 — LIQUIDITY ════════════════════════════════
with tabs[2]:
    st.markdown("<div class='section-header'>💧 Liquidity Analysis</div>",
                unsafe_allow_html=True)
    my = data[my_company]
    c1,c2,c3,c4 = st.columns(4)
    for col, (lbl, val, alrt_fn) in zip([c1,c2,c3,c4],[
        ("Net Working Capital", my["nwc"],
         lambda v: alert(v,(0,-1),("🟢 Positive","🟢 Positive","🔴 Negative"))),
        ("Net Cash Position",   my["net_cash"],
         lambda v: alert(v,(0,-1),("🟢 Positive","🟢 Positive","🔴 Negative"))),
        ("Current Ratio",       my["current_ratio"],
         lambda v: alert(v,(1.5,1.0))),
        ("Quick Ratio",         my["quick_ratio"],
         lambda v: alert(v,(1.0,0.8))),
    ]):
        with col:
            disp = fmt_usd(val) if abs(val) > 100 else f"{val:.2f}"
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value'>{disp}</div>
              <div class='kpi-sub'>{alrt_fn(val)}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(bar_chart("Current Ratio", companies, "current_ratio",
                                   data, lambda v: f"{v:.2f}"), use_container_width=True)
        st.plotly_chart(bar_chart("Stock Holding Period (days)", companies, "stock_days",
                                   data, lambda v: f"{v:.0f}d"), use_container_width=True)
    with col_r:
        st.plotly_chart(bar_chart("Quick Ratio", companies, "quick_ratio",
                                   data, lambda v: f"{v:.2f}"), use_container_width=True)
        st.plotly_chart(bar_chart("Cash Cycle (days)", companies, "cash_cycle",
                                   data, lambda v: f"{v:.0f}d"), use_container_width=True)

    st.markdown("---")
    fig_cc = go.Figure()
    for c in companies:
        d = data[c]
        fig_cc.add_trace(go.Bar(
            name=c, marker_color=COLOR_MAP[c],
            x=["Stock Days","Receivables Days","Payables Days","Cash Cycle"],
            y=[d["stock_days"], d["rec_days"], d["pay_days"], d["cash_cycle"]],
        ))
    fig_cc.update_layout(barmode="group", title="Cash Cycle Components",
                          height=340, plot_bgcolor="white", paper_bgcolor="white",
                          yaxis=dict(title="Days", gridcolor="#E8EDF8"),
                          font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20))
    st.plotly_chart(fig_cc, use_container_width=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Detailed Liquidity Table</div>",
                unsafe_allow_html=True)
    liq_metrics = [
        ("Current Assets",          "current_assets", fmt_usd,             "—"),
        ("Current Liabilities",     "current_liab",   fmt_usd,             "—"),
        ("Net Working Capital",     "nwc",            fmt_usd,             "Positive = good"),
        ("Net Cash Position",       "net_cash",       fmt_usd,             "Positive = good"),
        ("Current Ratio",           "current_ratio",  lambda v: f"{v:.2f}","1.0–1.5 normal"),
        ("Quick Ratio",             "quick_ratio",    lambda v: f"{v:.2f}","≥ 0.8 good"),
        ("Stock Holding Period",    "stock_days",     lambda v: f"{v:.0f}d","Shorter = leaner"),
        ("Receivables Days",        "rec_days",       lambda v: f"{v:.0f}d","Shorter = faster cash in"),
        ("Payables Days",           "pay_days",       lambda v: f"{v:.0f}d","Longer = free supplier credit"),
        ("Cash Cycle",              "cash_cycle",     lambda v: f"{v:.0f}d","Negative = favourable"),
        ("Cash & Equivalents",      "cash",           fmt_usd,             "Higher = more buffer"),
        ("Inventory",               "inventory",      fmt_usd,             "—"),
    ]
    ht = _table_header(companies, my_company, ["Benchmark"])
    for lbl, fld, fn, bench in liq_metrics:
        ht += _table_row(lbl, companies, data, fld,
                         lambda v, f=fn: fmt_opt(v, f),
                         my_company, [f"<small>{bench}</small>"])
    st.markdown(ht + "</tbody></table>", unsafe_allow_html=True)


# ══════════════════════════ TAB 3 — SOLVENCY ═════════════════════════════════
with tabs[3]:
    st.markdown("<div class='section-header'>🏦 Solvency & Financial Structure</div>",
                unsafe_allow_html=True)
    my = data[my_company]
    c1,c2,c3,c4 = st.columns(4)
    kpis_solv = [
        ("Financial Independence", my["fin_independence"], alert(my["fin_independence"],(40,30))),
        ("Debt Ratio",             my["debt_ratio"],       alert_low(my["debt_ratio"],60,70)),
        ("Financing Stability",    my["fin_stability"],    alert(my["fin_stability"],(80,60))),
        ("Net Gearing",            my["gearing"],          alert_low(my["gearing"],80,150)),
    ]
    for col, (lbl, val, alrt) in zip([c1,c2,c3,c4], kpis_solv):
        with col:
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value'>{val:.1f}%</div>
              <div class='kpi-sub'>{alrt}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(bar_chart("Financial Independence %", companies, "fin_independence",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
        st.plotly_chart(bar_chart("Net Gearing % (↓ better)", companies, "gearing",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
    with col_r:
        st.plotly_chart(bar_chart("Financing Stability %", companies, "fin_stability",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
        fig_stk = go.Figure()
        for lbl, fld, clr in [("Equity","total_equity","#3B6FE8"),
                               ("LT Debt","lt_debt","#F59E0B"),
                               ("ST Debt","st_debt","#EF4444")]:
            fig_stk.add_trace(go.Bar(name=lbl, x=companies,
                                      y=[data[c][fld] for c in companies],
                                      marker_color=clr))
        fig_stk.update_layout(barmode="stack", title="Capital Structure ($)",
                               height=320, plot_bgcolor="white", paper_bgcolor="white",
                               yaxis=dict(gridcolor="#E8EDF8"),
                               font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20))
        st.plotly_chart(fig_stk, use_container_width=True)

    st.markdown("---")
    ir_rows = []
    for c in companies:
        d = data[c]
        ir_rows.append({
            "Company": f"⭐ {c}" if c == my_company else c,
            "USA LT": f"{d['int_usa_long']:.2f}%",
            "USA ST": f"{d['int_usa_short']:.2f}%",
            "China ST": f"{d['int_cn_short']:.2f}%",
            "Europe ST": f"{d['int_eu_short']:.2f}%",
            "Avg Cost Debt": f"{d['avg_cost_debt']:.2f}%",
        })
    st.dataframe(pd.DataFrame(ir_rows).set_index("Company"), use_container_width=True)
    st.info("💡 Interest rates rise as leverage increases. Keep avg cost of debt below ROA.")

    st.markdown("---")
    st.markdown("<div class='section-header'>Detailed Solvency Table</div>", unsafe_allow_html=True)
    solv_html = ['<table class="styled-table"><thead><tr><th>Metric</th>']
    for c in companies:
        style = "background:#2A3F7A;color:white;" if c == my_company else ""
        solv_html.append(f'<th style="{style}">{c}</th>')
    solv_html.append('<th>Benchmark</th></tr></thead><tbody>')
    solv_metrics = [
        ("Total Equity",          lambda d: fmt_usd(d["total_equity"]),       "—"),
        ("LT Debt",               lambda d: fmt_usd(d["lt_debt"]),            "—"),
        ("ST Debt (unplanned)",   lambda d: fmt_usd(d["st_debt"]),            "0 = best"),
        ("Total Assets",          lambda d: fmt_usd(d["total_assets"]),       "—"),
        ("Financial Independence",lambda d: pct_str(d["fin_independence"]),   "≥ 40% healthy"),
        ("Debt Ratio",            lambda d: pct_str(d["debt_ratio"]),         "≤ 60% healthy"),
        ("Financing Stability",   lambda d: pct_str(d["fin_stability"]),     "≥ 80% good"),
        ("Net Debt",              lambda d: fmt_usd(d["net_debt"]),           "Lower = better"),
        ("Net Gearing",           lambda d: pct_str(d["gearing"]),            "≤ 100% manageable"),
        ("Avg Cost of Debt",      lambda d: f"{d['avg_cost_debt']:.2f}%",     "Lower = cheaper"),
        ("Fin. Leverage (sum)",   lambda d: pct_str(d["fin_lev_sum"]),       "> 0 = shareholders benefit"),
        ("Fin. Leverage (mult.)", lambda d: f"{d['fin_lev_mult']:.2f}x",     "> 1 = positive leverage"),
    ]
    for label, fn, bench in solv_metrics:
        solv_html.append("<tr>")
        solv_html.append(f"<td><b>{label}</b></td>")
        for c in companies:
            cls = "my-company" if c == my_company else ""
            solv_html.append(f'<td class="{cls}">{fn(data[c])}</td>')
        solv_html.append(f"<td><small>{bench}</small></td></tr>")
    st.markdown("".join(solv_html) + "</tbody></table>", unsafe_allow_html=True)


# ══════════════════════════ TAB 4 — MARKET ═══════════════════════════════════
with tabs[4]:
    st.markdown("<div class='section-header'>📉 Market & Shareholder Value</div>",
                unsafe_allow_html=True)
    my = data[my_company]
    c1,c2,c3,c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>Share Price</div>
          <div class='kpi-value'>${my['share_price']:.2f}</div>
          <div class='kpi-sub'>TSR: {my['tsr']:.1f}%</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>Market Cap</div>
          <div class='kpi-value'>{fmt_usd(my['market_cap'])}</div>
          <div class='kpi-sub'>EPS: ${my['eps']:.2f}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        pe_str = f"{my['pe_ratio']:.1f}×" if my['pe_ratio'] else "N/A"
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>P/E Ratio</div>
          <div class='kpi-value'>{pe_str}</div>
          <div class='kpi-sub'>P/B: {fmt_opt(my.get('adv_pb'), lambda v: f'{v:.2f}×')}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>Global Market Share</div>
          <div class='kpi-value'>{my['ms_total']:.1f}%</div>
          <div class='kpi-sub'>Combustion: {my['ms_combustion']:.1f}%</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(bar_chart("Share Price ($)", companies, "share_price",
                                   data, lambda v: f"${v:.2f}"), use_container_width=True)
        fig_pie = go.Figure(go.Pie(labels=companies,
                                    values=[data[c]["ms_total"] for c in companies],
                                    marker_colors=[COLOR_MAP[c] for c in companies],
                                    hole=0.4))
        fig_pie.update_layout(title="Global Market Share %", height=320,
                               paper_bgcolor="white", margin=dict(t=40,b=10,l=10,r=10))
        st.plotly_chart(fig_pie, use_container_width=True)
    with col_r:
        st.plotly_chart(bar_chart("Market Cap", companies, "market_cap",
                                   data, fmt_usd), use_container_width=True)
        fig_ms = go.Figure()
        for i, (fld, lbl) in enumerate(zip(
            ["ms_combustion","ms_hybrid","ms_electric","ms_hydrogen"],
            ["Combustion","Hybrid","Electric","Hydrogen"]
        )):
            fig_ms.add_trace(go.Bar(name=lbl, x=companies,
                                     y=[data[c][fld] for c in companies],
                                     marker_color=COLORS[i]))
        fig_ms.update_layout(barmode="group", title="Global Market Share by Technology",
                              height=320, plot_bgcolor="white", paper_bgcolor="white",
                              yaxis=dict(gridcolor="#E8EDF8"),
                              font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20))
        st.plotly_chart(fig_ms, use_container_width=True)

    st.plotly_chart(bar_chart("Cumulative TSR % p.a.", companies, "tsr",
                               data, lambda v: f"{v:.1f}%"), use_container_width=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Detailed Market Table</div>", unsafe_allow_html=True)
    market_html = ['<table class="styled-table"><thead><tr><th>Metric</th>']
    for c in companies:
        style = "background:#2A3F7A;color:white;" if c == my_company else ""
        market_html.append(f'<th style="{style}">{c}</th>')
    market_html.append('<th>Comment</th></tr></thead><tbody>')
    market_rows = [
        ("Share Price",          lambda d: f"${d['share_price']:.2f}",                  "Current trading price"),
        ("Average Trading Price",lambda d: f"${d['avg_price']:.2f}",                    "Avg executed price"),
        ("Market Cap",           lambda d: fmt_usd(d["market_cap"]),                     "Equity value"),
        ("EPS",                  lambda d: f"${d['eps']:.2f}",                           "Earnings per share"),
        ("P/E Ratio",            lambda d: f"{d['pe_ratio']:.1f}x" if d["pe_ratio"] else "N/A", "Lower = cheaper"),
        ("Dividend Yield",       lambda d: f"{d['div_yield']:.1f}%",                    "From Cesim KPI block"),
        ("TSR",                  lambda d: f"{d['tsr']:.1f}%",                          "Cumulative return"),
        ("Global Market Share",  lambda d: f"{d['ms_total']:.1f}%",                     "All technologies"),
        ("Combustion Share",     lambda d: f"{d['ms_combustion']:.1f}%",                "Legacy segment"),
        ("Hybrid Share",         lambda d: f"{d['ms_hybrid']:.1f}%",                    "Transition segment"),
        ("Electric Share",       lambda d: f"{d['ms_electric']:.1f}%",                  "Growth segment"),
        ("Hydrogen Share",       lambda d: f"{d['ms_hydrogen']:.1f}%",                  "Optional segment"),
    ]
    for label, fn, comment in market_rows:
        market_html.append("<tr>")
        market_html.append(f"<td><b>{label}</b></td>")
        for c in companies:
            cls = "my-company" if c == my_company else ""
            market_html.append(f'<td class="{cls}">{fn(data[c])}</td>')
        market_html.append(f"<td><small>{comment}</small></td></tr>")
    st.markdown("".join(market_html) + "</tbody></table>", unsafe_allow_html=True)


# ══════════════════════════ TAB 5 — COST BREAKDOWN ═══════════════════════════
with tabs[5]:
    st.markdown("<div class='section-header'>💰 Cost Structure & Breakdown</div>",
                unsafe_allow_html=True)
    my = data[my_company]
    cost_items = [
        ("In-house Manufacturing","inhouse_pct","inhouse_mfg"),
        ("Contract Manufacturing","contract_pct","contract_mfg"),
        ("Feature Costs",         "feature_pct","feature_costs"),
        ("Transport & Tariffs",   "transport_pct","transport_tariffs"),
        ("R&D",                   "rd_pct","rd_costs"),
        ("Promotion",             "promo_pct","promotion"),
        ("Warranty & Data",       "warranty_pct","warranty"),
        ("Administration",        "admin_pct","administration"),
    ]
    stk_colors = ["#3B6FE8","#10B981","#F59E0B","#EF4444","#8B5CF6","#EC4899","#14B8A6","#6366F1"]

    col_l, col_r = st.columns(2)
    with col_l:
        fig_pie_c = go.Figure(go.Pie(
            labels=[l for l,_,_ in cost_items],
            values=[my[p] for _,p,_ in cost_items],
            marker_colors=stk_colors, hole=0.4, textinfo="label+percent",
        ))
        fig_pie_c.update_layout(title=f"Cost Mix — {my_company}", height=380,
                                 paper_bgcolor="white", showlegend=False,
                                 margin=dict(t=40,b=10,l=10,r=10))
        st.plotly_chart(fig_pie_c, use_container_width=True)

    with col_r:
        wf_labels = ["Revenue","In-house","Contract","Features","Transport",
                     "R&D","Promo","Warranty","Admin","EBITDA","Deprec","EBIT",
                     "Fin Exp","Tax","Net Profit"]
        wf_measure = ["absolute","relative","relative","relative","relative",
                      "relative","relative","relative","relative","total",
                      "relative","total","relative","relative","total"]
        wf_vals = [my["sales_rev"], -my["inhouse_mfg"], -my["contract_mfg"],
                   -my["feature_costs"], -my["transport_tariffs"], -my["rd_costs"],
                   -my["promotion"], -my["warranty"], -my["administration"], 0,
                   -my["depreciation"], 0, -my["net_fin_exp"], -my["income_taxes"], 0]
        fig_wf = go.Figure(go.Waterfall(
            measure=wf_measure, x=wf_labels, y=wf_vals,
            connector={"line":{"color":"#D1D5DB"}},
            increasing={"marker":{"color":"#10B981"}},
            decreasing={"marker":{"color":"#EF4444"}},
            totals={"marker":{"color":"#3B6FE8"}},
            texttemplate="%{y:,.0f}", textposition="outside",
        ))
        fig_wf.update_layout(title=f"P&L Waterfall — {my_company}", height=380,
                              plot_bgcolor="white", paper_bgcolor="white",
                              yaxis=dict(gridcolor="#E8EDF8"),
                              xaxis=dict(tickangle=-35),
                              font=dict(color="#1E2A4A", size=11),
                              margin=dict(t=40,b=60,l=20,r=20))
        st.plotly_chart(fig_wf, use_container_width=True)

    st.markdown("---")
    fig_stk_c = go.Figure()
    for i,(lbl,pf,_) in enumerate(cost_items):
        fig_stk_c.add_trace(go.Bar(name=lbl, x=companies,
                                    y=[data[c][pf] for c in companies],
                                    marker_color=stk_colors[i]))
    fig_stk_c.update_layout(barmode="stack", title="Cost % of Revenue — All Companies",
                              height=400, plot_bgcolor="white", paper_bgcolor="white",
                              yaxis=dict(title="% of Revenue", gridcolor="#E8EDF8"),
                              font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20))
    st.plotly_chart(fig_stk_c, use_container_width=True)


# ══════════════════════════ TAB 6 — REGIONAL ANALYSIS (NEW v3) ═══════════════
with tabs[6]:
    st.markdown("<div class='section-header'>🌐 Regional Analysis — USA · China · Europe</div>",
                unsafe_allow_html=True)

    my    = data[my_company]
    regs  = list(REGIONS.keys())           # ["USA", "China", "Europe"]
    rclrs = [REGIONS[r]["color"] for r in regs]

    # ── 6a. Your-company regional KPIs ──────────────────────────────────────
    st.markdown(f"#### {my_company} — Regional P&L Snapshot", unsafe_allow_html=False)
    kpi_cols = st.columns(3)
    for col, rname in zip(kpi_cols, regs):
        rd = my["regions"][rname]
        pill = REGIONS[rname]["pill"]
        rev_pct = rd["rev_pct_global"]
        with col:
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{pill} Revenue</div>
              <div class='kpi-value'>{fmt_usd(rd['sales_rev'])}</div>
              <div class='kpi-sub'>{rev_pct:.1f}% of global | EBITDA margin: {rd['ebitda_margin']:.1f}%</div>
            </div>""", unsafe_allow_html=True)

    kpi2_cols = st.columns(3)
    for col, rname in zip(kpi2_cols, regs):
        rd = my["regions"][rname]
        pill = REGIONS[rname]["pill"]
        margin_color = "#10B981" if rd["net_margin"] >= 10 else \
                       "#F59E0B"  if rd["net_margin"] >= 0   else "#EF4444"
        with col:
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{pill} Net Profit</div>
              <div class='kpi-value' style='color:{margin_color}'>{fmt_usd(rd['net_profit'])}</div>
              <div class='kpi-sub'>Net margin: {rd['net_margin']:.1f}% | Assets: {fmt_usd(rd['total_assets'])}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── 6b. Revenue contribution charts ─────────────────────────────────────
    st.markdown("<div class='section-header'>Revenue by Region — All Companies</div>",
                unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        # Stacked absolute revenue
        fig_rev_stk = go.Figure()
        for rname, rclr in zip(regs, rclrs):
            fig_rev_stk.add_trace(go.Bar(
                name=rname,
                x=companies,
                y=[data[c]["regions"][rname]["sales_rev"] for c in companies],
                marker_color=rclr,
                text=[fmt_usd(data[c]["regions"][rname]["sales_rev"]) for c in companies],
                textposition="inside",
            ))
        fig_rev_stk.update_layout(
            barmode="stack", title="Revenue by Region ($)",
            height=360, plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(title="USD", gridcolor="#E8EDF8"),
            font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
        )
        st.plotly_chart(fig_rev_stk, use_container_width=True)

    with col_r:
        # Stacked % contribution
        fig_rev_pct = go.Figure()
        for rname, rclr in zip(regs, rclrs):
            fig_rev_pct.add_trace(go.Bar(
                name=rname,
                x=companies,
                y=[data[c]["regions"][rname]["rev_pct_global"] for c in companies],
                marker_color=rclr,
                text=[f"{data[c]['regions'][rname]['rev_pct_global']:.1f}%" for c in companies],
                textposition="inside",
            ))
        fig_rev_pct.update_layout(
            barmode="stack", title="Revenue Contribution % by Region",
            height=360, plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(title="%", gridcolor="#E8EDF8"),
            font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
        )
        st.plotly_chart(fig_rev_pct, use_container_width=True)

    st.markdown("---")

    # ── 6c. EBITDA & net margin by region ───────────────────────────────────
    st.markdown("<div class='section-header'>EBITDA & Net Margin by Region</div>",
                unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        fig_ebitda = go.Figure()
        for rname, rclr in zip(regs, rclrs):
            fig_ebitda.add_trace(go.Bar(
                name=rname, x=companies,
                y=[data[c]["regions"][rname]["ebitda_margin"] for c in companies],
                marker_color=rclr,
                text=[f"{data[c]['regions'][rname]['ebitda_margin']:.1f}%" for c in companies],
                textposition="outside",
            ))
        fig_ebitda.update_layout(
            barmode="group", title="EBITDA Margin % by Region",
            height=360, plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(title="%", gridcolor="#E8EDF8"),
            font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
        )
        st.plotly_chart(fig_ebitda, use_container_width=True)

    with col_r:
        fig_nm = go.Figure()
        for rname, rclr in zip(regs, rclrs):
            fig_nm.add_trace(go.Bar(
                name=rname, x=companies,
                y=[data[c]["regions"][rname]["net_margin"] for c in companies],
                marker_color=rclr,
                text=[f"{data[c]['regions'][rname]['net_margin']:.1f}%" for c in companies],
                textposition="outside",
            ))
        fig_nm.update_layout(
            barmode="group", title="Net Margin % by Region",
            height=360, plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(title="%", gridcolor="#E8EDF8"),
            font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
        )
        st.plotly_chart(fig_nm, use_container_width=True)

    # Absolute EBITDA by region
    fig_ebitda_abs = go.Figure()
    for rname, rclr in zip(regs, rclrs):
        fig_ebitda_abs.add_trace(go.Bar(
            name=rname, x=companies,
            y=[data[c]["regions"][rname]["ebitda"] for c in companies],
            marker_color=rclr,
            text=[fmt_usd(data[c]["regions"][rname]["ebitda"]) for c in companies],
            textposition="outside",
        ))
    fig_ebitda_abs.update_layout(
        barmode="group", title="EBITDA $ by Region",
        height=340, plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(title="USD", gridcolor="#E8EDF8"),
        font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
    )
    st.plotly_chart(fig_ebitda_abs, use_container_width=True)

    st.markdown("---")

    # ── 6d. Regional Market Shares ─────────────────────────────────────────
    st.markdown("<div class='section-header'>Market Shares by Region & Technology</div>",
                unsafe_allow_html=True)

    ms_tab_cols = st.columns(3)
    tech_order  = ["ms_combustion","ms_hybrid","ms_electric","ms_hydrogen"]
    tech_labels = ["Combustion","Hybrid","Electric","Hydrogen"]

    for col, rname in zip(ms_tab_cols, regs):
        pill = REGIONS[rname]["pill"]
        with col:
            st.markdown(f"**{pill} Market Shares**", unsafe_allow_html=True)
            fig_ms_r = go.Figure()
            for i, (fld, lbl) in enumerate(zip(tech_order, tech_labels)):
                fig_ms_r.add_trace(go.Bar(
                    name=lbl, x=companies,
                    y=[data[c]["regions"][rname][fld] for c in companies],
                    marker_color=COLORS[i],
                ))
            fig_ms_r.update_layout(
                barmode="group",
                height=300,
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis=dict(title="%", gridcolor="#E8EDF8"),
                xaxis=dict(tickangle=-20),
                font=dict(color="#1E2A4A", size=11),
                margin=dict(t=20,b=20,l=10,r=10),
                showlegend=(rname == regs[-1]),
            )
            st.plotly_chart(fig_ms_r, use_container_width=True)

    st.markdown("---")

    # Total market share comparison (all 3 regions side by side)
    fig_ms_total = go.Figure()
    for rname, rclr in zip(regs, rclrs):
        fig_ms_total.add_trace(go.Bar(
            name=rname, x=companies,
            y=[data[c]["regions"][rname]["ms_total"] for c in companies],
            marker_color=rclr,
            text=[f"{data[c]['regions'][rname]['ms_total']:.1f}%" for c in companies],
            textposition="outside",
        ))
    fig_ms_total.update_layout(
        barmode="group", title="Total Market Share % by Region",
        height=360, plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(title="%", gridcolor="#E8EDF8"),
        font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
    )
    st.plotly_chart(fig_ms_total, use_container_width=True)

    st.markdown("---")

    # ── 6e. Regional P&L table (per selected company) ──────────────────────
    st.markdown("<div class='section-header'>Regional P&L — Detailed Table</div>",
                unsafe_allow_html=True)

    sel_company = st.selectbox(
        "Select company for regional P&L detail",
        companies,
        index=companies.index(my_company),
        key="reg_company_selector",
    )
    sel_d = data[sel_company]

    reg_pl_html = ['<table class="styled-table"><thead><tr>',
                   '<th>P&L Line</th>',
                   '<th>🌍 Global</th>']
    for rname in regs:
        pill = REGIONS[rname]["pill"]
        reg_pl_html.append(f'<th>{pill}</th>')
    reg_pl_html.append('<th>Note</th></tr></thead><tbody>')

    def _rval(rname, field):
        v = sel_d["regions"][rname].get(field, 0)
        return fmt_usd(v) if abs(v) >= 100 else f"${v:.0f}"

    pl_lines = [
        ("Sales Revenue",        "sales_rev",  fmt_usd,   ""),
        ("Total Costs",          "total_costs",fmt_usd,   ""),
        ("EBITDA",               "ebitda",     fmt_usd,   ""),
        ("EBITDA Margin",        "ebitda_margin", lambda v: pct_str(v), "% of rev"),
        ("Depreciation",         "depreciation",fmt_usd,  ""),
        ("EBIT",                 "ebit",       fmt_usd,   ""),
        ("EBIT Margin",          "ebit_margin",lambda v: pct_str(v), "% of rev"),
        ("Net Fin. Expenses",    "net_fin_exp",fmt_usd,   ""),
        ("Profit Before Tax",    "profit_before_tax",fmt_usd,""),
        ("Income Taxes",         "income_taxes",fmt_usd,  ""),
        ("Net Profit",           "net_profit", fmt_usd,   ""),
        ("Net Margin",           "net_margin", lambda v: pct_str(v), "% of rev"),
        ("Rev. % of Global",     "rev_pct_global", lambda v: pct_str(v), "contribution"),
    ]

    global_field_map = {
        "sales_rev": "sales_rev", "total_costs": "total_costs",
        "ebitda": "ebitda", "ebitda_margin": "gross_sales_margin",
        "depreciation": "depreciation", "ebit": "ebit",
        "ebit_margin": "net_sales_margin", "net_fin_exp": "net_fin_exp",
        "profit_before_tax": "profit_before_tax", "income_taxes": "income_taxes",
        "net_profit": "net_profit", "net_margin": "net_sales_margin",
        "rev_pct_global": None,
    }

    for lbl, fld, fn, note in pl_lines:
        reg_pl_html.append("<tr>")
        reg_pl_html.append(f"<td><b>{lbl}</b></td>")
        # Global column
        gfld = global_field_map.get(fld)
        gval = sel_d.get(gfld, 0) if gfld else 100.0
        reg_pl_html.append(f"<td>{fn(gval)}</td>")
        # Regional columns
        for rname in regs:
            v = sel_d["regions"][rname].get(fld, 0)
            reg_pl_html.append(f"<td>{fn(v)}</td>")
        reg_pl_html.append(f"<td><small>{note}</small></td>")
        reg_pl_html.append("</tr>")

    st.markdown("".join(reg_pl_html) + "</tbody></table>", unsafe_allow_html=True)

    st.markdown("---")

    # ── 6f. Regional Balance Sheet snapshot ────────────────────────────────
    st.markdown("<div class='section-header'>Regional Balance Sheet Snapshot</div>",
                unsafe_allow_html=True)

    bs_rows = []
    for c in companies:
        d = data[c]
        row = {"Company": f"⭐ {c}" if c == my_company else c,
               "Global Assets": fmt_usd(d["total_assets"]),
               "Global Equity": fmt_usd(d["total_equity"])}
        for rname in regs:
            rd = d["regions"][rname]
            row[f"{rname} Assets"] = fmt_usd(rd["total_assets"])
            row[f"{rname} Equity"] = fmt_usd(rd["total_equity"])
            row[f"{rname} St.Debt"] = fmt_usd(rd["st_debt"])
        bs_rows.append(row)

    st.dataframe(pd.DataFrame(bs_rows).set_index("Company"), use_container_width=True)

    # Assets by region (stacked bar)
    fig_assets = go.Figure()
    for rname, rclr in zip(regs, rclrs):
        fig_assets.add_trace(go.Bar(
            name=rname, x=companies,
            y=[data[c]["regions"][rname]["total_assets"] for c in companies],
            marker_color=rclr,
        ))
    fig_assets.update_layout(
        barmode="stack", title="Total Assets by Region ($)",
        height=340, plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(title="USD", gridcolor="#E8EDF8"),
        font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
    )
    st.plotly_chart(fig_assets, use_container_width=True)

    st.info(
        "💡 **Note on regional financials**: USA and China are full subsidiaries with "
        "manufacturing; Europe is a distribution-only entity that imports from USA/China. "
        "Internal transfers between subsidiaries are eliminated at the global consolidation level."
    )


# ══════════════════════════ TAB 7 — ESG & HR ═════════════════════════════════
with tabs[7]:
    st.markdown("<div class='section-header'>🌍 ESG & HR Metrics</div>",
                unsafe_allow_html=True)
    my = data[my_company]
    c1,c2,c3,c4 = st.columns(4)
    for col, (lbl, fld) in zip([c1,c2,c3,c4],[
        ("Environmental (E)","esg_e"),("Social (S)","esg_s"),
        ("Governance (G)","esg_g"),("ESG Reputation","esg_rep"),
    ]):
        with col:
            val = my[fld]
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value'>{val:.2f}/5.0</div>
              <div class='kpi-sub'>{alert(val,(3.5,2.8))}</div>
            </div>""", unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        esg_cats = ["Environmental","Social","Governance","Reputation"]
        esg_v = {c:[data[c]["esg_e"],data[c]["esg_s"],data[c]["esg_g"],data[c]["esg_rep"]]
                 for c in companies}
        st.plotly_chart(radar_chart("ESG Scores (out of 5)", esg_cats, esg_v, companies),
                        use_container_width=True)
    with col_r:
        hr_df = []
        for c in companies:
            d = data[c]
            hr_df.append({
                "Company": f"⭐ {c}" if c == my_company else c,
                "Wage/Month ($)": f"${d['wage_month']:,.0f}",
                "Training ($)": f"${d['train_budget']:,.0f}",
                "HR Efficiency": f"{d['hr_eff']:.3f}",
                "R&D Staff": f"{d['rdstaff_now']:,.0f}",
                "Turnover %": f"{d['vol_turnover']:.2f}%",
            })
        st.dataframe(pd.DataFrame(hr_df).set_index("Company"), use_container_width=True)


# ══════════════════════════ TAB 8 — VALUATION ════════════════════════════════
with tabs[8]:
    st.markdown("<div class='section-header'>💎 Valuation Multiples & Intrinsic Value</div>",
                unsafe_allow_html=True)

    st.markdown(f"""<div class='assume-box'>
    <b>⚙️ Active assumptions:</b> &nbsp;
    Re = {Re_pct:.1f}% &nbsp;|&nbsp;
    g = {g_pct:.1f}% &nbsp;|&nbsp;
    Tax = {"auto" if tax_override_pct == 0 else f"{tax_override_pct:.0f}%"} &nbsp;|&nbsp;
    Forecast DPS = ${forecast_dps:.2f} &nbsp;|&nbsp;
    WACC equity = {"Market Cap" if use_mktcap else "Book Equity"}
    <br><small>Change these in the sidebar ⚙️ section.</small>
    </div>""", unsafe_allow_html=True)

    my = data[my_company]
    c1,c2,c3,c4,c5 = st.columns(5)
    kpis_val = [
        ("P/E",      "adv_pe",       lambda v: f"{v:.1f}×"),
        ("P/B",      "adv_pb",       lambda v: f"{v:.2f}×"),
        ("P/Sales",  "adv_ps",       lambda v: f"{v:.2f}×"),
        ("EV/EBITDA","adv_ev_ebitda",lambda v: f"{v:.1f}×"),
        ("EV/Sales", "adv_ev_sales", lambda v: f"{v:.2f}×"),
    ]
    for col, (lbl, fld, fn) in zip([c1,c2,c3,c4,c5], kpis_val):
        with col:
            val = my.get(fld)
            disp = fmt_opt(val, fn)
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value' style='font-size:22px'>{disp}</div>
              <div class='kpi-sub'>{my_company}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Valuation Multiples — All Companies</div>",
                unsafe_allow_html=True)
    ht = _table_header(companies, my_company, ["Benchmark"])
    mult_rows = [
        ("Enterprise Value ($)", "adv_ev",       fmt_usd,                  "EV = Mkt Cap + Net Debt"),
        ("Book Value/Share ($)",  "adv_bv_per_share",lambda v:f"${v:.2f}", "Equity / Shares"),
        ("P/E Ratio",             "adv_pe",       lambda v:fmt_opt(v,lambda x:f"{x:.1f}×"), "Lower = cheaper vs. earnings"),
        ("P/B Ratio",             "adv_pb",       lambda v:fmt_opt(v,lambda x:f"{x:.2f}×"), "< 1 = below book value"),
        ("P/Sales Ratio",         "adv_ps",       lambda v:fmt_opt(v,lambda x:f"{x:.2f}×"), "Lower = cheaper vs. revenue"),
        ("EV / EBITDA",           "adv_ev_ebitda",lambda v:fmt_opt(v,lambda x:f"{x:.1f}×"), "< 10× generally cheap"),
        ("EV / Sales",            "adv_ev_sales", lambda v:fmt_opt(v,lambda x:f"{x:.2f}×"), "Context-dependent"),
        ("Effective Tax Rate",    "adv_eff_tax_rate",lambda v:pct_str(v),   "Used in WACC"),
    ]
    for lbl, fld, fn, bench in mult_rows:
        ht += _table_row(lbl, companies, data, fld, fn, my_company,
                         [f"<small>{bench}</small>"])
    st.markdown(ht + "</tbody></table>", unsafe_allow_html=True)

    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        fig_pe = bar_chart_optional("P/E Ratio (×)", companies, "adv_pe", data,
                                     lambda v: f"{v:.1f}×")
        if fig_pe: st.plotly_chart(fig_pe, use_container_width=True)
        fig_ev = bar_chart_optional("EV / EBITDA (×)", companies, "adv_ev_ebitda", data,
                                     lambda v: f"{v:.1f}×")
        if fig_ev: st.plotly_chart(fig_ev, use_container_width=True)
    with col_r:
        fig_pb = bar_chart_optional("P/B Ratio (×)", companies, "adv_pb", data,
                                     lambda v: f"{v:.2f}×")
        if fig_pb: st.plotly_chart(fig_pb, use_container_width=True)
        fig_ps = bar_chart_optional("P/Sales Ratio (×)", companies, "adv_ps", data,
                                     lambda v: f"{v:.2f}×")
        if fig_ps: st.plotly_chart(fig_ps, use_container_width=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Intrinsic Value Models</div>",
                unsafe_allow_html=True)

    has_dps = any(data[c].get("adv_dps", 0) > 0 for c in companies)
    if not has_dps:
        st.markdown("""<div class='assume-box'>
        ℹ️ <b>DDM models require dividends.</b> No company paid dividends this round.
        Enter a <b>Forecast DPS</b> in the sidebar to unlock DDM estimates.
        RIM (Residual Income Model) works without dividends and is shown below.
        </div>""", unsafe_allow_html=True)

    ht2 = _table_header(companies, my_company, ["Formula"])
    iv_rows = [
        ("Current Share Price ($)",  "share_price",     lambda v:f"${v:.2f}", "Market"),
        ("DDM Zero Growth ($)",      "adv_ddm_zero",    lambda v:fmt_opt(v,lambda x:f"${x:.2f}"), "DPS / Re"),
        ("DDM Constant Growth ($)",  "adv_ddm_const",   lambda v:fmt_opt(v,lambda x:f"${x:.2f}"), f"DPS×(1+{g_pct:.0f}%) / (Re−g)"),
        ("DDM SGR-based ($)",        "adv_ddm_sgr",     lambda v:fmt_opt(v,lambda x:f"${x:.2f}"), "DPS×(1+SGR) / (Re−SGR)"),
        ("RIM Intrinsic Value ($)",  "adv_rim_iv",      lambda v:fmt_opt(v,lambda x:f"${x:.2f}"), "BVps + (EPS1 − Re×BVps)/(Re−g)"),
        ("Actual DPS ($)",           "adv_actual_dps",  lambda v:f"${v:.2f}", "From Cesim div_yield"),
        ("Forecast DPS ($)",         "adv_dps",         lambda v:f"${v:.2f}", "Sidebar override"),
        ("SGR %",                    "adv_sgr",         lambda v:pct_str(v),  "ROE × Retention"),
        ("Payout Ratio",             "adv_payout_ratio",lambda v:pct_str(v),  "Dividends / Net Profit"),
        ("Retention Ratio",          "adv_retention_ratio",lambda v:pct_str(v),"1 − Payout"),
        ("Residual Income/Share ($)","adv_ri_per_share", lambda v:fmt_opt(v,lambda x:f"${x:.2f}"),
         f"EPS1 − Re×BVps"),
    ]
    for lbl, fld, fn, formula in iv_rows:
        ht2 += _table_row(lbl, companies, data, fld, fn, my_company,
                          [f"<small>{formula}</small>"])
    st.markdown(ht2 + "</tbody></table>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Market Price vs RIM Intrinsic Value</div>",
                unsafe_allow_html=True)
    fig_iv = go.Figure()
    fig_iv.add_trace(go.Bar(
        name="Market Price ($)", x=companies,
        y=[data[c]["share_price"] for c in companies],
        marker_color=[COLOR_MAP[c] for c in companies], opacity=0.85,
    ))
    rim_vals = [data[c].get("adv_rim_iv") for c in companies]
    rim_present = [v for v in rim_vals if v is not None]
    if rim_present:
        fig_iv.add_trace(go.Scatter(
            name="RIM Intrinsic Value ($)",
            x=[c for c, v in zip(companies, rim_vals) if v is not None],
            y=[v for v in rim_vals if v is not None],
            mode="markers",
            marker=dict(symbol="diamond", size=16, color="#1E2A4A",
                        line=dict(width=2, color="white")),
        ))
    fig_iv.update_layout(
        title="Market Price vs RIM Intrinsic Value (◆ = RIM)",
        height=380, plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(title="USD", gridcolor="#E8EDF8"),
        font=dict(color="#1E2A4A"), margin=dict(t=50,b=20,l=20,r=20),
    )
    st.plotly_chart(fig_iv, use_container_width=True)
    st.caption(
        "📌 When ◆ (RIM) > bar (market price), the stock may be undervalued; "
        "when ◆ < bar, it may be overvalued — relative to the Re and g assumptions."
    )


# ══════════════════════════ TAB 9 — CAPITAL & RISK ═══════════════════════════
with tabs[9]:
    st.markdown("<div class='section-header'>📐 Capital Structure, WACC & Bankruptcy Risk</div>",
                unsafe_allow_html=True)
    st.markdown(f"""<div class='assume-box'>
    <b>⚙️ WACC assumptions:</b> Re = {Re_pct:.1f}% &nbsp;|&nbsp;
    Rd = company-specific (from Cesim) &nbsp;|&nbsp;
    Tax = {"auto-calc" if tax_override_pct == 0 else f"{tax_override_pct:.0f}%"} &nbsp;|&nbsp;
    Equity value = {"Market Cap" if use_mktcap else "Book Equity"}
    </div>""", unsafe_allow_html=True)

    my = data[my_company]
    c1,c2,c3,c4 = st.columns(4)
    wacc_val   = my.get("adv_wacc", 0)
    spread_val = my.get("adv_roce_wacc_spread", 0)
    z_val      = my.get("adv_altman_z")
    z_zone     = my.get("adv_altman_zone", ("N/A","amber"))

    with c1:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>WACC</div>
          <div class='kpi-value'>{wacc_val:.2f}%</div>
          <div class='kpi-sub'>Minimum required return on capital</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        spread_color = "kpi-card-ok" if spread_val > 0 else "kpi-card-bad"
        st.markdown(f"""<div class='kpi-card {spread_color}'>
          <div class='kpi-label'>ROCE − WACC Spread</div>
          <div class='kpi-value'>{spread_val:+.2f}%</div>
          <div class='kpi-sub'>{"✅ Value creation" if spread_val > 0 else "⚠️ Value destruction"}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class='kpi-card'>
          <div class='kpi-label'>Effective Tax Rate</div>
          <div class='kpi-value'>{my.get("adv_eff_tax_rate",0):.1f}%</div>
          <div class='kpi-sub'>Used in after-tax cost of debt</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        z_cls = "kpi-card-ok" if z_zone[1]=="green" else \
                "kpi-card-warn" if z_zone[1]=="amber" else "kpi-card-bad"
        z_disp = f"{z_val:.2f}" if z_val is not None else "N/A"
        st.markdown(f"""<div class='kpi-card {z_cls}'>
          <div class='kpi-label'>Altman Z-Score</div>
          <div class='kpi-value'>{z_disp}</div>
          <div class='kpi-sub'>{zone_badge(z_zone)}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>WACC Decomposition — All Companies</div>",
                unsafe_allow_html=True)
    fig_wacc = go.Figure()
    for c in companies:
        d       = data[c]
        E_wt    = d.get("adv_wacc_equity_wt", 0)
        D_wt    = d.get("adv_wacc_debt_wt",   0)
        Re_comp = E_wt / 100 * Re_pct
        Rd_comp = d.get("adv_wacc_after_tax_rd", 0) * D_wt / 100
        fig_wacc.add_trace(go.Bar(
            name=c,
            x=["Equity Cost Contrib.", "Debt Cost Contrib.", "WACC Total"],
            y=[Re_comp, Rd_comp, d.get("adv_wacc", 0)],
            marker_color=COLOR_MAP[c],
            text=[f"{Re_comp:.2f}%", f"{Rd_comp:.2f}%", f"{d.get('adv_wacc',0):.2f}%"],
            textposition="outside",
        ))
    fig_wacc.update_layout(
        barmode="group", height=360, plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(title="% ", gridcolor="#E8EDF8"),
        font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
        title="WACC Components by Company",
    )
    st.plotly_chart(fig_wacc, use_container_width=True)

    col_l, col_r = st.columns(2)
    with col_l:
        fig_spread = go.Figure()
        fig_spread.add_trace(go.Bar(
            name="ROCE %", x=companies,
            y=[data[c]["roce"] for c in companies],
            marker_color=[COLOR_MAP[c] for c in companies], opacity=0.8,
        ))
        fig_spread.add_trace(go.Scatter(
            name="WACC %", x=companies,
            y=[data[c].get("adv_wacc", 0) for c in companies],
            mode="lines+markers",
            line=dict(color="#EF4444", width=3, dash="dash"),
            marker=dict(size=10, color="#EF4444"),
        ))
        fig_spread.update_layout(
            title="ROCE vs WACC (bars above the line = value creation)",
            height=340, plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(title="%", gridcolor="#E8EDF8"),
            font=dict(color="#1E2A4A"), margin=dict(t=50,b=20,l=20,r=20),
        )
        st.plotly_chart(fig_spread, use_container_width=True)

    with col_r:
        fig_spread2 = bar_chart_optional(
            "ROCE − WACC Spread % (positive = value creation)",
            companies, "adv_roce_wacc_spread", data, lambda v: f"{v:+.1f}%",
        )
        if fig_spread2:
            vals_s = [data[c].get("adv_roce_wacc_spread", 0) for c in companies]
            clrs_s = ["#10B981" if v >= 0 else "#EF4444" for v in vals_s]
            fig_spread2.data[0].marker.color = clrs_s
            st.plotly_chart(fig_spread2, use_container_width=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>WACC Detail Table</div>",
                unsafe_allow_html=True)
    ht_w = _table_header(companies, my_company)
    wacc_detail = [
        ("Equity Weight %",       "adv_wacc_equity_wt",    lambda v: pct_str(v)),
        ("Debt Weight %",         "adv_wacc_debt_wt",      lambda v: pct_str(v)),
        ("Cost of Equity Re %",   None,                    lambda _: pct_str(Re_pct)),
        ("Avg Cost of Debt Rd %", "avg_cost_debt",         lambda v: pct_str(v)),
        ("After-tax Rd %",        "adv_wacc_after_tax_rd", lambda v: pct_str(v)),
        ("Effective Tax Rate %",  "adv_eff_tax_rate",      lambda v: pct_str(v)),
        ("WACC %",                "adv_wacc",              lambda v: pct_str(v, 3)),
        ("ROCE %",                "roce",                  lambda v: pct_str(v)),
        ("ROCE−WACC Spread",      "adv_roce_wacc_spread",  lambda v: fmt_pct(v, 2)),
    ]
    for lbl, fld, fn in wacc_detail:
        if fld is None:
            row = f"<tr><td>{lbl}</td>"
            for c in companies:
                cls = "my-company" if c == my_company else ""
                row += f'<td class="{cls}">{fn(None)}</td>'
            row += "</tr>"
            ht_w += row
        else:
            ht_w += _table_row(lbl, companies, data, fld, fn, my_company)
    st.markdown(ht_w + "</tbody></table>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Altman Z-Score (Bankruptcy Risk)</div>",
                unsafe_allow_html=True)
    st.markdown("""
    | Zone | Z-Score | Interpretation |
    |------|---------|---------------|
    | 🟢 Safe Zone | > 2.99 | Low bankruptcy probability |
    | 🟡 Grey Zone | 1.81 – 2.99 | Some risk; monitor closely |
    | 🔴 Distress Zone | < 1.81 | High bankruptcy risk |

    *Model: Altman (1968) public manufacturing firm.*
    """)

    z_present = [(c, data[c].get("adv_altman_z")) for c in companies
                 if data[c].get("adv_altman_z") is not None]
    if z_present:
        fig_z = go.Figure()
        fig_z.add_trace(go.Bar(
            x=[c for c,_ in z_present],
            y=[v for _,v in z_present],
            marker_color=[{"green":"#10B981","amber":"#F59E0B","red":"#EF4444"}
                          .get(data[c].get("adv_altman_zone",("N/A","amber"))[1],"#9CA3AF")
                          for c, _ in z_present],
            text=[f"{v:.2f}" for _,v in z_present],
            textposition="outside",
        ))
        fig_z.add_hline(y=2.99, line_dash="dash", line_color="#10B981",
                         annotation_text="Safe (2.99)", annotation_position="right")
        fig_z.add_hline(y=1.81, line_dash="dash", line_color="#EF4444",
                         annotation_text="Distress (1.81)", annotation_position="right")
        fig_z.update_layout(
            title="Altman Z-Score by Company", height=380,
            plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(title="Z-Score", gridcolor="#E8EDF8"),
            xaxis=dict(tickangle=-20),
            font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=80),
        )
        st.plotly_chart(fig_z, use_container_width=True)

    ht_z = _table_header(companies, my_company, ["Zone"])
    z_comp_rows = [
        ("Z-Score",                    "adv_altman_z",  lambda v: fmt_opt(v, lambda x: f"{x:.3f}")),
        ("X1 = NWC/Total Assets",      "adv_altman_x1", lambda v: fmt_opt(v, lambda x: f"{x:.4f}")),
        ("X2 = Ret.Earnings/TA",       "adv_altman_x2", lambda v: fmt_opt(v, lambda x: f"{x:.4f}")),
        ("X3 = EBIT/Total Assets",     "adv_altman_x3", lambda v: fmt_opt(v, lambda x: f"{x:.4f}")),
        ("X4 = Mkt Cap/Book Liab.",    "adv_altman_x4", lambda v: fmt_opt(v, lambda x: f"{x:.4f}")),
        ("X5 = Sales/Total Assets",    "adv_altman_x5", lambda v: fmt_opt(v, lambda x: f"{x:.4f}")),
    ]
    for lbl, fld, fn in z_comp_rows:
        r = f"<tr><td><b>{lbl}</b></td>"
        for c in companies:
            cls = "my-company" if c == my_company else ""
            r  += f'<td class="{cls}">{fn(data[c].get(fld))}</td>'
        r += "<td>"
        if fld == "adv_altman_z":
            r += "&nbsp;".join(zone_badge(data[c].get("adv_altman_zone",("N/A","amber")))
                               for c in companies)
        r += "</td></tr>"
        ht_z += r
    st.markdown(ht_z + "</tbody></table>", unsafe_allow_html=True)


# ══════════════════════════ TAB 10 — GROWTH & TRENDS ═════════════════════════
with tabs[10]:
    st.markdown("<div class='section-header'>📊 Growth Metrics & Historical Trends</div>",
                unsafe_allow_html=True)

    has_prior = prior_data is not None
    if not has_prior:
        st.info("ℹ️ Upload the **prior round** file to unlock growth metrics and trend charts.")

    my = data[my_company]
    c1,c2,c3,c4,c5 = st.columns(5)
    for col, (lbl, fld) in zip([c1,c2,c3,c4,c5],[
        ("Revenue Growth",    "adv_rev_growth"),
        ("Earnings Growth",   "adv_earnings_growth"),
        ("EBITDA Growth",     "adv_ebitda_growth"),
        ("EPS Growth",        "adv_eps_growth"),
        ("Share Price Growth","adv_price_growth"),
    ]):
        with col:
            val = my.get(fld)
            val_str = f"{val:+.1f}%" if val is not None else "N/A"
            color   = "#10B981" if (val or 0) >= 0 else "#EF4444"
            st.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value' style='color:{color};font-size:22px'>{val_str}</div>
              <div class='kpi-sub'>vs prior round</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Sustainable Growth Rate (SGR)</div>",
                unsafe_allow_html=True)
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(bar_chart("Sustainable Growth Rate %", companies, "adv_sgr",
                                   data, lambda v: f"{v:.1f}%"), use_container_width=True)
    with col_r:
        sgr_df = []
        for c in companies:
            d = data[c]
            sgr_df.append({
                "Company": f"⭐ {c}" if c == my_company else c,
                "ROE (AT) %": pct_str(d["roe_at"]),
                "Payout %":   fmt_opt(d.get("adv_payout_ratio"),  pct_str),
                "Retention %":fmt_opt(d.get("adv_retention_ratio"),pct_str),
                "SGR %":      pct_str(d.get("adv_sgr", 0)),
                "Actual DPS ($)": fmt_opt(d.get("adv_actual_dps"),lambda v:f"${v:.2f}"),
            })
        st.dataframe(pd.DataFrame(sgr_df).set_index("Company"), use_container_width=True)

    st.markdown("---")
    st.markdown("<div class='section-header'>Year-on-Year Growth — All Companies</div>",
                unsafe_allow_html=True)
    ht_g = _table_header(companies, my_company, ["Benchmark"])
    growth_metrics = [
        ("Revenue Growth",     "adv_rev_growth",      "Positive = growing top line"),
        ("Earnings Growth",    "adv_earnings_growth", "High volatility when base < 0"),
        ("EBITDA Growth",      "adv_ebitda_growth",   "Operating profit trend"),
        ("EPS Growth",         "adv_eps_growth",      "Direct driver of share price"),
        ("Book Equity Growth", "adv_equity_growth",   "Internal capital accumulation"),
        ("Share Price Growth", "adv_price_growth",    "Market perception"),
        ("Market Share Growth","adv_ms_growth",       "Competitive position"),
    ]
    for lbl, fld, bench in growth_metrics:
        ht_g += _table_row(lbl, companies, data, fld,
                           lambda v: fmt_growth(v),
                           my_company, [f"<small>{bench}</small>"])
    st.markdown(ht_g + "</tbody></table>", unsafe_allow_html=True)

    if len(sorted(rounds_data.keys())) > 1:
        st.markdown("---")
        st.markdown("<div class='section-header'>Multi-Round Trend Charts</div>",
                    unsafe_allow_html=True)
        all_rounds_sorted = sorted(rounds_data.keys())

        def trend_chart(title, field, fmt_fn=None, ylab=""):
            fig = go.Figure()
            for c in companies:
                y_vals = []
                for r in all_rounds_sorted:
                    val = rounds_data[r].get(c, {}).get(field)
                    y_vals.append(val if val is not None else np.nan)
                fig.add_trace(go.Scatter(
                    x=[f"R{r}" for r in all_rounds_sorted], y=y_vals, name=c,
                    mode="lines+markers",
                    line=dict(color=COLOR_MAP[c], width=2),
                    marker=dict(size=8),
                ))
            fig.update_layout(
                title=title, height=320, plot_bgcolor="white", paper_bgcolor="white",
                yaxis=dict(title=ylab, gridcolor="#E8EDF8"),
                font=dict(color="#1E2A4A"), margin=dict(t=40,b=20,l=20,r=20),
            )
            return fig

        col_l, col_r = st.columns(2)
        with col_l:
            st.plotly_chart(trend_chart("Net Profit Trend", "net_profit", fmt_usd, "$"),
                            use_container_width=True)
            st.plotly_chart(trend_chart("ROCE % Trend", "roce", pct_str, "%"),
                            use_container_width=True)
            st.plotly_chart(trend_chart("Market Share % Trend", "ms_total", pct_str, "%"),
                            use_container_width=True)
        with col_r:
            st.plotly_chart(trend_chart("Revenue Trend", "sales_rev", fmt_usd, "$"),
                            use_container_width=True)
            st.plotly_chart(trend_chart("Share Price Trend", "share_price",
                                        lambda v: f"${v:.2f}", "$"),
                            use_container_width=True)
            st.plotly_chart(trend_chart("EPS Trend", "eps",
                                        lambda v: f"${v:.2f}", "$"),
                            use_container_width=True)
    else:
        st.info("📈 Upload 2+ round files to see trend charts.")


# ══════════════════════════ TAB 11 — CFO CHEAT SHEET ═════════════════════════
with tabs[11]:
    st.markdown("<div class='section-header'>📄 CFO Cheat Sheet — All Ratios Explained</div>",
                unsafe_allow_html=True)

    def cheat(title, formula, body, benchmark, assumption_note=None):
        note = f'<br><span class="assume-tag">Assumption: {assumption_note}</span>' \
               if assumption_note else ""
        st.markdown(f"""<div class='cheat-box'>
          <div class='cheat-title'>📐 {title}</div>
          <div class='cheat-formula'>Formula: {formula}</div>
          <div class='cheat-body'>{body}</div>
          <div class='cheat-body'><b>Benchmark:</b> {benchmark}{note}</div>
        </div>""", unsafe_allow_html=True)

    sections = {
        "💧 LIQUIDITY": [
            ("Net Working Capital (NWC)",
             "(Equity + LT Debt) − Fixed Assets  OR  Current Assets − Current Liabilities",
             "Liquidity buffer: how much long-term funding remains after covering fixed assets.",
             "Positive = ✅ | Negative = ⚠️", None),
            ("Net Cash Position",
             "Cash − Short-term Financial Debt",
             "Immediate cash surplus. Positive = can meet obligations from cash alone.",
             "Positive = ✅", None),
            ("Current Ratio",
             "Current Assets ÷ Current Liabilities",
             "Times current assets cover current debts.",
             "1.0–1.5 = normal | > 1.5 = safe | < 1.0 = risky", None),
            ("Quick Ratio (Acid Test)",
             "(Current Assets − Inventory) ÷ Current Liabilities",
             "Stricter: excludes slow-to-liquidate inventory.",
             "≥ 1.0 = good | < 0.8 = watch", None),
            ("Cash Cycle",
             "Stock Days + Receivables Days − Payables Days",
             "Net days the company must self-finance its operations.",
             "Negative = ✅ favourable | Positive = ⚠️ must finance the gap", None),
        ],
        "📈 PROFITABILITY": [
            ("EBITDA Margin",
             "EBITDA ÷ Sales × 100%",
             "Operating efficiency before non-cash charges.",
             "≥ 15% strong | 8–15% acceptable | < 8% weak", None),
            ("Net Sales Margin (ROS)",
             "EBIT ÷ Sales × 100%",
             "Profit from operations after depreciation.",
             "> 10% good | < 5% thin", None),
            ("ROE (before/after tax)",
             "Profit (before/after tax) ÷ Equity × 100%",
             "Return earned for shareholders. Must exceed cost of equity to create value.",
             "≥ Re (cost of equity) = value-creating", None),
            ("ROCE",
             "EBIT ÷ (Total Assets − Current Liabilities) × 100%",
             "Return on all long-term capital employed.",
             "> 20% excellent | 15–20% good | < 15% review", None),
            ("ROCE − WACC Spread",
             "ROCE − WACC",
             "Positive spread = company earns more than its cost of capital = value creation.",
             "> 0% = value creation | < 0% = value destruction", None),
        ],
        "🏦 SOLVENCY": [
            ("Financial Independence",
             "Equity ÷ Total Assets × 100%",
             "Owner-funded proportion of the balance sheet.",
             "≥ 40% healthy | < 30% risky", None),
            ("Net Gearing",
             "(LT Debt + ST Debt − Cash) ÷ Equity × 100%",
             "Net debt as a multiple of equity.",
             "< 50% low | 50–100% moderate | > 100% high risk", None),
        ],
        "🌐 REGIONAL ANALYSIS (NEW v3)": [
            ("Regional Income Statement",
             "Same structure as Global IS, isolated per subsidiary",
             "USA and China are manufacturing subsidiaries with full P&Ls. "
             "Europe is a distribution entity (no inhouse/contract mfg, no R&D). "
             "Internal transfer revenues are included in regional IS and eliminated "
             "at global consolidation.",
             "Compare regional margins to global; weak regions dilute group performance",
             "Parsed from rows 55–112 (USA), 143–196 (China), 224–271 (Europe)"),
            ("Regional Market Share",
             "Units sold in region ÷ Total units sold in region (all companies)",
             "Shows competitive strength in each geography and technology segment. "
             "A company may dominate globally but be weak in one specific region.",
             "Higher regional share = stronger local competitive position",
             "Parsed from rows 364–369 (USA), 402–407 (China), 440–445 (Europe)"),
            ("Revenue Contribution %",
             "Regional Revenue ÷ Global Revenue × 100%",
             "Shows geographic diversification. Over-reliance on one region "
             "increases risk; balanced split signals robust distribution strategy.",
             "Balanced across regions reduces geographic risk",
             "Derived from regional sales_rev / global sales_rev"),
        ],
        "💎 VALUATION (v2)": [
            ("Enterprise Value (EV)",
             "Market Cap + LT Debt + ST Debt − Cash",
             "Total acquisition cost — what you pay to own the whole business.",
             "Compare across rounds; falling EV may signal concerns",
             "Book value of debt used"),
            ("P/B Ratio",
             "Share Price ÷ Book Value per Share",
             "Market premium over accounting net worth. "
             "< 1 = market values company below book; > 1 = market expects returns above Re.",
             "< 1 potential bargain | > 3 growth premium priced in",
             "Book equity used; excludes off-balance-sheet items"),
            ("P/Sales Ratio",
             "Market Cap ÷ Sales Revenue",
             "Market value relative to revenue. Useful when earnings are volatile.",
             "Depends on industry margins", None),
            ("EV / EBITDA",
             "Enterprise Value ÷ EBITDA",
             "Capital-structure-neutral profitability multiple; most-used M&A benchmark.",
             "< 8× cheap | 8–15× typical | > 15× expensive",
             "Uses global EBITDA; negative EV or EBITDA → N/A"),
            ("EV / Sales",
             "Enterprise Value ÷ Sales Revenue",
             "Revenue-based multiple; useful for unprofitable companies.",
             "< 1× very cheap | 1–3× typical", None),
            ("RIM Intrinsic Value",
             "IV = BVps + (EPS − Re × BVps) ÷ (Re − g)",
             "Intrinsic value equals book value plus discounted residual income. "
             "Works without dividends.",
             "IV > Market Price → potential undervaluation",
             f"Single-stage Ohlson 1995; Re = {Re_pct:.1f}%, g = {g_pct:.1f}%"),
        ],

        "📐 WACC & RISK (NEW)": [
            ("WACC",
             "(E/V)×Re + (D/V)×Rd×(1−t)",
             "Weighted Average Cost of Capital: the minimum return the company "
             "must earn on all capital employed to satisfy investors and lenders. "
             "E = equity value, D = debt, Re = cost of equity, Rd = cost of debt, t = tax rate. "
             "When ROCE > WACC, the company creates value.",
             "Compare against ROCE. Positive spread = value creation.",
             f"Re = {Re_pct:.1f}% (user input); Rd = Cesim company-specific; "
             "E = market cap (toggle to book in sidebar)"),
            ("Altman Z-Score",
             "1.2X1 + 1.4X2 + 3.3X3 + 0.6X4 + 1.0X5",
             "Bankruptcy predictor for public manufacturing firms (Altman 1968). "
             "X1=NWC/TA, X2=RetainedEarnings/TA, X3=EBIT/TA, "
             "X4=MarketCap/BookLiabilities, X5=Sales/TA.",
             "Z > 2.99 = Safe | 1.81–2.99 = Grey zone | < 1.81 = Distress",
             "Calibrated for public manufacturing. X4 uses market cap / book liabilities."),]
    }

    for section_title, items in sections.items():
        st.markdown(f"### {section_title}")
        for title, formula, body, bench, assume in items:
            cheat(title, formula, body, bench, assume)
        st.markdown("")

    st.markdown("---")
    st.markdown("### 🎯 Quick Decision Framework")
    st.markdown("""
| Situation | Key ratio | Action |
|-----------|-----------|--------|
| Cash drying up | Net Cash, Cash Cycle | Reduce inventory, collect faster, take LT loan |
| Margins shrinking | EBITDA Margin, ROS | Cut contract mfg costs, raise price, reduce promo |
| Share price falling | EPS, RIM vs Price, TSR | Improve profitability or consider dividends |
| Debt getting expensive | Gearing, Int. Rates, WACC | Repay debt with operating cash flow |
| WACC > ROCE | ROCE−WACC Spread | Stop value-destroying investments; restructure |
| Altman Z < 1.81 | Z-Score components | Urgently reduce debt, improve EBIT/assets ratio |
| SGR < target growth | SGR, ROE, Payout | Retain more earnings, or issue equity |
| Weak region found | Regional tab | Investigate pricing, MS, and cost structure per region |
| Region over-relied on | Revenue Contribution % | Diversify sales strategy to other regions |
""")
    st.markdown("---")
    st.markdown("### 📌 Three Pillars Summary")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""**💧 Liquidity**
Pay bills today:
- Current Ratio ≥ 1
- Positive Net Cash
- Negative Cash Cycle""")
    with col2:
        st.markdown("""**📈 Profitability**
Earn money efficiently:
- EBITDA Margin ≥ 10%
- ROCE > WACC (spread > 0)
- SGR covers growth plan""")
    with col3:
        st.markdown("""**🌐 Regional Balance**
Grow sustainably:
- No single-region dependency
- Positive margins in all regions
- Growing regional market shares""")