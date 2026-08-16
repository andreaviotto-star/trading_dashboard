"""
Quant Trading System Analytics Dashboard  v25.1
=================================================
Usage: streamlit run app.py

═══════════════════════════════════════════════════════════════
CHANGES IN v25.0
═══════════════════════════════════════════════════════════════
  • Correlation audit uses overlapping daily dollar P&L observations; no invented zero P&L for pre-launch periods
  • Portfolio volatility, sum of individual volatility and diversification ratio
  • System and correlation-cluster risk contributions
  • Rolling average-correlation diagnostic
  • Correlation-aware Risk Parity retained alongside volatility-only Risk Parity
  • Regime-conditioned Monte Carlo using empirical volatility-regime transitions
  • Tooltip descriptions formatted as separate Formula / Purpose / In plain English paragraphs

CHANGES IN v25.1 — 2026-08-16
═══════════════════════════════════════════════════════════════
  • Fixed correlation clustering by restoring fcluster and surfacing analytical exceptions
  • Portfolio Impact now compares current vs advisory sizing on the same estimator/window/system set
  • Tab 2 risk follows the selected system set and selected date window
  • Removed tab-level st.stop() flow; empty selections now return from the tab renderer only
  • Added canonical portfolio_risk() engine for diversified daily risk and risk contributions
  • Added constant-composition analytics toggle (default ON) and live-system-count diagnostic
  • Added daily/weekly/monthly correlation comparison; monthly is the primary diversification horizon
  • Added Ledoit-Wolf covariance shrinkage and covariance-quality diagnostics for sizing
  • Correlation-aware Risk Parity now requires 252 complete observations and uses shrunk covariance
  • Corrected stale Ward labels and misleading cluster-risk wording
  • Added golden_master.py for frozen-output regression testing
  • Added standalone global matching and empirical-slippage modules; integration is gated on available source data


1. DYNAMIC SIZING (Tab 1)
   - Per-system number_input in Tab 1 System Explorer.
   - Sizing panel above the overview table: compact number_inputs
     for all systems in a responsive grid.
   - Default = qty from ReadMe.txt; user can override.
   - All metrics and equity curves rescale live:
       scaled_P&L = raw_P&L × (current_n / default_n)
   - "Reset to defaults" button clears session_state overrides.
   - Raw P&L from TradeStation is NEVER changed — only displayed
     P&L is multiplied by the scaling ratio.

2. SYSTEM NAMES FROM README.TXT (everywhere)
   - si["display_name"] = name exactly as written left of ":" in ReadMe.
   - 16 / 18 files matched; 2 files absent from ReadMe show
     their filename + ⚠️ badge.
   - ReadMe names used in: overview table, equity legend,
     correlation heatmap labels, portfolio bar chart, Excel export.

3. SIZING DEFAULTS FROM README.TXT (improved parsing)
   - Regex r'(N+)\\s*([A-Z]{2,3})' on text before first comma.
   - ES reversal: first qty = 3 (conservative L2), label "3+4 MES (L2+L3)".
   - GC donchian: qty = 1 GC (strips "prima erano 5 MGC…" note).
   - NQ breakout sessione continua: 1 NQ (full contract, not micro).
   - CL donchian: 1 CL (full contract).

═══════════════════════════════════════════════════════════════
RETAINED FROM v10.0
═══════════════════════════════════════════════════════════════
  • Raw-trade P&L engine (no TS multiplier errors)
  • Light theme (plotly_white)
  • 10-year default lookback
  • Risk Parity sizing (Tab 3)
  • Drawdown Decomposition (Tab 2)
  • Clustered Correlation Heatmap (Tab 3)
  • Excel Export (sidebar)
  • What-If Sizing Simulator (Tab 2)
  • Monte Carlo Projection (Tab 4)
  • Regime Analysis (Tab 4)
  • Rolling Health Monitor (Tab 4)

COLUMN MAP — "Trades List" tab (verified all 18 xlsx):
  Entry row: col[0]=trade#, col[1]=direction, col[2]=entry_dt,
             col[6]=n_contracts, col[7]=trade_pnl
  Exit row:  col[0]=None, col[1]=exit_type, col[2]=exit_dt,
             col[6]=trade P&L ($, full position) ← canonical
             col[7]=cumulative P&L ($)
  Dollar values = full-position dollars from TradeStation.
  NEVER multiply by contract count again.
"""

import io
import os
import re
import warnings
from pathlib import Path
from datetime import datetime
from collections import defaultdict

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from scipy.cluster.hierarchy import linkage, leaves_list, fcluster
from scipy.spatial.distance import squareform
from scipy.optimize import minimize
from sklearn.covariance import LedoitWolf

warnings.filterwarnings("ignore")

# ── Trade-matching engine (new, v11.1) — imported defensively so the rest of
# the dashboard keeps working even if the module or its extra dependency
# (pdfplumber, only needed for the optional PDF reconciliation) is missing.
try:
    import tradestation_matcher as tsm
    _TSM_IMPORT_ERROR = None
except Exception as _tsm_exc:
    tsm = None
    _TSM_IMPORT_ERROR = str(_tsm_exc)

try:
    import empirical_slippage as _emp_slip
    _EMP_SLIP_IMPORT_ERROR = None
except Exception as _emp_slip_exc:
    _emp_slip = None
    _EMP_SLIP_IMPORT_ERROR = str(_emp_slip_exc)

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
APP_VERSION = "v25.1"
DATA_DIR = Path(os.path.dirname(os.path.abspath(__file__)))

# $/contract round-trip commission (NOT flat per-trade) -- see build_net_equity().
# Defaults below are blended from Andrea's real June 2026 statement, weighted
# by actual contract volume (not a simple average of per-root rates):
#   micro: (MES $401.94 + MGC $29.28 + MNQ $230.55) / (462+24+266 contracts) = $0.88/contract
#   mini/full: (CL $19.62 + GC $68.4 + NQ $37.8) / (6+20+12 contracts) = $3.31/contract
# Re-derived 2026-07 after Andrea flagged the reconciliation table's
# statement_fee ($787.59) vs fills_fee_reconstructed ($338.80) gap -- the
# PREVIOUS defaults here ($1.75/$6.60) were almost exactly 2x these correct
# values (ratio 1.99 on both), meaning they'd been calibrated against the
# wrong side of that reconciliation at some point. The order-history file's
# own Commission column ($0.40/contract micro, $1.00 full-size) is NOT the
# real cost -- it's only the flat exchange/clearing fee; the statement's
# real total also includes NFA/broker markup, which is why these defaults
# must come from statement_fee, not fills_fee_reconstructed.
# Legacy fallbacks (kept for compatibility if a mapping isn't provided)
COMMISSION_PER_MICRO = 0.88
COMMISSION_PER_MINI  = 3.31

# Per-contract round-trip commission rates ($/contract RT)
# Updated to values provided by the user.
COMMISSION_RATES = {
    "MES": 1.74,
    "MNQ": 1.73,
    "MGC": 2.45,
    "NQ":  5.00,
    "GC":  5.54,
    "CL":  5.31,
}

MICRO_PREFIXES = {"MES", "MNQ", "MGC", "MCL"}
FALLBACK_CTYPE = {"ES": "MES", "NQ": "MNQ", "GC": "MGC", "CL": "CL"}

COLORS = px.colors.qualitative.Plotly + px.colors.qualitative.Dark24
THEME  = "plotly_white"


# ─────────────────────────────────────────────────────────────────────────────
# TABLE DISPLAY / FORMATTING
# Presentation only: calculations retain full precision.
# ─────────────────────────────────────────────────────────────────────────────

def _fmt_money(v):
    if pd.isna(v): return "—"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"-${abs(v):,.0f}" if v < 0 else f"${v:,.0f}"


def _fmt_money2(v):
    if pd.isna(v): return "—"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"-${abs(v):,.2f}" if v < 0 else f"${v:,.2f}"


def _fmt_number(v):
    if pd.isna(v): return "—"
    try:
        return f"{float(v):,.0f}"
    except (TypeError, ValueError):
        return str(v)


def _fmt_decimal(v):
    if pd.isna(v): return "—"
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return str(v)


def _fmt_one_decimal(v):
    if pd.isna(v): return "—"
    try:
        return f"{float(v):,.1f}"
    except (TypeError, ValueError):
        return str(v)


def _fmt_percent(v):
    if pd.isna(v): return "—"
    try:
        return f"{float(v):.1%}"
    except (TypeError, ValueError):
        return str(v)


def _fmt_price(v):
    if pd.isna(v): return "—"
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return str(v)


METRIC_HELP = {
    'Account Balance': '**Purpose:** reference capital for risk/equity calculations. It does not change historical P&L.\n\n**In plain English:** this is the pot of money against which we judge how hard the portfolio is leaning on the account.',
    'Net Profit ($)': '**Formula:** total net P&L after modeled commissions.\n\n**Purpose:** measure the actual dollars generated over the selected period.\n\n**In plain English:** how much money did the strategy put in the till?',
    'Portfolio P&L': "**Formula:** sum of all selected systems' net P&L.\n\n**Purpose:** show the combined portfolio bottom line.\n\n**In plain English:** what did the whole trading desk make, not just one strategy?",
    'Ann. P&L': '**Formula:** total P&L ÷ elapsed years.\n\n**Purpose:** put histories of different lengths onto a common annualized scale.\n\n**In plain English:** if this pace continued, roughly how much would the strategy make per year?',
    'Max Drawdown': '**Formula:** largest peak-to-trough decline in cumulative P&L.\n\n**Purpose:** quantify the worst historical pain.\n\n**In plain English:** the number the risk committee remembers long after everyone has forgotten the Sharpe ratio.',
    'P&L Sharpe': '**Formula:** mean daily P&L ÷ daily P&L volatility × √252.\n\n**Purpose:** measure P&L efficiency, not return on invested capital.\n\n**In plain English:** how much daily P&L are we getting for each unit of daily pain? A higher number means a smoother money-making machine.',
    'Ann. P&L / Max DD': '**Formula:** annualized P&L ÷ absolute maximum drawdown.\n\n**Purpose:** measure annualized P&L generated per dollar of peak-to-trough pain.\n\n**In plain English:** how much money did we make for every dollar of historical suffering? Informally: the spleen-sanity ratio.',
    'Profit Factor': '**Formula:** gross winning P&L ÷ gross losing P&L.\n\n**Purpose:** measure the payoff efficiency of the trade distribution.\n\n**In plain English:** for every $1 we lost, how many dollars did the winners make back? Above 1.0 means the winners paid the rent.',
    'Win Rate': '**Formula:** profitable trades ÷ total trades.\n\n**Purpose:** show how often trades make money.\n\n**In plain English:** how often do we get to be right? A 40% win rate can still be excellent if the winners are much larger than the losers.',
    'Current Daily Risk': '**Formula:** portfolio daily P&L volatility at the current contract sizing.\n\n**Purpose:** estimate normal one-day fluctuation, not a worst-case loss or margin requirement.\n\n**In plain English:** on a typical day, how much financial turbulence should we expect?',
    'Daily Risk / Equity': "**Formula:** current daily P&L volatility ÷ account balance.\n\n**Purpose:** express daily portfolio risk as a percentage of capital.\n\n**In plain English:** what percentage of the account's equity is represented by one day's normal risk?",
    'Annualized Risk / Equity': "**Formula:** annualized P&L volatility ÷ account balance.\n\n**Purpose:** provide a risk-leverage proxy relative to your capital.\n\n**In plain English:** if today's risk profile persisted, how volatile is this book relative to the size of the account?\n\n- This is **not** notional futures leverage.",
    'Current Sizing': "**Reference point:** 100% means the contract counts currently selected in the Systems tab.\n\n**Purpose:** compare alternative allocations with the book you actually trade today.\n\n**In plain English:** today's portfolio is the benchmark; everything else is a what-if.",
    'Current daily risk': "**Formula:** one-day portfolio P&L volatility at the current contract sizes.\n\n**Purpose:** quantify current risk budget usage.\n\n**In plain English:** how much turbulence are we buying with today's positions?",
    'Current risk / equity': "**Formula:** current daily P&L volatility ÷ account balance.\n\n**Purpose:** compare today's risk with available capital.\n\n**In plain English:** how hard are we leaning on the account right now?",
    'RP advisory daily risk': "**Formula:** portfolio daily P&L volatility if all Risk Parity Advisory contract counts were implemented.\n\n**Purpose:** show the portfolio's normal daily turbulence after the proposed sizing changes.\n\n**In plain English:** what would the book normally wobble by if we followed the recommendation?",
    'RP risk / equity': "**Formula:** Risk Parity Advisory daily P&L volatility ÷ account balance.\n\n**Purpose:** compare the proposed allocation with available capital.\n\n**In plain English:** how much of the account's risk budget would the proposed book consume?",
    'Current → RP annual risk': '**Formula:** annualized P&L volatility ÷ account balance for current and advisory sizing.\n\n**Purpose:** show the risk change before changing any contracts.\n\n**In plain English:** does the proposed makeover actually make the portfolio calmer, or are we just moving deck chairs?',
     'Avg Pairwise ρ': '**Formula:** average correlation across valid overlapping daily P&L observations.\n\n**Purpose:** gauge broad co-movement between systems.\n\n**In plain English:** when one system has a bad day, how many of its friends tend to join the party?',
    'Portfolio Daily Vol': '**Formula:** √(wᵀΣw), using the current-size daily P&L covariance matrix.\n\n**Purpose:** measure the actual one-day volatility of the combined portfolio after correlations are taken into account.\n\n**In plain English:** how much can the **whole book** normally wobble by?',
    'Correlation-aware RP Daily Risk': '**Formula:** daily P&L volatility of the proposed integer contract allocation, including correlations.\n\n**Purpose:** show the actual portfolio risk produced by correlation-aware Risk Parity.\n\n**In plain English:** how much should the proposed whole book normally wobble by in a day?',
    'Target daily risk': '**Formula:** account balance × target annualized risk ÷ √252.\n\n**Purpose:** convert the annual portfolio-risk target into a one-day P&L volatility target.\n\n**In plain English:** the daily risk budget implied by your 25% target.',
    'Target annual risk': '**Formula:** annualized portfolio P&L volatility ÷ account balance.\n\n**Purpose:** show the annualized risk target used by the correlation-aware sizing calculation.\n\n**In plain English:** how much annual turbulence are we deliberately budgeting?',
    'Resulting annual risk': '**Formula:** proposed portfolio daily volatility × √252 ÷ account balance.\n\n**Purpose:** show where the rounded whole-contract allocation actually lands relative to the target.\n\n**In plain English:** did whole contracts get us close to the risk budget?',
    'Sum Individual Vol': "**Formula:** sum of each system's daily P&L volatility.\n\n**Purpose:** show the risk if every system moved in the same direction at the same time.\n\n**In plain English:** the scary version before diversification gets any credit.",
    'Diversification Ratio': '**Formula:** sum of individual volatilities ÷ portfolio volatility.\n\n**Purpose:** quantify how much diversification reduces aggregate volatility.\n\n**In plain English:** how much risk reduction are we getting because the systems do **not** all move together?\n\n- **1.00** = essentially no diversification. Higher = more diversification.',
    'Risk Contribution': "**Formula:** each system's marginal contribution to portfolio volatility × its current exposure.\n\n**Purpose:** identify which systems actually drive total portfolio risk after correlations.\n\n**In plain English:** who is paying most of the risk bill?\n\n- A system can be volatile on its own but contribute little if it diversifies the portfolio.",
    'Cluster Risk %': '**Formula:** risk contribution of all systems in the cluster ÷ total portfolio risk.\n\n**Purpose:** identify concentration at the **group** level.\n\n**In plain English:** several different strategies can still be one trade wearing different hats.',
    'Rolling Avg ρ': '**Formula:** average pairwise daily P&L correlation over a rolling window.\n\n**Purpose:** show whether diversification is stable or disappearing in recent market conditions.\n\n**In plain English:** are the strategies still minding their own business lately?',
    'Regime-conditioned bootstrap': '**Method:** simulate a sequence of historical volatility regimes using their observed transition probabilities, then sample daily portfolio P&L from the corresponding regime.\n\n**Purpose:** account for the fact that markets tend to stay in similar volatility states for several days rather than behaving as completely independent observations.\n\n**In plain English:** if the portfolio is currently in a stormy period, the simulation gives stormy days a better chance of sticking around.\n\n- This is still a historical bootstrap, not a forecast of the next regime.',
    'Current Vol Regime': '**Formula:** current portfolio realized volatility classified against the historical low/high percentile thresholds.\n\n**Purpose:** show the volatility state from which the regime-conditioned simulation starts.\n\n**In plain English:** what kind of market weather are we starting with?',
    'Regime Persistence': '**Formula:** probability that the current volatility regime remains the same on the next simulated day.\n\n**Purpose:** quantify how sticky the regime is in the historical data.\n\n**In plain English:** once the weather turns rough, how often does it stay rough tomorrow?',
    'P(loss)': '**Formula:** share of Monte Carlo simulations ending with negative portfolio P&L.\n\n**Purpose:** estimate downside frequency under the selected simulation assumptions.\n\n**In plain English:** how often does the simulated path finish in the red?',
    'Median': '**Formula:** 50th percentile of simulated ending P&L.\n\n**Purpose:** show the middle Monte Carlo outcome.\n\n**In plain English:** if we ran the future 1,000 times, this is the outcome sitting right in the middle.',
    'Mean': '**Formula:** average simulated ending P&L across all simulated paths.\n\n**Purpose:** show the arithmetic average outcome.\n\n**In plain English:** what the simulations make on average — remembering that a few extreme paths can move the average.',
    '5th pct': '**Formula:** 5th percentile of simulated ending P&L.\n\n**Purpose:** represent a downside-tail scenario.\n\n**In plain English:** roughly 1 simulation out of 20 finishes worse than this.',
    '95th': '**Formula:** 95th percentile of simulated ending P&L.\n\n**Purpose:** represent an upside-tail scenario.\n\n**In plain English:** roughly 1 simulation out of 20 finishes better than this.',
    '5% Terminal ES': '**Formula:** average ending P&L of simulations in the worst 5% of outcomes.\n\n**Purpose:** measure the average severity of the bad tail, not just the cutoff.\n\n**In plain English:** if things land in the worst 5%, how bad is the average outcome?',
    'Historical days': '**Formula:** number of valid historical daily P&L observations used by the simulation.\n\n**Purpose:** show how much historical evidence the Monte Carlo has available.\n\n**In plain English:** how many days of history are we asking the simulation to learn from?',
}

def metric_help(label):
    return METRIC_HELP.get(label)

def metric_with_help(container, label, value, delta=None, **kwargs):
    help_text = kwargs.pop("help", None) or metric_help(label)
    return container.metric(label, value, delta=delta, help=help_text, **kwargs)


def _infer_table_formats(df: pd.DataFrame) -> dict:
    """Infer readable formats safely. Never apply numeric formatters to text columns.

    Formatting is presentation-only. A semantic name is not enough to make a
    column numeric: e.g. ``System`` contains text, while some imported numeric
    columns may have object dtype. We therefore require the column to be
    genuinely numeric (or safely coercible to numeric) before assigning a
    numeric formatter.
    """
    formats = {}
    for col in df.columns:
        series = df[col]

        if pd.api.types.is_datetime64_any_dtype(series):
            formats[col] = lambda v: "—" if pd.isna(v) else pd.Timestamp(v).strftime("%Y-%m-%d %H:%M")
            continue

        # Never infer numeric formatting from the column name alone.
        # Text/object columns such as System, Market, Status, File, etc.
        # must pass through untouched. For object columns containing numbers,
        # allow formatting only when every non-null value is numeric.
        if not pd.api.types.is_numeric_dtype(series):
            non_null = series.dropna()
            if non_null.empty:
                continue
            converted = pd.to_numeric(non_null, errors="coerce")
            if converted.isna().any():
                continue

        name = str(col).strip().lower()
        if any(k in name for k in ["commission", "fee"]):
            formats[col] = _fmt_money2
        elif any(k in name for k in ["p&l", "pnl", "profit", "drawdown", "daily vol", "vol /", "volatility", "risk", "impact", "delta", "cash", "capital", "equity"]):
            formats[col] = _fmt_money
        elif any(k in name for k in ["win %", "win rate", "weight", "probability", "p(loss)", "%"]):
            formats[col] = _fmt_percent
        elif "price" in name:
            formats[col] = _fmt_price
        elif any(k in name for k in ["trades", "contracts", "current n", "suggested n", "advisory n", "raw n", "qty", "quantity", "count"]):
            formats[col] = _fmt_number
        elif any(k in name for k in ["sharpe", "calmar", "profit factor", "pf", "corr", "correlation", "scale", "slippage", "ticks"]):
            formats[col] = _fmt_decimal
        # Unknown numeric columns are deliberately left alone rather than
        # guessing a display precision. This prevents accidental formatting
        # of identifiers and keeps the renderer safe for arbitrary tables.

    return formats


def render_table(data, **kwargs):
    """Render raw DataFrames consistently; pass existing Styler objects through."""
    if isinstance(data, pd.DataFrame):
        return st.dataframe(data.style.format(_infer_table_formats(data), na_rep="—"), **kwargs)
    return st.dataframe(data, **kwargs)


def style_overview_table(df: pd.DataFrame):
    styled = df.style.format({
        "Net Profit": _fmt_money, "Max DD": _fmt_money,
        "P&L Sharpe": _fmt_decimal, "Ann. P&L / Max DD": _fmt_decimal,
        "PF": _fmt_decimal, "Win %": _fmt_percent, "# Trades": _fmt_number,
    }, na_rep="—")

    def pnl_colour(v):
        if pd.isna(v): return ""
        return "background-color: rgba(46, 125, 50, 0.10)" if v > 0 else ("background-color: rgba(198, 40, 40, 0.10)" if v < 0 else "")
    def pf_colour(v):
        if pd.isna(v): return ""
        if v < 1.0: return "background-color: rgba(198, 40, 40, 0.14)"
        if v < 1.3: return "background-color: rgba(245, 158, 11, 0.12)"
        return "background-color: rgba(46, 125, 50, 0.12)"
    def sharpe_colour(v):
        if pd.isna(v): return ""
        if v < 0: return "background-color: rgba(198, 40, 40, 0.14)"
        if v < 1: return "background-color: rgba(245, 158, 11, 0.10)"
        return "background-color: rgba(46, 125, 50, 0.12)"
    if "Net Profit" in df.columns: styled = styled.map(pnl_colour, subset=["Net Profit"])
    if "PF" in df.columns: styled = styled.map(pf_colour, subset=["PF"])
    if "P&L Sharpe" in df.columns: styled = styled.map(sharpe_colour, subset=["P&L Sharpe"])
    return styled


# ─────────────────────────────────────────────────────────────────────────────
# README PARSING  —  v11: returns display_name + conservative first_qty
# ─────────────────────────────────────────────────────────────────────────────

def _build_contract_label(alloc_part: str) -> str:
    """
    Convert the right-hand side of a ReadMe allocation line into a clean
    display label.

    Examples
    --------
    "3 MES"                                         → "3 MES"
    "1GC, prima erano 5 MGC …"                      → "1 GC"
    "1 MES - inattivo"                              → "1 MES  ⏸ inattivo"
    "3  MES sul secondo livello + 4 MES sul terzo"  → "3+4 MES (L2+L3)"
    """
    # Strip everything after the first comma (historical notes)
    part = alloc_part.split(",")[0].strip()

    # Normalise spacing: "1GC" → "1 GC"
    part = re.sub(r"(\d)\s*([A-Z]{2,3})", r"\1 \2", part)

    # Detect multi-level pattern  "3 MES sul secondo livello + 4 MES sul terzo livello"
    if re.search(r"sul\s+secondo\s+livello", part, re.I) and \
       re.search(r"sul\s+terzo\s+livello",   part, re.I):
        nums   = re.findall(r"(\d+)\s+([A-Z]{2,3})", part)
        ctypes = list({ct for _, ct in nums})
        ctype  = ctypes[0] if ctypes else "MES"
        qtys   = [q for q, _ in nums]
        return "+".join(qtys) + f" {ctype} (L2+L3)"

    # Single-level verbose descriptions
    part = re.sub(r"\s+sul\s+secondo\s+livello", " (L2)", part, flags=re.I)
    part = re.sub(r"\s+sul\s+terzo\s+livello",   " (L3)", part, flags=re.I)
    part = re.sub(r"\s+sul\s+primo\s+livello",   " (L1)", part, flags=re.I)

    # Inactive flag
    if re.search(r"inattiv", part, re.I):
        part = re.sub(r"\s*-?\s*inattiv\w*", "", part, flags=re.I).strip()
        part = part + "  ⏸ inattivo"

    return re.sub(r"  +", "  ", part).strip()


def parse_readme(readme_path: str) -> dict:
    """
    Parse ReadMe.txt and return a dict keyed by *normalised* name.

    Each entry::

        norm_name → {
            "display_name":   str,   # exactly as written left of ":"
            "contract_label": str,   # e.g. "3+4 MES (L2+L3)"
            "default_n":      int,   # conservative first qty (e.g. 3 for reversal)
            "total_n":        int,   # sum of all qty (e.g. 7 for reversal)
            "ctype":          str,   # primary contract type token
            "is_inactive":    bool,
        }
    """
    alloc_re = re.compile(r"(\d+)\s*([A-Z]{2,3})")
    entries  = {}
    sort_idx = 0   # preserve ReadMe line order

    try:
        with open(readme_path, "r", encoding="utf-8", errors="ignore") as fh:
            for raw in fh:
                line = raw.strip()
                if ":" not in line:
                    continue
                colon      = line.index(":")
                name_part  = line[:colon].strip()
                alloc_part = line[colon + 1:].strip()
                if not name_part or not alloc_part:
                    continue

                clean   = alloc_part.split(",")[0]
                matches = alloc_re.findall(clean)
                if not matches:
                    continue

                by_type = defaultdict(int)
                for qty_s, ctype in matches:
                    by_type[ctype] += int(qty_s)

                primary_ctype = matches[0][1]
                first_qty     = int(matches[0][0])   # conservative default
                total_qty     = by_type[primary_ctype]

                label       = _build_contract_label(alloc_part)
                is_inactive = "inattivo" in label

                norm = name_part.lower().replace("-", " ")
                norm = re.sub(r"\s+", " ", norm).strip()

                entries[norm] = {
                    "display_name":   name_part,
                    "contract_label": label,
                    "default_n":      first_qty,
                    "total_n":        total_qty,
                    "ctype":          primary_ctype,
                    "is_inactive":    is_inactive,
                    "sort_order":     sort_idx,
                }
                sort_idx += 1
    except FileNotFoundError:
        pass

    return entries


_README_PATH    = str(DATA_DIR / "ReadMe.txt")
_README_ENTRIES = parse_readme(_README_PATH)


def get_alloc(stem: str) -> dict:
    """
    Match an xlsx filename stem to a ReadMe entry via word-set containment.

    Strategy: normalise both the ReadMe entry name and the xlsx stem by
    lower-casing and replacing punctuation/underscores with spaces, then
    check whether ALL words in the ReadMe name appear in the stem's word-set.
    Pick the longest (most-specific) match to avoid ambiguity.

    This correctly handles cases like:
      ES_breakout_sessione_regolare__long  →  'ES breakout regolare - long'
      (stem has extra word 'sessione' not in ReadMe name — still matches
       because ReadMe words are a subset of stem words)

    Falls back to filename-derived name + ⚠️ if no ReadMe match found.
    """
    symbol = stem.split("_")[0].upper()

    # Normalise stem: underscores → spaces, keep all words
    norm_stem  = stem.lower().replace("__", " ").replace("_", " ")
    stem_words = set(norm_stem.split())

    best     = None
    best_len = 0
    for norm_name, entry in _README_ENTRIES.items():
        # norm_name already has dashes converted to spaces by parse_readme
        kw = set(norm_name.split())
        if kw.issubset(stem_words) and len(kw) > best_len:
            best_len = len(kw)
            best     = entry.copy()

    if best:
        ctype    = best["ctype"]
        is_micro = ctype in MICRO_PREFIXES
        return {
            "symbol":         symbol,
            "display_name":   best["display_name"],
            "contract_label": best["contract_label"],
            "default_n":      best["default_n"],
            "ctype":          ctype,
            "is_micro":       is_micro,
            "is_inactive":    best["is_inactive"],
            "matched":        True,
        }

    # No ReadMe match → show filename with warning badge; default to 0 contracts
    fallback_ctype = FALLBACK_CTYPE.get(symbol, "MES")
    fallback_name  = stem.replace("__", " ").replace("_", " ")
    return {
        "symbol":         symbol,
        "display_name":   f"⚠️ {fallback_name}",
        "contract_label": f"0 {fallback_ctype} (not in ReadMe)",
        "default_n":      0,
        "ctype":          fallback_ctype,
        "is_micro":       fallback_ctype in MICRO_PREFIXES,
        "is_inactive":    False,
        "matched":        False,
    }


# ─────────────────────────────────────────────────────────────────────────────
# DATE PARSING
# ─────────────────────────────────────────────────────────────────────────────

def parse_ts_date(val):
    """
    Parse TradeStation Italian-locale dates from openpyxl values.

    openpyxl reads DD/MM/YYYY dates where DD ≤ 12 as datetime(YYYY, DD, MM)
    because Excel auto-converted them from Italian strings to US-format dates.
    Fix: always swap month ↔ day for datetime objects coming from openpyxl.
    Strings are parsed as DD/MM/YYYY.
    """
    if val is None:
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        ts = pd.Timestamp(val)
        try:
            ts = ts.replace(month=ts.day, day=ts.month)
        except ValueError:
            pass
        return ts
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y",
                "%m/%d/%Y %H:%M", "%m/%d/%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            continue
    try:
        return pd.Timestamp(s)
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# RAW TRADE EXTRACTION  (unchanged from v10)
# ─────────────────────────────────────────────────────────────────────────────

def extract_trades_raw(ws) -> pd.DataFrame:
    """
    Parse the 'Trades List' worksheet.

    Columns (verified across all 18 files):
      Entry row → col[0]=trade#(int), col[1]=direction, col[2]=entry_dt,
                  col[6]=n_contracts, col[7]=trade_pnl
      Exit row  → col[0]=None, col[2]=exit_dt,
                  col[6]=trade P&L ($, full position)  ← canonical
                  col[7]=cumulative P&L ($)

    P&L values already reflect TradeStation's own contract-size multiplication.
    We never multiply again.
    """
    records      = []
    header_found = False
    current      = {}

    for row in ws.iter_rows(values_only=True):
        if not header_found:
            if row[0] == "#":
                header_found = True
            continue

        if (row[0] is not None
                and isinstance(row[0], (int, float))
                and float(row[0]) == int(float(row[0]))):
            current = {
                "trade_id":    int(row[0]),
                "entry_date":  parse_ts_date(row[2]),
                "direction":   str(row[1]).strip() if row[1] else "",
                "n_contracts": int(row[6]) if row[6] is not None else 1,
            }

        elif row[0] is None and row[1] is not None and current:
            try:
                pnl     = float(row[6]) if row[6] is not None else float("nan")
                cum_pnl = float(row[7]) if row[7] is not None else float("nan")
            except (TypeError, ValueError):
                current = {}
                continue

            records.append({
                "trade_id":    current["trade_id"],
                "entry_date":  current["entry_date"],
                "exit_date":   parse_ts_date(row[2]),
                "direction":   current["direction"],
                "n_contracts": current["n_contracts"],
                "pnl":         pnl,
                "cum_pnl":     cum_pnl,
            })
            current = {}

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["pnl"]     = pd.to_numeric(df["pnl"],     errors="coerce")
    df["cum_pnl"] = pd.to_numeric(df["cum_pnl"], errors="coerce")
    return (df.dropna(subset=["exit_date", "pnl"])
              .sort_values("exit_date")
              .reset_index(drop=True))


def _snap_to_bday(dt: pd.Timestamp) -> pd.Timestamp:
    """Snap weekend dates to next Monday (handles GC weekend-bias exits)."""
    return dt if dt.dayofweek < 5 else dt + pd.offsets.BDay(1)


def build_equity_from_trades(trades: pd.DataFrame) -> pd.Series:
    """
    Build daily equity curve (cumsum of daily P&L) from raw trades.
    Weekend exit dates are snapped to the next business day.
    Forward-fills over business-day calendar for continuity.
    """
    if trades.empty:
        return pd.Series(dtype=float)
    t = trades.copy()
    t["exit_day"] = pd.to_datetime(t["exit_date"]).dt.normalize().apply(_snap_to_bday)
    daily  = t.groupby("exit_day")["pnl"].sum().sort_index()
    equity = daily.cumsum()
    if len(equity) < 2:
        return equity
    all_days = pd.date_range(equity.index.min(), equity.index.max(), freq="B")
    return equity.reindex(all_days).ffill()


# ─────────────────────────────────────────────────────────────────────────────
# LOAD ALL SYSTEMS
# ─────────────────────────────────────────────────────────────────────────────

_LOAD_VERSION = 6   # bump this integer to force a full cache clear
                    # (bumped 5→6: adds .csv strategy-file support below)

@st.cache_data(show_spinner=False, ttl=86400)
def load_all_systems(data_dir: str, comm_rates: dict | None,
                     _version: int = _LOAD_VERSION):
    """
    Cached with st.cache_data (TTL 24 h).
    Changing _LOAD_VERSION forces a fresh load.
    Sizing lives in session_state and is applied as a ratio at render time —
    so slider changes never trigger xlsx re-reads.

    Supports both .xlsx (openpyxl "Trades List" sheet) and .csv (flat report
    export — used when TradeStation's xlsx report generator is unavailable/
    crashing) per system, same convention as the Trade Matching tab: if a
    system has BOTH an .xlsx and a .csv with the same filename stem, the
    .csv is used and the .xlsx is skipped, since .csv is the newer/likely
    more current export path for this account.
    """
    path          = Path(data_dir)
    systems       = {}
    load_warnings = []

    xlsx_files = {f.stem: f for f in sorted(path.glob("*.xlsx"))}
    csv_files  = {f.stem: f for f in sorted(path.glob("*.csv"))}
    stems      = sorted(set(xlsx_files) | set(csv_files))

    for stem in stems:
        f = csv_files.get(stem) or xlsx_files[stem]
        from_csv = stem in csv_files
        try:
            b5_val      = None
            data_source = "trades"
            trades      = pd.DataFrame()

            if from_csv:
                if tsm is None:
                    load_warnings.append(
                        f"❌ {f.name}: found a .csv strategy file but "
                        f"tradestation_matcher.py couldn't be imported "
                        f"({_TSM_IMPORT_ERROR}) — skipped.")
                    continue
                trades = tsm.extract_theoretical_trades_csv(str(f))
                b5_val = tsm.extract_csv_total_net_profit(str(f))
                if trades.empty:
                    load_warnings.append(f"⚠️ {f.name}: 'Trades List' section is empty.")
                    data_source = "empty"
            else:
                wb = openpyxl.load_workbook(str(f), data_only=True)

                if "Performance Summary" in wb.sheetnames:
                    try:
                        b5_val = float(wb["Performance Summary"]["B5"].value)
                    except (TypeError, ValueError):
                        pass

                if "Trades List" in wb.sheetnames:
                    trades = extract_trades_raw(wb["Trades List"])
                    if trades.empty:
                        load_warnings.append(f"⚠️ {f.name}: 'Trades List' is empty.")
                        data_source = "empty"
                else:
                    load_warnings.append(
                        f"⚠️ {f.name}: no 'Trades List' tab — using summary only.")
                    data_source = "summary_fallback"

            equity = build_equity_from_trades(trades)

            b5_ok = False
            if not trades.empty and b5_val is not None:
                last_cum = trades["cum_pnl"].iloc[-1]
                b5_ok = abs(float(b5_val) - float(last_cum)) < 0.5
                if not b5_ok:
                    load_warnings.append(
                        f"⚠️ {stem}: last cum_pnl ≠ B5 ({last_cum} vs {b5_val})")

            alloc    = get_alloc(stem)
            ctype    = alloc["ctype"]
            is_micro = ctype in MICRO_PREFIXES
            # Determine per-contract round-trip commission: prefer explicit
            # mapping passed in `comm_rates`, otherwise fall back to the
            # legacy micro/mini constants.
            comm_rt = None
            if isinstance(comm_rates, dict):
                # Exact token match first
                comm_rt = comm_rates.get(ctype)
                # If not found and it's a micro prefix like 'MES', allow
                # matching the short token (e.g., 'MES' already)
            if comm_rt is None:
                comm_rt = COMMISSION_PER_MICRO if is_micro else COMMISSION_PER_MINI

            systems[stem] = {
                # ── Identity ──────────────────────────────────────────────
                "stem":           stem,
                "display_name":   alloc["display_name"],   # from ReadMe
                "symbol":         alloc["symbol"],
                "contract_type":  ctype,
                "contract_label": alloc["contract_label"],
                "default_n":      alloc["default_n"],       # ReadMe first qty
                "is_micro":       is_micro,
                "is_inactive":    alloc["is_inactive"],
                "matched":        alloc["matched"],
                # ── Trade data ────────────────────────────────────────────
                "trades":         trades,
                "equity":         equity,               # raw, 1× factor
                # ── Costs ─────────────────────────────────────────────────
                "comm_per_trade": comm_rt,
                # ── Validation ────────────────────────────────────────────
                "b5_total":       b5_val,
                "b5_match":       b5_ok,
                "data_source":    data_source,
                "from_csv":       from_csv,
            }
        except Exception as exc:
            load_warnings.append(f"❌ Could not load {f.name}: {exc}")

    # ── Sort by ReadMe.txt line order (sort_order baked into each entry) ──────
    def _sort_key(stem):
        norm_stem  = stem.lower().replace("__", " ").replace("_", " ")
        stem_words = set(norm_stem.split())
        best_order = 9999
        best_len   = 0
        for norm_name, entry in _README_ENTRIES.items():
            kw = set(norm_name.split())
            if kw.issubset(stem_words) and len(kw) > best_len:
                best_len   = len(kw)
                best_order = entry.get("sort_order", 9999)
        return best_order

    systems = dict(sorted(systems.items(), key=lambda kv: _sort_key(kv[0])))

    return systems, load_warnings


# ─────────────────────────────────────────────────────────────────────────────
# SIZING HELPERS  (v11 — session_state sizing multiplier)
# ─────────────────────────────────────────────────────────────────────────────

def get_current_n(stem: str, systems: dict) -> int:
    """Return current user-set contract count from session_state, or default."""
    return st.session_state.get("sizing", {}).get(
        stem, systems[stem]["default_n"])


def get_sizing_ratio(stem: str, systems: dict) -> float:
    """
    Return scaling ratio = current_n / default_n.

    This is the ONLY place a multiplier is applied to raw P&L.
    It scales the displayed equity curve to reflect a different
    number of contracts than what TradeStation ran.

    If default_n == 0, return 0 to avoid division by zero.
    """
    default = systems[stem]["default_n"]
    if default == 0:
        return 0.0
    return get_current_n(stem, systems) / default


# ─────────────────────────────────────────────────────────────────────────────
# METRICS
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False, ttl=3600)
def build_net_equity(trades: pd.DataFrame, comm_per_trade: float,
                     scale: float = 1.0) -> pd.Series:
    """
    Build daily equity after subtracting commission.
    `comm_per_trade` is a $/CONTRACT round-trip rate (not a flat per-trade
    cost) -- multiplied by each trade's actual n_contracts below. Confirmed
    against Andrea's real statement: commissions are charged per contract,
    so a flat per-trade deduction badly underestimated cost on multi-lot
    trades (and overestimated it on the single-lot full-size systems).
    'scale' is the sizing ratio (current_n / default_n), applied
    multiplicatively to the net P&L series on top of that.
    Cached: re-computes only when trades, comm or scale actually change.
    """
    if trades.empty:
        return pd.Series(dtype=float)
    t = trades.copy()
    t["exit_day"] = pd.to_datetime(t["exit_date"]).dt.normalize().apply(_snap_to_bday)
    n_ct = t["n_contracts"] if "n_contracts" in t.columns else 1
    t["pnl_net"]  = (t["pnl"] - comm_per_trade * n_ct) * scale
    daily = t.groupby("exit_day")["pnl_net"].sum().sort_index()
    eq    = daily.cumsum()
    if len(eq) < 2:
        return eq
    all_days = pd.date_range(eq.index.min(), eq.index.max(), freq="B")
    return eq.reindex(all_days).ffill()


def compute_metrics(trades: pd.DataFrame, comm_per_trade: float,
                    scale: float = 1.0) -> dict:
    """Compute system performance metrics using daily net P&L.

    P&L Sharpe is deliberately based on daily P&L rather than a return series,
    because futures systems do not have a single meaningful invested-capital
    denominator. The equity curve starts at zero so the first trading day's P&L
    is included in Net Profit and the drawdown calculation.
    """
    if trades.empty:
        return {}

    eq_net = build_net_equity(trades, comm_per_trade, scale)
    if eq_net.empty or len(eq_net) < 5:
        return {}

    # The first observation is already the first day's P&L. Use it directly
    # rather than diff(), which would silently discard that first day.
    daily = eq_net.diff().dropna()
    first_day_pnl = eq_net.iloc[0]
    daily_pnl = pd.concat([pd.Series([first_day_pnl], index=[eq_net.index[0]]), daily]).sort_index()
    run_max = eq_net.cummax()
    max_dd = (eq_net - run_max).min()
    n_yr = max((eq_net.index[-1] - eq_net.index[0]).days / 365.25, 0.01)
    net = float(eq_net.iloc[-1])
    ann_pnl = net / n_yr
    ann_vol = daily_pnl.std() * np.sqrt(252)
    pnl_sharpe = (daily_pnl.mean() / daily_pnl.std() * np.sqrt(252)) if daily_pnl.std() > 0 else 0.0
    ann_pnl_dd = ann_pnl / abs(max_dd) if max_dd < 0 else 0.0

    n_ct_t = trades["n_contracts"] if "n_contracts" in trades.columns else 1
    pnl_net_t = (trades["pnl"] - comm_per_trade * n_ct_t) * scale
    wins_sum = pnl_net_t[pnl_net_t > 0].sum()
    loss_sum = abs(pnl_net_t[pnl_net_t < 0].sum())
    pf = wins_sum / loss_sum if loss_sum > 0 else float("inf")
    win_rate = (pnl_net_t > 0).mean()

    return {
        "Net Profit ($)":      round(net, 0),
        "Ann. P&L ($)":        round(ann_pnl, 0),
        "Max Drawdown ($)":    round(max_dd, 0),
        "P&L Sharpe":          round(pnl_sharpe, 2),
        "Ann. P&L / Max DD":   round(ann_pnl_dd, 2),
        "Profit Factor":       round(pf, 2),
        "Win Rate":            round(win_rate, 3),
        "Ann. Volatility ($)": round(ann_vol, 0),
        "# Trades":            len(trades),
        "Total Comm ($)":      round(float((n_ct_t * comm_per_trade).sum()) * scale, 0),
    }


def get_net_equity_trimmed(si: dict, lookback_years: float,
                            scale: float = 1.0) -> pd.Series:
    """Return lookback-clipped, rebased-to-zero net equity for a system."""
    eq = build_net_equity(si["trades"], si["comm_per_trade"], scale)
    if eq.empty:
        return pd.Series(dtype=float)
    cutoff = eq.index.max() - pd.Timedelta(days=int(lookback_years * 365))
    eq_t   = eq[eq.index >= cutoff]
    if eq_t.empty:
        return pd.Series(dtype=float)
    return eq_t - eq_t.iloc[0]


def compute_portfolio_metrics(port_eq: pd.Series) -> dict:
    """Portfolio metrics using daily portfolio P&L and a zero-based equity curve."""
    if port_eq.empty or len(port_eq) < 5:
        return {}
    daily = port_eq.diff().dropna()
    first_day_pnl = port_eq.iloc[0]
    daily_pnl = pd.concat([pd.Series([first_day_pnl], index=[port_eq.index[0]]), daily]).sort_index()
    net = float(port_eq.iloc[-1])
    n_yr = max((port_eq.index[-1] - port_eq.index[0]).days / 365.25, 0.01)
    ann_pnl = net / n_yr
    ann_vol = daily_pnl.std() * np.sqrt(252)
    pnl_sharpe = (daily_pnl.mean() / daily_pnl.std() * np.sqrt(252)) if daily_pnl.std() > 0 else 0.0
    dd = (port_eq - port_eq.cummax()).min()
    ann_pnl_dd = ann_pnl / abs(dd) if dd < 0 else 0.0
    return {
        "Net Profit ($)":      round(net, 0),
        "Ann. P&L ($)":        round(ann_pnl, 0),
        "Max Drawdown ($)":    round(dd, 0),
        "Ann. Volatility ($)": round(ann_vol, 0),
        "P&L Sharpe":          round(pnl_sharpe, 2),
        "Ann. P&L / Max DD":   round(ann_pnl_dd, 2),
    }


def combine_equity_curves(curves: dict, lookback_years: float) -> pd.DataFrame:
    if not curves:
        return pd.DataFrame()
    df     = pd.concat(curves.values(), axis=1, keys=curves.keys()).sort_index().ffill()
    cutoff = df.index.max() - pd.Timedelta(days=int(lookback_years * 365))
    return df[df.index >= cutoff].ffill().fillna(0)


def resolve_portfolio_date_range(min_date: pd.Timestamp,
                                 max_date: pd.Timestamp,
                                 preset: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    today = pd.Timestamp.today().normalize()
    if preset == "MTD":
        start = today.replace(day=1)
        end = min(today, max_date)
    elif preset == "YTD":
        start = pd.Timestamp(today.year, 1, 1)
        end = min(today, max_date)
    elif preset == "Past week":
        start = today - pd.Timedelta(days=7)
        end = min(today, max_date)
    elif preset == "Past 1 month":
        start = today - pd.Timedelta(days=30)
        end = min(today, max_date)
    elif preset == "Past 3 months":
        start = today - pd.Timedelta(days=90)
        end = min(today, max_date)
    elif preset == "Past 6 months":
        start = today - pd.Timedelta(days=180)
        end = min(today, max_date)
    elif preset == "Past year":
        start = today - pd.Timedelta(days=365)
        end = min(today, max_date)
    elif preset == "Past 2 years":
        start = today - pd.Timedelta(days=730)
        end = min(today, max_date)
    elif preset == "2025":
        start = pd.Timestamp("2025-01-01")
        end = min(pd.Timestamp("2025-12-31"), max_date)
    elif preset == "2024":
        start = pd.Timestamp("2024-01-01")
        end = min(pd.Timestamp("2024-12-31"), max_date)
    else:
        start, end = min_date, max_date

    start = max(start, min_date)
    end = min(end, max_date)
    return start, end


def filter_equity_by_date_range(eq_df: pd.DataFrame,
                                start_date: pd.Timestamp,
                                end_date: pd.Timestamp) -> pd.DataFrame:
    if eq_df.empty:
        return eq_df
    dates = pd.date_range(start=start_date, end=end_date, freq="B")
    return eq_df.reindex(dates).ffill().fillna(0)


# ─────────────────────────────────────────────────────────────────────────────
# ROLLING HEALTH
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False, ttl=3600)
def compute_rolling_health(trades: pd.DataFrame, comm_per_trade: float,
                            window_days: int = 63) -> pd.DataFrame:
    eq_net = build_net_equity(trades, comm_per_trade)
    if eq_net.empty or len(eq_net) < window_days + 10:
        return pd.DataFrame()
    daily = eq_net.diff().dropna()
    if len(daily) < window_days:
        return pd.DataFrame()

    roll_sharpe = (daily.rolling(window_days).mean()
                   / daily.rolling(window_days).std()
                   * np.sqrt(252)).dropna()
    wins_r = (daily > 0).astype(float).rolling(window_days).mean().dropna()
    gains  = daily.clip(lower=0).rolling(window_days).sum()
    losses = (-daily).clip(lower=0).rolling(window_days).sum()
    roll_pf  = (gains / losses.replace(0, np.nan)).dropna()
    roll_cum = daily.rolling(window_days).sum().dropna()

    common = roll_sharpe.index.intersection(wins_r.index)
    if common.empty:
        return pd.DataFrame()

    return pd.DataFrame({
        "Sharpe":        roll_sharpe.reindex(common),
        "Win Rate":      wins_r.reindex(common),
        "Profit Factor": roll_pf.reindex(common) if not roll_pf.empty else np.nan,
        "Period P&L":    roll_cum.reindex(common),
    }).dropna()


def health_traffic_light(sharpe: float, pf: float, win_rate: float) -> str:
    score  = 0
    score += 2 if sharpe   > 0.8  else (1 if sharpe   > 0.2  else (-1 if sharpe   < -0.3 else 0))
    score += 2 if pf       > 1.3  else (1 if pf       > 1.0  else (-1 if pf       < 0.8  else 0))
    score += 1 if win_rate > 0.55 else (-1 if win_rate < 0.4 else 0)
    return "🟢" if score >= 3 else ("🟡" if score >= 1 else "🔴")


# ─────────────────────────────────────────────────────────────────────────────
# PORTFOLIO HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _safe_date_range_state(key: str, min_date: pd.Timestamp, max_date: pd.Timestamp):
    """Clamp persisted date-input state to the currently available data range."""
    lo, hi = pd.Timestamp(min_date).date(), pd.Timestamp(max_date).date()
    raw = st.session_state.get(key)
    if not raw:
        fixed = [lo, hi]
    else:
        vals = list(raw) if isinstance(raw, (list, tuple)) else [raw]
        vals = vals[:2]
        while len(vals) < 2:
            vals.append(vals[-1] if vals else lo)
        try:
            d1, d2 = pd.Timestamp(vals[0]).date(), pd.Timestamp(vals[1]).date()
        except Exception as exc:
            st.warning(f"⚠️ Could not parse saved date range '{key}': {exc}. Resetting it.")
            d1, d2 = lo, hi
        d1, d2 = min(max(d1, lo), hi), min(max(d2, lo), hi)
        if d1 > d2: d1, d2 = d2, d1
        fixed = [d1, d2]
    st.session_state[key] = fixed
    return fixed

def _effective_independent_bets(corr: pd.DataFrame) -> float:
    if corr.empty or len(corr) < 2: return float(len(corr))
    upper = corr.where(np.triu(np.ones(corr.shape), k=1).astype(bool)).stack()
    if upper.empty: return float(len(corr))
    rho = float(upper.mean()); n = len(corr); denom = 1.0 + (n - 1) * rho
    if denom <= 0: return float(n)
    return float(np.clip(n / denom, 1.0, float(n)))

def _resampled_corr(daily_pnl: pd.DataFrame, freq: str, min_periods: int = 12) -> pd.DataFrame:
    if daily_pnl.empty: return pd.DataFrame()
    if freq == "W": grouped = daily_pnl.resample("W-FRI").sum(min_count=1)
    elif freq == "M": grouped = daily_pnl.resample("ME").sum(min_count=1)
    else: return daily_pnl.corr(min_periods=min_periods)
    return grouped.corr(min_periods=min_periods)

def _build_periodic_trade_pnl(systems: dict, stems: list, lookback_years: float, sizing_map: dict, freq: str) -> pd.DataFrame:
    """Weekly/monthly P&L from actual trade exits; no-trade periods remain NaN."""
    series = {}
    for stem in stems:
        si = systems.get(stem)
        if si is None or si.get("default_n", 0) <= 0 or si["trades"].empty: continue
        t = si["trades"].copy()
        ratio = float(sizing_map.get(stem, si["default_n"])) / float(si["default_n"])
        t["exit_day"] = pd.to_datetime(t["exit_date"]).dt.normalize().apply(_snap_to_bday)
        t["pnl_net"] = (pd.to_numeric(t["pnl"], errors="coerce") - float(si["comm_per_trade"]) * pd.to_numeric(t.get("n_contracts", 1), errors="coerce").fillna(1)) * ratio
        t = t.dropna(subset=["exit_day", "pnl_net"])
        if t.empty: continue
        if freq == "W": g = t.groupby(pd.Grouper(key="exit_day", freq="W-FRI"))["pnl_net"].sum()
        elif freq == "M": g = t.groupby(pd.Grouper(key="exit_day", freq="ME"))["pnl_net"].sum()
        else: continue
        cutoff = g.index.max() - pd.Timedelta(days=int(lookback_years*365))
        series[stem] = g[g.index >= cutoff].rename(stem)
    return pd.concat(series.values(), axis=1).sort_index() if series else pd.DataFrame()

def portfolio_risk(systems: dict, stems: list, sizing_map: dict, lookback_years: float,
                   window_days: int | None = None, date_start: pd.Timestamp | None = None,
                   date_end: pd.Timestamp | None = None, min_overlap_days: int = 252) -> dict:
    """Canonical portfolio risk: complete-case, constant-composition daily P&L."""
    curves = {}
    for stem in stems:
        si = systems.get(stem)
        if si is None or si.get("default_n", 0) <= 0: continue
        cur_n = float(sizing_map.get(stem, si["default_n"]))
        if cur_n <= 0: continue
        eq = get_net_equity_trimmed(si, lookback_years, scale=cur_n / float(si["default_n"]))
        if not eq.empty: curves[stem] = eq
    if len(curves) < 2: return {}
    daily = build_correlation_daily_pnl(curves, lookback_years).dropna(how="any")
    if date_start is not None: daily = daily[daily.index >= pd.Timestamp(date_start)]
    if date_end is not None: daily = daily[daily.index <= pd.Timestamp(date_end)]
    if window_days is not None: daily = daily.tail(window_days)
    if len(daily) < 20: return {"daily_pnl": daily}
    port_daily = daily.sum(axis=1); daily_vol = float(port_daily.std())
    cov_complete = daily.cov(); corr_complete = daily.corr(); vols = daily.std()
    sum_vol = float(vols.sum()); div_ratio = sum_vol / daily_vol if daily_vol > 0 else np.nan
    sigma = cov_complete.values; w = np.ones(len(cov_complete))
    component = (sigma @ w) / daily_vol if daily_vol > 0 else np.zeros(len(w))
    rc = pd.Series(component, index=cov_complete.index); rc_pct = rc / daily_vol if daily_vol > 0 else pd.Series(0.0, index=cov_complete.index)
    pair_cov = daily.cov(min_periods=min_overlap_days); pair_corr = daily.corr(min_periods=min_overlap_days)
    overlap = daily.notna().astype(int).T.dot(daily.notna().astype(int))
    pair_psd = _nearest_psd(pair_cov) if not pair_cov.empty else pair_cov
    raw_fill = pair_cov.fillna(0) if not pair_cov.empty else pair_cov
    psd_norm = float(np.linalg.norm(pair_psd.values - raw_fill.values, ord="fro")) if not pair_cov.empty else 0.0
    lw = LedoitWolf().fit(daily.values)
    cov_shrunk = pd.DataFrame(lw.covariance_, index=daily.columns, columns=daily.columns)
    
    if not overlap.empty and len(overlap) > 1:
        _ov = overlap.values.astype(float)
        _mask = ~np.eye(len(overlap), dtype=bool)
        min_pairwise = int(np.nanmin(_ov[_mask])) if np.isfinite(_ov[_mask]).any() else 0
    else:
        min_pairwise = 0
    offdiag = pair_corr.where(np.triu(np.ones(pair_corr.shape), k=1).astype(bool)).stack()
    return {"daily_pnl": daily, "portfolio_daily_vol": daily_vol, "portfolio_ann_vol": daily_vol*np.sqrt(252),
            "vols": vols, "covariance": cov_complete, "covariance_pairwise": pair_cov, "covariance_shrunk": cov_shrunk,
            "shrinkage_intensity": float(lw.shrinkage_), "corr": corr_complete, "corr_pairwise": pair_corr, "overlap": overlap,
            "min_pairwise_overlap": min_pairwise, "psd_correction_frobenius": psd_norm, "sum_individual_vol": sum_vol,
            "diversification_ratio": div_ratio, "avg_pairwise_corr": float(offdiag.mean()) if not offdiag.empty else np.nan,
            "risk_contribution": rc, "risk_pct": rc_pct, "observations": len(daily), "start": daily.index.min(), "end": daily.index.max()}


def build_portfolio_equity(systems: dict, stems: list,
                            lookback_years: float,
                            sizing_ratios: dict | None = None) -> pd.Series:
    """
    Sum per-system net-equity curves.
    sizing_ratios = {stem: float}, defaults to 1.0 (ReadMe default).
    """
    curves = {}
    for stem in stems:
        si    = systems.get(stem)
        if si is None:
            continue
        ratio = (sizing_ratios or {}).get(stem, 1.0)
        if ratio == 0.0:
            continue   # system switched off — skip entirely
        eq    = get_net_equity_trimmed(si, lookback_years, scale=ratio)
        if not eq.empty:
            curves[stem] = eq
    if not curves:
        return pd.Series(dtype=float)
    df = pd.concat(curves.values(), axis=1,
                   keys=curves.keys()).sort_index().ffill().fillna(0)
    return df.sum(axis=1)


# ─────────────────────────────────────────────────────────────────────────────
# RISK PARITY
# ─────────────────────────────────────────────────────────────────────────────

def compute_risk_parity_sizing(systems: dict, stems: list, lookback_years: float,
                                target_daily_risk_usd: float = 500.0,
                                max_contracts: int = 20) -> tuple:
    sizing, detail = {}, {}
    for stem in stems:
        si = systems.get(stem)
        if si is None or si["equity"].empty:
            sizing[stem] = 1
            continue
        eq     = get_net_equity_trimmed(si, lookback_years, scale=1.0)
        cutoff = eq.index.max() - pd.Timedelta(days=int(lookback_years * 365))
        eq_t   = eq[eq.index >= cutoff]
        # The stored equity curve is at the system's ReadMe/default sizing,
        # so its daily volatility is NOT one-contract volatility. Convert it
        # to a per-contract risk estimate before solving for N.
        default_n = si["default_n"]
        vol_default = eq_t.diff().dropna().std()
        if vol_default <= 0 or np.isnan(vol_default) or default_n <= 0:
            sizing[stem] = default_n if default_n > 0 else 1
            detail[stem] = {
                "vol_1ct": 0, "vol_default": float(vol_default) if np.isfinite(vol_default) else 0,
                "raw_n": default_n, "final_n": sizing[stem]
            }
            continue
        vol_1ct = vol_default / default_n
        raw_n   = target_daily_risk_usd / vol_1ct
        final_n = int(np.clip(round(raw_n), 0, max_contracts))
        sizing[stem] = final_n
        detail[stem] = {
            "vol_1ct": round(float(vol_1ct), 2),
            "vol_default": round(float(vol_default), 2),
            "raw_n": round(float(raw_n), 2),
            "final_n": final_n
        }
    return sizing, detail


def compute_recent_portfolio_daily_risk(systems: dict, stems: list, lookback_years: float,
                                        sizing_ratios: dict, recent_days: int = 126,
                                        date_start: pd.Timestamp | None = None, date_end: pd.Timestamp | None = None) -> tuple:
    sizing_map = {stem: systems[stem]["default_n"] * sizing_ratios.get(stem, 1.0) for stem in stems if stem in systems}
    r = portfolio_risk(systems, stems, sizing_map, lookback_years, window_days=recent_days, date_start=date_start, date_end=date_end)
    d = r.get("daily_pnl", pd.DataFrame()) if r else pd.DataFrame()
    if not r or "portfolio_daily_vol" not in r: return 0.0, len(d), d.index.min() if len(d) else None, d.index.max() if len(d) else None
    return r["portfolio_daily_vol"], r["observations"], r["start"], r["end"]


# ─────────────────────────────────────────────────────────────────────────────
# DRAWDOWN DECOMPOSITION
# ─────────────────────────────────────────────────────────────────────────────

def decompose_drawdown(eq_df: pd.DataFrame, port_eq: pd.Series, top_n: int = 3) -> list:
    if port_eq.empty or eq_df.empty:
        return []
    peak_idx, peak_val = 0, port_eq.iloc[0]
    episodes = []
    for i in range(1, len(port_eq)):
        val = port_eq.iloc[i]
        if val > peak_val:
            peak_val, peak_idx = val, i
        else:
            dd = val - peak_val
            if not episodes or episodes[-1]["peak_idx"] != peak_idx:
                episodes.append({"peak_idx": peak_idx, "trough_idx": i,
                                  "dd_abs": dd,
                                  "peak_date":   port_eq.index[peak_idx],
                                  "trough_date": port_eq.index[i]})
            elif dd < episodes[-1]["dd_abs"]:
                episodes[-1].update({"trough_idx": i, "dd_abs": dd,
                                      "trough_date": port_eq.index[i]})
    episodes = sorted(episodes, key=lambda x: x["dd_abs"])[:top_n]
    for ep in episodes:
        pd_, td = ep["peak_date"], ep["trough_date"]
        contribs = {}
        for col in eq_df.columns:
            s = eq_df[col]
            try:
                s_p = float(s.asof(pd_)) if pd_ >= s.index[0] else 0.0
                s_t = float(s.asof(td))  if td  >= s.index[0] else 0.0
                contribs[col] = s_t - s_p
            except Exception as exc:
                st.warning(f"⚠️ Drawdown contribution calculation failed for {col}: {exc}")
                contribs[col] = 0.0
        total_loss = abs(ep["dd_abs"])
        ep["contributions"] = contribs
        ep["blame_pct"]     = {k: abs(v) / total_loss * 100
                               for k, v in contribs.items()
                               if v < 0 and total_loss > 0}
    return episodes


# ─────────────────────────────────────────────────────────────────────────────
# CORRELATION / CLUSTERING — v24 quantitative audit
# ─────────────────────────────────────────────────────────────────────────────

def build_correlation_daily_pnl(curves: dict, lookback_years: float,
                                min_overlap_days: int = 63) -> pd.DataFrame:
    """Build daily P&L series for correlation without inventing zero P&L.

    Each system is allowed to have its own history. Missing observations remain
    NaN, so pairwise correlation/covariance uses only dates where both systems
    actually have data. This avoids treating a pre-launch strategy as if it
    existed and made exactly $0 every day.
    """
    if not curves:
        return pd.DataFrame()
    daily = {}
    max_date = max((s.index.max() for s in curves.values() if not s.empty),
                   default=None)
    if max_date is None:
        return pd.DataFrame()
    cutoff = max_date - pd.Timedelta(days=int(lookback_years * 365))
    for stem, eq in curves.items():
        if eq.empty:
            continue
        e = eq.sort_index()
        e = e[e.index >= cutoff]
        if len(e) < 2:
            continue
        d = e.diff()
        d = d[d.index >= cutoff].rename(stem)
        daily[stem] = d
    if not daily:
        return pd.DataFrame()
    return pd.concat(daily.values(), axis=1).sort_index()


def _pairwise_corr_cov(daily_pnl: pd.DataFrame, min_overlap_days: int = 63) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Return pairwise correlation, covariance and observation-count matrices."""
    if daily_pnl.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    corr = daily_pnl.corr(min_periods=min_overlap_days)
    cov = daily_pnl.cov(min_periods=min_overlap_days)
    counts = daily_pnl.notna().astype(int).T.dot(daily_pnl.notna().astype(int))
    return corr, cov, counts


def _nearest_psd(matrix: pd.DataFrame) -> pd.DataFrame:
    """Project a symmetric covariance matrix to positive semi-definite form."""
    if matrix.empty:
        return matrix
    a = matrix.astype(float).values
    a = np.nan_to_num((a + a.T) / 2.0, nan=0.0, posinf=0.0, neginf=0.0)
    diag = np.diag(matrix.astype(float).values)
    for i, v in enumerate(diag):
        if np.isfinite(v) and v > 0:
            a[i, i] = v
    vals, vecs = np.linalg.eigh(a)
    vals = np.clip(vals, 1e-10, None)
    psd = (vecs * vals) @ vecs.T
    psd = (psd + psd.T) / 2.0
    return pd.DataFrame(psd, index=matrix.index, columns=matrix.columns)


def correlation_risk_stats(daily_pnl: pd.DataFrame, corr_mat: pd.DataFrame,
                           cov_mat: pd.DataFrame) -> dict:
    """Portfolio volatility and risk contributions at the current sizing."""
    if daily_pnl.empty:
        return {}
    cols = [c for c in daily_pnl.columns if c in corr_mat.index and c in cov_mat.index]
    if len(cols) < 2:
        return {}
    d = daily_pnl[cols]
    vols = d.std().replace([np.inf, -np.inf], np.nan).fillna(0.0)
    cov = cov_mat.loc[cols, cols].copy()
    # Pairwise covariance can be non-PSD when histories differ. Missing
    # cross-covariances are treated as zero for the risk calculation, then the
    # matrix is projected to PSD so portfolio volatility is mathematically valid.
    for c in cols:
        if not np.isfinite(cov.loc[c, c]) or cov.loc[c, c] <= 0:
            cov.loc[c, c] = vols[c] ** 2
    cov = cov.fillna(0.0)
    cov_psd = _nearest_psd(cov)
    w = np.ones(len(cols))
    sigma = cov_psd.values
    port_var = float(w @ sigma @ w)
    port_vol = float(np.sqrt(max(port_var, 0.0)))
    marginal = (sigma @ w) / port_vol if port_vol > 0 else np.zeros(len(cols))
    component = w * marginal
    rc_pct = component / port_vol if port_vol > 0 else np.zeros(len(cols))
    sum_vol = float(vols.sum())
    div_ratio = sum_vol / port_vol if port_vol > 0 else np.nan
    return {
        "vols": vols,
        "cov_psd": cov_psd,
        "portfolio_vol": port_vol,
        "sum_individual_vol": sum_vol,
        "diversification_ratio": div_ratio,
        "marginal": pd.Series(marginal, index=cols),
        "component": pd.Series(component, index=cols),
        "rc_pct": pd.Series(rc_pct, index=cols),
    }


def compute_correlation_aware_rp(unit_daily_pnl: pd.DataFrame, account_balance: float,
                                 target_annual_risk_pct: float, max_contracts: int = 20) -> dict:
    """Solve equal-risk-contribution sizing using unit-contract covariance."""
    if unit_daily_pnl.empty or unit_daily_pnl.shape[1] < 2:
        return {}
    cols = unit_daily_pnl.columns.tolist()
    complete = unit_daily_pnl[cols].dropna(how="any")
    if len(complete) < 252:
        return {}
    vols = complete.std().reindex(cols).fillna(0.0)
    lw = LedoitWolf().fit(complete.values)
    cov = pd.DataFrame(lw.covariance_, index=cols, columns=cols)
    S = cov.values
    n = len(cols)

    def rc_pct(w):
        pv = float(w @ S @ w)
        if pv <= 0:
            return np.zeros(n)
        return w * (S @ w) / pv

    target_rc = 1.0 / n
    def objective(w):
        pct = rc_pct(w)
        penalty = np.sum(np.minimum(pct, 0.0) ** 2) * 100.0
        return float(np.sum((pct - target_rc) ** 2) + penalty)

    x0 = np.repeat(1.0 / n, n)
    res = minimize(objective, x0, method="SLSQP", bounds=[(1e-7, 1.0)] * n,
                   constraints=[{"type": "eq", "fun": lambda w: np.sum(w) - 1.0}],
                   options={"maxiter": 2000, "ftol": 1e-12})
    weights = np.asarray(res.x if res.success else x0, dtype=float)
    weights = np.clip(weights, 1e-9, None); weights /= weights.sum()
    target_daily = max(float(account_balance) * float(target_annual_risk_pct) / 100.0 / np.sqrt(252), 0.0)
    base_vol = float(np.sqrt(max(weights @ S @ weights, 0.0)))
    continuous = weights * (target_daily / base_vol if base_vol > 0 else 0.0)
    integer = np.clip(np.rint(continuous).astype(int), 0, max_contracts)
    if target_daily > 0 and integer.sum() == 0:
        integer[np.argmax(continuous)] = 1

    def stats(q):
        pv = float(q @ S @ q); sig = np.sqrt(max(pv, 0.0))
        comp = q * (S @ q) / sig if sig > 0 else np.zeros(n)
        return sig, comp, comp / sig if sig > 0 else np.zeros(n)
    def score(q):
        sig, _, pct = stats(q)
        if sig <= 0: return 1e9
        active = q > 0; desired = 1.0 / active.sum()
        return float(np.sum((pct[active] - desired) ** 2) + 0.05 * ((sig - target_daily) / max(target_daily, 1.0)) ** 2)
    current = integer.copy()
    for _ in range(6 * n):
        best, best_score = current.copy(), score(current)
        for i in range(n):
            for delta in (-1, 1):
                cand = current.copy(); cand[i] = int(np.clip(cand[i] + delta, 0, max_contracts))
                sc = score(cand)
                if sc < best_score - 1e-12: best, best_score = cand, sc
        if np.array_equal(best, current): break
        current = best
    final_vol, final_comp, final_pct = stats(current)
    return {"cov": cov, "unit_vol": vols, "continuous": pd.Series(continuous, index=cols),
            "integer": pd.Series(current, index=cols, dtype=int), "target_daily": target_daily,
            "portfolio_vol": final_vol, "risk_contribution": pd.Series(final_comp, index=cols),
            "risk_pct": pd.Series(final_pct, index=cols), "continuous_success": bool(res.success), "shrinkage_intensity": float(lw.shrinkage_), "observations": len(complete), "min_overlap_days": len(complete)}


def extract_correlation_clusters(corr_mat: pd.DataFrame, threshold: float = 0.70) -> dict:
    """Group systems using average-linkage correlation distance.

    Missing pairwise correlations are not treated as zero correlation. Systems
    without a complete correlation row become singleton clusters.
    """
    if corr_mat.empty:
        return {}
    clean = corr_mat.copy().clip(-1, 1)
    clean = clean.dropna(axis=0, how="all").dropna(axis=1, how="all")
    if len(clean) < 2:
        return {c: i + 1 for i, c in enumerate(clean.columns)}
    complete = clean.dropna(axis=0, how="any").dropna(axis=1, how="any")
    labels_out = {}
    next_label = 1
    if len(complete) >= 2:
        vals = complete.values.astype(float)
        vals = (vals + vals.T) / 2.0
        np.fill_diagonal(vals, 1.0)
        dist = np.clip(1.0 - vals, 0.0, 2.0)
        try:
            z = linkage(squareform(dist, checks=False), method="average")
            labels = fcluster(z, t=max(0.05, 1 - threshold), criterion="distance")
            labels_out.update(dict(zip(complete.columns, labels)))
            next_label = max(labels_out.values(), default=0) + 1
        except Exception as exc:
            st.warning(f"⚠️ Correlation clustering failed; affected systems will be singletons: {exc}")
    for c in clean.columns:
        if c not in labels_out:
            labels_out[c] = next_label
            next_label += 1
    return labels_out


# ─────────────────────────────────────────────────────────────────────────────
# CORRELATION / CLUSTERING
# ─────────────────────────────────────────────────────────────────────────────

def cluster_correlation_matrix(corr_matrix: pd.DataFrame) -> tuple:
    # Drop any rows/cols that are all-NaN (systems with no data / n=0)
    corr_matrix = corr_matrix.dropna(axis=0, how="all").dropna(axis=1, how="all")
    n = len(corr_matrix)
    if n < 3:
        return corr_matrix, list(range(n))
    # Ensure symmetry: average with transpose to fix floating-point asymmetry
    vals = corr_matrix.values.clip(-1, 1)
    vals = (vals + vals.T) / 2
    np.fill_diagonal(vals, 1.0)
    dist = np.maximum(1 - vals, 0)
    # squareform expects a condensed vector; pass the full symmetric matrix
    try:
        from scipy.spatial.distance import squareform as sf
        dist_condensed = sf(dist, checks=False)
        order = leaves_list(linkage(dist_condensed, method="average"))
    except Exception as exc:
        st.warning(f"⚠️ Correlation heatmap ordering failed; using original order: {exc}")
        order = list(range(n))
    cols = corr_matrix.columns[order]
    return corr_matrix.loc[cols, cols], order


# ─────────────────────────────────────────────────────────────────────────────
# PORTFOLIO OPTIMISATION HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def recommend_portfolios(eq_df: pd.DataFrame, systems: dict) -> list:
    daily = eq_df.diff().dropna()
    if daily.empty or daily.shape[1] < 2:
        return []
    means   = daily.mean()
    stds    = daily.std().replace(0, np.nan)
    sharpes = (means / stds).fillna(0)
    corr    = daily.corr()
    recs    = []

    top = sharpes.nlargest(min(8, len(sharpes))).index.tolist()
    sel = []
    for s in top:
        if not sel or max(corr.loc[s, x] for x in sel if x in corr.columns) < 0.85:
            sel.append(s)
    recs.append({"name": "🏆 Max Sharpe",
                 "description": "Top-Sharpe, pairwise ρ < 0.85", "systems": sel})

    n    = min(6, len(corr))
    seed = sharpes.idxmax()
    grp  = [seed]
    rem  = [x for x in corr.columns if x != seed]
    while len(grp) < n and rem:
        avg_c = {r: corr.loc[r, grp].mean() for r in rem}
        nxt   = min(avg_c, key=avg_c.get)
        grp.append(nxt); rem.remove(nxt)
    recs.append({"name": "🌐 Max Diversification",
                 "description": "Min avg pairwise correlation", "systems": grp})

    scores = {}
    for col in daily.columns:
        cum = daily[col].cumsum()
        dd  = (cum - cum.cummax()).min()
        scores[col] = means[col] / abs(dd) if dd < 0 else 0.0
    recs.append({"name": "🛡 Min Drawdown",
                 "description": "Best return/drawdown ratio",
                 "systems": sorted(scores, key=scores.get, reverse=True)[:6]})
    return recs


# ─────────────────────────────────────────────────────────────────────────────
# MONTE CARLO
# ─────────────────────────────────────────────────────────────────────────────

def monte_carlo_simulation(systems: dict, stems: list, lookback_years: float,
                            n_sims: int = 1000, forward_days: int = 126,
                            sizing_ratios: dict | None = None,
                            method: str = "Regime-conditioned bootstrap",
                            block_days: int = 5,
                            regime_window: int = 20,
                            low_pct: int = 33, high_pct: int = 66,
                            account_balance: float = 130000,
                            seed: int = 42) -> dict | None:
    """Monte Carlo of historical daily portfolio P&L.

    The simulation operates on the already-combined portfolio daily P&L, so
    cross-system correlation is preserved in every historical observation.
    IID bootstrap resamples individual days. Block bootstrap resamples short
    consecutive blocks, preserving some serial dependence (runs of wins/losses
    and volatility clustering). Regime-conditioned bootstrap first classifies
    historical portfolio volatility into Low/Medium/High states, estimates
    state-to-state transition probabilities, then samples daily P&L from the
    simulated state. This captures regime persistence without pretending we
    know the future regime.
    """
    curves = {}
    for stem in stems:
        si = systems.get(stem)
        if si is None or si["trades"].empty:
            continue
        ratio = (sizing_ratios or {}).get(stem, 1.0)
        eq = get_net_equity_trimmed(si, lookback_years, scale=ratio)
        if not eq.empty:
            curves[stem] = eq
    if not curves:
        return None

    # Only use the period where the portfolio has valid observations.
    # Missing dates are not silently converted to zero before the first
    # observation of a system.
    df = pd.concat(curves.values(), axis=1, keys=curves.keys()).sort_index().ffill()
    daily = df.diff().dropna(how="all")
    # Do not mix early periods with fewer live systems into the forward pool.
    daily = daily.dropna(how="any")
    if daily.empty:
        return None
    pool = daily.sum(axis=1).dropna().values.astype(float)
    if len(pool) < 20:
        return None

    rng = np.random.default_rng(seed)
    paths = np.zeros((forward_days + 1, n_sims), dtype=float)

    regime_info = {}
    if method == "IID bootstrap" or (block_days <= 1 and method == "Block bootstrap"):
        draws = rng.integers(0, len(pool), size=(forward_days, n_sims))
        sampled = pool[draws]
    elif method == "Block bootstrap":
        # Moving-block bootstrap. Start positions are random; blocks may wrap
        # around the historical sample so every path reaches the horizon.
        n_blocks = int(np.ceil(forward_days / block_days))
        starts = rng.integers(0, len(pool), size=(n_blocks, n_sims))
        offsets = np.arange(block_days)[:, None]
        pieces = []
        for j in range(n_blocks):
            idx = (starts[j][None, :] + offsets) % len(pool)
            pieces.append(pool[idx])
        sampled = np.concatenate(pieces, axis=0)[:forward_days, :]
    else:
        # Regime-conditioned bootstrap: classify historical portfolio daily P&L
        # by rolling realized volatility, estimate an empirical Markov transition
        # matrix, and draw P&L from the regime selected for each simulated day.
        # The simulated path starts in the regime observed at the end of history.
        s = pd.Series(pool, index=daily.index[-len(pool):])
        roll_vol = s.rolling(regime_window).std() * np.sqrt(252)
        valid = roll_vol.dropna()
        if len(valid) < max(60, regime_window * 3):
            return None
        lo = float(np.percentile(valid.values, low_pct))
        hi = float(np.percentile(valid.values, high_pct))
        labels = pd.Series(index=s.index, dtype="object")
        labels.loc[roll_vol <= lo] = "Low Vol"
        labels.loc[(roll_vol > lo) & (roll_vol < hi)] = "Medium Vol"
        labels.loc[roll_vol >= hi] = "High Vol"
        labels = labels.dropna()
        aligned = s.reindex(labels.index).dropna()
        labels = labels.reindex(aligned.index).dropna()
        states = ["Low Vol", "Medium Vol", "High Vol"]
        pools = {state_name: aligned[labels == state_name].values.astype(float) for state_name in states}
        if any(len(v) < 10 for v in pools.values()):
            return None

        trans = pd.DataFrame(0.0, index=states, columns=states)
        counts = pd.Series(0.0, index=states)
        arr = labels.values
        for i in range(len(arr) - 1):
            a, b = arr[i], arr[i + 1]
            if a in states and b in states:
                trans.loc[a, b] += 1
                counts.loc[a] += 1
        for state_name in states:
            if counts.loc[state_name] > 0:
                trans.loc[state_name] /= counts.loc[state_name]
            else:
                trans.loc[state_name, state_name] = 1.0

        current_state = labels.iloc[-1]
        sampled = np.empty((forward_days, n_sims), dtype=float)
        state_arr = np.empty((forward_days, n_sims), dtype=object)
        prev_states = np.full(n_sims, current_state, dtype=object)
        for d in range(forward_days):
            for j in range(n_sims):
                probs = trans.loc[prev_states[j]].values
                next_state = rng.choice(states, p=probs / probs.sum())
                state_arr[d, j] = next_state
                sampled[d, j] = rng.choice(pools[next_state])
            prev_states = state_arr[d, :].copy()

        persist = float(trans.loc[current_state, current_state])
        regime_info = {
            "current_regime": current_state,
            "regime_low_threshold": lo,
            "regime_high_threshold": hi,
            "regime_persistence": persist,
            "regime_transition": trans,
        }

    paths[1:, :] = np.cumsum(sampled, axis=0)

    pctiles = pd.DataFrame({
        k: np.percentile(paths[1:, :], p, axis=1)
        for k, p in [("5th", 5), ("25th", 25), ("50th", 50),
                     ("75th", 75), ("95th", 95)]
    })

    final = paths[-1, :]
    peak = np.maximum.accumulate(paths, axis=0)
    drawdowns = paths - peak
    max_dd = drawdowns.min(axis=0)

    # Use meaningful, account-scaled thresholds rather than hard-coded
    # dollar values. They are expressed as fractions of account equity.
    account_balance = float(account_balance)
    dd_thresholds = [0.025, 0.05, 0.10, 0.15, 0.20, 0.25]
    dd_probs = {t: float((max_dd < -(account_balance * t)).mean())
                for t in dd_thresholds}

    terminal_var_5 = float(np.percentile(final, 5))
    tail = final[final <= terminal_var_5]
    terminal_es_5 = float(tail.mean()) if len(tail) else terminal_var_5

    return {
        "paths": pd.DataFrame(paths[1:]),
        "percentiles": pctiles,
        "dd_probs": dd_probs,
        "max_dd": max_dd,
        "terminal_var_5": terminal_var_5,
        "terminal_es_5": terminal_es_5,
        "n_sims": n_sims,
        "forward_days": forward_days,
        "method": method,
        "block_days": block_days,
        "historical_days": len(pool),
        **regime_info,
    }


# ─────────────────────────────────────────────────────────────────────────────
# REGIME ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

def compute_regime_series(equity_dict: dict, vol_window: int = 20,
                           low_pct: int = 33, high_pct: int = 66) -> tuple:
    if not equity_dict:
        return pd.DataFrame(), {}
    df    = pd.concat(equity_dict.values(), axis=1,
                      keys=equity_dict.keys()).sort_index().ffill().fillna(0)
    daily = df.sum(axis=1).diff().dropna()
    if len(daily) < vol_window + 10:
        return pd.DataFrame(), {}
    roll_vol = (daily.rolling(vol_window).std() * np.sqrt(252)).dropna()
    if roll_vol.empty:
        return pd.DataFrame(), {}
    lo, hi = np.percentile(roll_vol, low_pct), np.percentile(roll_vol, high_pct)
    def _cls(v):
        return "Low Vol" if v <= lo else ("High Vol" if v >= hi else "Medium Vol")
    return (pd.DataFrame({"vol": roll_vol, "regime": roll_vol.apply(_cls)}),
            {"low": lo, "high": hi})


def compute_system_regime_performance(trades: pd.DataFrame, regime_df: pd.DataFrame,
                                       comm_per_trade: float) -> dict:
    if trades.empty or regime_df.empty:
        return {}
    t = trades.copy()
    t["exit_day"] = pd.to_datetime(t["exit_date"]).dt.normalize()
    n_ct_r = t["n_contracts"] if "n_contracts" in t.columns else 1
    t["pnl_net"]  = t["pnl"] - comm_per_trade * n_ct_r

    def _nr(dt):
        if dt in regime_df.index:
            return regime_df.loc[dt, "regime"]
        for off in range(1, 6):
            prev = dt - pd.Timedelta(days=off)
            if prev in regime_df.index:
                return regime_df.loc[prev, "regime"]
        return None

    t["regime"] = t["exit_day"].apply(_nr)
    t = t.dropna(subset=["regime"])
    if t.empty:
        return {}

    results = {}
    for lbl in ["Low Vol", "Medium Vol", "High Vol"]:
        g = t[t["regime"] == lbl]
        if g.empty:
            results[lbl] = {"n_trades": 0, "total_pnl": 0, "avg_pnl": 0,
                             "win_rate": 0, "profit_factor": 0}
            continue
        wins  = g["pnl_net"] > 0
        gains = g.loc[wins, "pnl_net"].sum()
        loss  = abs(g.loc[~wins, "pnl_net"].sum())
        results[lbl] = {
            "n_trades":      len(g),
            "total_pnl":     round(g["pnl_net"].sum(), 0),
            "avg_pnl":       round(g["pnl_net"].mean(), 0),
            "win_rate":      round(wins.mean(), 3),
            "profit_factor": round(gains / loss if loss > 0 else (np.inf if gains > 0 else 0), 2),
        }
    return results


def compute_regime_transitions(regime_df: pd.DataFrame) -> pd.DataFrame:
    if regime_df.empty or len(regime_df) < 2:
        return pd.DataFrame()
    labels = ["Low Vol", "Medium Vol", "High Vol"]
    trans  = pd.DataFrame(0, index=labels, columns=labels, dtype=float)
    counts = pd.Series(0, index=labels, dtype=float)
    regs   = regime_df["regime"]
    for i in range(len(regs) - 1):
        c, n = regs.iloc[i], regs.iloc[i + 1]
        if c in labels and n in labels:
            trans.loc[c, n] += 1
            counts[c]       += 1
    for r in labels:
        if counts[r] > 0:
            trans.loc[r] = trans.loc[r] / counts[r]
    return trans.round(3)


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────

def build_excel_export(systems: dict, lookback_years: float,
                        sizing_ratios: dict | None = None,
                        corr_matrix=None) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    HDR_FILL = PatternFill("solid", fgColor="1F3864")
    HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
    ALT_FILL = PatternFill("solid", fgColor="F2F4F8")

    def _ws(name, headers, rows):
        ws = wb.create_sheet(name)
        ws.append(headers)
        for c in ws[1]:
            c.fill, c.font = HDR_FILL, HDR_FONT
            c.alignment = Alignment(horizontal="center", vertical="center")
        for i, row in enumerate(rows, 2):
            ws.append(row)
            if i % 2 == 0:
                for c in ws[i]:
                    c.fill = ALT_FILL
        for col in ws.columns:
            w = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(w + 3, 40)
        ws.freeze_panes = "A2"
        return ws

    ratios = sizing_ratios or {}

    # Summary sheet
    hdr1  = ["System", "Contract Label", "Default N", "Current N", "Scale",
             "Source", "# Trades", "Net Profit ($)", "Max DD ($)",
             "P&L Sharpe", "Ann. P&L / Max DD", "PF", "Win Rate", "B5 Match"]
    rows1 = []
    for stem, si in systems.items():
        ratio = ratios.get(stem, 1.0)
        cur_n = round(si["default_n"] * ratio)
        m     = compute_metrics(si["trades"], si["comm_per_trade"], ratio)
        rows1.append([
            si["display_name"][:40], si["contract_label"],
            si["default_n"], cur_n, f"{ratio:.2f}×",
            "📊 trades" if si["data_source"] == "trades" else "⚠️ fallback",
            m.get("# Trades", 0),
            m.get("Net Profit ($)", ""),    m.get("Max Drawdown ($)", ""),
            m.get("P&L Sharpe", ""),        m.get("Ann. P&L / Max DD", ""),
            m.get("Profit Factor", ""),
            f'{m.get("Win Rate", 0):.1%}' if m else "",
            "✅" if si["b5_match"] else "❌",
        ])
    _ws("Summary", hdr1, rows1)

    # Trade log
    hdr2  = ["System", "Trade#", "Direction", "Entry Date", "Exit Date",
             "N Contracts (TS)", "P&L ($) Raw", "Cum P&L ($)"]
    rows2 = []
    for stem, si in systems.items():
        for _, r in si["trades"].iterrows():
            rows2.append([
                si["display_name"][:30], r["trade_id"], r["direction"],
                str(r["entry_date"])[:16] if pd.notna(r.get("entry_date")) else "",
                str(r["exit_date"])[:16],
                r.get("n_contracts", 1),
                round(r["pnl"], 2), round(r["cum_pnl"], 2),
            ])
    _ws("Trade Log", hdr2, rows2)

    # Correlation
    if corr_matrix is not None and not corr_matrix.empty:
        ws3  = wb.create_sheet("Correlation")
        cols = [systems[c]["display_name"][:20] if c in systems else c[:20]
                for c in corr_matrix.columns]
        ws3.append([""] + cols)
        for idx, row_data in corr_matrix.iterrows():
            label = systems[idx]["display_name"][:20] if idx in systems else idx[:20]
            ws3.append([label] + [round(v, 3) for v in row_data.values])
        ws3.column_dimensions["A"].width = 28
        lc = get_column_letter(len(cols) + 1)
        lr = len(cols) + 1
        ws3.conditional_formatting.add(
            f"B2:{lc}{lr}",
            ColorScaleRule(start_type="num", start_value=-1, start_color="D73027",
                           mid_type="num",   mid_value=0,    mid_color="FFFFFF",
                           end_type="num",   end_value=1,    end_color="4575B4"))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# PLOTTING
# ─────────────────────────────────────────────────────────────────────────────

C = {
    "drawdown":  "rgba(220,50,47,0.6)",
    "dd_line":   "#dc322f",
    "peak":      "#93a1a1",
    "portfolio": "#073642",
    "green":     "#2aa198",
    "red":       "#dc322f",
    "amber":     "#b58900",
    "blue":      "#268bd2",
    "zero_line": "#93a1a1",
}


def plot_equity(eq: pd.Series, name: str, color: str = "#268bd2") -> go.Figure:
    run_max = eq.cummax()
    dd      = eq - run_max
    fig     = make_subplots(rows=2, cols=1, shared_xaxes=True,
                            row_heights=[0.7, 0.3], vertical_spacing=0.04)
    fig.add_trace(go.Scatter(x=eq.index, y=eq.values, name=name,
                             line=dict(color=color, width=2),
                             hovertemplate="%{x|%Y-%m-%d}<br>$%{y:,.0f}<extra></extra>"),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=eq.index, y=run_max.values, name="Peak",
                             line=dict(color=C["peak"], width=1, dash="dot"),
                             showlegend=False), row=1, col=1)
    fig.add_shape(type="line", x0=eq.index.min(), x1=eq.index.max(),
                  y0=0, y1=0, line=dict(color=C["zero_line"], width=1, dash="dash"),
                  xref="x", yref="y", row=1, col=1)
    fig.add_trace(go.Scatter(x=eq.index, y=dd.values, name="DD",
                             fill="tozeroy", line=dict(color=C["dd_line"], width=1),
                             fillcolor=C["drawdown"],
                             hovertemplate="%{x|%Y-%m-%d}<br>$%{y:,.0f}<extra></extra>"),
                  row=2, col=1)
    fig.update_layout(template=THEME, height=500,
                      margin=dict(l=0, r=0, t=10, b=0),
                      legend=dict(orientation="h", y=1.05),
                      yaxis_title="Cum. P&L ($)", yaxis2_title="Drawdown ($)")
    return fig


def plot_portfolio_equity(port_eq: pd.Series, eq_df: pd.DataFrame,
                           label_map: dict | None = None) -> go.Figure:
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        row_heights=[0.65, 0.35], vertical_spacing=0.04)
    for i, col in enumerate(eq_df.columns):
        rc     = COLORS[i % len(COLORS)]
        fill_c = rc.replace("rgb", "rgba").replace(")", ",0.45)") if rc.startswith("rgb") else rc
        lbl    = (label_map or {}).get(col, col.replace("_", " ")[:30])
        fig.add_trace(go.Scatter(
            x=eq_df.index, y=eq_df[col].values,
            name=lbl, stackgroup="one",
            line=dict(width=0.5), fillcolor=fill_c,
            hovertemplate=f"{lbl[:25]}<br>%{{x|%Y-%m-%d}}<br>$%{{y:,.0f}}<extra></extra>"),
            row=1, col=1)

    fig.add_trace(go.Scatter(
        x=eq_df.index, y=[0] * len(eq_df),
        mode="lines", name="Zero", showlegend=False,
        line=dict(color=C["zero_line"], width=1, dash="dash")),
        row=1, col=1)
    fig.update_yaxes(zeroline=True, zerolinewidth=1,
                     zerolinecolor=C["zero_line"], row=1, col=1)
    fig.update_layout(template=THEME, height=560,
                      margin=dict(l=0, r=0, t=10, b=0),
                      legend=dict(orientation="h", y=1.05),
                      yaxis_title="Cum. P&L ($)", yaxis2_title="Drawdown ($)")
    return fig


def plot_monthly_heatmap(equity: pd.Series, name: str) -> go.Figure:
    """
    Monthly P&L heatmap with an annual 'Total' column.

    The 12 monthly cells use RdYlGn coloring relative to their own range.
    The 'Total' column uses a *separate* heatmap trace with independent
    coloring so year totals don't distort the monthly color scale and vice versa.
    A thin vertical gap visually separates the two sections.
    """
    monthly = equity.resample("ME").last().diff().dropna()
    if monthly.empty:
        return go.Figure()

    df = pd.DataFrame({"val": monthly})
    df["year"]  = df.index.year
    df["month"] = df.index.month
    pivot = df.pivot(index="year", columns="month", values="val").fillna(0)

    mnames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    # Only keep columns that exist in the data
    month_labels = [mnames[c-1] for c in pivot.columns]
    years        = [str(y) for y in pivot.index.tolist()]

    monthly_z    = pivot.values.tolist()
    monthly_text = [[f"${v:,.0f}" for v in r] for r in pivot.values]

    # Annual totals
    totals      = pivot.sum(axis=1).values.tolist()
    totals_text = [[f"${v:,.0f}"] for v in totals]
    totals_z    = [[v] for v in totals]

    # x-positions: months at 0..N-1, gap, Total at N+0.6
    n_months  = len(month_labels)
    x_months  = list(range(n_months))
    x_total   = [n_months + 0.6]
    x_labels  = month_labels + [""] + ["Total"]   # blank spacer tick

    fig = go.Figure()

    # ── Monthly heatmap ───────────────────────────────────────────────────────
    # Colorbars are placed AFTER the Total column so they never overlap the
    # month cells. Monthly bar sits just right of Total; Annual bar further right.
    # thickness=12 keeps them slim; xpad=4 adds a small gap.
    fig.add_trace(go.Heatmap(
        z=monthly_z,
        x=x_months,
        y=years,
        text=monthly_text,
        texttemplate="%{text}",
        colorscale="RdYlGn",
        zmid=0,
        showscale=True,
        colorbar=dict(
            title=dict(text="Monthly", side="right"),
            x=1.06,           # well right of the Total column
            xpad=4,
            thickness=12,
            len=0.85,
        ),
        hovertemplate="Year %{y}<br>%{x}<br>$%{z:,.0f}<extra></extra>",
        name="Monthly",
    ))

    # ── Annual total heatmap (independent color scale) ────────────────────────
    fig.add_trace(go.Heatmap(
        z=totals_z,
        x=x_total,
        y=years,
        text=totals_text,
        texttemplate="%{text}",
        colorscale="RdYlGn",
        zmid=0,
        showscale=True,
        colorbar=dict(
            title=dict(text="Annual", side="right"),
            x=1.18,           # further right, no overlap with monthly bar
            xpad=4,
            thickness=12,
            len=0.85,
        ),
        hovertemplate="Year %{y}<br>Total: $%{z:,.0f}<extra></extra>",
        name="Annual Total",
    ))

    # Build tick labels: month names at integer positions, blank spacer, "Total"
    tickvals = x_months + [n_months + 0.0, n_months + 0.6]
    ticktext = month_labels + ["", "Total"]

    fig.update_layout(
        template=THEME,
        height=max(220, 34 * len(pivot) + 80),
        # right margin large enough to show both slim colorbars without overlap
        margin=dict(l=0, r=160, t=40, b=0),
        title=f"Monthly P&L — {name}",
        xaxis=dict(
            tickmode="array",
            tickvals=tickvals,
            ticktext=ticktext,
            tickangle=0,
        ),
    )
    return fig


def plot_clustered_correlation(corr_matrix: pd.DataFrame, systems: dict,
                                over_thresh: float = 0.70) -> go.Figure:
    reordered, _ = cluster_correlation_matrix(corr_matrix)
    cols   = reordered.columns.tolist()
    labels = [systems[c]["display_name"][:22] if c in systems else c[:22] for c in cols]
    shapes = []
    n      = len(cols)
    for i in range(n):
        for j in range(i+1, n):
            if reordered.iloc[i, j] > over_thresh:
                shapes.append(dict(type="rect",
                    x0=j-0.5, x1=j+0.5, y0=i-0.5, y1=i+0.5,
                    line=dict(color="rgba(220,50,47,0.9)", width=2)))
    colorscale = [[0.0,"#d73027"],[0.4,"#f7b99e"],
                  [0.5,"#ffffff"],[0.6,"#92c5de"],[1.0,"#4575b4"]]
    fig = go.Figure(go.Heatmap(
        z=reordered.values, x=labels, y=labels,
        text=[[f"{v:.2f}" for v in r] for r in reordered.values],
        texttemplate="%{text}", colorscale=colorscale,
        zmid=0, zmin=-1, zmax=1, colorbar=dict(title="ρ")))
    fig.update_layout(template=THEME, height=680,
                      margin=dict(l=0, r=0, t=50, b=0),
                      title=f"Clustered Correlation (Average linkage; monthly primary) — 🔴 ρ > {over_thresh}",
                      xaxis_tickangle=-45, shapes=shapes)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG & CSS
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title=f"Quant Dashboard {APP_VERSION}", page_icon="📈",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.main,[data-testid="stAppViewContainer"]{background:#f8f9fb;color:#1a1a2e}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e0e4ea}
[data-testid="stMetricValue"]{font-size:1.35rem;font-weight:700;color:#1a1a2e}
[data-testid="stMetricLabel"]{font-size:.75rem;color:#5a6070;font-weight:500}
div.stTabs [data-baseweb="tab"]{height:38px;padding:0 18px;border-radius:6px 6px 0 0;
    font-weight:600;color:#3a3f52}
div.stTabs [aria-selected="true"]{background:#fff;color:#0057ff;border-bottom:3px solid #0057ff}
[data-testid="stDataFrame"]{border:1px solid #e0e4ea;border-radius:8px}
[data-testid="stExpander"]{background:#fff;border:1px solid #e0e4ea;border-radius:8px}
hr{border-color:#e0e4ea}
h1,h2,h3{color:#1a1a2e!important}
.sizing-panel{background:#fff;border:1px solid #e0e4ea;border-radius:10px;
    padding:14px 18px 6px;margin-bottom:16px}
.sizing-label{font-size:.78rem;color:#5a6070;font-weight:600;margin-bottom:2px}
.sizing-default{font-size:.70rem;color:#93a1a1}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("⚙️ Settings")
    data_dir_input = st.text_input("Data directory", value=str(DATA_DIR))
    lookback       = st.slider("Lookback (years)", 1, 10, 10)
    account_balance = st.number_input(
        "Portfolio account balance ($)",
        min_value=1_000.0, max_value=10_000_000.0,
        value=130_000.0, step=5_000.0, format="%0.0f",
        help=("Capital you want the dashboard to use as the reference equity. "
              "This is an input for risk/leverage context, not historical P&L."))
    st.caption("Used for risk/equity and sizing context — not for calculating historical P&L.")
    constant_composition = st.checkbox("Constant composition for portfolio analytics", value=True, key="constant_composition_v251", help="When on, portfolio performance and sizing diagnostics use the same system set throughout the analysed window.")
    target_annual_risk = st.number_input(
        "Target annualized portfolio risk (%)",
        min_value=5.0, max_value=100.0, value=25.0, step=1.0,
        format="%0.0f",
        help=("Reference target for portfolio risk. It is not a forecast and not a margin requirement. "
              "In plain English: how much annualized P&L turbulence are you willing to tolerate relative to the account?"))
    st.divider()
    st.caption("**Commission rates ($/contract, round-trip)**")
    st.write("Edit per-contract round-trip commission rates below. These are applied as `rate * n_contracts` for each trade.")
    # UI order for contract types
    _contract_order = ["MES", "MNQ", "MGC", "NQ", "GC", "CL"]
    commission_rates = {}
    for token in _contract_order:
        default_val = COMMISSION_RATES.get(token, (COMMISSION_PER_MICRO if token in MICRO_PREFIXES else COMMISSION_PER_MINI))
        commission_rates[token] = st.number_input(f"{token} ($/contract)", value=float(default_val), step=0.01, key=f"comm_{token}")
    st.divider()
    st.divider()
    st.caption("**Risk Parity**")
    rp_target = st.number_input("Target daily risk/system ($)",
                                 value=500.0, min_value=50.0,
                                 max_value=5000.0, step=50.0)
    rp_max_ct = st.slider("Max contracts (RP)", 1, 30, 20)
    st.divider()
    st.caption(f"{APP_VERSION} — Correlation audit · Risk-aware portfolio sizing · Dynamic sizing · Raw-trade P&L")


# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────

with st.spinner("Loading trading systems from raw trades…"):
    # pass the user-edited commission mapping into the loader
    systems, load_warnings = load_all_systems(
        data_dir_input, commission_rates,
        _version=_LOAD_VERSION)

if not systems:
    st.error("No trading-system files (.xlsx or .csv) could be loaded from the selected data directory.")
    st.stop()

for w in load_warnings:
    st.warning(w)

n_sys        = len(systems)
total_trades = sum(len(si["trades"]) for si in systems.values())
n_active_readme = sum(1 for si in systems.values() if si["default_n"] > 0 and not si["trades"].empty)

st.title(f"📊 Quant Portfolio Dashboard {APP_VERSION}  —  {n_sys} files · {n_active_readme} active · {total_trades:,} trades")

# ── Initialise session_state sizing dict on first load ─────────────────────
if "sizing" not in st.session_state:
    st.session_state["sizing"] = {
        stem: systems[stem]["default_n"] for stem in systems
    }
# Ensure any newly discovered system is present in session state
for stem in systems:
    if stem not in st.session_state["sizing"]:
        st.session_state["sizing"][stem] = systems[stem]["default_n"]

# Only explicitly active, matched systems enter portfolio risk and sizing.
_active_stems = [
    stem for stem in systems
    if not systems[stem]["trades"].empty
    and not systems[stem]["equity"].empty
    and systems[stem]["default_n"] > 0
    and st.session_state["sizing"].get(stem, systems[stem]["default_n"]) > 0
]


def _cur_ratio(stem: str) -> float:
    """Sizing ratio for this stem: current_n / default_n."""
    default = systems[stem]["default_n"]
    if default == 0:
        return 0.0
    cur = st.session_state["sizing"].get(stem, default)
    return cur / default


# Risk parity (advisory)
_rp_sizing, _rp_detail = compute_risk_parity_sizing(
    systems, _active_stems, lookback,
    target_daily_risk_usd=rp_target, max_contracts=rp_max_ct)

# ── Sidebar Excel export ───────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    _export_corr = None
    _export_eq   = {}
    for stem in _active_stems:
        si = systems[stem]
        eq = get_net_equity_trimmed(si, lookback, _cur_ratio(stem))
        if not eq.empty:
            _export_eq[stem] = eq
    if len(_export_eq) >= 2:
        _export_daily_corr = build_correlation_daily_pnl(_export_eq, lookback)
        _export_corr, _, _ = _pairwise_corr_cov(_export_daily_corr, min_overlap_days=63)

    if st.button("📥 Generate Excel Report", use_container_width=True):
        with st.spinner("Building workbook…"):
            ratios = {s: _cur_ratio(s) for s in systems}
            xls = build_excel_export(systems, lookback,
                                      sizing_ratios=ratios,
                                      corr_matrix=_export_corr)
        st.download_button(
            "⬇️ Download Excel", data=xls,
            file_name=f"quant_{APP_VERSION}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────

def b5_is_none(si: dict) -> bool:
    """True if b5_total is absent → show '—' not '❌' in the overview table."""
    return si["b5_total"] is None


tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🖥 Systems", "📦 Portfolio",
    "🔬 Correlation & Optimisation", "🔮 Forward Analysis",
    "🔗 Trade Matching"])


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 — SYSTEMS
# ═════════════════════════════════════════════════════════════════════════════

with tab1:
    # ── SIZING PANEL ─────────────────────────────────────────────────────────
    st.subheader("⚖️ Position Sizing Editor")
    active_readme = sum(1 for si in systems.values() if si["default_n"] > 0)
    inactive_readme = sum(1 for si in systems.values() if si["default_n"] == 0)
    market_count = len({si["symbol"] for si in systems.values()})
    st.caption(f"{active_readme} active ReadMe allocations · {inactive_readme} zero-size/inactive · {market_count} markets")
    st.caption(
        "Adjust contract counts below. Metrics and equity curves rescale live as "
        "**P&L × (current / default)**. Default values come from ReadMe.txt. "
        "Raw TradeStation P&L is never altered — only the display is scaled.")

    # Reset button
    col_reset, _ = st.columns([1, 5])
    with col_reset:
        if st.button("↺ Reset all to ReadMe defaults", use_container_width=True):
            st.session_state["sizing"] = {
                stem: systems[stem]["default_n"] for stem in systems
            }
            st.rerun()

    # Render a compact grid: 4 systems per row
    stems_list = list(systems.keys())
    n_cols     = 4
    for row_start in range(0, len(stems_list), n_cols):
        row_stems = stems_list[row_start: row_start + n_cols]
        cols      = st.columns(n_cols)
        for col_widget, stem in zip(cols, row_stems):
            si      = systems[stem]
            default = si["default_n"]
            ctype   = si["contract_type"]
            cur_val = st.session_state["sizing"].get(stem, default)

            with col_widget:
                new_val = st.number_input(
                    label=si["display_name"][:30],
                    min_value=0,
                    max_value=50,
                    value=int(cur_val),
                    step=1,
                    key=f"sz_{stem}",
                    help=(f"ReadMe default: {default} {ctype}\n"
                          f"Contract label: {si['contract_label']}"),
                )
                st.caption(f"Default: {default} {ctype}")
                if new_val != cur_val:
                    st.session_state["sizing"][stem] = new_val

    st.divider()

    # ── OVERVIEW TABLE ───────────────────────────────────────────────────────
    st.subheader("System Overview")
    st.caption(
        "P&L is TradeStation full-position dollars scaled to current sizing. "
        "Values are rounded only for display; calculations retain full precision. "
        "Use the Data quality & validation section below to inspect source/coverage issues.")

    overview_dates = [si["trades"]["exit_date"] for si in systems.values() if not si["trades"].empty]
    if overview_dates:
        ov_min = min(series.min() for series in overview_dates)
        ov_max = max(series.max() for series in overview_dates)
    else:
        ov_min = pd.Timestamp.today()
        ov_max = pd.Timestamp.today()

    ov_presets = [
        "All history", "Past week", "Past 1 month", "Past 3 months", "Past 6 months",
        "Past year", "Past 2 years", "MTD", "YTD", "2025", "2024", "Custom"]
    if "ov_date_preset" not in st.session_state or st.session_state["ov_date_preset"] not in ov_presets:
        st.session_state["ov_date_preset"] = "All history"
    if "ov_custom_range" not in st.session_state:
        st.session_state["ov_custom_range"] = [ov_min.date(), ov_max.date()]

    with st.expander("Date filter", expanded=True):
        btn1 = st.columns(6)
        if btn1[0].button("All history", key="ov_all"):
            st.session_state["ov_date_preset"] = "All history"
        if btn1[1].button("Past week", key="ov_week"):
            st.session_state["ov_date_preset"] = "Past week"
        if btn1[2].button("Past 1 month", key="ov_1m"):
            st.session_state["ov_date_preset"] = "Past 1 month"
        if btn1[3].button("Past 3 months", key="ov_3m"):
            st.session_state["ov_date_preset"] = "Past 3 months"
        if btn1[4].button("Past 6 months", key="ov_6m"):
            st.session_state["ov_date_preset"] = "Past 6 months"
        if btn1[5].button("Past year", key="ov_1y"):
            st.session_state["ov_date_preset"] = "Past year"

        btn2 = st.columns(6)
        if btn2[0].button("Past 2 years", key="ov_2y"):
            st.session_state["ov_date_preset"] = "Past 2 years"
        if btn2[1].button("MTD", key="ov_mtd"):
            st.session_state["ov_date_preset"] = "MTD"
        if btn2[2].button("YTD", key="ov_ytd"):
            st.session_state["ov_date_preset"] = "YTD"
        if btn2[3].button("2025", key="ov_2025"):
            st.session_state["ov_date_preset"] = "2025"
        if btn2[4].button("2024", key="ov_2024"):
            st.session_state["ov_date_preset"] = "2024"
        if btn2[5].button("Custom", key="ov_custom"):
            st.session_state["ov_date_preset"] = "Custom"

        ov_custom_range = st.date_input(
            "Custom range",
            value=st.session_state["ov_custom_range"],
            min_value=ov_min.date(),
            max_value=ov_max.date(),
            key="ov_custom_range")

        st.caption(f"Active filter: {st.session_state['ov_date_preset']}")

    ov_preset = st.session_state["ov_date_preset"]
    if ov_preset == "Custom":
        if isinstance(ov_custom_range, tuple):
            ov_start = pd.Timestamp(ov_custom_range[0])
            ov_end = pd.Timestamp(ov_custom_range[1])
        else:
            ov_start = pd.Timestamp(ov_custom_range)
            ov_end = pd.Timestamp(ov_custom_range)
    else:
        ov_start, ov_end = resolve_portfolio_date_range(ov_min, ov_max, ov_preset)

    ov_caption = f"Showing {ov_start:%Y-%m-%d} → {ov_end:%Y-%m-%d}  |  Active filter: {ov_preset}"
    st.caption(ov_caption)

    ov_rows = []
    for stem, si in systems.items():
        ratio     = _cur_ratio(stem)
        cur_n     = st.session_state["sizing"].get(stem, si["default_n"])
        inactive  = " ⏸" if si["is_inactive"] else ""
        trades    = si["trades"]
        if not trades.empty:
            trades = trades[(trades["exit_date"] >= ov_start) &
                            (trades["exit_date"] <= ov_end)].copy()
        m         = compute_metrics(trades, si["comm_per_trade"], ratio)
        # Source/validation details belong in the expandable data-quality section,
        # not in the main overview table.
        _ = inactive  # retained for readability of the sizing logic above
        ov_rows.append({
            "System":     si["display_name"],
            "Market":     si["symbol"],
            "Position":   f"{cur_n} {si['contract_type']}" + (f"  (def {si['default_n']})" if cur_n != si["default_n"] else ""),
            "# Trades":   m.get("# Trades", len(trades) if not trades.empty else 0),
            "Net Profit": m.get("Net Profit ($)", np.nan) if m else np.nan,
            "Max DD":     m.get("Max Drawdown ($)", np.nan) if m else np.nan,
            "P&L Sharpe": m.get("P&L Sharpe", np.nan) if m else np.nan,
            "Ann. P&L / DD": m.get("Ann. P&L / Max DD", np.nan) if m else np.nan,
            "PF":         m.get("Profit Factor", np.nan) if m else np.nan,
            "Win %":      m.get("Win Rate", np.nan) if m else np.nan,
        })

    ov_df = pd.DataFrame(ov_rows)
    # ensure numeric columns keep numeric dtype for sorting, then style for display
    for c in ["Net Profit", "Max DD", "Sharpe", "Calmar", "PF", "Win %", "# Trades"]:
        if c in ov_df.columns:
            ov_df[c] = pd.to_numeric(ov_df[c], errors="coerce")

    render_table(
        style_overview_table(ov_df),
        use_container_width=True,
        height=min(42 + 36 * n_sys, 700), hide_index=True)

    with st.expander("🔍 Data quality & validation details"):
        dq_rows = []
        for stem, si in systems.items():
            last_cum = (si["trades"]["cum_pnl"].iloc[-1]
                        if not si["trades"].empty else None)
            dq_rows.append({
                "File":         stem[:45],
                "ReadMe match": "✅" if si["matched"] else "⚠️ no match",
                "Display name": si["display_name"][:40],
                "Contract":     si["contract_label"],
                "Default N":    si["default_n"],
                "Current N":    st.session_state["sizing"].get(stem, si["default_n"]),
                "Comm/contract": f"${si['comm_per_trade']:.2f}",
                "B5 (TS TNP)":  f"${si['b5_total']:,.2f}" if si["b5_total"] else "—",
                "Last cum $":   f"${last_cum:,.2f}" if last_cum is not None else "—",
                "B5 match":     "✅" if si["b5_match"] else "❌",
            })
        render_table(pd.DataFrame(dq_rows), hide_index=True, use_container_width=True)

        st.markdown("**Date ranges per system:**")
        dr_rows = []
        for stem, si in systems.items():
            if not si["trades"].empty:
                dates = si["trades"]["exit_date"].dropna()
                dr_rows.append({
                    "System":   si["display_name"][:40],
                    "First":    dates.min().strftime("%Y-%m-%d"),
                    "Last":     dates.max().strftime("%Y-%m-%d"),
                    "# Trades": len(si["trades"]),
                })
        render_table(pd.DataFrame(dr_rows), hide_index=True, use_container_width=True)

    with st.expander("🧮 Metric methodology & precision", expanded=False):
        st.markdown(
            """- **Net Profit / Max DD:** calculated from the daily net equity curve after the configured per-contract commission.
- **Profit Factor / Win Rate:** calculated from trade-level net P&L after commission.
- **Sharpe / Calmar:** current formulas are preserved in this release; this section documents the methodology only.
- **Sizing:** displayed P&L/equity is scaled by `current contracts / ReadMe default contracts`; raw TradeStation trade P&L is never modified.
- **Precision:** calculations use the underlying numeric values; rounding happens only when tables/metrics are rendered.

A deeper mathematical audit of portfolio risk, correlation, covariance quality, and Monte Carlo assumptions is now implemented in the risk engine and diagnostics below.""")

    st.divider()

    # ── SYSTEM EXPLORER ──────────────────────────────────────────────────────
    st.subheader("System Explorer")
    chosen = st.selectbox("Select system", list(systems.keys()),
                          format_func=lambda x: systems[x]["display_name"])

    si_c     = systems[chosen]
    trades_c = si_c["trades"]
    ratio_c  = _cur_ratio(chosen)
    cur_n_c  = st.session_state["sizing"].get(chosen, si_c["default_n"])

    if trades_c.empty:
        st.warning("No trades found for this system.")
    else:
        dates_min = trades_c["exit_date"].min()
        dates_max = trades_c["exit_date"].max()
        valid_se_presets = [
            "All history", "MTD", "YTD", "Past week", "Past 1 month", "Past 3 months",
            "Past 6 months", "Past year", "Past 2 years", "2025", "2024", "Custom"]
        if "se_date_preset" not in st.session_state or st.session_state["se_date_preset"] not in valid_se_presets:
            st.session_state["se_date_preset"] = "All history"
        if "se_custom_range" not in st.session_state or not isinstance(st.session_state["se_custom_range"], list):
            st.session_state["se_custom_range"] = [dates_min.date(), dates_max.date()]

        with st.expander("Date filter", expanded=True):
            c1 = st.columns(6)
            if c1[0].button("All history", key="se_all"):
                st.session_state["se_date_preset"] = "All history"
            if c1[1].button("MTD", key="se_mtd"):
                st.session_state["se_date_preset"] = "MTD"
            if c1[2].button("YTD", key="se_ytd"):
                st.session_state["se_date_preset"] = "YTD"
            if c1[3].button("Past week", key="se_week"):
                st.session_state["se_date_preset"] = "Past week"
            if c1[4].button("Past 1 month", key="se_1m"):
                st.session_state["se_date_preset"] = "Past 1 month"
            if c1[5].button("Past 3 months", key="se_3m"):
                st.session_state["se_date_preset"] = "Past 3 months"

            c2 = st.columns(6)
            if c2[0].button("Past 6 months", key="se_6m"):
                st.session_state["se_date_preset"] = "Past 6 months"
            if c2[1].button("Past year", key="se_1y"):
                st.session_state["se_date_preset"] = "Past year"
            if c2[2].button("Past 2 years", key="se_2y"):
                st.session_state["se_date_preset"] = "Past 2 years"
            if c2[3].button("2025", key="se_2025"):
                st.session_state["se_date_preset"] = "2025"
            if c2[4].button("2024", key="se_2024"):
                st.session_state["se_date_preset"] = "2024"
            if c2[5].button("Custom", key="se_custom"):
                st.session_state["se_date_preset"] = "Custom"

            custom = st.date_input(
                "Custom range",
                value=st.session_state["se_custom_range"],
                min_value=dates_min.date(),
                max_value=dates_max.date(),
                key="se_custom_range")

            preset = st.session_state["se_date_preset"]
            if preset == "Custom":
                if isinstance(custom, tuple) or isinstance(custom, list):
                    se_start = pd.Timestamp(custom[0])
                    se_end = pd.Timestamp(custom[1])
                else:
                    se_start = pd.Timestamp(custom)
                    se_end = pd.Timestamp(custom)
            else:
                se_start, se_end = resolve_portfolio_date_range(dates_min, dates_max, preset)

            st.caption(f"Showing {se_start:%Y-%m-%d} → {se_end:%Y-%m-%d}  |  Active filter: {preset}")

        # trim trades to selected date window
        tr_trim = trades_c.copy()
        tr_trim = tr_trim[(tr_trim["exit_date"] >= se_start) & (tr_trim["exit_date"] <= se_end)].copy()

        eq_net_c = build_net_equity(tr_trim, si_c["comm_per_trade"], ratio_c)
        if not eq_net_c.empty:
            eq_net_c = eq_net_c - eq_net_c.iloc[0]

        m_c = compute_metrics(tr_trim, si_c["comm_per_trade"], ratio_c)

        st.caption(
            f"**System:** {si_c['display_name']}  |  "
            f"**Contract:** {si_c['contract_label']}  |  "
            f"**Current sizing:** {cur_n_c} × {si_c['contract_type']} "
            f"(default: {si_c['default_n']}, scale: {ratio_c:.2f}×)  |  "
            f"**B5:** {'✅' if si_c['b5_match'] else '❌'}")
        st.caption(
            "**Metrics & equity curve:** scaled to current sizing.  "
            "**Trade list P&L:** raw TradeStation dollars. Display rounding never changes calculations.")

        cols5 = st.columns(5)
        for col, (lbl, key, fmt) in zip(cols5, [
            ("Net Profit",   "Net Profit ($)",   "$"),
            ("Max Drawdown", "Max Drawdown ($)",  "$"),
            ("P&L Sharpe",   "P&L Sharpe",        "f"),
            ("Ann. P&L / DD", "Ann. P&L / Max DD", "f"),
            ("Win Rate",     "Win Rate",          "pct"),
        ]):
            v = m_c.get(key, 0)
            col.metric(lbl,
                       f"${v:,.0f}" if fmt == "$" else
                       (f"{v:.1%}" if fmt == "pct" else f"{v:.2f}"))

        if not eq_net_c.empty:
            color_idx = list(systems.keys()).index(chosen)
            st.plotly_chart(
                plot_equity(eq_net_c, si_c["display_name"],
                            COLORS[color_idx % len(COLORS)]),
                use_container_width=True)
            st.plotly_chart(plot_monthly_heatmap(eq_net_c, si_c["display_name"]),
                            use_container_width=True)

        with st.expander("📋 Trade list"):
            st.caption("Raw TradeStation trade-level P&L; this table is intentionally not rescaled by the current sizing editor.")
            td = tr_trim[["trade_id","entry_date","exit_date",
                           "direction","n_contracts","pnl","cum_pnl"]].copy()
            td.columns = ["#","Entry","Exit","Dir","N Ctrts (TS)",
                          "Raw P&L ($)","Cum P&L ($)"]
            # coerce numeric columns then style for display so sorting remains numeric
            td["Entry"] = pd.to_datetime(td["Entry"], errors="coerce")
            td["Exit"] = pd.to_datetime(td["Exit"], errors="coerce")
            td["N Ctrts (TS)"] = pd.to_numeric(td["N Ctrts (TS)"], errors="coerce")
            td["Raw P&L ($)"] = pd.to_numeric(td["Raw P&L ($)"], errors="coerce")
            td["Cum P&L ($)"] = pd.to_numeric(td["Cum P&L ($)"], errors="coerce")
            render_table(td, use_container_width=True, height=320)


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 — PORTFOLIO
# ═════════════════════════════════════════════════════════════════════════════

def _render_tab2():
        st.subheader("Portfolio Builder")
        st.caption(
            "Portfolio equity = sum of per-system net-equity curves scaled by "
            "current sizing ratios (set in Tab 1).")

        selected_systems = st.multiselect(
            "Select systems to include",
            _active_stems,
            default=_active_stems,
            format_func=lambda x: systems[x]["display_name"])

        if not selected_systems:
            st.info("Select at least one system.")
            return

        dates_min = min(systems[s]["trades"]["exit_date"].min() for s in selected_systems)
        dates_max = max(systems[s]["trades"]["exit_date"].max() for s in selected_systems)

        valid_presets = [
            "All history", "MTD", "YTD", "Past week", "Past 1 month", "Past 3 months",
            "Past 6 months", "Past year", "Past 2 years", "2025", "2024", "Custom"]
        if "pf_date_preset" not in st.session_state or st.session_state["pf_date_preset"] not in valid_presets:
            st.session_state["pf_date_preset"] = "All history"

        _safe_date_range_state("pf_custom_range", dates_min, dates_max)

        with st.expander("Date filter", expanded=True):
            btn_cols1 = st.columns(6)
            if btn_cols1[0].button("All history", key="pf_all"):
                st.session_state["pf_date_preset"] = "All history"
            if btn_cols1[1].button("MTD", key="pf_mtd"):
                st.session_state["pf_date_preset"] = "MTD"
            if btn_cols1[2].button("YTD", key="pf_ytd"):
                st.session_state["pf_date_preset"] = "YTD"
            if btn_cols1[3].button("Past week", key="pf_week"):
                st.session_state["pf_date_preset"] = "Past week"
            if btn_cols1[4].button("Past 1 month", key="pf_1m"):
                st.session_state["pf_date_preset"] = "Past 1 month"
            if btn_cols1[5].button("Past 3 months", key="pf_3m"):
                st.session_state["pf_date_preset"] = "Past 3 months"

            btn_cols2 = st.columns(6)
            if btn_cols2[0].button("Past 6 months", key="pf_6m"):
                st.session_state["pf_date_preset"] = "Past 6 months"
            if btn_cols2[1].button("Past year", key="pf_1y"):
                st.session_state["pf_date_preset"] = "Past year"
            if btn_cols2[2].button("Past 2 years", key="pf_2y"):
                st.session_state["pf_date_preset"] = "Past 2 years"
            if btn_cols2[3].button("2025", key="pf_2025"):
                st.session_state["pf_date_preset"] = "2025"
            if btn_cols2[4].button("2024", key="pf_2024"):
                st.session_state["pf_date_preset"] = "2024"
            if btn_cols2[5].button("Custom", key="pf_custom"):
                st.session_state["pf_date_preset"] = "Custom"

            custom_range = st.date_input(
                "Custom range",
                value=st.session_state["pf_custom_range"],
                min_value=dates_min.date(),
                max_value=dates_max.date(),
                key="pf_custom_range")

            if st.button("Use custom range", key="pf_use_custom"):
                st.session_state["pf_date_preset"] = "Custom"

            preset = st.session_state["pf_date_preset"]
            if preset == "Custom":
                if isinstance(custom_range, tuple):
                    start_date = pd.Timestamp(custom_range[0])
                    end_date = pd.Timestamp(custom_range[1])
                else:
                    start_date = pd.Timestamp(custom_range)
                    end_date = pd.Timestamp(custom_range)
            else:
                start_date, end_date = resolve_portfolio_date_range(dates_min, dates_max, preset)

            st.caption(f"Showing {start_date:%Y-%m-%d} → {end_date:%Y-%m-%d}  |  Active filter: {preset}")

        # Build portfolio using current sizing from session_state
        # Skip systems where current sizing = 0 (they contribute nothing)
        curves     = {}
        label_map  = {}
        for stem in selected_systems:
            si    = systems[stem]
            ratio = _cur_ratio(stem)
            if ratio == 0.0:
                continue
            eq    = get_net_equity_trimmed(si, lookback, ratio)
            if not eq.empty:
                curves[stem]    = eq
                label_map[stem] = si["display_name"][:30]

        if not curves:
            st.warning("No valid equity curves.")
            return

        raw_eq_df = pd.concat(curves.values(), axis=1, keys=curves.keys()).sort_index()
        live_system_count = raw_eq_df.notna().sum(axis=1)
        eq_df = raw_eq_df.dropna(how="any") if constant_composition else combine_equity_curves(curves, lookback_years=lookback)
        if eq_df.empty:
            st.warning("No portfolio data available for the selected date range.")
            return

        if preset == "All history":
            start_date = eq_df.index.min()
            end_date = eq_df.index.max()
        else:
            start_date, end_date = resolve_portfolio_date_range(eq_df.index.min(), eq_df.index.max(), preset) if preset != "Custom" else (start_date, end_date)

        eq_df   = filter_equity_by_date_range(eq_df, start_date, end_date)
        if eq_df.empty:
            st.warning("No portfolio data available for the selected date range.")
            return
        if constant_composition:
            eq_df = eq_df.dropna(how="any")
            if eq_df.empty:
                st.warning("No common constant-composition period exists for the selected date range.")
                return

        eq_df   = eq_df - eq_df.iloc[0]
        port_eq = eq_df.sum(axis=1)
        pm      = compute_portfolio_metrics(port_eq)

        # Capital / risk context: this is the bridge between historical strategy P&L
        # and the actual account size you are using today.
        _recent_sizing = {stem: st.session_state["sizing"].get(stem, systems[stem]["default_n"]) for stem in selected_systems}
        _risk_engine = portfolio_risk(systems, selected_systems, _recent_sizing, lookback, window_days=126, date_start=start_date, date_end=end_date, min_overlap_days=252)
        portfolio_daily_risk = float(_risk_engine.get("portfolio_daily_vol", 0.0))
        st.session_state["_canonical_tab2_risk"] = {"vol": portfolio_daily_risk, "systems": tuple(sorted(selected_systems)), "start": _risk_engine.get("start"), "end": _risk_engine.get("end")}
        _risk_obs = int(_risk_engine.get("observations", 0)); _risk_start = _risk_engine.get("start"); _risk_end = _risk_engine.get("end")
        long_run_ann_vol_usd = float(pm.get("Ann. Volatility ($)", 0))
        if portfolio_daily_risk <= 0:
            portfolio_daily_risk = long_run_ann_vol_usd / np.sqrt(252)
            _risk_label = "long-run fallback (recent risk unavailable)"
        else:
            _risk_label = f"recent {_risk_obs}-day constant-composition estimate for selected systems"
        risk_equity = portfolio_daily_risk / account_balance if account_balance > 0 else 0.0
        ann_risk_equity = portfolio_daily_risk * np.sqrt(252) / account_balance if account_balance > 0 else 0.0
        target_risk = target_annual_risk / 100.0
        risk_budget_used = ann_risk_equity / target_risk if target_risk > 0 else 0.0
        portfolio_scale_to_target = target_risk / ann_risk_equity if ann_risk_equity > 0 else 0.0
        available_daily_risk = max(target_risk * account_balance / np.sqrt(252) - portfolio_daily_risk, 0.0)
        worst_day = float(port_eq.diff().dropna().min()) if len(port_eq) > 1 else 0.0

        st.subheader("💰 Capital & Risk Context")
        st.caption(
            "The account balance is a user-defined capital reference. It does not change historical P&L; "
            "it tells you how hard the current book is leaning on the capital you actually have available. "
            "The target-risk figures are sizing references, not forecasts or margin requirements.")
        cc1, cc2, cc3, cc4, cc5 = st.columns(5)
        metric_with_help(cc1, "Account Balance", f"${account_balance:,.0f}", help=METRIC_HELP.get("Account Balance"))
        metric_with_help(cc2, "Current Daily Risk", f"${portfolio_daily_risk:,.0f}",
                         help="""**Formula:** standard deviation of combined daily P&L over the most recent 126 valid trading days, using only currently active systems.

    **Purpose:** estimate recent diversified portfolio risk rather than a decade-long average.

    **In plain English:** what does the book look capable of wobbling by now?

    - This is a volatility estimate, not a worst-case loss or margin requirement.""")
        metric_with_help(cc3, "Daily Risk / Equity", f"{risk_equity:.2%}", help=METRIC_HELP.get("Daily Risk / Equity"))
        metric_with_help(cc4, "Annualized Risk / Equity", f"{ann_risk_equity:.1%}", help=METRIC_HELP.get("Annualized Risk / Equity"))
        metric_with_help(cc5, "Target Annual Risk", f"{target_risk:.0%}",
                         help="User-defined target for annualized portfolio P&L volatility relative to account equity. In plain English: your chosen risk speed limit.")

        rcx1, rcx2, rcx3, rcx4 = st.columns(4)
        metric_with_help(rcx1, "Risk Budget Used", f"{risk_budget_used:.0%}",
                         help="Formula: current annualized risk / target annualized risk. Purpose: show how much of the chosen risk budget the current book consumes. In plain English: 100% is on the speed limit; 130% means we are already 30% over it.")
        metric_with_help(rcx2, "Scale to Target", f"{portfolio_scale_to_target:.2f}×",
                         help="Formula: target annualized risk / current annualized risk. Purpose: estimate the uniform sizing multiplier that would bring current portfolio risk to the target. In plain English: 0.80× means cut every position by roughly 20%; 1.20× means the book could be about 20% larger. This preserves the current relative system weights.")
        metric_with_help(rcx3, "Available Daily Risk", f"${available_daily_risk:,.0f}",
                         help="Formula: target daily risk budget minus current daily risk, floored at zero. Purpose: show remaining risk capacity before hitting the chosen annualized-risk target. In plain English: how much more normal daily turbulence can we afford?")
        metric_with_help(rcx4, "Worst Historical Day", f"${worst_day:,.0f}",
                         help="Largest one-day portfolio P&L loss in the selected history. Purpose: show realized tail pain. In plain English: the ugliest single trading day in the book's history — not a forecast of the next one.")
        _risk_dates = f"{_risk_start:%Y-%m-%d} → {_risk_end:%Y-%m-%d}" if _risk_start is not None else "n/a"
        st.caption(f"Current-risk basis: **{_risk_label}** ({_risk_dates}). Full-window annualized volatility remains **${long_run_ann_vol_usd:,.0f}** for historical context.")

        if ann_risk_equity > 0:
            status = "🟢 Below target" if risk_budget_used <= 0.95 else ("🟡 Near target" if risk_budget_used <= 1.05 else "🔴 Above target")
            st.caption(
                f"**Risk budget:** {status}. Current book is using **{risk_budget_used:.0%}** of the "
                f"chosen {target_risk:.0%} annualized-risk target. A uniform **{portfolio_scale_to_target:.2f}×** "
                "scaling of all positions is the first-order sizing required to hit the target, because "
                "portfolio P&L volatility scales linearly when every position is scaled together.")

        mc5 = st.columns(5)
        for col, (lbl, key, fmt) in zip(mc5, [
            ("Portfolio P&L",   "Net Profit ($)",   "$"),
            ("Ann. P&L",        "Ann. P&L ($)",      "$"),
            ("Max Drawdown",    "Max Drawdown ($)",  "$"),
            ("P&L Sharpe",      "P&L Sharpe",        "f"),
            ("Ann. P&L / Max DD", "Ann. P&L / Max DD", "f"),
        ]):
            v = pm.get(key, 0)
            metric_with_help(col, lbl, f"${v:,.0f}" if fmt == "$" else f"{v:.2f}")

        st.plotly_chart(
            plot_portfolio_equity(port_eq, eq_df, label_map=label_map),
            use_container_width=True)
        live_subset = live_system_count[(live_system_count.index >= start_date) & (live_system_count.index <= end_date)]
        if not live_subset.empty:
            st.caption("Live-system count before the constant-composition filter. This shows where unfiltered portfolio statistics would change their risk base over time.")
            fig_live = go.Figure(go.Scatter(x=live_subset.index, y=live_subset.values, mode="lines", name="Live systems"))
            fig_live.update_layout(template=THEME, height=180, margin=dict(l=0,r=0,t=10,b=0), yaxis_title="# live systems")
            st.plotly_chart(fig_live, use_container_width=True, key="chart_live_system_count_v251")

        # Contribution bar
        st.subheader("System Contribution to Portfolio P&L")
        contrib  = {s: eq_df[s].iloc[-1] - eq_df[s].iloc[0] for s in eq_df.columns}
        cs       = pd.Series(contrib).sort_values(ascending=True)
        bar_labels = [systems[k]["display_name"][:38] for k in cs.index]
        fig_bar  = go.Figure(go.Bar(
            x=cs.values, y=bar_labels, orientation="h",
            marker_color=[C["red"] if v < 0 else C["green"] for v in cs.values],
            text=[f"${v:,.0f}" for v in cs.values],
            textposition="outside"))
        fig_bar.add_shape(
            type="line",
            x0=0, x1=0, y0=-0.5, y1=len(cs) - 0.5,
            line=dict(color="black", width=2))
        fig_bar.update_layout(template=THEME, height=max(300, 38 * len(cs)),
                               margin=dict(l=0, r=70, t=10, b=0),
                               xaxis_title="P&L Contribution ($)",
                               xaxis=dict(zeroline=True, zerolinewidth=2, zerolinecolor="black"))
        st.plotly_chart(fig_bar, use_container_width=True, key="chart_1")

        st.plotly_chart(plot_monthly_heatmap(port_eq, "Portfolio"), use_container_width=True)

        # ── Drawdown Decomposition ────────────────────────────────────────────────
        st.subheader("🔍 Drawdown Decomposition")
        st.caption("For each major portfolio drawdown, which system caused it?")
        dd_episodes = decompose_drawdown(eq_df, port_eq, top_n=3)
        if dd_episodes:
            for ep_i, ep in enumerate(dd_episodes):
                title = (f"DD #{ep_i+1}:  {ep['peak_date'].strftime('%Y-%m-%d')} → "
                         f"{ep['trough_date'].strftime('%Y-%m-%d')}  "
                         f"(Loss: ${ep['dd_abs']:,.0f})")
                with st.expander(title, expanded=(ep_i == 0)):
                    blame_s = sorted(ep["blame_pct"].items(),
                                      key=lambda x: x[1], reverse=True)
                    blame_s = [(k, v) for k, v in blame_s if v > 0][:8]
                    if blame_s:
                        bkeys = [systems[k]["display_name"][:30] if k in systems else k[:30]
                                 for k, _ in blame_s]
                        bvals = [v for _, v in blame_s]
                        fig_b = go.Figure(go.Bar(
                            y=bkeys, x=bvals, orientation="h",
                            marker_color=[C["red"] if v > 15 else C["amber"] for v in bvals],
                            text=[f"{v:.1f}%" for v in bvals],
                            textposition="outside"))
                        fig_b.update_layout(template=THEME,
                                             height=max(200, 35 * len(bkeys)),
                                             margin=dict(l=0, r=60, t=10, b=0),
                                             xaxis_title="Blame (%)")
                        st.plotly_chart(fig_b, use_container_width=True, key=f"chart_2_{ep_i}")
        else:
            st.info("No significant drawdown episodes in this lookback window.")

        with st.expander("📉 Overall Drawdown Analysis"):
            dd     = port_eq - port_eq.cummax()
            dd_neg = dd[dd < 0]
            if not dd_neg.empty:
                d1, d2, d3 = st.columns(3)
                metric_with_help(d1, "Max Drawdown",     f"${dd_neg.min():,.0f}")
                d2.metric("Avg Drawdown",     f"${dd_neg.mean():,.0f}")
                d3.metric("Time in Drawdown", f"{len(dd_neg)/len(dd)*100:.1f}%")
                fig_dd = go.Figure(go.Scatter(
                    x=dd.index, y=dd.values, fill="tozeroy",
                    line=dict(color=C["dd_line"]), fillcolor=C["drawdown"]))
                fig_dd.update_layout(template=THEME, height=240,
                                      margin=dict(l=0, r=0, t=0, b=0),
                                      yaxis_title="Drawdown ($)")
                st.plotly_chart(fig_dd, use_container_width=True, key="chart_3")


    # ═════════════════════════════════════════════════════════════════════════════
    # TAB 3 — CORRELATION & OPTIMISATION
    # ═════════════════════════════════════════════════════════════════════════════

_render_tab2()

def _render_tab3():
        _current_sizing_map = {
            stem: st.session_state["sizing"].get(stem, systems[stem]["default_n"])
            for stem in _active_stems
        }

        all_curves = {}
        for stem in _active_stems:
            si = systems[stem]
            eq = get_net_equity_trimmed(si, lookback, _cur_ratio(stem))
            if not eq.empty:
                all_curves[stem] = eq

        if len(all_curves) < 2:
            st.info("Need ≥ 2 systems with data.")
            return

        daily_r = build_correlation_daily_pnl(all_curves, lookback_years=lookback)
        corr_mat, cov_mat, overlap_mat = _pairwise_corr_cov(daily_r, min_overlap_days=63)
        _weekly_pnl = _build_periodic_trade_pnl(systems, _active_stems, lookback, _current_sizing_map, "W")
        _monthly_pnl = _build_periodic_trade_pnl(systems, _active_stems, lookback, _current_sizing_map, "M")
        weekly_corr = _weekly_pnl.corr(min_periods=12) if not _weekly_pnl.empty else pd.DataFrame()
        monthly_corr = _monthly_pnl.corr(min_periods=6) if not _monthly_pnl.empty else pd.DataFrame()

        if corr_mat.empty or corr_mat.shape[1] < 2:
            st.warning("Not enough overlapping history to calculate portfolio correlation.")
            return

        upper = corr_mat.where(np.triu(np.ones(corr_mat.shape), k=1).astype(bool))
        corr_vals = upper.stack()
        if corr_vals.empty:
            st.warning("No valid system pairs with enough overlapping history.")
            return

        risk_stats = portfolio_risk(systems, _active_stems, _current_sizing_map, lookback, min_overlap_days=252)
        over_exp_thresh = st.slider("Over-exposed threshold (ρ)", 0.5, 0.95, 0.70, 0.05)
        _primary_corr = monthly_corr if not monthly_corr.empty else corr_mat
        cluster_labels = extract_correlation_clusters(_primary_corr, over_exp_thresh)
        n_overexp = int((corr_vals > over_exp_thresh).sum())

        # ── Diversification summary ───────────────────────────────────────────────
        st.subheader("Portfolio Correlation & Diversification")
        st.info(
            "**Start here:** this section answers one simple question: **are my 16 systems actually diversifying each other, or are several of them taking the same risk?**\n\n"
            "- **Correlation (ρ):** how similarly two systems tend to make or lose money on the same day.\n"
            "- **Portfolio Daily Vol:** the volatility of the *combined* book after those relationships are taken into account.\n"
            "- **Risk Contribution:** which systems are responsible for the portfolio's total risk.\n"
            "- **Clusters:** groups of systems that behave similarly.\n\n"
            "The important distinction is: **system volatility tells us how risky a system is alone; correlation tells us how much that risk overlaps with the rest of the portfolio.**")
        st.caption(
            "Correlation is based on overlapping daily dollar P&L. Missing history is left missing, "
            "rather than treated as zero P&L. Current sizing is used for volatility and risk contribution.")

        d1, d2, d3, d4, d5, d6 = st.columns(6)
        metric_with_help(d1, "Portfolio Daily Vol", f"${risk_stats.get('portfolio_daily_vol', 0):,.0f}")
        metric_with_help(d2, "Sum Individual Vol", f"${risk_stats.get('sum_individual_vol', 0):,.0f}")
        metric_with_help(d3, "Diversification Ratio", f"{risk_stats.get('diversification_ratio', np.nan):.2f}")
        metric_with_help(d4, "Avg Daily ρ", f"{corr_vals.mean():.2f}")
        _weekly_vals = weekly_corr.where(np.triu(np.ones(weekly_corr.shape), k=1).astype(bool)).stack() if not weekly_corr.empty else pd.Series(dtype=float)
        _monthly_vals = monthly_corr.where(np.triu(np.ones(monthly_corr.shape), k=1).astype(bool)).stack() if not monthly_corr.empty else pd.Series(dtype=float)
        metric_with_help(d5, "Avg Monthly ρ", f"{_monthly_vals.mean():.2f}" if not _monthly_vals.empty else "n/a")
        metric_with_help(d6, "Effective Bets (monthly)", f"{_effective_independent_bets(monthly_corr):.1f}" if not monthly_corr.empty else "n/a")

        c1, c2 = st.columns(2)
        c1.metric(f"Daily pairs ρ > {over_exp_thresh:.2f}", f"{n_overexp}",
                  delta="⚠️ Concentration" if n_overexp > 3 else "OK",
                  delta_color="inverse")
        c2.metric("Active systems", f"{len(corr_mat)}")

        st.caption(
            "Portfolio Daily Vol is the volatility of the combined book after correlations. "
            "Sum Individual Vol is what the risk would look like if every system moved together. "
            "The gap between them is diversification doing useful work.")
        _horizon_rows = []
        for _label, _mat, _desc in [("Daily", corr_mat, "Short-term co-movement"), ("Weekly", weekly_corr, "Intermediate diversification"), ("Monthly", monthly_corr, "Primary structural-diversification measure")]:
            if _mat.empty: continue
            _vals = _mat.where(np.triu(np.ones(_mat.shape), k=1).astype(bool)).stack()
            _horizon_rows.append({"Horizon": _label, "Average pairwise ρ": float(_vals.mean()) if not _vals.empty else np.nan, "Effective independent bets": _effective_independent_bets(_mat), "Interpretation": _desc})
        if _horizon_rows: render_table(pd.DataFrame(_horizon_rows), hide_index=True, use_container_width=True)
        q1, q2, q3 = st.columns(3)
        q1.metric("LW shrinkage intensity", f"{risk_stats.get('shrinkage_intensity', np.nan):.1%}" if np.isfinite(risk_stats.get('shrinkage_intensity', np.nan)) else "n/a")
        q2.metric("Minimum pairwise overlap", f"{risk_stats.get('min_pairwise_overlap', 0):,} days")
        q3.metric("PSD correction (Frobenius)", f"{risk_stats.get('psd_correction_frobenius', 0):.2f}")
        _psd_correction = float(risk_stats.get("psd_correction_frobenius", 0) or 0)
        if _psd_correction > 1e-6:
            st.warning("⚠️ The pairwise covariance matrix required a material PSD correction. The canonical portfolio volatility is still computed directly from complete daily P&L observations.")

        # ── Clustered Correlation ─────────────────────────────────────────────────
        st.subheader("Clustered Correlation Heatmap")
        st.caption("Average-linkage clustering. Cluster assignment uses the monthly correlation matrix as the primary structural-diversification horizon; the heatmap shows daily correlation. Boxes mark pairs above the selected threshold.")
        st.plotly_chart(
            plot_clustered_correlation(corr_mat, systems, over_exp_thresh),
            use_container_width=True)

        _primary_upper = _primary_corr.where(np.triu(np.ones(_primary_corr.shape), k=1).astype(bool))
        high_corr = _primary_upper.stack()
        high_corr = high_corr[high_corr > over_exp_thresh].sort_values(ascending=False)
        if not high_corr.empty:
            with st.expander(f"⚠️ Over-exposed pairs (ρ > {over_exp_thresh:.2f})", expanded=False):
                hc_df = pd.DataFrame({
                    "System A":    [systems.get(i[0], {}).get("display_name", i[0])[:38]
                                    for i in high_corr.index],
                    "System B":    [systems.get(i[1], {}).get("display_name", i[1])[:38]
                                    for i in high_corr.index],
                    "Correlation": high_corr.values,
                    "Risk":        ["🔴 High" if v > 0.85 else "🟡 Watch"
                                    for v in high_corr.values],
                })
                render_table(hc_df, hide_index=True, use_container_width=True)

        # ── System risk contribution ──────────────────────────────────────────────
        st.subheader("Who Actually Drives Portfolio Risk?")
        rc_rows = []
        rc_pct = risk_stats.get("rc_pct", pd.Series(dtype=float))
        vols = risk_stats.get("vols", pd.Series(dtype=float))
        components = risk_stats.get("component", pd.Series(dtype=float))
        for stem in rc_pct.sort_values(ascending=False).index:
            rc_rows.append({
                "System": systems.get(stem, {}).get("display_name", stem)[:38],
                "Contract": systems.get(stem, {}).get("contract_label", ""),
                "Daily Vol": vols.get(stem, np.nan),
                "Risk Contribution": components.get(stem, np.nan),
                "Risk %": rc_pct.get(stem, np.nan),
                "Avg ρ": corr_mat.loc[stem].drop(labels=[stem], errors="ignore").mean(),
            })
        if rc_rows:
            render_table(pd.DataFrame(rc_rows), hide_index=True, use_container_width=True,
                         height=min(42 + 36 * len(rc_rows), 620))

        # ── Correlation clusters ──────────────────────────────────────────────────
        st.subheader("Correlation Clusters")
        cluster_rows = []
        for cl in sorted(set(cluster_labels.values())):
            members = [s for s, label in cluster_labels.items() if label == cl]
            cluster_rc = float(rc_pct.reindex(members).fillna(0).sum())
            cluster_vol = float(components.reindex(members).fillna(0).sum())
            pair_vals = []
            for i, a in enumerate(members):
                for b in members[i + 1:]:
                    v = corr_mat.loc[a, b]
                    if np.isfinite(v):
                        pair_vals.append(float(v))
            cluster_rows.append({
                "Cluster": f"Cluster {cl}",
                "Systems": len(members),
                "Avg Within-Cluster ρ": np.mean(pair_vals) if pair_vals else np.nan,
                "Risk Contribution": cluster_vol,
                "Risk %": cluster_rc,
                "Members": ", ".join(systems.get(m, {}).get("display_name", m)[:24] for m in members),
            })
        if cluster_rows:
            cluster_df = pd.DataFrame(cluster_rows).sort_values("Risk %", ascending=False)
            render_table(cluster_df, hide_index=True, use_container_width=True)

        with st.expander("ℹ️ How to read this", expanded=False):
            st.markdown(
                "**Daily Vol** measures each system on its own. **Risk Contribution** measures how much "
                "that system contributes to the combined portfolio after correlations. A system can have "
                "high standalone volatility but modest portfolio contribution if it diversifies the book — "
                "and the reverse is also true. Clusters show when several different strategies are effectively "
                "taking the same risk. This is the bridge to correlation-aware Risk Parity.")

        # ── Rolling diversification diagnostic ────────────────────────────────────
        with st.expander("📈 Rolling Correlation Diagnostic", expanded=False):
            roll_days = st.selectbox("Rolling window", [21, 63, 126], index=1,
                                     format_func=lambda x: {21: "1 month (21d)", 63: "3 months (63d)", 126: "6 months (126d)"}[x],
                                     key="corr_roll_window")
            if len(daily_r) >= roll_days + 5:
                roll_corr = daily_r.rolling(roll_days).corr()
                # Average off-diagonal correlation at each date.
                vals = []
                dates = []
                cols = daily_r.columns.tolist()
                for dt in daily_r.index:
                    try:
                        mat = roll_corr.loc[dt]
                    except KeyError:
                        continue
                    arr = mat.values.astype(float)
                    mask = np.triu(np.ones(arr.shape, dtype=bool), k=1)
                    x = arr[mask]
                    x = x[np.isfinite(x)]
                    if len(x):
                        vals.append(float(x.mean())); dates.append(dt)
                if vals:
                    fig_roll = go.Figure(go.Scatter(x=dates, y=vals, mode="lines", name="Avg ρ"))
                    fig_roll.add_hline(y=over_exp_thresh, line_dash="dash",
                                       line_color=C["red"], opacity=0.6,
                                       annotation_text=f"Threshold {over_exp_thresh:.2f}")
                    fig_roll.update_layout(template=THEME, height=300,
                                           margin=dict(l=0, r=0, t=10, b=0),
                                           yaxis_title="Average pairwise ρ",
                                           yaxis=dict(range=[-1, 1]))
                    st.plotly_chart(fig_roll, use_container_width=True, key="chart_corr_rolling")
                    metric_with_help(st, "Rolling Avg ρ", f"{vals[-1]:.2f}")
                else:
                    st.info("Not enough overlapping observations for the selected rolling window.")
            else:
                st.info("Not enough overlapping observations for the selected rolling window.")

        st.divider()

        # ── Correlation-aware Risk Parity ─────────────────────────────────────────
        st.subheader("🧠 Correlation-Aware Risk Parity")
        st.info(
            "**This is the next-generation Risk Parity.** It does not just equalise standalone volatility. "
            "It aims to equalise each system’s **contribution to total portfolio volatility**, including correlations.\n\n"
            "- **Volatility-only RP:** how risky is each system by itself?\n"
            "- **Correlation-aware RP:** how much risk does each system add to the whole book?\n"
            "- The result is scaled toward your target annualized portfolio risk and rounded to whole contracts.\n\n"
            "Two highly correlated systems may therefore receive less combined exposure than two independent systems."
        )
        _unit_daily = pd.DataFrame(index=daily_r.index)
        for stem in daily_r.columns:
            cur_n = st.session_state["sizing"].get(stem, systems[stem]["default_n"])
            # daily_r is already at current contract sizing; convert directly to one contract.
            _unit_daily[stem] = daily_r[stem] / cur_n if cur_n > 0 else np.nan
        _carp = compute_correlation_aware_rp(_unit_daily, account_balance, target_annual_risk, rp_max_ct)
        if _carp:
            _carp_n, _carp_rc = _carp["integer"], _carp["risk_pct"]
            _carp_vol, _carp_target = _carp["portfolio_vol"], _carp["target_daily"]
            _carp_ann = _carp_vol * np.sqrt(252) / account_balance if account_balance > 0 else 0.0
            cc = st.columns(5)
            metric_with_help(cc[0], "Target annual risk", f"{target_annual_risk:.0f}%")
            metric_with_help(cc[1], "Target daily risk", f"${_carp_target:,.0f}")
            metric_with_help(cc[2], "Correlation-aware RP Daily Risk", f"${_carp_vol:,.0f}")
            metric_with_help(cc[3], "Resulting annual risk", f"{_carp_ann:.1%}")
            metric_with_help(cc[4], "LW shrinkage", f"{_carp.get('shrinkage_intensity', np.nan):.1%}" if np.isfinite(_carp.get('shrinkage_intensity', np.nan)) else "n/a")
            rows=[]
            for stem in _carp_n.index:
                cur=st.session_state["sizing"].get(stem, systems[stem]["default_n"]); adv=int(_carp_n[stem])
                rows.append({"System":systems[stem]["display_name"][:38],"Contract":systems[stem]["contract_label"],
                             "Current N":cur,"Correlation-Aware RP N":adv,"Change":adv-cur,
                             "Unit Daily Vol":_carp["unit_vol"].get(stem,np.nan),"Risk %":_carp_rc.get(stem,np.nan)})
            render_table(pd.DataFrame(rows).sort_values("Risk %",ascending=False), hide_index=True, use_container_width=True, height=min(42+36*len(rows),620))
            with st.expander("ℹ️ How to read this", expanded=False):
                st.markdown("**The goal is not equal contracts.** The goal is more balanced **portfolio risk contribution**.\n\n"
                            "A system can have high standalone volatility but deserve more contracts if it diversifies the book. Conversely, a highly correlated system may deserve fewer contracts.\n\n"
                            "Because futures use whole contracts, the final integer allocation will not be perfectly equal-risk.")
        else:
            st.warning("Not enough valid covariance data to calculate correlation-aware Risk Parity.")

        st.divider()

        # ── Risk Parity ───────────────────────────────────────────────────────────
        st.subheader("⚖️ Risk Parity Sizing (Advisory)")
        st.info(
            "**What this section does:** it tries to give each system roughly the same *standalone* amount of daily P&L risk.\n\n"
            "For example, if one system moves about $250/day and another moves about $1,000/day, the first can receive more contracts and the second fewer contracts.\n\n"
            "**What it does NOT do yet:** it does not fully account for the fact that two systems may be highly correlated. Two systems can each have $500 of standalone risk and still create much more than $500 of combined portfolio risk if they tend to move together.\n\n"
            "**That is the next step:** correlation-aware Risk Parity will size the systems based on their contribution to the *whole portfolio*, not just their individual volatility.")
        st.caption(
            f"A sizing reference that answers: **how many contracts would target about ${rp_target:,.0f} of one-day P&L volatility per system?** "
            "It uses historical daily P&L volatility and is independent of the current Tab 1 sizing. "
            "It is not a margin requirement, stop-loss, or maximum-loss estimate. **v25.1 keeps the volatility-only RP as a benchmark and adds the correlation-aware advisory below.**")

        _rp_engine = portfolio_risk(systems, _active_stems, _current_sizing_map, lookback, window_days=126, min_overlap_days=252)
        _tab2_sig = st.session_state.get("_canonical_tab2_risk")
        if _tab2_sig and _tab2_sig.get("systems") == tuple(sorted(_active_stems)):
            _risk_diff = abs(float(_tab2_sig.get("vol", 0.0)) - float(_rp_engine.get("portfolio_daily_vol", 0.0)))
            if _risk_diff > 0.01:
                st.warning(f"⚠️ Canonical risk consistency check failed: Tab 2 vs Tab 3 current daily risk differ by ${_risk_diff:,.2f} for the same system set/window.")
            else:
                st.caption("✅ Canonical risk check: Tab 2 and Tab 3 current daily risk agree for the same system set/window.")
        elif _tab2_sig:
            st.caption("ℹ️ Canonical risk check: Tab 2 is filtered to a different system set, so its current-risk figure is intentionally different from the all-active Tab 3 figure.")
        _rp_current_risk = {stem: float(_rp_engine.get("vols", pd.Series()).get(stem, 0.0)) for stem in _active_stems}
        _rp_total_risk = float(_rp_engine.get("portfolio_daily_vol", 0.0))
        _rp_sum_individual_risk = float(_rp_engine.get("sum_individual_vol", 0.0))
        _advisory_sizing_map = {stem: int(_rp_sizing.get(stem, 0)) for stem in _active_stems}
        _rp_advisory_engine = portfolio_risk(systems, _active_stems, _advisory_sizing_map, lookback, min_overlap_days=252)
        _rp_advisory_risk = float(_rp_advisory_engine.get("portfolio_daily_vol", 0.0))
        _rp_current_risk_pct = _rp_total_risk / account_balance if account_balance > 0 else 0.0
        _rp_advisory_risk_pct = _rp_advisory_risk / account_balance if account_balance > 0 else 0.0
        _rp_current_ann_pct = _rp_total_risk * np.sqrt(252) / account_balance if account_balance > 0 else 0.0
        _rp_advisory_ann_pct = _rp_advisory_risk * np.sqrt(252) / account_balance if account_balance > 0 else 0.0
        rc1, rc2, rc3, rc4, rc5 = st.columns(5)
        metric_with_help(rc1, "Current daily risk", f"${_rp_total_risk:,.0f}",
                   help="""**Formula:** standard deviation of combined daily portfolio P&L at the current contract sizes, including correlations.

    **Purpose:** measure actual diversified portfolio volatility.

    **In plain English:** how much does the whole book normally wobble by?

    - The sum of standalone volatilities is shown separately; it deliberately gives no diversification credit.""")
        metric_with_help(rc2, "Current risk / equity", f"{_rp_current_risk_pct:.2%}")
        metric_with_help(rc3, "RP advisory daily risk", f"${_rp_advisory_risk:,.0f}",
                   help="Estimated daily P&L volatility if the advisory contract counts were used.")
        metric_with_help(rc4, "RP risk / equity", f"{_rp_advisory_risk_pct:.2%}")
        metric_with_help(rc5, "Current → RP annual risk", f"{_rp_current_ann_pct:.0%} → {_rp_advisory_ann_pct:.0%}",
                   help="Annualized P&L volatility divided by the account balance. This is a risk-leverage proxy, not notional futures leverage.")
        st.caption(f"Standalone-volatility sum: **${_rp_sum_individual_risk:,.0f}/day**. This is deliberately not used as portfolio risk because it gives the systems no diversification credit.")

        st.caption(
            f"With a **${account_balance:,.0f}** account, current sizing implies about "
            f"**{_rp_current_ann_pct:.0%} annualized P&L volatility**; the RP advisory sizing "
            f"implies about **{_rp_advisory_ann_pct:.0%}**. In banker dialect: current book is "
            f"running at {(_rp_current_ann_pct/_rp_advisory_ann_pct if _rp_advisory_ann_pct else 0):.1f}× "
            "the advisory risk budget. Notional leverage and margin leverage are deliberately "
            "not inferred here.")

        rp_rows = []
        for stem in _active_stems:
            si = systems[stem]
            d  = _rp_detail.get(stem, {})
            rp_rows.append({
                "System":        si["display_name"][:35],
                "Symbol":        si["symbol"],
                "Contract":      si["contract_label"],
                "Vol / contract / day": d.get("vol_1ct", 0),
                "Current Daily Risk": d.get("vol_1ct", 0) * st.session_state["sizing"].get(stem, si["default_n"]),
                "RP Advisory N": _rp_sizing.get(stem, si["default_n"]),
                "Current N":     st.session_state["sizing"].get(stem, si["default_n"]),
                "Risk / Equity": (d.get("vol_1ct", 0) * st.session_state["sizing"].get(stem, si["default_n"]) / account_balance),
                "Advisory Risk / Equity": (d.get("vol_1ct", 0) * _rp_sizing.get(stem, si["default_n"]) / account_balance),
                "Action": (
                    (lambda cur, adv: f"Increase +{adv-cur}" if adv > cur else (f"Reduce {adv-cur}" if adv < cur else "Maintain"))
                    (st.session_state["sizing"].get(stem, si["default_n"]), _rp_sizing.get(stem, si["default_n"]))
                ),
            })
        rp_df = pd.DataFrame(rp_rows)
        render_table(
            rp_df,
            hide_index=True, use_container_width=True)
        st.caption(
            "**Action** shows the exact contract change required: **Increase +2**, **Reduce -1**, "
            "or **Maintain**. Implementing all rows is simulated in the Portfolio Impact section below "
            "before you change anything in the Systems tab.")

        # ── Portfolio impact of implementing the full advisory sizing ────────────
        st.subheader("📈 Portfolio Impact — If You Implemented All RP Changes")
        st.caption(
            "A before/after view using the same systems and date window. It shows what the portfolio "
            "would have looked like if every Action above had been implemented simultaneously. "
            "This is a historical sizing simulation, not a forecast or a margin calculation.")

        advisory_curves = {}
        for stem in _active_stems:
            si = systems[stem]
            adv_n = _rp_sizing.get(stem, si["default_n"])
            if si["default_n"] <= 0 or adv_n <= 0:
                continue
            adv_ratio = adv_n / si["default_n"]
            adv_eq = get_net_equity_trimmed(si, lookback, adv_ratio)
            if not adv_eq.empty:
                advisory_curves[stem] = adv_eq

        if advisory_curves:
            adv_df = combine_equity_curves(advisory_curves, lookback_years=lookback)
            adv_df = filter_equity_by_date_range(adv_df, start_date, end_date)
            if not adv_df.empty:
                adv_df = adv_df - adv_df.iloc[0]
                advisory_eq = adv_df.sum(axis=1)
                advisory_pm = compute_portfolio_metrics(advisory_eq)
                _impact_current_sizing = {stem: st.session_state["sizing"].get(stem, systems[stem]["default_n"]) for stem in selected_systems}
                _impact_advisory_sizing = {stem: int(_rp_sizing.get(stem, _impact_current_sizing[stem])) for stem in _impact_current_sizing}
                _impact_current = portfolio_risk(systems, selected_systems, _impact_current_sizing, lookback, date_start=start_date, date_end=end_date, min_overlap_days=252)
                _impact_advisory = portfolio_risk(systems, selected_systems, _impact_advisory_sizing, lookback, date_start=start_date, date_end=end_date, min_overlap_days=252)
                current_daily = float(_impact_current.get("portfolio_daily_vol", 0.0))
                advisory_daily = float(_impact_advisory.get("portfolio_daily_vol", 0.0))
                current_ann_risk_pct = current_daily * np.sqrt(252) / account_balance if account_balance > 0 else 0
                advisory_ann_risk_pct = advisory_daily * np.sqrt(252) / account_balance if account_balance > 0 else 0

                impact_rows = [
                    {"Metric": "Account balance", "Current sizing": account_balance, "RP advisory sizing": account_balance, "Change": 0.0},
                    {"Metric": "Net P&L", "Current sizing": pm.get("Net Profit ($)", 0), "RP advisory sizing": advisory_pm.get("Net Profit ($)", 0), "Change": advisory_pm.get("Net Profit ($)", 0) - pm.get("Net Profit ($)", 0)},
                    {"Metric": "Max Drawdown", "Current sizing": pm.get("Max Drawdown ($)", 0), "RP advisory sizing": advisory_pm.get("Max Drawdown ($)", 0), "Change": advisory_pm.get("Max Drawdown ($)", 0) - pm.get("Max Drawdown ($)", 0)},
                    {"Metric": "Daily P&L risk", "Current sizing": current_daily, "RP advisory sizing": advisory_daily, "Change": advisory_daily - current_daily},
                    {"Metric": "Annualized risk / equity", "Current sizing": current_ann_risk_pct, "RP advisory sizing": advisory_ann_risk_pct, "Change": advisory_ann_risk_pct - current_ann_risk_pct},
                    {"Metric": "P&L Sharpe", "Current sizing": pm.get("P&L Sharpe", 0), "RP advisory sizing": advisory_pm.get("P&L Sharpe", 0), "Change": advisory_pm.get("P&L Sharpe", 0) - pm.get("P&L Sharpe", 0)},
                    {"Metric": "Ann. P&L / Max DD", "Current sizing": pm.get("Ann. P&L / Max DD", 0), "RP advisory sizing": advisory_pm.get("Ann. P&L / Max DD", 0), "Change": advisory_pm.get("Ann. P&L / Max DD", 0) - pm.get("Ann. P&L / Max DD", 0)},
                ]
                # This comparison intentionally uses display strings because the same column
                # contains dollars, percentages and ratios depending on the row.
                def _impact_fmt(metric, value):
                    if metric in {"Account balance", "Net P&L", "Max Drawdown", "Daily P&L risk"}:
                        return _fmt_money(value)
                    if metric == "Annualized risk / equity":
                        return f"{value:.1%}"
                    return f"{value:.2f}"

                impact_display = pd.DataFrame([
                    {
                        "Metric": r["Metric"],
                        "Current": _impact_fmt(r["Metric"], r["Current sizing"]),
                        "RP Advisory": _impact_fmt(r["Metric"], r["RP advisory sizing"]),
                        "Change": (
                            (f"{r['Change']:+.1%}" if r["Metric"] == "Annualized risk / equity"
                             else f"{r['Change']:+.2f}" if r["Metric"] in {"P&L Sharpe", "Ann. P&L / Max DD"}
                             else _fmt_money(r["Change"]))
                        ),
                    }
                    for r in impact_rows
                ])
                render_table(impact_display, hide_index=True, use_container_width=True)

                imp1, imp2, imp3 = st.columns(3)
                imp1.metric("Daily risk", f"${current_daily:,.0f} → ${advisory_daily:,.0f}",
                             help="Historical one-day P&L volatility under current sizing versus the full Risk Parity Advisory sizing.")
                imp2.metric("Annual risk / equity", f"{current_ann_risk_pct:.1%} → {advisory_ann_risk_pct:.1%}",
                             help="Annualized P&L volatility divided by account balance, comparing current and advisory sizing.")
                risk_change = advisory_ann_risk_pct - current_ann_risk_pct
                imp3.metric("Risk change", f"{risk_change:+.1%}",
                             help="Change in annualized portfolio risk relative to account equity if all advisory sizing changes were implemented.")

        st.divider()

        # ── Recommended Sub-Portfolios ────────────────────────────────────────────
        st.subheader("🎯 Recommended Sub-Portfolios")
        st.caption("Subsets selected by different risk/return criteria.")

        reccos = recommend_portfolios(corr_df, systems)
        for rec_i, rec in enumerate(reccos):
            with st.expander(f"{rec['name']} — {rec['description']}"):
                rec_stems = [s for s in rec["systems"] if s in all_curves]
                if not rec_stems:
                    st.write("Insufficient data.")
                    continue
                ratios_rec = {s: _cur_ratio(s) for s in rec_stems}
                rec_eq     = build_portfolio_equity(systems, rec_stems, lookback, ratios_rec)
                rec_m      = compute_portfolio_metrics(rec_eq)
                mc4 = st.columns(4)
                for col, (lbl, key, fmt) in zip(mc4, [
                    ("Net P&L", "Net Profit ($)", "$"), ("P&L Sharpe", "P&L Sharpe", "f"),
                    ("Max DD",  "Max Drawdown ($)", "$"), ("Ann. P&L / DD", "Ann. P&L / Max DD", "f"),
                ]):
                    v = rec_m.get(key, 0)
                    metric_with_help(col, lbl, f"${v:,.0f}" if fmt == "$" else f"{v:.2f}")
                if not rec_eq.empty:
                    norm = rec_eq - rec_eq.iloc[0]
                    fig_r = go.Figure(go.Scatter(x=norm.index, y=norm.values,
                                                  line=dict(color=C["blue"], width=2)))
                    fig_r.update_layout(template=THEME, height=320,
                                         margin=dict(l=0, r=0, t=10, b=0),
                                         yaxis_title="Cum. P&L ($)")
                    st.plotly_chart(fig_r, use_container_width=True, key=f"chart_4_{rec_i}")

        st.divider()

        # ── Risk / Return scatter ─────────────────────────────────────────────────
        st.subheader("Risk / Return Scatter")
        scatter_rows = []
        for stem in corr_df.columns:
            d   = daily_r[stem]
            col = corr_df[stem]
            net   = col.iloc[-1] - col.iloc[0]
            n_yr  = max((col.index[-1] - col.index[0]).days / 365.25, 0.01)
            ann_r = net / n_yr
            sharpe_v = (d.mean() / d.std() * np.sqrt(252)) if d.std() > 0 else 0
            max_dd_v = (col - col.cummax()).min()
            scatter_rows.append({
                "System":         systems.get(stem, {}).get("display_name", stem)[:32],
                "Symbol":         systems.get(stem, {}).get("symbol", ""),
                "Sharpe":         round(float(sharpe_v), 2),
                "Max DD ($)":     round(float(max_dd_v), 0),
                "Ann Return ($)": round(float(ann_r), 0),
                "_sz":            max(abs(float(ann_r)), 1.0),
            })
        if scatter_rows:
            sc_df  = pd.DataFrame(scatter_rows)
            fig_sc = px.scatter(sc_df, x="Max DD ($)", y="Sharpe",
                                text="System", color="Symbol",
                                size="_sz", size_max=42, template=THEME,
                                title="P&L Sharpe vs Max Drawdown  (bubble = |Ann. P&L|)",
                                hover_data={"Ann Return ($)": True, "_sz": False})
            fig_sc.update_traces(textposition="top center")
            fig_sc.update_layout(height=520, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig_sc, use_container_width=True, key="chart_5")


    # ═════════════════════════════════════════════════════════════════════════════
    # TAB 4 — FORWARD ANALYSIS
    # ═════════════════════════════════════════════════════════════════════════════

_render_tab3()

with tab4:

    # ── System Health Monitor ─────────────────────────────────────────────────
    st.subheader("🏥 System Health Monitor")
    st.caption("Rolling metrics using ReadMe default sizing (independent of Tab 1 overrides).")

    hw_choice = st.radio("Rolling window", ["3 months (63d)", "6 months (126d)"],
                          horizontal=True, key="hw")
    hw_days   = 63 if "3" in hw_choice else 126

    health_rows   = []
    health_charts = {}
    for stem in _active_stems:
        si  = systems[stem]
        hdf = compute_rolling_health(si["trades"], si["comm_per_trade"], hw_days)
        if hdf.empty:
            health_rows.append({
                "System": si["display_name"][:35], "Status": "⚪",
                "Sharpe": np.nan, "PF": np.nan, "Win %": np.nan,
                "Period P&L": np.nan})
            continue
        health_charts[stem] = hdf
        lt = hdf.iloc[-1]
        tl = health_traffic_light(lt["Sharpe"], lt["Profit Factor"], lt["Win Rate"])
        health_rows.append({
            "System":     si["display_name"][:35],
            "Status":     tl,
            "Sharpe":     lt["Sharpe"],
            "PF":         lt["Profit Factor"],
            "Win %":      lt["Win Rate"],
            "Period P&L": lt["Period P&L"],
        })

    health_df = pd.DataFrame(health_rows)
    # coerce numeric columns to preserve numeric sorting while formatting
    for c in ["Sharpe", "PF", "Win %", "Period P&L", "# Trades"]:
        if c in health_df.columns:
            health_df[c] = pd.to_numeric(health_df[c], errors="coerce")

    render_table(
        health_df,
        hide_index=True,
        use_container_width=True,
        height=min(42 + 36 * len(health_rows), 520))

    if health_charts:
        with st.expander("📈 Rolling Sharpe over time"):
            fig_rsh = go.Figure()
            for i, (stem, hdf) in enumerate(health_charts.items()):
                fig_rsh.add_trace(go.Scatter(
                    x=hdf.index, y=hdf["Sharpe"].values,
                    name=systems[stem]["display_name"][:25],
                    line=dict(color=COLORS[i % len(COLORS)], width=1.5),
                    hovertemplate="%{x|%Y-%m-%d}<br>Sharpe:%{y:.2f}<extra></extra>"))
            fig_rsh.add_hline(y=0,   line_dash="dot",  line_color=C["zero_line"])
            fig_rsh.add_hline(y=0.8, line_dash="dash", line_color=C["green"],
                              opacity=0.5, annotation_text="Good (0.8)")
            fig_rsh.update_layout(template=THEME, height=420,
                                   margin=dict(l=0, r=0, t=10, b=0),
                                   yaxis_title="Rolling Sharpe",
                                   legend=dict(font=dict(size=10)))
            st.plotly_chart(fig_rsh, use_container_width=True, key="chart_6")

        with st.expander("📈 Rolling Profit Factor over time"):
            fig_rpf = go.Figure()
            for i, (stem, hdf) in enumerate(health_charts.items()):
                if "Profit Factor" in hdf.columns:
                    fig_rpf.add_trace(go.Scatter(
                        x=hdf.index, y=hdf["Profit Factor"].values,
                        name=systems[stem]["display_name"][:25],
                        line=dict(color=COLORS[i % len(COLORS)], width=1.5)))
            fig_rpf.add_hline(y=1.0, line_dash="dot", line_color=C["red"],
                              annotation_text="Break-even (1.0)")
            fig_rpf.update_layout(template=THEME, height=420,
                                   margin=dict(l=0, r=0, t=10, b=0),
                                   yaxis_title="Rolling PF",
                                   legend=dict(font=dict(size=10)))
            st.plotly_chart(fig_rpf, use_container_width=True, key="chart_7")

    st.divider()

    # ── Monte Carlo ───────────────────────────────────────────────────────────
    st.subheader("🎲 Monte Carlo Forward Projection")
    st.caption(
        "Projects the current portfolio by resampling historical daily P&L. "
        "The simulation is a scenario tool, not a forecast."
    )

    mc_info = st.expander("ℹ️ How to read this", expanded=False)
    with mc_info:
        st.markdown(
            "**What it does**\n\n"
            "- Takes the historical **portfolio daily P&L** at your current sizing.\n"
            "- Randomly rearranges historical observations to create many possible futures.\n"
            "- Reports the distribution of possible P&L and drawdowns.\n\n"
            "**Why the block bootstrap matters**\n\n"
            "- **IID bootstrap** treats every day as independent.\n"
            "- **Block bootstrap** keeps short runs of consecutive days together, "
            "which better preserves streaks and volatility clustering.\n\n"
            "**Important:** Monte Carlo does not know what the market will do next. "
            "It asks: *if the future behaves broadly like the selected historical sample, "
            "what range of outcomes is plausible?*"
        )

    mc_c1, mc_c2, mc_c3 = st.columns(3)
    with mc_c1:
        mc_sims = st.selectbox("Simulations", [500, 1000, 2000, 5000], index=1,
                               key="mc_sims_v24")
    with mc_c2:
        mc_days_opt = st.selectbox(
            "Horizon", [63, 126, 252],
            format_func=lambda x: {63: "3 months", 126: "6 months", 252: "12 months"}[x],
            index=1, key="mc_days_v24")
    with mc_c3:
        mc_method = st.selectbox(
            "Method",
            ["Regime-conditioned bootstrap", "Block bootstrap", "IID bootstrap"],
            index=0, key="mc_method_v25")

    mc_block_days = 5
    if mc_method == "Block bootstrap":
        mc_block_days = st.selectbox(
            "Block length", [3, 5, 10, 20], index=1,
            format_func=lambda x: f"{x} trading days", key="mc_block_v25")

    mc_regime_window, mc_low_pct, mc_high_pct = 20, 33, 66
    if mc_method == "Regime-conditioned bootstrap":
        rc1, rc2, rc3 = st.columns(3)
        with rc1:
            mc_regime_window = st.selectbox("Regime vol window", [10, 15, 20, 30, 40], index=2, key="mc_regime_window_v25")
        with rc2:
            mc_low_pct = st.slider("Low-vol percentile", 20, 45, 33, key="mc_low_pct_v25")
        with rc3:
            mc_high_pct = st.slider("High-vol percentile", 55, 80, 66, key="mc_high_pct_v25")
        st.caption(
            "**Regime-conditioned bootstrap:** first simulates whether the portfolio stays in Low / Medium / High Vol, "
            "then samples historical daily P&L from that regime. This captures volatility-state persistence; it does not predict the future.")

    if st.button("🎲 Run Monte Carlo", type="primary", key="run_mc_v25"):
        with st.spinner(f"Running {mc_sims:,} simulations…"):
            ratios_mc = {s: _cur_ratio(s) for s in _active_stems}
            mc_res = monte_carlo_simulation(
                systems, _active_stems, lookback,
                n_sims=mc_sims, forward_days=mc_days_opt,
                sizing_ratios=ratios_mc,
                method=mc_method, block_days=mc_block_days,
                regime_window=mc_regime_window, low_pct=mc_low_pct, high_pct=mc_high_pct,
                account_balance=account_balance)

        if mc_res is None:
            st.warning("Not enough valid historical portfolio data for the selected lookback.")
        else:
            pct = mc_res["percentiles"]
            days_x = list(range(1, len(pct) + 1))
            hor_lbl = {63: "3 months", 126: "6 months", 252: "12 months"}.get(mc_days_opt, "")

            fig_mc = go.Figure()
            fig_mc.add_trace(go.Scatter(x=days_x, y=pct["95th"].values,
                                        mode="lines", line=dict(width=0), showlegend=False))
            fig_mc.add_trace(go.Scatter(x=days_x, y=pct["5th"].values,
                                        mode="lines", line=dict(width=0), fill="tonexty",
                                        fillcolor="rgba(38,139,210,0.12)", name="5–95th pct"))
            fig_mc.add_trace(go.Scatter(x=days_x, y=pct["75th"].values,
                                        mode="lines", line=dict(width=0), showlegend=False))
            fig_mc.add_trace(go.Scatter(x=days_x, y=pct["25th"].values,
                                        mode="lines", line=dict(width=0), fill="tonexty",
                                        fillcolor="rgba(38,139,210,0.22)", name="25–75th pct"))
            fig_mc.add_trace(go.Scatter(x=days_x, y=pct["50th"].values,
                                        name="Median", line=dict(color=C["blue"], width=2.5)))
            fig_mc.add_hline(y=0, line_dash="dot", line_color=C["zero_line"])
            fig_mc.update_layout(template=THEME, height=450,
                                 margin=dict(l=0, r=0, t=30, b=0),
                                 title=f"Monte Carlo: {mc_sims:,} × {hor_lbl} — {mc_method}",
                                 xaxis_title="Trading Days Forward",
                                 yaxis_title="Projected P&L ($)")
            st.plotly_chart(fig_mc, use_container_width=True, key="chart_mc_v25")

            final = mc_res["paths"].iloc[-1]
            ec = st.columns(7)
            metric_with_help(ec[0], "Median", f"${final.median():,.0f}")
            metric_with_help(ec[1], "Mean", f"${final.mean():,.0f}")
            metric_with_help(ec[2], "5th pct", f"${final.quantile(0.05):,.0f}")
            metric_with_help(ec[3], "95th", f"${final.quantile(0.95):,.0f}")
            metric_with_help(ec[4], "P(loss)", f"{(final < 0).mean():.1%}")
            metric_with_help(ec[5], "5% Terminal ES", f"${mc_res['terminal_es_5']:,.0f}")
            metric_with_help(ec[6], "Historical days", f"{mc_res['historical_days']:,}")

            if mc_method == "Regime-conditioned bootstrap" and "current_regime" in mc_res:
                st.markdown("**Starting volatility regime**")
                rg1, rg2 = st.columns(2)
                metric_with_help(rg1, "Current Vol Regime", mc_res["current_regime"])
                metric_with_help(rg2, "Regime Persistence", f"{mc_res['regime_persistence']:.1%}")

            st.markdown("**Maximum drawdown probabilities**")
            dd_rows = [{
                "Drawdown Threshold": f"{t:.1%} of account",
                "Probability": p,
                "Dollar Threshold": account_balance * t,
            } for t, p in sorted(mc_res["dd_probs"].items())]
            render_table(pd.DataFrame(dd_rows), hide_index=True,
                         use_container_width=True)

            st.caption(
                f"Simulation: **{mc_method}**; "
                f"{mc_res['historical_days']:,} historical daily observations; "
                f"{mc_sims:,} paths × {hor_lbl}. "
                "Block bootstrap preserves short sequences of historical days. "
                "Regime-conditioned bootstrap also models empirical volatility-regime persistence. "
                "Neither method predicts future regimes."
            )

    st.divider()

    # ── Sizing Recommendation Summary ─────────────────────────────────────────
    st.subheader("📋 Sizing Recommendation Summary")
    st.caption("Based on rolling health metrics (3-month window). Not financial advice.")

    portfolio_exposure = {
        stem: abs(eq_df[stem].iloc[-1]) if stem in eq_df.columns else 0.0
        for stem in _active_stems
    }
    exposure_sum = sum(portfolio_exposure.values()) or 1.0
    avg_corr = pd.Series({
        stem: (corr_df.loc[stem, _active_stems].drop(labels=[stem], errors='ignore').mean()
               if stem in corr_df.index else np.nan)
        for stem in _active_stems
    })
    regime_df, _ = compute_regime_series({stem: eq_df[stem] for stem in _active_stems if stem in eq_df},
                                         vol_window=20, low_pct=33, high_pct=66)
    cur_regime = regime_df["regime"].iloc[-1] if not regime_df.empty else "Unknown"

    rec_rows = []
    for stem in _active_stems:
        si  = systems[stem]
        nd  = si["default_n"]
        hdf = compute_rolling_health(si["trades"], si["comm_per_trade"], 63)
        cur_n = st.session_state["sizing"].get(stem, nd)
        weight = portfolio_exposure.get(stem, 0.0) / exposure_sum
        corr  = avg_corr.get(stem, np.nan)

        if hdf.empty:
            rec_rows.append({
                "System":        si["display_name"][:35],
                "Contract":      si["contract_label"],
                "Current N":     cur_n,
                "Suggested N":   cur_n,
                "Scale":         "1.0x",
                "Weight":        weight,
                "Avg Corr":      corr,
                "Health":        "⚪ No data",
                "Suggestion":     "Insufficient data",
                "Portfolio Regime": cur_regime,
            })
            continue

        lt  = hdf.iloc[-1]
        sh, pf_v, wr = lt["Sharpe"], lt["Profit Factor"], lt["Win Rate"]
        tl  = health_traffic_light(sh, pf_v, wr)

        suggested_n = cur_n
        if tl == "🟢":
            if sh > 1.5 and pf_v > 1.5 and corr < 0.75 and cur_regime != "High Vol":
                sugg = "Scale up — strong momentum + diversified"
                suggested_n = max(cur_n + 1, int(np.ceil(cur_n * 1.2)))
            else:
                sugg = "Maintain — strong core holding"
        elif tl == "🟡":
            if weight > 0.20 or corr > 0.75 or cur_regime == "High Vol":
                sugg = "Reduce exposure — watch portfolio concentration"
                suggested_n = max(cur_n - 1, 1)
            else:
                sugg = "Hold — monitor correlation and regime"
        else:
            sugg = "Reduce or pause — weak health"
            suggested_n = max(cur_n - 1, 1)

        if cur_regime == "High Vol" and tl == "🟢":
            sugg += " (volatility regime is high)"
        if corr > 0.8:
            sugg += " (high correlation to portfolio)"

        scale_label = f"{suggested_n / cur_n:.1f}x" if cur_n else "—"
        rec_rows.append({
            "System":        si["display_name"][:35],
            "Contract":      si["contract_label"],
            "Current N":     cur_n,
            "Suggested N":   suggested_n,
            "Scale":         scale_label,
            "Weight":        weight,
            "Avg Corr":      corr,
            "Health":        f"{tl} Sh={sh:.1f} PF={pf_v:.1f}",
            "Suggestion":     sugg,
            "Portfolio Regime": cur_regime,
        })

    rec_df = pd.DataFrame(rec_rows)
    # coerce numeric columns so sorting works; display via Styler
    for c in ["Current N", "Suggested N", "Weight", "Avg Corr"]:
        if c in rec_df.columns:
            rec_df[c] = pd.to_numeric(rec_df[c], errors="coerce")

    render_table(
        rec_df,
        hide_index=True,
        use_container_width=True,
        height=min(42 + 36 * len(rec_rows), 520))
    st.info("⚠️ Suggestions based on rolling historical metrics. "
            "Not financial advice. Past performance ≠ future results.")

    st.divider()

    # ── Regime Analysis ───────────────────────────────────────────────────────
    st.subheader("🌡️ Regime Analysis")
    st.caption("Classifies market conditions via portfolio realized volatility.")

    reg_c1, reg_c2, reg_c3 = st.columns(3)
    with reg_c1:
        reg_vw = st.selectbox("Vol window (days)", [10, 15, 20, 30, 40], index=2, key="r_vw")
    with reg_c2:
        reg_lp = st.slider("Low-vol pct", 10, 45, 33, key="r_lp")
    with reg_c3:
        reg_hp = st.slider("High-vol pct", 55, 90, 66, key="r_hp")

    _reg_curves = {}
    for stem in _active_stems:
        si = systems[stem]
        eq = get_net_equity_trimmed(si, lookback, _cur_ratio(stem))
        if not eq.empty:
            _reg_curves[stem] = eq

    if len(_reg_curves) < 2:
        st.info("Need ≥ 2 active systems.")
    else:
        regime_df, regime_thresholds = compute_regime_series(
            _reg_curves, vol_window=reg_vw, low_pct=reg_lp, high_pct=reg_hp)

        if regime_df.empty:
            st.warning("Not enough data for regime classification.")
        else:
            cur_regime = regime_df["regime"].iloc[-1]
            reg_emoji  = {"Low Vol": "😴", "Medium Vol": "😐", "High Vol": "🔥"}
            st.markdown(
                f"### Current Regime: {reg_emoji.get(cur_regime, '')} **{cur_regime}**")

            rm1, rm2, rm3, rm4 = st.columns(4)
            rm1.metric("Current Ann. Vol", f"${regime_df['vol'].iloc[-1]:,.0f}")
            rm2.metric("Thresholds",
                       f"${regime_thresholds['low']:,.0f} / ${regime_thresholds['high']:,.0f}")
            rc  = regime_df["regime"].value_counts()
            tot = len(regime_df)
            rm3.metric("% Low Vol",  f"{rc.get('Low Vol',  0)/tot*100:.0f}%")
            rm4.metric("% High Vol", f"{rc.get('High Vol', 0)/tot*100:.0f}%")

            rcmap = {"Low Vol": C["green"], "Medium Vol": C["amber"], "High Vol": C["red"]}
            fig_reg = go.Figure()
            for label, color in rcmap.items():
                mask = regime_df["regime"] == label
                sub  = regime_df[mask]
                if not sub.empty:
                    fig_reg.add_trace(go.Scatter(
                        x=sub.index, y=sub["vol"].values, mode="markers",
                        marker=dict(color=color, size=3, opacity=0.6), name=label))
            fig_reg.add_hline(y=regime_thresholds["low"],  line_dash="dash",
                              line_color=C["green"], opacity=0.5)
            fig_reg.add_hline(y=regime_thresholds["high"], line_dash="dash",
                              line_color=C["red"], opacity=0.5)
            fig_reg.update_layout(template=THEME, height=300,
                                   margin=dict(l=0, r=0, t=10, b=0),
                                   yaxis_title="Ann. Vol ($)",
                                   legend=dict(orientation="h", y=1.05))
            st.plotly_chart(fig_reg, use_container_width=True, key="chart_9")

            with st.expander("🔄 Regime Transition Probabilities"):
                trans_df = compute_regime_transitions(regime_df)
                if not trans_df.empty:
                    st.caption("P(next = column | today = row)")
                    fig_t = go.Figure(go.Heatmap(
                        z=trans_df.values,
                        x=trans_df.columns.tolist(), y=trans_df.index.tolist(),
                        text=[[f"{v:.1%}" for v in row] for row in trans_df.values],
                        texttemplate="%{text}",
                        colorscale="Blues", zmin=0, zmax=1,
                        colorbar=dict(title="P")))
                    fig_t.update_layout(template=THEME, height=280,
                                         margin=dict(l=0, r=0, t=10, b=0))
                    st.plotly_chart(fig_t, use_container_width=True, key="chart_10")
                    persist = {r: trans_df.loc[r, r] for r in trans_df.index}
                    st.markdown(
                        f"**Persistence:** Low Vol stays {persist.get('Low Vol',0):.0%} · "
                        f"Medium stays {persist.get('Medium Vol',0):.0%} · "
                        f"High Vol stays {persist.get('High Vol',0):.0%}")

            st.markdown("##### System Performance by Regime")
            rp_table = []
            for stem in _active_stems:
                si   = systems[stem]
                cutoff = (si["equity"].index.max() - pd.Timedelta(days=int(lookback * 365))
                          if not si["equity"].empty else pd.Timestamp("2000-01-01"))
                tr_t = (si["trades"][si["trades"]["exit_date"] >= cutoff]
                        if not si["trades"].empty else si["trades"])
                rp = compute_system_regime_performance(
                    tr_t, regime_df, si["comm_per_trade"])
                for rl in ["Low Vol", "Medium Vol", "High Vol"]:
                    m = rp.get(rl, {})
                    if m.get("n_trades", 0) == 0:
                        continue
                    rp_table.append({
                        "System":    si["display_name"][:32],
                        "Regime":    rl,
                        "# Trades":  m["n_trades"],
                        "Total P&L": m["total_pnl"],
                        "Avg P&L":   m["avg_pnl"],
                        "Win Rate":  m["win_rate"],
                        "PF":        m["profit_factor"],
                    })
            if rp_table:
                rp_df = pd.DataFrame(rp_table)
                for c in ["Total P&L", "Avg P&L", "Win Rate", "PF", "# Trades"]:
                    if c in rp_df.columns:
                        rp_df[c] = pd.to_numeric(rp_df[c], errors="coerce")

                render_table(
                    rp_df,
                    hide_index=True, use_container_width=True,
                    height=min(42 + 36 * len(rp_table), 520))
            st.info("💡 Systems robust across all regimes are the strongest core holdings.")


# ═════════════════════════════════════════════════════════════════════════
# TAB 5 — TRADE MATCHING  (theoretical XLSX trades vs. actual TradeStation
# executions). Uses tradestation_matcher.py — see that module's docstring
# for the data-quality caveats this tab surfaces (order-history coverage
# gaps, no per-strategy tag on fills, continuous-contract price offsets).
# ═════════════════════════════════════════════════════════════════════════

with tab5:
    st.subheader("🔗 Trade Matching — Theoretical (XLSX) vs Actual (TradeStation)")
    st.caption(
        "Matches each system's theoretical trades against real executions from a "
        "TradeStation order-history export, within date/price tolerance. The monthly "
        "PDF statement is used only as an independent P&L/commission cross-check "
        "(it has no per-trade prices, so it can't be used for matching itself). "
        "Strategy trade files can be .xlsx or .csv (flat report export) — drop either "
        "into the Data directory set in the sidebar; if both exist for the same system, "
        "the .csv is used.")

    if tsm is None:
        st.error(
            "`tradestation_matcher.py` could not be imported "
            f"({_TSM_IMPORT_ERROR}). Make sure the file is in the same folder as "
            "app.py, and that `openpyxl`/`pdfplumber` are installed.")
    else:
        OVERRIDE_LOG_PATH = str(DATA_DIR / "trade_overrides_log.csv")

        with st.expander("⚙️ Matching configuration", expanded=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                order_hist_path = st.text_input(
                    "Order-history file OR folder path (.xlsx or .ods)",
                    value=str(DATA_DIR / "trade_data" / "order_history"),
                    key="tsm_order_hist_path",
                    help="TradeStation order/execution export (Entered, Filled/Canceled, "
                         "Symbol, Type, Qty Filled, Filled Price, Order Status, Commission, "
                         "Contract Exp Date). Both .xlsx and .ods exports are supported. "
                         "Point this at a FOLDER (e.g. trade_data/order_history/) to always "
                         "auto-use the most recently added file in it — no need to retype a "
                         "filename each time you drop in a fresh export. Or point it at one "
                         "specific file if you want to pin an exact export.")
            with c2:
                stmt_pdf_path = st.text_input(
                    "Monthly statement PDF path (optional)",
                    value=str(DATA_DIR / "trade_data" / "statements"),
                    key="tsm_pdf_path",
                    help="Used only for the reconciliation check, not trade matching. "
                         "Point this at a FOLDER (e.g. trade_data/statements/) and it will "
                         "pick the PDF whose filename contains the selected year+month "
                         "(TradeStation's own naming, e.g. '..._202601.pdf' for Jan 2026). "
                         "If no filename matches, it falls back to the newest PDF in the "
                         "folder and flags that in the reconciliation section. Or point it "
                         "at one specific file if you want to pin an exact statement.")
            with c3:
                cM, cY = st.columns(2)
                with cM:
                    match_month = st.selectbox("Month", list(range(1, 13)),
                                                index=datetime.now().month - 1,
                                                format_func=lambda m: datetime(2000, m, 1).strftime("%b"),
                                                key="tsm_month")
                with cY:
                    match_year = st.number_input("Year", value=datetime.now().year,
                                                  min_value=2015, max_value=2100, step=1,
                                                  key="tsm_year")

            c4, c5 = st.columns(2)
            with c4:
                date_tol = st.slider("Time tolerance (± minutes)", 0, 60, 10, key="tsm_date_tol",
                                      help="Theoretical entry/exit timestamps are candle-CLOSE "
                                           "times, so the real execution can lag by up to one "
                                           "candle (~2.5-5 min on a 5-min ORB signal, confirmed "
                                           "on real June 2026 fills). Theoretical times are also "
                                           "shifted +7h automatically to align with Chicago fill "
                                           "time — this slider only covers the remaining small "
                                           "execution lag, not timezone. Was a 1-business-day "
                                           "window before; that was masking the (now-fixed) "
                                           "timezone gap and, once fixed, became too wide — it let "
                                           "unrelated same-day fills collide as false matches.")
            with c5:
                tick_tol = st.slider("Price tolerance (± ticks)", 0, 10, 2, key="tsm_tick_tol")

            match_portfolio_tab = st.checkbox(
                "Apply current Tab 1 sizing + commission (match Portfolio tab numbers)",
                value=True, key="tsm_match_portfolio",
                help="When on, 'Theoretical P&L' is scaled by each system's current "
                     "Tab 1 sizing ratio and reduced by the sidebar's commission rate — "
                     "the exact same formula the Portfolio tab uses — so the totals "
                     "below are directly comparable to what Tab 2 shows for this month. "
                     "When off, 'Theoretical P&L' is the raw, unscaled number straight "
                     "from the trade files.")

            run_clicked = st.button("▶️ Run matching engine", type="primary", key="tsm_run_btn")

        if run_clicked:
            if not os.path.exists(order_hist_path):
                st.error(f"Order-history file not found: {order_hist_path}")
            else:
                pdf_arg = stmt_pdf_path if os.path.exists(stmt_pdf_path) else None
                if pdf_arg is None:
                    st.info("No statement PDF/folder found at that path — running without "
                             "the reconciliation cross-check.")
                sizing_ratios_arg   = None
                commission_rates_arg = None
                if match_portfolio_tab:
                    sizing_ratios_arg    = {s: _cur_ratio(s) for s in systems}
                    commission_rates_arg = {s: systems[s]["comm_per_trade"] for s in systems}
                with st.spinner("Matching theoretical trades to actual executions…"):
                    cfg = tsm.MatchConfig(date_tol_minutes=date_tol, tick_tol=tick_tol)
                    tsm_result = tsm.run_matching_engine(
                        str(DATA_DIR), _README_PATH, order_hist_path,
                        int(match_year), int(match_month), pdf_arg, cfg,
                        sizing_ratios_arg, commission_rates_arg)
                st.session_state["tsm_result"] = tsm_result

        tsm_result = st.session_state.get("tsm_result")

        if tsm_result is None:
            st.info("Set the file paths above and click **Run matching engine** to begin.")
        else:
            cov = tsm_result["coverage"]
            month_label = datetime(int(match_year), int(match_month), 1).strftime("%B %Y")

            # ── Data coverage banner — surfaces gaps rather than hiding them ──
            if cov["n"] == 0:
                st.error("Order-history file has no recognizable 'Filled' rows.")
            else:
                cov_lo, cov_hi = cov["min"], cov["max"]
                month_start = pd.Timestamp(int(match_year), int(match_month), 1)
                month_end = month_start + pd.offsets.MonthEnd(1)
                if cov_lo > month_start or cov_hi < month_end:
                    st.warning(
                        f"⚠️ Order-history file covers **{cov_lo:%Y-%m-%d} → {cov_hi:%Y-%m-%d}** "
                        f"({cov['n']:,} filled orders), which does not fully cover "
                        f"**{month_label}** ({month_start:%Y-%m-%d} → {month_end:%Y-%m-%d}). "
                        "Trades outside the covered window will show as unmatched/unreconciled "
                        "— that's a data gap, not a real discrepancy.")
                else:
                    st.success(
                        f"✅ Order-history file fully covers {month_label} "
                        f"({cov_lo:%Y-%m-%d} → {cov_hi:%Y-%m-%d}, {cov['n']:,} filled orders).")

            # ── Systems excluded entirely (not in ReadMe, or ReadMe size is 0) ──
            # Per Andrea: a system with no ReadMe entry or an explicit 0-contract
            # allocation isn't part of the real portfolio, so its trades should
            # count as zero -- NOT show up as large phantom "unmatched" P&L for a
            # system she isn't actually running. load_all_theoretical_trades()
            # drops these before matching even starts; shown here so an excluded
            # system is still visible (in case it's actually a rename/typo that
            # needs fixing in ReadMe.txt, not a system that should stay excluded).
            excluded = tsm_result.get("excluded_systems", [])
            if excluded:
                with st.expander(f"🚫 {len(excluded)} system file(s) excluded from all totals below"):
                    st.caption(
                        "Not in ReadMe.txt, or listed there with 0 contracts — treated as not "
                        "part of the real portfolio, so none of their trades count toward "
                        "Theoretical P&L / Unmatched Impact / Portfolio Totals. If one of these "
                        "is actually an active system (e.g. a rename ReadMe.txt hasn't caught "
                        "up with), add/fix its ReadMe.txt line and it'll be picked up normally.")
                    render_table(pd.DataFrame(excluded), hide_index=True, use_container_width=True)

            # ── Portfolio totals — the number that actually answers "where did
            # the money go", but ONLY once (most of) the 18 systems are loaded ──
            n_loaded = tsm_result.get("n_systems_loaded", 0)
            totals = tsm_result.get("portfolio_totals", {})
            st.subheader(f"💰 Portfolio Totals — {month_label}")
            expected_files = len(systems)
            if n_loaded < expected_files:
                st.warning(
                    f"⚠️ Only **{n_loaded} of {expected_files} loaded systems** have a trade file "
                    f"available for the selected month. Missing files can understate the portfolio gap.")
            if not totals:
                st.caption("No theoretical trades exited in this month for any loaded system.")
            else:
                # ── Statement P&L/fee, pulled up here as the headline "how much
                # money do I actually have" figure — per Andrea: "consider the
                # monthly statements as the bible... the rest is less important."
                # The matching-engine's own "Actual P&L (net)" below only sums
                # MATCHED + pending-override trades (107 of ~314 June fills);
                # it structurally excludes unmatched actual fills (real P&L
                # that never got attributed to a theoretical signal) and any
                # system whose trade file isn't loaded yet, so it will never
                # equal the statement and shouldn't be read as "cash on hand."
                _recon = tsm_result.get("reconciliation")
                _stmt_row = None
                if _recon is not None and not _recon.empty and "TOTAL" in _recon["root"].values:
                    _stmt_row = _recon[_recon["root"] == "TOTAL"].iloc[0]
                if _stmt_row is not None:
                    st.metric("💵 Statement P&L (real, all fills, ground truth)",
                              f"${_stmt_row['statement_pnl']:,.2f}",
                              help="From the broker statement PDF — the actual cash result "
                                   "for the month, covering every fill on these 6 roots "
                                   "(matched, unmatched, and any system not yet loaded here). "
                                   "This is the number to trust for 'how much do I have.'")
                    st.caption(
                        f"Statement fee: ${_stmt_row['statement_fee']:,.2f} — see the fee "
                        f"explanation in the Reconciliation section below for why this is "
                        f"higher than the order-history-reconstructed commission figure.")
                else:
                    st.caption("No statement PDF reconciled for this month — showing the "
                               "matching-engine's own totals below only (matched + pending-"
                               "override trades, NOT the full real P&L for the month).")
                st.caption("Matching-engine totals below (matched + pending-override trades "
                           "only — a subset of the statement total above):")
                tcol1, tcol2, tcol3 = st.columns(3)
                tcol1.metric("Total Theoretical P&L", f"${totals.get('Theoretical P&L', 0):,.0f}",
                             help="Matches the Portfolio tab's methodology if the sizing/"
                                  "commission toggle above was on when you ran this.")
                tcol2.metric("Total Actual P&L (net, matched trades only)",
                             f"${totals.get('Actual P&L', 0):,.0f}",
                             help="Only matched + pending-override trades — NOT the full "
                                  "month's real P&L. See the Statement P&L figure above for that.")
                tcol3.metric("Total Delta", f"${totals.get('Delta', 0):,.0f}",
                             delta=f"${totals.get('Delta', 0):,.0f}")
                dcol1, dcol2, dcol3, dcol4 = st.columns(4)
                dcol1.metric("Commissions", f"${totals.get('Commissions ($)', 0):,.0f}")
                dcol2.metric("Slippage ($)", f"${totals.get('Slippage ($)', 0):,.0f}",
                             help="Dollar P&L gap (price gap x point value x contracts), "
                                  "not a points/ticks figure -- a small, tolerable price gap "
                                  "on a high-point-value root like GC or NQ can still show as "
                                  "a large dollar number here.")
                dcol3.metric("Override Impact ($)", f"${totals.get('Override Impact ($)', 0):,.0f}")
                dcol4.metric("Unmatched Impact ($)", f"${totals.get('Unmatched Impact ($)', 0):,.0f}")
                st.caption(
                    f"{totals.get('# Matched', 0)} matched · "
                    f"{totals.get('# Overrides (pending review)', 0)} pending override review · "
                    f"{totals.get('# Unmatched Theo', 0)} unmatched theoretical trades, "
                    f"across {n_loaded} system(s).")

            # ── Price-offset calibration transparency ──
            # price_offsets is dict[stem] -> LIST of segments (one per
            # contract-month the system traded, e.g. pre-roll/post-roll --
            # see calibrate_price_offsets()), not a single dict per system,
            # since a mid-month contract roll needs its own offset on each
            # side of the roll.
            offsets = tsm_result.get("price_offsets", {})
            off_rows = []
            for stem, segments in offsets.items():
                for seg in segments:
                    off_rows.append({
                        "System": stem,
                        "Contract": (f"{seg['contract_key'][0]}"
                                     f"{seg['contract_key'][1]:02d}/{seg['contract_key'][2]}"
                                     if seg.get("contract_key") else ""),
                        "Offset (pts)": round(seg["offset"] if seg["applied"] else seg.get("raw_offset", 0.0), 2),
                        "Trusted (auto-matched)": seg["applied"],
                        "# Anchors": seg["n_anchors"],
                        "Spread (pts)": round(seg["spread"], 3),
                        "Date range": f"{seg['date_min']:%Y-%m-%d} → {seg['date_max']:%Y-%m-%d}",
                    })
            if off_rows:
                n_trusted = sum(1 for r in off_rows if r["Trusted (auto-matched)"])
                with st.expander(f"📐 Price offset calibrated for {len(off_rows)} system/contract "
                                  f"segment(s) ({n_trusted} trusted for auto-matching)", expanded=False):
                    st.caption(
                        "Theoretical prices come from a back-adjusted continuous contract, offset "
                        "from the real traded contract — recalibrated separately per contract-month "
                        "a system actually traded, since a mid-month roll (e.g. MNQ M26→U26) resets "
                        "the offset. 'Trusted' segments were tight enough to auto-match on; "
                        "untrusted ones are still offered as override-review candidates (never "
                        "silently auto-matched) rather than discarded.")
                    render_table(pd.DataFrame(off_rows), hide_index=True, use_container_width=True)

            st.divider()

            # ── Per-system breakdown ──
            st.subheader(f"📊 System Breakdown — {month_label}")
            breakdown = tsm_result["breakdown"]
            if breakdown.empty:
                st.info("No theoretical trades exited in this month for any loaded system.")
            else:
                render_table(breakdown, hide_index=True, use_container_width=True,
                             height=min(42 + 36 * len(breakdown), 500))

                fig_bd = go.Figure()
                fig_bd.add_trace(go.Bar(x=breakdown["System"], y=breakdown["Theoretical P&L"],
                                         name="Theoretical P&L", marker_color=C["blue"]))
                fig_bd.add_trace(go.Bar(x=breakdown["System"], y=breakdown["Actual P&L"],
                                         name="Actual P&L", marker_color=C["portfolio"]))
                fig_bd.update_layout(template=THEME, barmode="group", height=380,
                                      margin=dict(l=0, r=0, t=30, b=0),
                                      title="Theoretical vs Actual P&L by system",
                                      yaxis_title="$")
                st.plotly_chart(fig_bd, use_container_width=True, key="chart_tsm_breakdown_bar")

                fig_delta = go.Figure()
                comp_cols = ["Commissions ($)", "Slippage ($)", "Override Impact ($)", "Unmatched Impact ($)"]
                comp_colors = [C["amber"], C["red"], C["green"], C["peak"]]
                for i, (col, color) in enumerate(zip(comp_cols, comp_colors)):
                    fig_delta.add_trace(go.Bar(x=breakdown["System"], y=breakdown[col],
                                                name=col, marker_color=color))
                fig_delta.update_layout(template=THEME, barmode="relative", height=380,
                                         margin=dict(l=0, r=0, t=30, b=0),
                                         title="Delta decomposition by system (sums to Actual − Theoretical)",
                                         yaxis_title="$")
                st.plotly_chart(fig_delta, use_container_width=True, key="chart_tsm_delta_bar")

            st.divider()

            # ── Matched trades ──
            matched = tsm_result["matched"]
            with st.expander(f"✅ Matched trades ({len(matched)})", expanded=False):
                if matched.empty:
                    st.caption("No confident matches within tolerance.")
                else:
                    flagged = matched[matched["had_alt_candidates"]]
                    if not flagged.empty:
                        st.warning(
                            f"{len(flagged)} of these had more than one similarly-good candidate "
                            "fill — double-check them (column `had_alt_candidates`).")
                    if "high_slippage" in matched.columns:
                        slippy = matched[matched["high_slippage"]]
                        if not slippy.empty:
                            st.info(
                                f"{len(slippy)} matched at essentially the same time as the "
                                "signal but with a bigger price gap than the base tolerance — "
                                "treated as slippage, not a manual override (column "
                                "`high_slippage`). NQ/MNQ allow up to 13pts, ES/MES up to 4pts "
                                "before this flag no longer applies at all and it falls to "
                                "override review instead.")
                    render_table(matched.drop(columns=["theo_id"]), hide_index=True,
                                 use_container_width=True, height=min(42 + 36 * len(matched), 600))

                    if _emp_slip is None:
                        st.warning(f"⚠️ Empirical slippage module unavailable: {_EMP_SLIP_IMPORT_ERROR}")
                    else:
                        _slip_summary = _emp_slip.estimate_slippage_ticks(matched)
                        st.subheader("📏 Empirical Slippage — Diagnostic")
                        if _slip_summary.empty:
                            st.warning("⚠️ No usable entry/exit slippage-tick fields were found in the matched sample. No slippage haircut is applied.")
                        else:
                            st.caption("Observed slippage is shown in ticks by contract root and strategy family. Dollar conversion and historical P&L haircuts are deliberately NOT applied until an explicit tick-value mapping is supplied; no placeholder market values are invented.")
                            render_table(_slip_summary, hide_index=True, use_container_width=True)

            # ── Override candidates — interactive human review ──
            overrides = tsm_result["override_candidates"]
            st.subheader(f"✏️ Manual override review ({len(overrides)} pending)")
            if overrides.empty:
                st.caption("No trades matched outside tight tolerance this month.")
            else:
                st.caption(
                    "These matched on symbol/qty/side within a wider window, but price or date "
                    "fell outside the normal tolerance — consistent with an early/late manual "
                    "exit. No override history exists yet, so nothing is assumed: confirm or "
                    "correct the reason for each, and it's saved to the override log below.")
                override_log = tsm.load_override_log(OVERRIDE_LOG_PATH)
                for i, row in overrides.reset_index(drop=True).iterrows():
                    theo_id = row["theo_id"]
                    prior = override_log[override_log["theo_id"] == theo_id] if not override_log.empty else pd.DataFrame()
                    prior_reason = prior["override_reason"].iloc[0] if not prior.empty else ""
                    with st.container():
                        st.markdown(
                            f"**{row['system_name']}** — {row['symbol']} × {row['n_contracts']} "
                            f"({row['direction']}) — theo P&L ${row['theo_pnl']:,.0f} vs "
                            f"actual net ${row['actual_pnl_net']:,.0f}  \n"
                            f"Entry: theo {row['theo_entry_price']:.2f} @ {row['theo_entry_date']:%Y-%m-%d %H:%M} "
                            f"→ actual {row['actual_entry_price']:.2f} @ {row['actual_entry_time']:%Y-%m-%d %H:%M} "
                            f"({row['entry_diff_ticks']:+.1f} ticks)  \n"
                            f"Exit: theo {row['theo_exit_price']:.2f} @ {row['theo_exit_date']:%Y-%m-%d %H:%M} "
                            f"→ actual {row['actual_exit_price']:.2f} @ {row['actual_exit_time']:%Y-%m-%d %H:%M} "
                            f"({row['exit_diff_ticks']:+.1f} ticks)")
                        rcol1, rcol2 = st.columns([3, 1])
                        with rcol1:
                            reason = st.text_input(
                                "Override reason", value=prior_reason,
                                key=f"tsm_override_reason_{i}_{theo_id}_{row['system_stem']}",
                                placeholder="e.g. exited early on discretionary risk call")
                        with rcol2:
                            if st.button("💾 Save", key=f"tsm_override_save_{i}_{theo_id}_{row['system_stem']}",
                                         use_container_width=True):
                                entry = row.to_dict()
                                entry["override_reason"] = reason
                                entry["match_type"] = "override"
                                entry["reviewed_by"] = "Andrea"
                                entry["reviewed_at"] = datetime.now().isoformat(timespec="seconds")
                                tsm.save_override_entry(OVERRIDE_LOG_PATH, entry)
                                st.success("Saved.")
                        st.divider()

            # ── Unmatched / out-of-scope ──
            uc1, uc2, uc3 = st.columns(3)
            with uc1:
                unmatched_theo = tsm_result["unmatched_theoretical"]
                with st.expander(f"❔ Unmatched theoretical ({len(unmatched_theo)})"):
                    st.caption("Strategy signal with no corresponding execution found — "
                               "either not taken, or outside the order-history's covered dates. "
                               "Its theoretical P&L still counts toward this system's totals "
                               "above (that's what drives the 'Unmatched Impact' column) — "
                               "tagging a reason below is just for your own record, it doesn't "
                               "change the numbers.")
                    if not unmatched_theo.empty:
                        by_sys = (unmatched_theo.groupby("system_name")
                                  .agg(n_trades=("pnl", "size"), total_theo_pnl=("pnl", "sum"))
                                  .sort_values("n_trades", ascending=False).reset_index())
                        by_sys["total_theo_pnl"] = by_sys["total_theo_pnl"].round(2)
                        st.caption("By system (which ones are driving this count):")
                        render_table(by_sys, hide_index=True, use_container_width=True)
                        st.caption("Full list:")
                        render_table(
                            unmatched_theo[["system_name", "entry_date", "entry_price",
                                            "exit_date", "exit_price", "n_contracts", "pnl"]],
                            hide_index=True, use_container_width=True)
                        skip_log = tsm.load_override_log(OVERRIDE_LOG_PATH)
                        st.caption("Tag why a specific signal wasn't taken (optional):")
                        for j, urow in unmatched_theo.reset_index(drop=True).iterrows():
                            u_theo_id = urow["theo_id"]
                            u_prior = skip_log[skip_log["theo_id"] == u_theo_id] if not skip_log.empty else pd.DataFrame()
                            u_prior_reason = (u_prior["override_reason"].iloc[0]
                                               if not u_prior.empty and "override_reason" in u_prior.columns
                                               else "")
                            ucol1, ucol2 = st.columns([3, 1])
                            with ucol1:
                                skip_reason = st.text_input(
                                    f"{urow['system_name']} — {urow['entry_date']:%Y-%m-%d} "
                                    f"(theo P&L ${urow['pnl']:,.0f})",
                                    value=u_prior_reason,
                                    key=f"tsm_skip_reason_{j}_{u_theo_id}_{urow['system_stem']}",
                                    placeholder="e.g. system was off, deliberately skipped")
                            with ucol2:
                                st.write("")
                                if st.button("💾 Save", key=f"tsm_skip_save_{j}_{u_theo_id}_{urow['system_stem']}",
                                             use_container_width=True):
                                    u_entry = urow.to_dict()
                                    u_entry["theo_pnl"] = urow.get("pnl")
                                    u_entry["actual_entry_time"] = ""
                                    u_entry["actual_entry_price"] = ""
                                    u_entry["actual_exit_time"] = ""
                                    u_entry["actual_exit_price"] = ""
                                    u_entry["actual_pnl_net"] = 0.0
                                    u_entry["entry_diff_ticks"] = ""
                                    u_entry["exit_diff_ticks"] = ""
                                    u_entry["override_reason"] = skip_reason
                                    u_entry["match_type"] = "skipped"
                                    u_entry["reviewed_by"] = "Andrea"
                                    u_entry["reviewed_at"] = datetime.now().isoformat(timespec="seconds")
                                    tsm.save_override_entry(OVERRIDE_LOG_PATH, u_entry)
                                    st.success("Saved.")
            with uc2:
                unmatched_actual = tsm_result["unmatched_actual"]
                with st.expander(f"❔ Unmatched actual fills ({len(unmatched_actual)})"):
                    st.caption("Real executions (in the 18 systems' symbols) that no "
                               "theoretical trade claimed — extra/manual trades, or a "
                               "system not currently loaded.")
                    if not unmatched_actual.empty:
                        render_table(
                            unmatched_actual[["root", "side", "qty_filled", "price",
                                              "fill_time", "commission"]].sort_values("fill_time"),
                            hide_index=True, use_container_width=True, height=300)
            with uc3:
                out_scope = tsm_result["out_of_scope_actual"]
                with st.expander(f"🚫 Out-of-scope fills ({len(out_scope)})"):
                    st.caption("Symbols outside the 18 systems (e.g. RTY, HG/MHG) — shown "
                               "for visibility only, never matched.")
                    if not out_scope.empty:
                        render_table(
                            out_scope[["root", "side", "qty_filled", "price", "fill_time"]]
                            .sort_values("fill_time"),
                            hide_index=True, use_container_width=True, height=300)

            st.divider()

            # ── Reconciliation vs monthly statement ──
            st.subheader("🧾 Reconciliation vs monthly statement")
            recon = tsm_result["reconciliation"]
            resolved_pdf = tsm_result.get("resolved_pdf_path")
            if recon.empty:
                if resolved_pdf:
                    st.caption(f"Found `{resolved_pdf}` but couldn't parse/reconcile it — "
                               "reconciliation skipped.")
                else:
                    st.caption("No statement PDF provided (or none found at the given path) — "
                               "reconciliation skipped.")
            else:
                if not tsm_result.get("pdf_is_stamped_match", True):
                    st.warning(
                        f"No PDF filename in the statements folder matched "
                        f"{int(match_year)}-{int(match_month):02d} — falling back to the "
                        f"most recently added PDF instead: `{resolved_pdf}`. These numbers "
                        "may be for the wrong month; add the correct month's statement PDF "
                        "to be sure.")
                else:
                    st.caption(f"Statement used: `{resolved_pdf}`")
                st.caption(
                    "Coarse per-symbol total check: the PDF statement's daily P&L/commission "
                    "totals vs. the order-history reconstruction, for the whole month. This "
                    "does not use trade-level matching (the PDF doesn't support it — see the "
                    "module docstring) so treat P&L mismatches as a prompt to double-check the "
                    "order-history file's coverage, not as a matching error.")
                st.caption(
                    "⚠️ `fee_match` will almost always read False and `fills_fee_reconstructed` "
                    "will almost always run below `statement_fee` — confirmed not fixable from "
                    "this data: the order-history file's Commission column only carries the "
                    "exchange/clearing fee attached at execution (a flat $0.40/contract on "
                    "micros, $1.00/contract on full-size, checked across an entire month with "
                    "zero variance), while the statement's fee total also includes NFA/broker "
                    "charges that never appear in that export. Trust `statement_fee` as the "
                    "real cost — it's what the sidebar's commission defaults are calibrated "
                    "from.")
                render_table(recon, hide_index=True, use_container_width=True)

            st.divider()
            with st.expander("📜 Override log (all systems, all months)"):
                full_log = tsm.load_override_log(OVERRIDE_LOG_PATH)
                if full_log.empty:
                    st.caption("No overrides saved yet.")
                else:
                    render_table(full_log, hide_index=True, use_container_width=True)
