"""
Youth Consumer Spend Index.
Cloud-safe: detects shared hosting (Streamlit Cloud etc.) and runs session-only.
Local (Mac .app / Windows .exe): persists config to ~/.youth_index/user_config.json.
"""

import os
import time
import io
import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import timedelta, datetime
from pathlib import Path
import json

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------- Defaults ----------

DEFAULT_INDEX_CONSTITUENTS = {
    "ETERNAL.NS":   0.20,
    "PAYTM.NS":     0.15,
    "NYKAA.NS":     0.15,
    "DELHIVERY.NS": 0.15,
    "TITAN.NS":     0.20,
    "ABFRL.NS":     0.15,
}

DEFAULT_WATCHLIST = ["RELIANCE.NS"]

BENCHMARKS = {
    "^NSEI":  "Nifty 50",
    "^BSESN": "Sensex",
}

DEFAULT_DISPLAY_NAMES = {
    "ETERNAL.NS":   "Eternal (Zomato)",
    "PAYTM.NS":     "Paytm",
    "NYKAA.NS":     "Nykaa",
    "DELHIVERY.NS": "Delhivery",
    "TITAN.NS":     "Titan",
    "ABFRL.NS":     "ABFRL",
    "RELIANCE.NS":  "Reliance",
    "^NSEI":        "Nifty 50 (benchmark)",
    "^BSESN":       "Sensex (benchmark)",
}

START_DATE = "2023-01-01"
BASE_VALUE = 1000
INDEX_NAME = "Youth Consumer Spend Index"

TV_BG        = "#0e1117"
TV_GRID      = "#1e222d"
TV_TEXT      = "#d1d4dc"
TV_UP        = "#26a69a"
TV_DOWN      = "#ef5350"
TV_UP_FILL   = "rgba(38, 166, 154, 0.15)"
TV_DOWN_FILL = "rgba(239, 83, 80, 0.15)"

CONFIG_DIR = Path.home() / ".youth_index"
CONFIG_PATH = CONFIG_DIR / "user_config.json"


# ---------- Cloud detection ----------

def is_cloud_deployment():
    """
    Detect if running on a shared cloud server (Streamlit Cloud, HF Spaces, etc.).
    On cloud, we MUST NOT write to disk because the filesystem is shared across visitors.
    """
    for var in ["STREAMLIT_SHARING_MODE", "STREAMLIT_CLOUD", "SPACE_ID", "RENDER", "RAILWAY_ENVIRONMENT"]:
        if os.environ.get(var):
            return True
    hostname = os.environ.get("HOSTNAME", "").lower()
    if "streamlit" in hostname or "cloud" in hostname:
        return True
    return False


IS_CLOUD = is_cloud_deployment()


# ---------- Persistence ----------

def load_config():
    if IS_CLOUD or not CONFIG_PATH.exists():
        return {
            "index_constituents": dict(DEFAULT_INDEX_CONSTITUENTS),
            "watchlist": list(DEFAULT_WATCHLIST),
            "display_names": dict(DEFAULT_DISPLAY_NAMES),
        }
    try:
        with open(CONFIG_PATH, "r") as f:
            cfg = json.load(f)
        assert "index_constituents" in cfg
        assert "watchlist" in cfg
        assert "display_names" in cfg
        total = sum(cfg["index_constituents"].values())
        if abs(total - 1.0) > 0.01:
            cfg["index_constituents"] = {
                t: w / total for t, w in cfg["index_constituents"].items()
            }
        return cfg
    except Exception:
        return {
            "index_constituents": dict(DEFAULT_INDEX_CONSTITUENTS),
            "watchlist": list(DEFAULT_WATCHLIST),
            "display_names": dict(DEFAULT_DISPLAY_NAMES),
        }


def save_config():
    if IS_CLOUD:
        return
    cfg = {
        "index_constituents": st.session_state.index_constituents,
        "watchlist": st.session_state.watchlist,
        "display_names": st.session_state.display_names,
    }
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        st.warning(f"Could not save config: {e}")


def reset_to_defaults():
    st.session_state.index_constituents = dict(DEFAULT_INDEX_CONSTITUENTS)
    st.session_state.watchlist = list(DEFAULT_WATCHLIST)
    st.session_state.display_names = dict(DEFAULT_DISPLAY_NAMES)
    if not IS_CLOUD and CONFIG_PATH.exists():
        try:
            CONFIG_PATH.unlink()
        except Exception:
            pass


def quit_app():
    if not IS_CLOUD:
        save_config()
        os._exit(0)


# ---------- Session state ----------

if "config_loaded" not in st.session_state:
    cfg = load_config()
    st.session_state.index_constituents = cfg["index_constituents"]
    st.session_state.watchlist = cfg["watchlist"]
    st.session_state.display_names = cfg["display_names"]
    st.session_state.config_loaded = True

if "validated_ticker" not in st.session_state:
    st.session_state.validated_ticker = None
if "validated_info" not in st.session_state:
    st.session_state.validated_info = None


# ---------- Data ----------

@st.cache_data(ttl=3600)
def fetch_prices(tickers, start):
    data = yf.download(tickers, start=start, progress=False, auto_adjust=True)["Close"]
    if isinstance(data, pd.Series):
        data = data.to_frame()
    return data


@st.cache_data(ttl=3600)
def fetch_ohlcv(ticker, start):
    df = yf.download(ticker, start=start, progress=False, auto_adjust=True)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df.dropna()


@st.cache_data(ttl=3600)
def validate_ticker(ticker):
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period="1mo")
        if hist.empty:
            return False, "No data returned for this ticker."
        info = {"last_price": round(hist["Close"].iloc[-1], 2)}
        try:
            i = t.info
            info["longName"] = i.get("longName") or i.get("shortName") or ticker
        except Exception:
            info["longName"] = ticker
        return True, info
    except Exception as e:
        return False, f"Fetch failed: {str(e)[:100]}"


# ---------- Core math ----------

def build_index(prices, weights, base_value=1000):
    normalised = prices / prices.iloc[0] * 100
    weight_series = pd.Series(weights)
    normalised = normalised[weight_series.index]
    index_level = (normalised * weight_series).sum(axis=1)
    index_level = index_level / index_level.iloc[0] * base_value
    return index_level


def compute_stats(series, name):
    returns = series.pct_change().dropna()
    total_return = (series.iloc[-1] / series.iloc[0] - 1) * 100
    ann_return = ((1 + returns.mean()) ** 252 - 1) * 100
    ann_std = returns.std() * np.sqrt(252) * 100
    sharpe = ann_return / ann_std if ann_std > 0 else 0
    return {
        "Name": name,
        "Total Return %": round(total_return, 2),
        "Annualised Return %": round(ann_return, 2),
        "Annualised Std Dev %": round(ann_std, 2),
        "Sharpe": round(sharpe, 2),
        "High": round(series.max(), 2),
        "Low": round(series.min(), 2),
    }


def compute_drawdown(series):
    running_max = series.cummax()
    drawdown = (series - running_max) / running_max * 100
    max_dd = drawdown.min()
    return drawdown, max_dd


def compute_beta(index_series, benchmark_series):
    df = pd.concat([index_series, benchmark_series], axis=1).dropna()
    df.columns = ["idx", "bench"]
    idx_returns = df["idx"].pct_change().dropna()
    bench_returns = df["bench"].pct_change().dropna()
    if bench_returns.var() == 0:
        return 0.0
    covariance = np.cov(idx_returns, bench_returns)[0, 1]
    beta = covariance / bench_returns.var()
    return round(beta, 3)


def compute_correlation(index_series, benchmark_series):
    df = pd.concat([index_series, benchmark_series], axis=1).dropna()
    df.columns = ["idx", "bench"]
    return round(df["idx"].pct_change().corr(df["bench"].pct_change()), 3)


def compute_rolling_volatility(series, window=30):
    returns = series.pct_change().dropna()
    rolling_std = returns.rolling(window).std() * np.sqrt(252) * 100
    return rolling_std.dropna()


def compute_daily_return_stats(series):
    returns = series.pct_change().dropna() * 100
    return {
        "Best Day %": round(returns.max(), 2),
        "Worst Day %": round(returns.min(), 2),
        "Mean Daily %": round(returns.mean(), 4),
        "Median Daily %": round(returns.median(), 4),
        "Positive Days %": round((returns > 0).sum() / len(returns) * 100, 2),
    }


def filter_by_timeframe(series_or_df, timeframe):
    end = series_or_df.index[-1]
    if timeframe == "1M":
        start = end - timedelta(days=30)
    elif timeframe == "6M":
        start = end - timedelta(days=182)
    elif timeframe == "YTD":
        start = pd.Timestamp(year=end.year, month=1, day=1)
    elif timeframe == "1Y":
        start = end - timedelta(days=365)
    else:
        return series_or_df
    return series_or_df.loc[series_or_df.index >= start]


def period_change_pct(series, timeframe):
    sliced = filter_by_timeframe(series, timeframe)
    if len(sliced) < 2:
        return 0.0
    return (sliced.iloc[-1] / sliced.iloc[0] - 1) * 100


def rebalance_add(existing_weights, new_ticker, new_weight):
    if new_weight <= 0 or new_weight >= 1:
        raise ValueError("Weight must be between 0 and 1 exclusive.")
    scale = 1.0 - new_weight
    rebalanced = {t: w * scale for t, w in existing_weights.items()}
    rebalanced[new_ticker] = new_weight
    return rebalanced


def rebalance_remove(existing_weights, ticker_to_remove):
    if ticker_to_remove not in existing_weights:
        return existing_weights
    remaining = {t: w for t, w in existing_weights.items() if t != ticker_to_remove}
    total = sum(remaining.values())
    if total == 0:
        return remaining
    return {t: w / total for t, w in remaining.items()}


# ---------- Charts ----------

def tv_style_index_chart(series, timeframe, height=500):
    sliced = filter_by_timeframe(series, timeframe)
    change = period_change_pct(series, timeframe)
    color = TV_UP if change >= 0 else TV_DOWN
    fill = TV_UP_FILL if change >= 0 else TV_DOWN_FILL
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=sliced.index, y=sliced.values, mode="lines",
        line=dict(width=2.5, color=color), fill="tozeroy", fillcolor=fill,
        hovertemplate="<b>%{x|%d %b %Y}</b><br>%{y:,.2f}<extra></extra>",
    ))
    if timeframe == "All":
        prev_close = series.iloc[-2] if len(series) >= 2 else series.iloc[-1]
        fig.add_hline(y=prev_close, line_dash="dot", line_color="gray",
                      annotation_text=f"Prev close {prev_close:,.2f}", annotation_position="top left")
    fig.update_layout(
        height=height, hovermode="x unified", showlegend=False,
        plot_bgcolor=TV_BG, paper_bgcolor=TV_BG, font=dict(color=TV_TEXT),
        xaxis=dict(gridcolor=TV_GRID, showspikes=True, spikemode="across", spikecolor="gray", spikethickness=1),
        yaxis=dict(gridcolor=TV_GRID, side="right"),
        margin=dict(l=0, r=0, t=10, b=0),
    )
    return fig


def tv_style_stock_chart(df, timeframe, height=550):
    sliced = filter_by_timeframe(df, timeframe)
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        vertical_spacing=0.02, row_heights=[0.75, 0.25])
    fig.add_trace(go.Candlestick(
        x=sliced.index, open=sliced["Open"], high=sliced["High"],
        low=sliced["Low"], close=sliced["Close"],
        increasing_line_color=TV_UP, decreasing_line_color=TV_DOWN,
        increasing_fillcolor=TV_UP, decreasing_fillcolor=TV_DOWN,
        name="Price", showlegend=False,
    ), row=1, col=1)
    vol_colors = [TV_UP if c >= o else TV_DOWN
                  for o, c in zip(sliced["Open"], sliced["Close"])]
    fig.add_trace(go.Bar(
        x=sliced.index, y=sliced["Volume"],
        marker_color=vol_colors, name="Volume", showlegend=False, opacity=0.6,
    ), row=2, col=1)
    fig.update_layout(
        height=height, hovermode="x unified",
        plot_bgcolor=TV_BG, paper_bgcolor=TV_BG, font=dict(color=TV_TEXT),
        xaxis_rangeslider_visible=False, margin=dict(l=0, r=0, t=10, b=0),
        showlegend=False,
    )
    fig.update_xaxes(gridcolor=TV_GRID, showspikes=True, spikemode="across",
                     spikecolor="gray", spikethickness=1)
    fig.update_yaxes(gridcolor=TV_GRID, side="right")
    return fig


def drawdown_chart(dd_series, height=350):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=dd_series.index, y=dd_series.values, mode="lines",
        line=dict(width=1.5, color=TV_DOWN), fill="tozeroy", fillcolor=TV_DOWN_FILL,
        hovertemplate="<b>%{x|%d %b %Y}</b><br>%{y:.2f}%<extra></extra>",
    ))
    fig.update_layout(
        height=height, hovermode="x unified", showlegend=False,
        plot_bgcolor=TV_BG, paper_bgcolor=TV_BG, font=dict(color=TV_TEXT),
        xaxis=dict(gridcolor=TV_GRID),
        yaxis=dict(gridcolor=TV_GRID, title="Drawdown %", side="right"),
        margin=dict(l=0, r=0, t=10, b=0),
    )
    return fig


def rolling_vol_chart(rvol_series, window, height=350):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=rvol_series.index, y=rvol_series.values, mode="lines",
        line=dict(width=2, color=TV_UP),
        hovertemplate="<b>%{x|%d %b %Y}</b><br>%{y:.2f}%<extra></extra>",
    ))
    fig.update_layout(
        height=height, hovermode="x unified", showlegend=False,
        plot_bgcolor=TV_BG, paper_bgcolor=TV_BG, font=dict(color=TV_TEXT),
        xaxis=dict(gridcolor=TV_GRID),
        yaxis=dict(gridcolor=TV_GRID, title=f"Rolling {window}D Volatility %", side="right"),
        margin=dict(l=0, r=0, t=10, b=0),
    )
    return fig


# ---------- Report generator ----------

def _png_bytes_from_fig(fig):
    return fig.to_image(format="png", width=1400, height=600, scale=2)


def _add_heading(doc, text, level=1):
    return doc.add_heading(text, level=level)


def _add_table_from_df(doc, df):
    df_reset = df.reset_index() if df.index.name else df.copy()
    table = doc.add_table(rows=1, cols=len(df_reset.columns))
    table.style = "Light Grid Accent 1"
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(df_reset.columns):
        hdr_cells[i].text = str(col)
    for _, row in df_reset.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)
    return table


def build_report(index_level, bench_rebased, constituents, display_names,
                 stats_df, monthly_returns, max_dd, dd_series, beta_dict, corr_dict,
                 daily_stats, rolling_vol, common_start):
    doc = Document()
    title = doc.add_heading(INDEX_NAME, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Insights Report\n").bold = True
    p.add_run(f"Period: {common_start.date()} to {index_level.index[-1].date()}\n")
    p.add_run(f"Generated: {datetime.now().strftime('%d %B %Y')}\n")
    doc.add_page_break()

    _add_heading(doc, "1. Executive Summary", level=1)
    idx_stats = stats_df.loc[INDEX_NAME]
    summary = doc.add_paragraph()
    summary.add_run(
        f"The {INDEX_NAME} is a capped market-cap index of India's youth-facing consumer brands. "
        f"Over the tracked period, the index returned "
    )
    summary.add_run(f"{idx_stats['Total Return %']}%").bold = True
    summary.add_run(
        f" with an annualised standard deviation of {idx_stats['Annualised Std Dev %']}% "
        f"and a Sharpe ratio of {idx_stats['Sharpe']}.\n\n"
    )
    for bench_name in BENCHMARKS.values():
        b = stats_df.loc[bench_name]
        excess = idx_stats["Total Return %"] - b["Total Return %"]
        summary.add_run(
            f"Excess return vs {bench_name}: {excess:+.2f}% "
            f"({bench_name} returned {b['Total Return %']}% over the same window).\n"
        )
    summary.add_run(f"\nMaximum drawdown over the period: {max_dd:.2f}%.\n")
    doc.add_paragraph()
    caveats = doc.add_paragraph()
    caveats.add_run("Honest caveats:").bold = True
    caveats.add_run(
        f"\n- The index is notably more volatile than the benchmarks "
        f"({idx_stats['Annualised Std Dev %']}% vs "
        f"{stats_df.loc['Nifty 50', 'Annualised Std Dev %']}% for Nifty 50). "
        f"Higher return has come with higher risk.\n"
        f"- Start-date sensitivity: several constituents IPO'd in 2021-2022 with weak initial performance. "
        f"Starting the index in early 2023 captures much of their post-drawdown recovery. "
        f"A different start date would give a different picture.\n"
        f"- The index composition may change over time as constituents are added or removed. "
        f"Only conclusions drawn over a stable composition period are strictly comparable.\n"
    )
    doc.add_page_break()

    _add_heading(doc, "2. Constituents and Methodology", level=1)
    doc.add_paragraph(
        "The index uses capped market-cap weighting: each constituent has a fixed weight, "
        "chosen so no single company dominates. Weights sum to 100% and are rebalanced whenever "
        "a constituent is added or removed. Benchmarks (Nifty 50 and Sensex) are tracked in parallel "
        "for comparison but are not part of the index."
    )
    const_df = pd.DataFrame([
        {"Ticker": t, "Company": display_names.get(t, t), "Weight %": round(w * 100, 2)}
        for t, w in constituents.items()
    ])
    _add_table_from_df(doc, const_df)
    doc.add_page_break()

    _add_heading(doc, "3. Performance", level=1)
    perf_fig = go.Figure()
    perf_fig.add_trace(go.Scatter(x=index_level.index, y=index_level.values,
                                  mode="lines", name=INDEX_NAME, line=dict(width=2.5)))
    for ticker, name in BENCHMARKS.items():
        perf_fig.add_trace(go.Scatter(x=bench_rebased.index, y=bench_rebased[ticker],
                                      mode="lines", name=name, line=dict(width=1.5, dash="dot")))
    perf_fig.update_layout(template="plotly_white", yaxis_title=f"Rebased to {BASE_VALUE}",
                           legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                           margin=dict(l=40, r=40, t=40, b=40))
    try:
        img_bytes = _png_bytes_from_fig(perf_fig)
        doc.add_picture(io.BytesIO(img_bytes), width=Inches(6.5))
    except Exception as e:
        doc.add_paragraph(f"[Chart could not be rendered: {e}. Ensure `kaleido` is installed.]")

    doc.add_paragraph()
    _add_heading(doc, "Monthly Returns", level=2)
    _add_table_from_df(doc, monthly_returns)
    doc.add_page_break()

    _add_heading(doc, "4. Risk Analysis", level=1)
    _add_heading(doc, "Return and Risk Statistics", level=2)
    _add_table_from_df(doc, stats_df)
    _add_heading(doc, "Daily Return Distribution", level=2)
    ds_df = pd.DataFrame([(k, v) for k, v in daily_stats.items()], columns=["Metric", "Value"])
    _add_table_from_df(doc, ds_df)
    _add_heading(doc, "Beta and Correlation vs Benchmarks", level=2)
    bc_rows = []
    for bench_name in BENCHMARKS.values():
        bc_rows.append({"Benchmark": bench_name,
                        "Beta": beta_dict.get(bench_name, "n/a"),
                        "Correlation": corr_dict.get(bench_name, "n/a")})
    _add_table_from_df(doc, pd.DataFrame(bc_rows))
    doc.add_paragraph(
        "Beta measures how much the index moves per unit move in the benchmark. "
        "Beta of 1.0 means it moves in step; higher means more sensitive, lower means less. "
        "Correlation measures how tightly the two move together (1.0 = perfect, 0 = unrelated)."
    )

    _add_heading(doc, "Drawdown Over Time", level=2)
    dd_fig = go.Figure()
    dd_fig.add_trace(go.Scatter(x=dd_series.index, y=dd_series.values, mode="lines",
                                line=dict(width=1.5, color="crimson"), fill="tozeroy"))
    dd_fig.update_layout(template="plotly_white", yaxis_title="Drawdown %",
                         margin=dict(l=40, r=40, t=40, b=40))
    try:
        img_bytes = _png_bytes_from_fig(dd_fig)
        doc.add_picture(io.BytesIO(img_bytes), width=Inches(6.5))
    except Exception as e:
        doc.add_paragraph(f"[Chart could not be rendered: {e}]")

    _add_heading(doc, "30-Day Rolling Volatility", level=2)
    rv_fig = go.Figure()
    rv_fig.add_trace(go.Scatter(x=rolling_vol.index, y=rolling_vol.values, mode="lines",
                                line=dict(width=2)))
    rv_fig.update_layout(template="plotly_white", yaxis_title="Annualised Volatility %",
                         margin=dict(l=40, r=40, t=40, b=40))
    try:
        img_bytes = _png_bytes_from_fig(rv_fig)
        doc.add_picture(io.BytesIO(img_bytes), width=Inches(6.5))
    except Exception as e:
        doc.add_paragraph(f"[Chart could not be rendered: {e}]")

    doc.add_page_break()
    _add_heading(doc, "5. Reflection", level=1)
    doc.add_paragraph("This section is for your own commentary. Consider addressing:")
    doc.add_paragraph("What surprised you most about the numbers?", style="List Bullet")
    doc.add_paragraph("Which constituent drove the biggest gains or losses?", style="List Bullet")
    doc.add_paragraph("Does the excess return justify the extra volatility?", style="List Bullet")
    doc.add_paragraph("What would you change about the index if you were to build it again?", style="List Bullet")
    doc.add_paragraph("What does the drawdown chart tell you about risk during the period?", style="List Bullet")
    doc.add_paragraph()
    doc.add_paragraph("[Write your reflection here]")
    for _ in range(15):
        doc.add_paragraph()

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ---------- Streamlit page ----------

st.set_page_config(page_title=INDEX_NAME, layout="wide")

header_col, quit_col = st.columns([5, 1])
with header_col:
    st.title(INDEX_NAME)
    st.caption("A capped market-cap index of India's youth-facing consumer brands, benchmarked against Nifty 50 and Sensex.")
with quit_col:
    st.write("")
    if not IS_CLOUD:
        if st.button("Quit App", type="secondary", width='stretch'):
            st.warning("Shutting down. You can close this tab.")
            time.sleep(1)
            quit_app()

index_constituents = st.session_state.index_constituents
watchlist = st.session_state.watchlist
display_names = st.session_state.display_names

with st.spinner("Fetching market data..."):
    all_tickers = list(index_constituents.keys()) + list(BENCHMARKS.keys()) + watchlist
    raw = fetch_prices(all_tickers, START_DATE)

valid_constituents = [t for t in index_constituents.keys()
                      if t in raw.columns and not raw[t].dropna().empty]
if len(valid_constituents) < len(index_constituents):
    dropped = [t for t in index_constituents if t not in valid_constituents]
    st.warning(f"Dropping tickers with no data: {', '.join(dropped)}")
    for t in dropped:
        del st.session_state.index_constituents[t]
    total = sum(st.session_state.index_constituents.values())
    if total > 0:
        st.session_state.index_constituents = {
            t: w / total for t, w in st.session_state.index_constituents.items()
        }
    save_config()
    index_constituents = st.session_state.index_constituents

constituent_prices = raw[list(index_constituents.keys())].dropna()
common_start = constituent_prices.index[0]
benchmark_prices = raw[list(BENCHMARKS.keys())].loc[common_start:].dropna()

index_level = build_index(constituent_prices, index_constituents, BASE_VALUE)
bench_rebased = benchmark_prices / benchmark_prices.iloc[0] * BASE_VALUE

latest_val = index_level.iloc[-1]
first_val = index_level.iloc[0]
total_ret = (latest_val / first_val - 1) * 100
day_change = (index_level.iloc[-1] / index_level.iloc[-2] - 1) * 100

col1, col2, col3, col4 = st.columns(4)
col1.metric("Index Level", f"{latest_val:,.2f}", f"{day_change:+.2f}% today")
col2.metric("Total Return", f"{total_ret:+.2f}%")
col3.metric("Window Start", f"{common_start.date()}")
col4.metric("Trading Days", f"{len(index_level)}")

st.divider()

view_col, tf_col = st.columns([1, 3])
with view_col:
    view_mode = st.radio("Chart view", ["Index Only", "vs Benchmarks"], horizontal=True)

if view_mode == "Index Only":
    with tf_col:
        timeframe = st.radio("Timeframe", ["1M", "6M", "YTD", "1Y", "All"],
                             horizontal=True, index=4, key="idx_tf")
    sliced = filter_by_timeframe(index_level, timeframe)
    change = period_change_pct(index_level, timeframe)
    hcol1, _ = st.columns([1, 3])
    hcol1.metric(f"Change ({timeframe})", f"{sliced.iloc[-1]:,.2f}", f"{change:+.2f}%")
    st.plotly_chart(tv_style_index_chart(index_level, timeframe), width='stretch')
    st.markdown("**Change across periods**")
    period_cols = st.columns(5)
    for i, tf in enumerate(["1M", "6M", "YTD", "1Y", "All"]):
        period_cols[i].metric(tf, f"{period_change_pct(index_level, tf):+.2f}%")
else:
    st.subheader("Index vs Benchmarks")
    chart_df = pd.DataFrame({
        INDEX_NAME: index_level,
        "Nifty 50": bench_rebased["^NSEI"],
        "Sensex":   bench_rebased["^BSESN"],
    }).dropna()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=chart_df.index, y=chart_df[INDEX_NAME], mode="lines",
                             name="Youth Index", line=dict(width=2.5, color=TV_UP)))
    fig.add_trace(go.Scatter(x=chart_df.index, y=chart_df["Nifty 50"], mode="lines",
                             name="Nifty 50", line=dict(width=1.5, dash="dot", color="#5d9cec")))
    fig.add_trace(go.Scatter(x=chart_df.index, y=chart_df["Sensex"], mode="lines",
                             name="Sensex", line=dict(width=1.5, dash="dot", color="#f7b955")))
    fig.update_layout(
        height=500, hovermode="x unified",
        plot_bgcolor=TV_BG, paper_bgcolor=TV_BG, font=dict(color=TV_TEXT),
        xaxis=dict(gridcolor=TV_GRID),
        yaxis=dict(gridcolor=TV_GRID, title=f"Rebased to {BASE_VALUE}"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=0, r=0, t=10, b=0),
    )
    st.plotly_chart(fig, width='stretch')

st.divider()

st.subheader("Individual Stock View")
st.caption("Candlestick with volume, styled TradingView-like.")

dropdown_order = list(index_constituents.keys()) + watchlist + list(BENCHMARKS.keys())
dropdown_labels = {display_names.get(t, t): t for t in dropdown_order}

sel_col, sel_tf_col = st.columns([1, 3])
with sel_col:
    selected_label = st.selectbox("Select a stock or benchmark", list(dropdown_labels.keys()))
with sel_tf_col:
    stock_tf = st.radio("Timeframe", ["1M", "6M", "YTD", "1Y", "All"],
                        horizontal=True, index=3, key="stock_tf")

selected_ticker = dropdown_labels[selected_label]

with st.spinner(f"Loading {selected_label}..."):
    ohlcv = fetch_ohlcv(selected_ticker, START_DATE)

if ohlcv.empty:
    st.warning(f"No data returned for {selected_ticker}.")
else:
    close_series = ohlcv["Close"]
    tf_slice = filter_by_timeframe(close_series, stock_tf)
    tf_change_pct = period_change_pct(close_series, stock_tf)
    latest_price = close_series.iloc[-1]
    day_pct = (close_series.iloc[-1] / close_series.iloc[-2] - 1) * 100 if len(close_series) >= 2 else 0

    h1, h2, h3 = st.columns(3)
    h1.metric(f"{selected_label} - Latest", f"{latest_price:,.2f}", f"{day_pct:+.2f}% today")
    h2.metric(f"Change ({stock_tf})", f"{tf_slice.iloc[-1]:,.2f}", f"{tf_change_pct:+.2f}%")
    h3.metric("Volume (latest)", f"{int(ohlcv['Volume'].iloc[-1]):,}")

    st.plotly_chart(tv_style_stock_chart(ohlcv, stock_tf), width='stretch')

    s_stats = compute_stats(close_series, selected_label)
    s_df = pd.DataFrame([s_stats]).set_index("Name")
    st.markdown("**Statistics (full data window)**")
    st.dataframe(s_df, width='stretch')

st.divider()

st.subheader("Performance and Risk Statistics")
stats = [compute_stats(index_level, INDEX_NAME)]
for ticker, name in BENCHMARKS.items():
    stats.append(compute_stats(bench_rebased[ticker], name))
stats_df = pd.DataFrame(stats).set_index("Name")
st.dataframe(stats_df, width='stretch')

st.subheader("Excess Return vs Benchmarks")
idx_return = stats_df.loc[INDEX_NAME, "Total Return %"]
excess_cols = st.columns(len(BENCHMARKS))
for i, name in enumerate(BENCHMARKS.values()):
    bench_return = stats_df.loc[name, "Total Return %"]
    excess = idx_return - bench_return
    excess_cols[i].metric(f"vs {name}", f"{excess:+.2f}%")

st.divider()
st.subheader("Extended Risk Analysis")

dd_series, max_dd = compute_drawdown(index_level)
dd_col1, dd_col2 = st.columns([1, 3])
dd_col1.metric("Max Drawdown", f"{max_dd:.2f}%")
dd_col2.caption("The biggest peak-to-trough fall in the index over the tracked period.")
st.plotly_chart(drawdown_chart(dd_series), width='stretch')

st.markdown("**Beta and Correlation vs Benchmarks**")
beta_dict = {}
corr_dict = {}
bc_rows = []
for ticker, bench_name in BENCHMARKS.items():
    b = compute_beta(index_level, bench_rebased[ticker])
    c = compute_correlation(index_level, bench_rebased[ticker])
    beta_dict[bench_name] = b
    corr_dict[bench_name] = c
    bc_rows.append({"Benchmark": bench_name, "Beta": b, "Correlation": c})
bc_df = pd.DataFrame(bc_rows).set_index("Benchmark")
st.dataframe(bc_df, width='stretch')
st.caption(
    "Beta > 1 means the index moves more than the benchmark; < 1 means less. "
    "Correlation near 1 means they move together; near 0 means independent."
)

st.markdown("**Daily Return Distribution**")
daily_stats = compute_daily_return_stats(index_level)
ds_df = pd.DataFrame([daily_stats])
st.dataframe(ds_df, width='stretch', hide_index=True)

st.markdown("**Rolling 30-Day Volatility (Annualised)**")
rolling_vol = compute_rolling_volatility(index_level, window=30)
st.plotly_chart(rolling_vol_chart(rolling_vol, window=30), width='stretch')

st.divider()

st.subheader("Insights Report")
st.caption("Generate a Word document with the full index summary, charts, risk analysis, and a blank section for your written reflection.")

report_col1, report_col2 = st.columns([1, 3])
with report_col1:
    if st.button("Generate Report", type="primary", width='stretch'):
        with st.spinner("Building report... (this takes ~10 seconds)"):
            monthly = index_level.resample("ME").last().pct_change().dropna() * 100
            monthly_df = pd.DataFrame({
                "Month": monthly.index.strftime("%b %Y"),
                "Return %": monthly.round(2).values,
            })
            report_bytes = build_report(
                index_level=index_level, bench_rebased=bench_rebased,
                constituents=index_constituents, display_names=display_names,
                stats_df=stats_df, monthly_returns=monthly_df,
                max_dd=max_dd, dd_series=dd_series,
                beta_dict=beta_dict, corr_dict=corr_dict,
                daily_stats=daily_stats, rolling_vol=rolling_vol,
                common_start=common_start,
            )
            st.session_state.report_bytes = report_bytes
            st.session_state.report_filename = (
                f"YouthIndex_Report_{datetime.now().strftime('%Y%m%d')}.docx"
            )
        st.success("Report ready. Click Download.")

with report_col2:
    if "report_bytes" in st.session_state:
        st.download_button(
            label="Download Report (.docx)",
            data=st.session_state.report_bytes,
            file_name=st.session_state.report_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            width='stretch',
        )

st.divider()

st.subheader("Manage Constituents")

status_col, reset_col = st.columns([3, 1])
with status_col:
    if IS_CLOUD:
        st.caption("Running on the cloud. Your changes are kept for this session only and reset when you reload.")
    elif CONFIG_PATH.exists():
        st.caption(f"Changes are saved to `{CONFIG_PATH}` and persist across restarts.")
    else:
        st.caption(f"Using shipped defaults. Changes will be saved to `{CONFIG_PATH}`.")
with reset_col:
    if st.button("Reset to defaults", width='stretch'):
        reset_to_defaults()
        st.cache_data.clear()
        st.rerun()

st.markdown("**Add a new stock**")
search_col, btn_col = st.columns([3, 1])
with search_col:
    search_input = st.text_input(
        "Ticker (yfinance format, e.g. SWIGGY.NS, IRCTC.NS, JUBLFOOD.NS)",
        key="search_input",
        placeholder="Type a ticker and click Validate",
    )
with btn_col:
    st.write("")
    st.write("")
    validate_btn = st.button("Validate", width='stretch')

if validate_btn and search_input.strip():
    ticker_clean = search_input.strip().upper()
    with st.spinner(f"Validating {ticker_clean}..."):
        ok, info = validate_ticker(ticker_clean)
    if ok:
        st.session_state.validated_ticker = ticker_clean
        st.session_state.validated_info = info
    else:
        st.session_state.validated_ticker = None
        st.session_state.validated_info = None
        st.error(f"Could not validate {ticker_clean}: {info}")

if st.session_state.validated_ticker:
    vt = st.session_state.validated_ticker
    vi = st.session_state.validated_info
    st.success(f"Validated: **{vt}** ({vi['longName']}) - last price ₹{vi['last_price']}")

    already_in_index = vt in st.session_state.index_constituents
    already_in_watchlist = vt in st.session_state.watchlist

    if already_in_index:
        st.info(f"{vt} is already in the index.")
    elif already_in_watchlist:
        st.info(f"{vt} is already in the watchlist.")
    else:
        add_col1, add_col2 = st.columns(2)
        with add_col1:
            st.markdown("**Add to Index**")
            new_weight_pct = st.slider(
                "Weight for new stock (%)", min_value=1, max_value=50, value=10,
                key="new_weight_slider",
            )
            st.caption(f"Existing weights will be scaled by {(100 - new_weight_pct) / 100:.2f} to keep total at 100%.")
            if st.button("Add to Index", key="add_to_index", width='stretch'):
                try:
                    new_weights = rebalance_add(
                        st.session_state.index_constituents, vt, new_weight_pct / 100,
                    )
                    st.session_state.index_constituents = new_weights
                    st.session_state.display_names[vt] = vi["longName"]
                    st.session_state.validated_ticker = None
                    st.session_state.validated_info = None
                    save_config()
                    st.cache_data.clear()
                    st.success(f"Added {vt} to the index at {new_weight_pct}%.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to add: {e}")
        with add_col2:
            st.markdown("**Add to Watchlist**")
            st.caption("Displayed in the Individual Stock View but not weighted into the index.")
            st.write("")
            if st.button("Add to Watchlist", key="add_to_watchlist", width='stretch'):
                st.session_state.watchlist.append(vt)
                st.session_state.display_names[vt] = f"{vi['longName']} (watchlist)"
                st.session_state.validated_ticker = None
                st.session_state.validated_info = None
                save_config()
                st.cache_data.clear()
                st.success(f"Added {vt} to the watchlist.")
                st.rerun()

st.markdown("---")

st.markdown("**Current Index Constituents**")
for ticker, weight in list(st.session_state.index_constituents.items()):
    c1, c2, c3, c4 = st.columns([2, 3, 2, 1])
    c1.write(f"`{ticker}`")
    c2.write(display_names.get(ticker, ticker))
    c3.write(f"{weight * 100:.2f}%")
    if c4.button("Remove", key=f"rm_idx_{ticker}"):
        if len(st.session_state.index_constituents) <= 1:
            st.error("Cannot remove the last constituent.")
        else:
            st.session_state.index_constituents = rebalance_remove(
                st.session_state.index_constituents, ticker,
            )
            save_config()
            st.cache_data.clear()
            st.rerun()

st.markdown("**Current Watchlist**")
if not st.session_state.watchlist:
    st.caption("Watchlist is empty.")
else:
    for ticker in list(st.session_state.watchlist):
        c1, c2, c3 = st.columns([2, 5, 1])
        c1.write(f"`{ticker}`")
        c2.write(display_names.get(ticker, ticker))
        if c3.button("Remove", key=f"rm_wl_{ticker}"):
            st.session_state.watchlist.remove(ticker)
            save_config()
            st.cache_data.clear()
            st.rerun()

st.caption(
    "Weights are capped market-cap style, chosen so no single constituent dominates the basket. "
    "Adding to the index proportionally scales existing weights so the total stays at 100%."
)
