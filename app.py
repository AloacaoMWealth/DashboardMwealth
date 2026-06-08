from __future__ import annotations

import re
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ==============================
# CONFIGURAÇÕES GERAIS
# ==============================

APP_TITLE = "Dashboard Macro M Wealth"
DATA_DIR = Path("data")
DEFAULT_FILE = DATA_DIR / "Controle de Clientes MWealth 2026.xlsx"
ASSETS_DIR = Path("assets")
LOGO_FILE = ASSETS_DIR / "Logo-M-Wealth.png"
MAIN_SHEET = "Controle Contas"
META_PL_EMPRESA = 500_000_000.00
DATA_META = date(2026, 12, 31)

AZUL_ESCURO = "#003348"
AZUL_MEDIO = "#0B5C7A"
AZUL_CLARO = "#2A9FD6"
LARANJA = "#D99C3A"
VERDE = "#167A3C"
VERMELHO = "#B42318"
CINZA_FUNDO = "#F4F7FA"
CINZA_BORDA = "#D9E2EC"
CINZA_TEXTO = "#111827"

PALETA_INSTITUCIONAL = [
    "#003348",  # azul escuro
    "#0B5C7A",  # azul médio
    "#2A9FD6",  # azul claro
    "#6B8796",  # azul acinzentado
    "#D99C3A",  # dourado institucional
    "#8FA7B3",  # cinza azulado
    "#4B5563",  # cinza escuro
    "#A7C7D9",  # azul suave
    "#B8A06A",  # dourado discreto
    "#CBD5E1",  # cinza claro
]

MAPA_CORES_FIXAS = {
    "Onshore": "#003348",
    "Offshore": "#2A9FD6",
    "Wealth": "#0B5C7A",
    "Tesouraria": "#6B8796",
    "Não informado": "#CBD5E1",
    "Outros": "#8FA7B3",
}

CORRETORAS_OFFSHORE = {
    "AVENUE",
    "CHARLES SCHWAB",
    "CHARLES SHWAB",
    "SCHWAB",
    "XP US",
    "XP INTERNACIONAL",
    "XP INTERNATIONAL",
    "INTERACTIVE BROKERS",
    "IBKR",
    "PERSHING",
    "ATIVORE",
}

# ==============================
# CONFIGURAÇÃO DA PÁGINA
# ==============================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================
# ESTILO INSTITUCIONAL
# ==============================

st.markdown(
    f"""
    <style>
        :root {{
            --primary: {AZUL_ESCURO};
            --secondary: {AZUL_MEDIO};
            --text: {CINZA_TEXTO};
            --muted: #4B5563;
        }}

        .stApp {{
            background: {CINZA_FUNDO};
            color: var(--text) !important;
        }}
        .stApp * {{
            color: inherit;
        }}
        section[data-testid="stSidebar"] {{
            background: {AZUL_ESCURO};
        }}
        section[data-testid="stSidebar"] * {{
            color: white !important;
        }}
        .block-container {{
            padding-top: 1.0rem;
            padding-bottom: 1.5rem;
            max-width: 1500px;
        }}
        h1, h2, h3, h4, h5, h6, p, span, label {{
            color: {CINZA_TEXTO};
        }}
        h1, h2, h3 {{
            color: {AZUL_ESCURO} !important;
            letter-spacing: -0.02em;
        }}
        .main-title {{
            background: linear-gradient(90deg, {AZUL_ESCURO}, {AZUL_MEDIO});
            color: white !important;
            padding: 18px 26px;
            border-radius: 18px;
            margin-bottom: 14px;
            box-shadow: 0 8px 24px rgba(0, 51, 72, 0.16);
        }}
        .main-title h1 {{
            color: white !important;
            margin: 0;
            font-size: 30px;
        }}
        .main-title p {{
            color: rgba(255,255,255,0.88) !important;
            margin: 6px 0 0 0;
            font-size: 14px;
        }}
        .kpi-card {{
            background: white;
            border: 1px solid {CINZA_BORDA};
            border-radius: 16px;
            padding: 14px 16px;
            min-height: 102px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
        }}
        .kpi-label {{
            font-size: 13px;
            color: #4B5563 !important;
            font-weight: 700;
            margin-bottom: 10px;
        }}
        .kpi-value {{
            font-size: 23px;
            color: {AZUL_ESCURO} !important;
            font-weight: 800;
            line-height: 1.1;
        }}
        .kpi-sub {{
            font-size: 12px;
            color: #6B7280 !important;
            margin-top: 8px;
        }}
        .kpi-separator {{
            width: 100%;
            height: 1px;
            background: linear-gradient(
                90deg,
                rgba(0, 51, 72, 0.00),
                rgba(0, 51, 72, 0.18),
                rgba(0, 51, 72, 0.00)
            );
            margin: 14px 0 16px 0;
        }}
        .section-card {{
            background: white;
            border: 1px solid {CINZA_BORDA};
            border-radius: 18px;
            padding: 14px;
            margin-bottom: 12px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.045);
        }}
        div[data-testid="stMetric"] {{
            background: white;
            border: 1px solid {CINZA_BORDA};
            border-radius: 16px;
            padding: 14px 16px;
            box-shadow: 0 5px 16px rgba(15, 23, 42, 0.04);
        }}
        div[data-testid="stMetricLabel"] p {{
            color: #4B5563 !important;
        }}
        div[data-testid="stMetricValue"] {{
            color: {AZUL_ESCURO} !important;
        }}
        .stDataFrame {{
            border-radius: 12px;
            overflow: hidden;
        }}


        .delta-positive {{
            color: {VERDE} !important;
            font-weight: 800;
        }}
        .delta-negative {{
            color: {VERMELHO} !important;
            font-weight: 800;
        }}
        .kpi-value .delta-arrow {{
            font-size: 20px;
            margin-left: 6px;
            vertical-align: 2px;
        }}
        div[data-testid="stVerticalBlock"] {{
            gap: 0.75rem;
        }}

        /* Inputs e multiselects legíveis no tema claro */
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        textarea,
        input {{
            background-color: white !important;
            color: {CINZA_TEXTO} !important;
            border-color: {CINZA_BORDA} !important;
        }}
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] div,
        div[data-baseweb="popover"] span,
        div[data-baseweb="popover"] div {{
            color: {CINZA_TEXTO} !important;
        }}
        div[data-baseweb="tag"] {{
            background-color: #E8F2F7 !important;
            color: {AZUL_ESCURO} !important;
        }}
        div[data-baseweb="tag"] span {{
            color: {AZUL_ESCURO} !important;
        }}
        .stCaption, [data-testid="stCaptionContainer"] {{
            color: #4B5563 !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================
# FUNÇÕES UTILITÁRIAS
# ==============================

def br_money(value: float | int | None) -> str:
    if pd.isna(value):
        value = 0
    value = float(value)
    sign = "-" if value < 0 else ""
    value = abs(value)
    txt = f"{value:,.2f}"
    txt = txt.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}R$ {txt}"


def br_percent(value: float | int | None) -> str:
    if pd.isna(value):
        value = 0
    txt = f"{float(value):,.2%}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


def br_number(value: float | int | None, decimals: int = 0) -> str:
    if pd.isna(value):
        value = 0
    txt = f"{float(value):,.{decimals}f}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


def normalize_text(value) -> str:
    if pd.isna(value):
        return "Não informado"
    value = str(value).replace("\xa0", " ").strip()
    return value if value else "Não informado"


def normalize_key(value) -> str:
    return normalize_text(value).upper().replace("  ", " ")


def make_unique_columns(columns) -> List[str]:
    """Evita quebra quando o Excel possui colunas iguais após strip."""
    seen = {}
    result = []
    for col in columns:
        col = str(col).replace("\xa0", " ").strip()
        if col not in seen:
            seen[col] = 0
            result.append(col)
        else:
            seen[col] += 1
            result.append(f"{col}.{seen[col]}")
    return result


def parse_money_value(value) -> float:
    """Converte número vindo do Excel sem multiplicar indevidamente por 100.

    O problema principal da base é misturar números reais do Excel com textos em BR
    (1.177.994,77) e textos em formato americano (878446.54). A regra abaixo trata
    cada célula individualmente.
    """
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)

    s = str(value)
    s = s.replace("R$", "")
    s = s.replace("US$", "")
    s = s.replace("$", "")
    s = s.replace("\xa0", "")
    s = s.replace(" ", "")
    s = s.strip()

    if s in {"", "-", "nan", "None"}:
        return 0.0

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    if s.startswith("-"):
        negative = True
        s = s[1:]

    has_comma = "," in s
    has_dot = "." in s

    if has_comma and has_dot:
        # O último separador define o decimal.
        if s.rfind(",") > s.rfind("."):
            # Brasil: 1.234.567,89
            s = s.replace(".", "").replace(",", ".")
        else:
            # EUA: 1,234,567.89
            s = s.replace(",", "")
    elif has_comma:
        # Vírgula decimal quando há 1 ou 2 casas ao final; caso contrário, separador de milhar.
        decimals = len(s.split(",")[-1])
        if decimals in (1, 2):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_dot:
        parts = s.split(".")
        if len(parts) > 2:
            # 1.234.567 ou 1.234.567.89; mantém a última como decimal só se fizer sentido
            if len(parts[-1]) in (1, 2):
                s = "".join(parts[:-1]) + "." + parts[-1]
            else:
                s = "".join(parts)
        else:
            # 878446.54 deve continuar decimal, não virar 87.844.654
            pass

    try:
        out = float(s)
    except ValueError:
        out = 0.0
    return -out if negative else out


def clean_numeric(series: pd.Series) -> pd.Series:
    return series.apply(parse_money_value).astype(float).fillna(0)


def extract_pl_columns(df: pd.DataFrame) -> Dict[str, pd.Timestamp]:
    pl_cols = {}
    pattern = re.compile(r"^PL\s+(\d{2}/\d{2}/\d{4})$", flags=re.IGNORECASE)
    for col in df.columns:
        match = pattern.match(str(col).strip())
        if match:
            dt = pd.to_datetime(match.group(1), dayfirst=True, errors="coerce")
            if pd.notna(dt):
                pl_cols[col] = dt
    return dict(sorted(pl_cols.items(), key=lambda item: item[1]))


def classify_region(corretora: str) -> str:
    corretora_clean = normalize_key(corretora)
    if any(term in corretora_clean for term in CORRETORAS_OFFSHORE):
        return "Offshore"
    return "Onshore"


def make_line_meta(months: List[pd.Timestamp], first_value: float, target_value: float) -> pd.Series:
    if len(months) == 0:
        return pd.Series(dtype=float)
    first_date = min(months)
    last_date = pd.Timestamp(DATA_META)
    total_days = max((last_date - first_date).days, 1)
    values = []
    for dt in months:
        progress = min(max((dt - first_date).days / total_days, 0), 1)
        values.append(first_value + (target_value - first_value) * progress)
    return pd.Series(values, index=months)


def kpi_card(label: str, value: str, sub: str = ""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def variation_card(label: str, value: float, pct: float, sub: str = ""):
    positive = value >= 0
    arrow = "▲" if positive else "▼"
    klass = "delta-positive" if positive else "delta-negative"
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value {klass}">{br_money(value)}<span class="delta-arrow">{arrow}</span></div>
            <div class="kpi-sub">{sub} | {br_percent(pct)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def separator():
    st.markdown('<div class="kpi-separator"></div>', unsafe_allow_html=True)

def standard_layout(fig: go.Figure, height: int = 360, legend: bool = True) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=24, r=24, t=50, b=30),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color=CINZA_TEXTO, size=12),
        title_font=dict(color=CINZA_TEXTO, size=16),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color=CINZA_TEXTO, size=11),
            bgcolor="rgba(255,255,255,0.85)",
        ) if legend else None,
    )
    fig.update_xaxes(showgrid=False, linecolor="#D1D5DB", tickfont=dict(color=CINZA_TEXTO))
    fig.update_yaxes(gridcolor="#E5E7EB", linecolor="#D1D5DB", tickfont=dict(color=CINZA_TEXTO))
    return fig


def bar_chart(df: pd.DataFrame, x: str, y: str, title: str, horizontal: bool = False, height: int = 360):
    if df.empty:
        st.info("Sem dados para exibir neste gráfico.")
        return
    plot_df = df.copy()
    if horizontal:
        fig = px.bar(plot_df, x=y, y=x, orientation="h", title=title, text=y)
        fig.update_traces(
            marker_color=AZUL_MEDIO,
            texttemplate="R$ %{x:,.0f}",
            textposition="auto",
            cliponaxis=False,
            hovertemplate="%{y}<br>PL: R$ %{x:,.2f}<extra></extra>",
        )
        fig.update_layout(yaxis=dict(autorange="reversed"), margin=dict(l=24, r=110, t=50, b=30))
        fig.update_xaxes(tickprefix="R$ ", tickformat=",.0f")
    else:
        fig = px.bar(plot_df, x=x, y=y, title=title, text=y)
        fig.update_traces(
            marker_color=AZUL_MEDIO,
            texttemplate="R$ %{y:,.0f}",
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{x}<br>PL: R$ %{y:,.2f}<extra></extra>",
        )
        fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
    fig = standard_layout(fig, height=height, legend=False)
    st.plotly_chart(fig, use_container_width=True)


def value_bar_chart(df: pd.DataFrame, category: str, value: str, title: str, height: int = 340, money: bool = True):
    if df.empty or df[value].fillna(0).abs().sum() == 0:
        st.info("Sem dados para exibir neste gráfico.")
        return
    d = df.copy()
    if money:
        total = d[value].sum()
        d["Participação"] = np.where(total != 0, d[value] / total, 0)
        d["Texto"] = d.apply(lambda r: f"{br_money(r[value])} | {br_percent(r['Participação'])}", axis=1)
        hover = "%{y}<br>Valor: R$ %{x:,.2f}<br>Participação: %{customdata:.2%}<extra></extra>"
        tickprefix = "R$ "
        customdata = d["Participação"]
    else:
        total = d[value].sum()
        d["Participação"] = np.where(total != 0, d[value] / total, 0)
        d["Texto"] = d.apply(lambda r: f"{br_number(r[value])} | {br_percent(r['Participação'])}", axis=1)
        hover = "%{y}<br>Valor: %{x:,.0f}<br>Participação: %{customdata:.2%}<extra></extra>"
        tickprefix = ""
        customdata = d["Participação"]
    d = d.sort_values(value, ascending=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=d[value], y=d[category], orientation="h", marker_color=AZUL_MEDIO,
        text=d["Texto"], textposition="auto", cliponaxis=False,
        customdata=customdata, hovertemplate=hover,
    ))
    fig.update_layout(title=title, margin=dict(l=24, r=150, t=50, b=30))
    fig.update_xaxes(tickprefix=tickprefix, tickformat=",.0f")
    fig = standard_layout(fig, height=height, legend=False)
    st.plotly_chart(fig, use_container_width=True)


def historical_bar_chart(
    df: pd.DataFrame,
    category: str,
    title: str,
    top_n: int | None = None,
    percent: bool = False,
    height: int = 420,
):
    """Gráfico histórico de barras empilhadas por mês."""
    if df.empty:
        st.info("Sem dados para exibir neste gráfico.")
        return

    if category not in df.columns:
        st.info(f"Coluna {category} não encontrada na base.")
        return

    d = df[["Data", category, "PL"]].copy()
    d[category] = d[category].apply(normalize_text)
    d = d[d["PL"].fillna(0) != 0]

    if d.empty:
        st.info("Sem PL histórico para exibir neste gráfico.")
        return

    if top_n is not None:
        top_categories = (
            d.groupby(category)["PL"]
            .sum()
            .sort_values(ascending=False)
            .head(top_n)
            .index
        )
        d[category] = np.where(d[category].isin(top_categories), d[category], "Outras")

    grouped = (
        d.groupby(["Data", category], as_index=False)["PL"]
        .sum()
        .sort_values("Data")
    )
    grouped["Mês"] = grouped["Data"].dt.strftime("%m/%Y")

    category_order = {
        "Mês": grouped.drop_duplicates("Data").sort_values("Data")["Mês"].tolist()
    }

    if percent:
        total_mes = grouped.groupby("Data")["PL"].transform("sum")
        grouped["Participação"] = np.where(total_mes != 0, grouped["PL"] / total_mes, 0)

        fig = px.bar(
            grouped,
            x="Mês",
            y="Participação",
            color=category,
            category_orders=category_order,
            title=None,
        )
        fig.update_yaxes(tickformat=".0%")
        fig.update_traces(
            hovertemplate=(
                "Mês: %{x}<br>"
                f"{category}: %{{legendgroup}}<br>"
                "Participação: %{y:.2%}<extra></extra>"
            )
        )
    else:
        fig = px.bar(
            grouped,
            x="Mês",
            y="PL",
            color=category,
            category_orders=category_order,
            title=None,
        )
        fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
        fig.update_traces(
            hovertemplate=(
                "Mês: %{x}<br>"
                f"{category}: %{{legendgroup}}<br>"
                "PL: R$ %{y:,.2f}<extra></extra>"
            )
        )

    fig.update_layout(
        barmode="stack",
        xaxis_title="Mês",
        yaxis_title="Participação" if percent else "PL",
        legend_title_text=category,
        margin=dict(l=24, r=24, t=30, b=60),
    )
    fig = standard_layout(fig, height=height, legend=True)
    st.plotly_chart(fig, use_container_width=True)

def donut_chart(df: pd.DataFrame, names: str, values: str, title: str, height: int = 340):
    if df.empty or df[values].sum() == 0:
        st.info("Sem dados para exibir neste gráfico.")
        return

    fig = px.pie(
        df,
        names=names,
        values=values,
        hole=0.56,
        title=None,
        color=names,
        color_discrete_map=MAPA_CORES_FIXAS,
        color_discrete_sequence=PALETA_INSTITUCIONAL,
    )

    fig.update_traces(
        textposition="inside",
        textinfo="percent+label",
        marker=dict(line=dict(color="white", width=2)),
    )

    fig = standard_layout(fig, height=height, legend=True)
    st.plotly_chart(fig, use_container_width=True)

def _extract_close_series(data: pd.DataFrame) -> pd.Series:
    """Extrai uma série de fechamento mesmo quando o yfinance devolve MultiIndex."""
    if data is None or data.empty:
        return pd.Series(dtype=float)

    if isinstance(data.columns, pd.MultiIndex):
        close_cols = [col for col in data.columns if str(col[0]).lower() == "close"]
        if not close_cols:
            close_cols = [col for col in data.columns if "close" in str(col).lower()]
        if not close_cols:
            return pd.Series(dtype=float)
        close = data[close_cols[0]]
    elif "Close" in data.columns:
        close = data["Close"]
    elif "Adj Close" in data.columns:
        close = data["Adj Close"]
    else:
        return pd.Series(dtype=float)

    if isinstance(close, pd.DataFrame):
        close = close.iloc[:, 0]

    close.index = pd.to_datetime(close.index).tz_localize(None)
    close = pd.to_numeric(close, errors="coerce").dropna().astype(float)
    close = close[close > 0]
    return close.sort_index()


@st.cache_data(show_spinner="Buscando cotações USDBRL pelo yfinance...")
def get_usdbrl_quotes(reference_dates: Tuple[str, ...]) -> Dict[str, float]:
    """Busca fechamento USDBRL para cada data de PL usando apenas yfinance.

    Regras:
    - tenta mais de um ticker aceito pelo Yahoo Finance;
    - usa o último fechamento disponível até a data de referência;
    - ignora meses sem PL preenchido, para não tentar buscar datas futuras/vazias;
    - se a data de referência for futura, usa o último fechamento disponível até hoje.
    """
    try:
        import yfinance as yf
    except Exception as exc:
        raise RuntimeError(
            "Biblioteca yfinance não instalada. Rode: pip install yfinance"
        ) from exc

    if not reference_dates:
        return {}

    dates = sorted({pd.to_datetime(d).date() for d in reference_dates})
    today = date.today()
    start = min(dates) - timedelta(days=15)
    # yfinance usa end exclusivo. Não adianta pedir muito no futuro.
    end_limit = min(max(dates) + timedelta(days=5), today + timedelta(days=1))

    tickers = ["USDBRL=X", "BRL=X"]
    errors = []
    close = pd.Series(dtype=float)

    for ticker in tickers:
        try:
            data = yf.download(
                ticker,
                start=start.isoformat(),
                end=end_limit.isoformat(),
                progress=False,
                auto_adjust=False,
                threads=False,
                timeout=20,
            )
            close = _extract_close_series(data)
            if not close.empty:
                break
        except Exception as exc:
            errors.append(f"{ticker}/download: {exc}")

        try:
            hist = yf.Ticker(ticker).history(
                start=start.isoformat(),
                end=end_limit.isoformat(),
                auto_adjust=False,
                timeout=20,
            )
            close = _extract_close_series(hist)
            if not close.empty:
                break
        except Exception as exc:
            errors.append(f"{ticker}/history: {exc}")

    if close.empty:
        detalhe = " | ".join(errors[-4:]) if errors else "sem detalhe retornado pelo Yahoo Finance"
        raise RuntimeError(
            "Não foi possível baixar USDBRL pelo yfinance. "
            "Confira se o ambiente tem internet e se yfinance está instalado/atualizado. "
            f"Detalhe: {detalhe}"
        )

    quotes = {}
    for d in dates:
        ref = min(pd.Timestamp(d), close.index.max())
        available = close[close.index <= ref]
        if available.empty:
            available = close[close.index >= ref]
        if available.empty:
            raise RuntimeError(f"Não encontrei cotação USDBRL próxima de {d.strftime('%d/%m/%Y')}.")
        quotes[pd.Timestamp(d).strftime("%Y-%m-%d")] = float(available.iloc[-1])

    return quotes


def get_latest_valid_pl_col(df: pd.DataFrame, pl_cols: Dict[str, pd.Timestamp]) -> str:
    """Usa a última coluna de PL que tenha pelo menos algum valor preenchido e diferente de zero."""
    for col in reversed(list(pl_cols.keys())):
        series = df[col].fillna(0)
        if series.abs().sum() > 0:
            return col
    return list(pl_cols.keys())[-1]


@st.cache_data(show_spinner="Carregando e tratando a base de clientes...")
def load_data(file_path: str, file_mtime: float) -> Tuple[pd.DataFrame, pd.DataFrame, List[str], Dict[str, pd.Timestamp], pd.DataFrame, str]:
    df = pd.read_excel(file_path, sheet_name=MAIN_SHEET)
    df.columns = make_unique_columns(df.columns)
    df["_row_id"] = np.arange(len(df))

    pl_cols = extract_pl_columns(df)
    if not pl_cols:
        raise ValueError("Não encontrei colunas de PL no padrão 'PL dd/mm/aaaa'.")

    text_cols = [
        "Corretora", "Grupo Geral", "Grupo Familiar", "Cliente", "PF/ PJ", "Canal",
        "Cliente - Corretora", "Conta", "UF", "Consultor", "Tipo de Marcação",
        "Perfil Carteira/ Renda", "Perfil de Investidor", "Consolidador", "Liquidez",
        "Observações", "Possui Crédito", "Possui Previdência", "Calculadora de IR", "E-mail",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(normalize_text)

    if "Conta" in df.columns:
        df["Conta"] = df["Conta"].astype(str).str.replace(".0", "", regex=False).str.strip()

    # 1) Limpa os PLs sem destruir casas decimais.
    for col in pl_cols.keys():
        df[col] = clean_numeric(df[col])

    # 2) Classifica offshore/onshore pela corretora.
    df["Região"] = df.get("Corretora", pd.Series(["Não informado"] * len(df))).apply(classify_region)

    # 3) Busca USDBRL exclusivamente no yfinance e converte apenas contas offshore.
    # Usa somente colunas de PL preenchidas. Isso evita distorção por mês futuro/vazio.
    active_pl_cols = {col: dt for col, dt in pl_cols.items() if df[col].abs().sum() > 0}
    if not active_pl_cols:
        raise ValueError("As colunas de PL foram identificadas, mas todas estão zeradas/vazias.")

    reference_dates = tuple(pd.Timestamp(dt).strftime("%Y-%m-%d") for dt in active_pl_cols.values())
    usdbrl_quotes = get_usdbrl_quotes(reference_dates)

    fx_records = []
    converted_pl_cols = {}
    for col, dt in active_pl_cols.items():
        key = pd.Timestamp(dt).strftime("%Y-%m-%d")
        fx = float(usdbrl_quotes[key])
        converted_col = f"{col} BRL"
        df[converted_col] = np.where(df["Região"].eq("Offshore"), df[col] * fx, df[col])
        converted_pl_cols[converted_col] = dt
        fx_records.append({"Data": pd.Timestamp(dt), "Coluna PL": col, "USDBRL usado": fx})

    latest_col = get_latest_valid_pl_col(df, converted_pl_cols)

    df["PL Atual"] = df[latest_col].fillna(0)
    df["Data PL Atual"] = converted_pl_cols[latest_col]

    id_cols = [c for c in df.columns if c not in converted_pl_cols.keys()]
    long_df = df.melt(
        id_vars=id_cols,
        value_vars=list(converted_pl_cols.keys()),
        var_name="Coluna PL Convertida",
        value_name="PL",
    )
    long_df["Data"] = long_df["Coluna PL Convertida"].map(converted_pl_cols)
    long_df["Mês"] = long_df["Data"].dt.to_period("M").astype(str)
    long_df["Ano"] = long_df["Data"].dt.year
    long_df["PL"] = long_df["PL"].fillna(0)

    fx_df = pd.DataFrame(fx_records)
    return df, long_df, list(converted_pl_cols.keys()), converted_pl_cols, fx_df, latest_col


def get_base_file() -> Path:
    DATA_DIR.mkdir(exist_ok=True)
    return DEFAULT_FILE


def apply_filters(base: pd.DataFrame, filtros: Dict[str, List[str]]) -> pd.DataFrame:
    out = base.copy()
    for col, selected in filtros.items():
        if col in out.columns and selected:
            out = out[out[col].isin(selected)]
    return out


def filtered_multiselect(label: str, df: pd.DataFrame, col: str):
    if col not in df.columns:
        return []
    options = sorted([x for x in df[col].dropna().unique().tolist() if str(x).strip()])
    return st.multiselect(label, options=options, default=[], placeholder="Todos")



def get_col_by_date(pl_col_dates: Dict[str, pd.Timestamp], selected_date: pd.Timestamp) -> str:
    selected_date = pd.Timestamp(selected_date)
    for col, dt in pl_col_dates.items():
        if pd.Timestamp(dt) == selected_date:
            return col
    raise KeyError(f"Data não encontrada nas colunas de PL: {selected_date}")


def month_display(dt: pd.Timestamp) -> str:
    return pd.Timestamp(dt).strftime("%d/%m/%Y")


def calc_delta_from_mensal(mensal: pd.DataFrame, end_dt: pd.Timestamp, start_dt: pd.Timestamp) -> Tuple[float, float, float, float]:
    if mensal.empty:
        return 0.0, 0.0, 0.0, 0.0
    lookup = mensal.set_index("Data")["PL"]
    end_value = float(lookup.get(pd.Timestamp(end_dt), 0.0))
    start_value = float(lookup.get(pd.Timestamp(start_dt), 0.0))
    delta = end_value - start_value
    pct = delta / start_value if start_value else 0.0
    return start_value, end_value, delta, pct

def calc_meses_restantes_meta(data_base: pd.Timestamp, data_meta: date) -> int:
    """
    Calcula quantos meses ainda restam para atingir a meta final.
    Usa a data-base do PL atual como referência.
    """
    data_base = pd.Timestamp(data_base).date()

    meses = (data_meta.year - data_base.year) * 12 + (data_meta.month - data_base.month)
    return max(meses, 1)

def format_table_money(df: pd.DataFrame, money_cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    for col in money_cols:
        if col in out.columns:
            out[col] = out[col].apply(br_money)
    return out

# ==============================
# SIDEBAR
# ==============================

base_file = get_base_file()

if LOGO_FILE.exists():
    st.sidebar.image(str(LOGO_FILE), use_container_width=True)
else:
    st.sidebar.markdown("# M Wealth")

st.sidebar.markdown("Dashboard macro do escritório")

if not base_file.exists():
    st.error("Arquivo base não encontrado. Inclua a planilha em data/Controle de Clientes MWealth 2026.xlsx ou envie pelo upload lateral.")
    st.stop()

try:
    raw_df, long_df, pl_cols, pl_col_dates, fx_df, latest_col = load_data(str(base_file), base_file.stat().st_mtime)
except Exception as exc:
    st.error(f"Erro ao carregar a base: {exc}")
    st.stop()

latest_date = pl_col_dates[latest_col]

pagina = st.sidebar.radio(
    "Navegação",
    ["Dashboard Macro", "Análise Dinâmica", "Histórico de Informações", "Base de Dados"],
    index=0,
)

st.sidebar.divider()
st.sidebar.caption(f"Base: {base_file.name}")
st.sidebar.caption(f"PL Atual: {latest_col.replace(' BRL', '')}")

# ==============================
# HEADER
# ==============================

st.markdown(
    f"""
    <div class="main-title">
        <h1>{APP_TITLE}</h1>
        <p>Acompanhamento institucional de patrimônio, meta comercial, corretoras, canais e segmentação da base.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==============================
# DASHBOARD MACRO
# ==============================

if pagina == "Dashboard Macro":
    current_pl = raw_df["PL Atual"].sum()
    pct_meta = current_pl / META_PL_EMPRESA if META_PL_EMPRESA else 0
    gap_meta = META_PL_EMPRESA - current_pl

    today = date.today()
    meses_restantes = max((DATA_META.year - today.year) * 12 + (DATA_META.month - today.month), 0)
    semanas_restantes = max((DATA_META - today).days // 7, 0)

    meses_restantes_meta = calc_meses_restantes_meta(latest_date, DATA_META)
    meta_captacao_mensal_futura = gap_meta / meses_restantes_meta if meses_restantes_meta else 0

    grupos = raw_df.groupby("Grupo Familiar", dropna=False)["PL Atual"].sum().reset_index()
    grupos_validos = grupos[grupos["Grupo Familiar"] != "Não informado"]
    qtd_grupos = int(grupos_validos["Grupo Familiar"].nunique())
    pl_medio_grupo = grupos_validos["PL Atual"].sum() / qtd_grupos if qtd_grupos else 0

    mensal = long_df.groupby("Data", as_index=False)["PL"].sum().sort_values("Data")
    mensal["Meta Linear"] = make_line_meta(mensal["Data"].tolist(), mensal["PL"].iloc[0], META_PL_EMPRESA).values
    mensal["Percentual da Meta"] = mensal["PL"] / META_PL_EMPRESA

    meta_linear_atual = float(mensal.loc[mensal["Data"].eq(pd.Timestamp(latest_date)), "Meta Linear"].iloc[0]) if not mensal.empty else 0.0
    gap_meta_linear = meta_linear_atual - current_pl
    status_linear = "abaixo" if gap_meta_linear > 0 else "acima"

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        kpi_card("PL total sob gestão", br_money(current_pl), f"Data-base: {latest_date.strftime('%d/%m/%Y')}")
    with col2:
        kpi_card("Meta patrimonial final", br_money(META_PL_EMPRESA), f"Realizado: {br_percent(pct_meta)}")
    with col3:
        kpi_card("Desvio para meta final", br_money(gap_meta), f"Meses restantes: {meses_restantes} | Semanas: {semanas_restantes}")
    with col4:
        kpi_card("Meta linear atual", br_money(meta_linear_atual), f"Executado: {br_percent(current_pl / meta_linear_atual if meta_linear_atual else 0)}")
    with col5:
        kpi_card("Desvio para meta linear", br_money(abs(gap_meta_linear)), f"{status_linear} da meta linear")

    separator()

    datas_mensais = mensal["Data"].tolist()
    latest_idx = datas_mensais.index(pd.Timestamp(latest_date)) if pd.Timestamp(latest_date) in datas_mensais else len(datas_mensais) - 1
    prev_idx = max(latest_idx - 1, 0)
    start_prev, end_latest, delta_latest, pct_delta_latest = calc_delta_from_mensal(mensal, datas_mensais[latest_idx], datas_mensais[prev_idx])

    hoje_ts = pd.Timestamp(date.today())
    meses_fechados = mensal[mensal["Data"] < hoje_ts.replace(day=1)]
    if len(meses_fechados) >= 2:
        closed_end = meses_fechados["Data"].iloc[-1]
        closed_start = meses_fechados["Data"].iloc[-2]
    elif len(mensal) >= 2:
        closed_end = mensal["Data"].iloc[-1]
        closed_start = mensal["Data"].iloc[-2]
    else:
        closed_end = mensal["Data"].iloc[-1]
        closed_start = mensal["Data"].iloc[-1]
    _, _, delta_closed, pct_delta_closed = calc_delta_from_mensal(mensal, closed_end, closed_start)

        # Segunda linha: variações, PL anterior, captação necessária e grupos
    v1, v2, v3, v4, v5 = st.columns(5, gap="medium")

    with v1:
        variation_card(
            "Variação período atual",
            delta_latest,
            pct_delta_latest,
            f"{month_display(datas_mensais[prev_idx])} até {month_display(datas_mensais[latest_idx])}"
        )

    with v2:
        variation_card(
            "Variação último mês fechado",
            delta_closed,
            pct_delta_closed,
            f"{month_display(closed_start)} até {month_display(closed_end)}"
        )

    with v3:
        kpi_card(
            "PL mês anterior",
            br_money(start_prev),
            f"Referência: {month_display(datas_mensais[prev_idx])}"
        )

    with v4:
        kpi_card(
            "Captação mensal necessária",
            br_money(meta_captacao_mensal_futura),
            f"Para atingir a meta em {meses_restantes_meta} meses"
        )

    with v5:
        kpi_card(
            "Grupos familiares",
            br_number(qtd_grupos),
            f"PL médio por grupo: {br_money(pl_medio_grupo)}"
        )

    separator()

    st.subheader("Evolução do PL vs. Meta Patrimonial")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=mensal["Data"], y=mensal["PL"], mode="lines+markers", name="PL Realizado",
        line=dict(color=AZUL_CLARO, width=3), marker=dict(size=7),
        hovertemplate="%{x|%d/%m/%Y}<br>PL: R$ %{y:,.2f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=mensal["Data"], y=mensal["Meta Linear"], mode="lines", name="Meta Linear",
        line=dict(color=AZUL_ESCURO, width=3),
        hovertemplate="%{x|%d/%m/%Y}<br>Meta: R$ %{y:,.2f}<extra></extra>",
    ))
    fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
    fig = standard_layout(fig, height=420, legend=True)
    st.plotly_chart(fig, use_container_width=True)

    separator()

    c1, c2 = st.columns([1.25, 1])

    with c1:
        st.subheader("PL por Corretora")
        corretora = raw_df.groupby("Corretora", as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False)
        bar_chart(corretora, "Corretora", "PL Atual", "Distribuição por Corretora", horizontal=False, height=390)
    
    with c2:
        st.subheader("Onshore vs. Offshore")
        regiao = raw_df.groupby("Região", as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False)
        donut_chart(regiao, "Região", "PL Atual", "Proporção de PL", height=390)
    
    c3, c4 = st.columns(2)
    with c3:
        st.subheader("Proporção por Canal")
        if "Canal" in raw_df.columns:
            canal = raw_df.groupby("Canal", as_index=False).agg(
                **{"PL Atual": ("PL Atual", "sum"), "Clientes": ("Cliente", "nunique"), "Grupos": ("Grupo Familiar", "nunique")}
            ).sort_values("PL Atual", ascending=False)
            value_bar_chart(canal, "Canal", "PL Atual", "Participação de PL por Canal", height=285, money=True)
            canal_display = canal.copy()
            total_canal = canal_display["PL Atual"].sum()
            canal_display["Participação"] = np.where(total_canal != 0, canal_display["PL Atual"] / total_canal, 0)
            canal_display["PL Atual"] = canal_display["PL Atual"].apply(br_money)
            canal_display["Participação"] = canal_display["Participação"].apply(br_percent)
            st.dataframe(canal_display, use_container_width=True, hide_index=True)
        else:
            st.info("Coluna Canal não encontrada na base.")
    
    with c4:
        st.subheader("Segmentação da Base por Faixa de PL")

        grupo_pl = (
            raw_df
            .groupby("Grupo Familiar", as_index=False)
            .agg(
                PL=("PL Atual", "sum"),
                Clientes=("Cliente", "nunique")
            )
        )

        grupo_pl = grupo_pl[grupo_pl["Grupo Familiar"] != "Não informado"].copy()

        bins = [-0.01, 5_000_000, 10_000_000, 15_000_000, 30_000_000, 100_000_000, np.inf]
        labels = [
            "Até R$ 5 MM",
            "R$ 5 MM a R$ 10 MM",
            "R$ 10 MM a R$ 15 MM",
            "R$ 15 MM a R$ 30 MM",
            "R$ 30 MM a R$ 100 MM",
            "Acima de R$ 100 MM",
        ]

        grupo_pl["Faixa"] = pd.cut(grupo_pl["PL"], bins=bins, labels=labels)

        faixa = (
            grupo_pl
            .groupby("Faixa", observed=False)
            .agg(
                Grupos=("Grupo Familiar", "nunique"),
                Clientes=("Clientes", "sum"),
                PL=("PL", "sum")
            )
            .reset_index()
        )

        value_bar_chart(
            faixa,
            "Faixa",
            "Grupos",
            "Quantidade de Grupos Familiares por Faixa",
            height=285,
            money=False
        )

        faixa_display = faixa.copy()
        faixa_display["PL"] = faixa_display["PL"].apply(br_money)

        st.dataframe(
            faixa_display,
            use_container_width=True,
            hide_index=True
        )
    
    c5, c6 = st.columns(2)
    with c5:
        st.subheader("Top 10 Clientes")
        top_clientes = raw_df.groupby("Cliente", as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False).head(10)
        bar_chart(top_clientes, "Cliente", "PL Atual", "Maiores Clientes por PL", horizontal=True, height=450)
    
    with c6:
        st.subheader("Top 10 Grupos Familiares")
        top_grupos = raw_df.groupby("Grupo Familiar", as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False).head(10)
        bar_chart(top_grupos, "Grupo Familiar", "PL Atual", "Maiores Grupos por PL", horizontal=True, height=450)
    
# ==============================
# ANÁLISE DETALHADA
# ==============================

elif pagina == "Análise Dinâmica":
    st.subheader("Análise Dinâmica da Base")
    st.caption("Use os filtros e o período para validar PL, variação patrimonial e evolução por consultor, cliente ou grupo familiar.")

    datas_disponiveis = list(pl_col_dates.values())
    labels_datas = {month_display(dt): pd.Timestamp(dt) for dt in datas_disponiveis}
    label_lista = list(labels_datas.keys())
    idx_jan_26 = next((i for i, dt in enumerate(datas_disponiveis) if pd.Timestamp(dt).year == 2026 and pd.Timestamp(dt).month == 1), 0)

    p1, p2 = st.columns(2)
    with p1:
        data_inicio_label = st.selectbox("Mês inicial da análise", label_lista, index=idx_jan_26)
    with p2:
        data_fim_label = st.selectbox("Mês final da análise", label_lista, index=len(label_lista) - 1)

    data_inicio = labels_datas[data_inicio_label]
    data_fim = labels_datas[data_fim_label]
    if data_inicio > data_fim:
        st.warning("O mês inicial está depois do mês final. Inverti o período para calcular corretamente.")
        data_inicio, data_fim = data_fim, data_inicio

    col_pl_inicio = get_col_by_date(pl_col_dates, data_inicio)
    col_pl_fim = get_col_by_date(pl_col_dates, data_fim)

    f1, f2, f3 = st.columns(3)
    f4, f5, f6 = st.columns(3)

    with f1:
        consultores = filtered_multiselect("Consultor", raw_df, "Consultor")
    with f2:
        clientes = filtered_multiselect("Cliente", raw_df, "Cliente")
    with f3:
        grupos_familiares = filtered_multiselect("Grupo Familiar", raw_df, "Grupo Familiar")
    with f4:
        corretoras = filtered_multiselect("Corretora", raw_df, "Corretora")
    with f5:
        canais = filtered_multiselect("Canal", raw_df, "Canal")
    with f6:
        tipo_pessoa = filtered_multiselect("PF ou PJ", raw_df, "PF/ PJ")

    filtros = {
        "Consultor": consultores,
        "Cliente": clientes,
        "Grupo Familiar": grupos_familiares,
        "Corretora": corretoras,
        "Canal": canais,
        "PF/ PJ": tipo_pessoa,
    }

    base_filtrada = apply_filters(raw_df, filtros).copy()
    base_filtrada["PL Inicial Período"] = base_filtrada[col_pl_inicio].fillna(0)
    base_filtrada["PL Final Período"] = base_filtrada[col_pl_fim].fillna(0)
    base_filtrada["Variação PL"] = base_filtrada["PL Final Período"] - base_filtrada["PL Inicial Período"]
    base_filtrada["Variação %"] = np.where(base_filtrada["PL Inicial Período"] != 0, base_filtrada["Variação PL"] / base_filtrada["PL Inicial Período"], 0)

    contas_filtradas = base_filtrada["Conta"].nunique() if "Conta" in base_filtrada.columns else len(base_filtrada)
    clientes_filtrados = base_filtrada["Cliente"].nunique() if "Cliente" in base_filtrada.columns else len(base_filtrada)
    grupos_filtrados = base_filtrada["Grupo Familiar"].nunique() if "Grupo Familiar" in base_filtrada.columns else 0
    pl_inicial_filtrado = base_filtrada["PL Inicial Período"].sum()
    pl_final_filtrado = base_filtrada["PL Final Período"].sum()
    delta_filtrado = pl_final_filtrado - pl_inicial_filtrado
    pct_delta_filtrado = delta_filtrado / pl_inicial_filtrado if pl_inicial_filtrado else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        kpi_card("PL inicial", br_money(pl_inicial_filtrado), f"Referência: {month_display(data_inicio)}")
    with k2:
        kpi_card("PL final", br_money(pl_final_filtrado), f"Referência: {month_display(data_fim)}")
    with k3:
        kpi_card("Variação patrimonial", br_money(delta_filtrado), f"{br_percent(pct_delta_filtrado)} no período")
    with k4:
        kpi_card("Clientes", br_number(clientes_filtrados), f"Contas: {br_number(contas_filtradas)}")
    with k5:
        kpi_card("Grupos familiares", br_number(grupos_filtrados), "Grupos únicos na seleção")

    row_ids = base_filtrada["_row_id"].tolist() if not base_filtrada.empty else []
    long_filtrado = long_df[long_df["_row_id"].isin(row_ids)] if row_ids else long_df.iloc[0:0]
    long_filtrado = long_filtrado[(long_filtrado["Data"] >= data_inicio) & (long_filtrado["Data"] <= data_fim)]

    st.subheader("Evolução Histórica da Seleção")
    evo = long_filtrado.groupby("Data", as_index=False)["PL"].sum().sort_values("Data")
    if evo.empty:
        st.info("Nenhum dado encontrado para os filtros selecionados.")
    else:
        fig = px.line(evo, x="Data", y="PL", markers=True, title="PL Mensal da Seleção")
        fig.update_traces(line_color=AZUL_CLARO, line_width=3)
        fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
        fig = standard_layout(fig, height=360, legend=False)
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("PL por Consultor")
        if "Consultor" in base_filtrada.columns:
            por_consultor = base_filtrada.groupby("Consultor", as_index=False).agg({"PL Final Período": "sum", "Variação PL": "sum"}).reset_index().sort_values("PL Final Período", ascending=False).head(20)
            bar_chart(por_consultor.rename(columns={"PL Final Período": "PL"}), "Consultor", "PL", "Ranking de Consultores por PL Final", horizontal=True, height=500)
    
    with c2:
        st.subheader("PL por Grupo Familiar")
        por_grupo = base_filtrada.groupby("Grupo Familiar", as_index=False).agg({"PL Final Período": "sum", "Variação PL": "sum"}).reset_index().sort_values("PL Final Período", ascending=False).head(20)
        bar_chart(por_grupo.rename(columns={"PL Final Período": "PL"}), "Grupo Familiar", "PL", "Ranking de Grupos por PL Final", horizontal=True, height=500)
    
    st.subheader("Maiores Variações Patrimoniais no Período")
    if "Consultor" in base_filtrada.columns and not base_filtrada.empty:
        variacao_consultor = base_filtrada.groupby("Consultor", as_index=False)["Variação PL"].sum().sort_values("Variação PL", ascending=False).head(15)
        bar_chart(variacao_consultor.rename(columns={"Variação PL": "Variação"}), "Consultor", "Variação", "Variação por Consultor", horizontal=True, height=420)

    st.subheader("Tabela Analítica")
    cols_show = [
        "Corretora", "Grupo Geral", "Grupo Familiar", "Cliente", "PF/ PJ", "Canal",
        "Conta", "UF", "Consultor", "Perfil Carteira/ Renda", "Perfil de Investidor", "Região",
        "PL Inicial Período", "PL Final Período", "Variação PL", "Variação %",
    ]
    cols_show = [c for c in cols_show if c in base_filtrada.columns]
    tabela = base_filtrada[cols_show].sort_values("PL Final Período", ascending=False).copy()
    tabela_display = format_table_money(tabela, ["PL Inicial Período", "PL Final Período", "Variação PL"])
    if "Variação %" in tabela_display.columns:
        tabela_display["Variação %"] = tabela["Variação %"].apply(br_percent)
    st.dataframe(tabela_display, use_container_width=True, hide_index=True)
    csv = tabela.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("Baixar tabela filtrada em CSV", csv, "base_filtrada_mwealth.csv", "text/csv")


# ==============================
# HISTÓRICO DE INFORMAÇÕES
# ==============================

elif pagina == "Histórico de Informações":
    st.subheader("Histórico de Informações")
    st.caption(
        "Evolução mês a mês da distribuição de PL por corretora, região "
        "(Onshore vs. Offshore) e canal."
    )

    datas_disponiveis = sorted(long_df["Data"].dropna().unique())
    labels_datas = {month_display(dt): pd.Timestamp(dt) for dt in datas_disponiveis}
    label_lista = list(labels_datas.keys())

    p1, p2, p3 = st.columns([1, 1, 1])
    with p1:
        data_inicio_label = st.selectbox(
            "Mês inicial",
            label_lista,
            index=0,
            key="hist_data_inicio"
        )
    with p2:
        data_fim_label = st.selectbox(
            "Mês final",
            label_lista,
            index=len(label_lista) - 1,
            key="hist_data_fim"
        )
    with p3:
        tipo_visualizacao = st.radio(
            "Visualização",
            ["PL financeiro", "Participação percentual"],
            horizontal=False,
            key="hist_tipo_visualizacao"
        )

    data_inicio = labels_datas[data_inicio_label]
    data_fim = labels_datas[data_fim_label]

    if data_inicio > data_fim:
        data_inicio, data_fim = data_fim, data_inicio
        st.info("O mês inicial estava depois do mês final. O período foi ajustado automaticamente.")

    historico_df = long_df[
        (long_df["Data"] >= data_inicio) &
        (long_df["Data"] <= data_fim)
    ].copy()

    usar_percentual = tipo_visualizacao == "Participação percentual"

    if historico_df.empty:
        st.info("Não há dados históricos para o período selecionado.")
    else:
        st.subheader("Histórico por Corretora")
        historical_bar_chart(
            historico_df,
            category="Corretora",
            title="Distribuição Histórica por Corretora",
            top_n=12,
            percent=usar_percentual,
            height=460,
        )

        st.subheader("Histórico Onshore vs. Offshore")
        historical_bar_chart(
            historico_df,
            category="Região",
            title="Distribuição Histórica Onshore vs. Offshore",
            top_n=None,
            percent=usar_percentual,
            height=420,
        )

        st.subheader("Histórico por Canal")
        historical_bar_chart(
            historico_df,
            category="Canal",
            title="Distribuição Histórica por Canal",
            top_n=None,
            percent=usar_percentual,
            height=420,
        )

        with st.expander("Tabela histórica consolidada", expanded=False):
            resumo_historico = (
                historico_df
                .groupby(["Data", "Corretora", "Região", "Canal"], as_index=False)["PL"]
                .sum()
                .sort_values(["Data", "PL"], ascending=[True, False])
            )
            resumo_historico["Data"] = resumo_historico["Data"].apply(month_display)
            resumo_historico["PL"] = resumo_historico["PL"].apply(br_money)
            st.dataframe(
                resumo_historico,
                use_container_width=True,
                hide_index=True,
                height=520,
            )


# ==============================
# BASE DE DADOS
# ==============================

else:
    st.subheader("Base de Dados")
    st.caption("Visualização de controle para conferência da planilha carregada.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Linhas", br_number(len(raw_df)))
    c2.metric("Colunas", br_number(len(raw_df.columns)))
    c3.metric("Colunas de PL", br_number(len(pl_cols)))
    c4.metric("Última data de PL", latest_date.strftime("%d/%m/%Y"))

    st.subheader("Colunas de PL Identificadas")
    mapa_pl = pd.DataFrame({"Coluna Convertida": list(pl_col_dates.keys()), "Data": list(pl_col_dates.values())})
    st.dataframe(mapa_pl, use_container_width=True, hide_index=True)

    st.subheader("Cotações USDBRL Usadas na Conversão Offshore")
    fx_show = fx_df.copy()
    fx_show["USDBRL usado"] = fx_show["USDBRL usado"].map(lambda x: br_number(x, 4))
    st.dataframe(fx_show, use_container_width=True, hide_index=True)

    st.subheader("Prévia da Base")
    preview = raw_df.head(200).copy()
    if "PL Atual" in preview.columns:
        preview["PL Atual"] = preview["PL Atual"].apply(br_money)
    st.dataframe(preview, use_container_width=True, hide_index=True)
