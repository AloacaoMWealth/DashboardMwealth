from __future__ import annotations

import re
import unicodedata
from datetime import date
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==============================
# CONFIGURAÇÕES GERAIS
# ==============================

APP_TITLE = "Dashboard Macro M Wealth"
DATA_DIR = Path("data")
DEFAULT_FILE = DATA_DIR / "Controle de Clientes MWealth 2026.xlsx"
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
TEXTO = "#111827"

CORRETORAS_OFFSHORE = {
    "AVENUE",
    "CHARLES SCHWAB",
    "CHARLES SHWAB",  # grafia observada na base
    "SCHWAB",
    "SHWAB",
    "XP US",
    "XP INTERNACIONAL",
    "INTERACTIVE BROKERS",
    "IBKR",
    "PERSHING",
    "ATIVORE",
}

# ==============================
# PÁGINA E ESTILO
# ==============================

st.set_page_config(page_title=APP_TITLE, page_icon=None, layout="wide", initial_sidebar_state="expanded")

st.markdown(
    f"""
    <style>
        .stApp {{
            background: {CINZA_FUNDO};
            color: {TEXTO};
        }}
        .stApp p, .stApp span, .stApp div, .stApp label {{
            color: {TEXTO};
        }}
        section[data-testid="stSidebar"] {{
            background: {AZUL_ESCURO};
        }}
        section[data-testid="stSidebar"], section[data-testid="stSidebar"] * {{
            color: white !important;
        }}
        .block-container {{
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            max-width: 1500px;
        }}
        h1, h2, h3 {{
            color: {AZUL_ESCURO};
            letter-spacing: -0.02em;
        }}
        .main-title {{
            background: linear-gradient(90deg, {AZUL_ESCURO}, {AZUL_MEDIO});
            color: white !important;
            padding: 22px 28px;
            border-radius: 18px;
            margin-bottom: 18px;
            box-shadow: 0 8px 24px rgba(0, 51, 72, 0.16);
        }}
        .main-title h1, .main-title p {{
            color: white !important;
            margin: 0;
        }}
        .main-title h1 {{ font-size: 30px; }}
        .main-title p {{ color: rgba(255,255,255,0.88) !important; margin-top: 6px; font-size: 14px; }}
        .kpi-card {{
            background: white;
            border: 1px solid {CINZA_BORDA};
            border-radius: 16px;
            padding: 18px 18px;
            min-height: 118px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
        }}
        .kpi-label {{
            font-size: 13px;
            color: #5B677A !important;
            font-weight: 600;
            margin-bottom: 10px;
        }}
        .kpi-value {{
            font-size: 25px;
            color: {AZUL_ESCURO} !important;
            font-weight: 800;
            line-height: 1.1;
        }}
        .kpi-sub {{
            font-size: 12px;
            color: #6B7280 !important;
            margin-top: 8px;
        }}
        .section-card {{
            background: white;
            border: 1px solid {CINZA_BORDA};
            border-radius: 18px;
            padding: 18px;
            margin-bottom: 18px;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.045);
        }}
        div[data-testid="stMetric"] {{
            background: white;
            border: 1px solid {CINZA_BORDA};
            border-radius: 16px;
            padding: 14px 16px;
            box-shadow: 0 5px 16px rgba(15, 23, 42, 0.04);
        }}
        div[data-testid="stMetricLabel"] {{ color: #5B677A !important; }}
        div[data-testid="stMetricValue"] {{ color: {AZUL_ESCURO} !important; }}
        .stDataFrame {{ border-radius: 12px; overflow: hidden; }}

        /* Legibilidade de inputs/selects no tema claro */
        div[data-baseweb="select"] > div {{
            background-color: #FFFFFF !important;
            border-color: #CBD5E1 !important;
        }}
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] input,
        div[data-baseweb="popover"] * {{
            color: {TEXTO} !important;
        }}
        div[data-baseweb="tag"] {{ background-color: #E6F0F4 !important; }}
        div[data-baseweb="tag"] span {{ color: {AZUL_ESCURO} !important; }}
        .stTextInput input, .stNumberInput input, .stDateInput input {{
            color: {TEXTO} !important;
            background-color: #FFFFFF !important;
        }}
        .stAlert, .stAlert * {{ color: {TEXTO} !important; }}
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
    txt = f"R$ {value:,.2f}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


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


def make_unique_columns(columns):
    seen = {}
    out = []
    for col in columns:
        col = str(col).strip()
        if col not in seen:
            seen[col] = 0
            out.append(col)
        else:
            seen[col] += 1
            out.append(f"{col}.{seen[col]}")
    return out


def normalize_text(value) -> str:
    if pd.isna(value):
        return "Não informado"
    value = str(value).strip()
    return value if value else "Não informado"


def normalize_key(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Z0-9]", "", value.upper())


def find_col(df: pd.DataFrame, candidates: List[str]) -> str | None:
    lookup = {normalize_key(c): c for c in df.columns}
    for candidate in candidates:
        key = normalize_key(candidate)
        if key in lookup:
            return lookup[key]
    return None


def clean_numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    cleaned = (
        series.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace("R$", "", regex=False)
        .str.replace("US$", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )
    cleaned = cleaned.replace({"": "0", "nan": "0", "None": "0", "-": "0"})
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


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
    corretora_clean = normalize_text(corretora).upper().replace("SHWAB", "SCHWAB")
    offshore_terms = {x.replace("SHWAB", "SCHWAB") for x in CORRETORAS_OFFSHORE}
    if any(term in corretora_clean for term in offshore_terms):
        return "Offshore"
    return "Onshore"


def make_line_meta(months: List[pd.Timestamp], first_value: float, target_value: float) -> pd.Series:
    if not months:
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


def standard_layout(fig: go.Figure, height: int = 360, legend: bool = True) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=45, b=25),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color=TEXTO, size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1) if legend else None,
    )
    fig.update_xaxes(showgrid=False, linecolor="#E5E7EB")
    fig.update_yaxes(gridcolor="#E5E7EB", linecolor="#E5E7EB")
    return fig


def bar_chart(df: pd.DataFrame, x: str, y: str, title: str, horizontal: bool = False, height: int = 360):
    if df.empty or y not in df.columns or df[y].sum() == 0:
        st.info("Sem dados para exibir neste gráfico.")
        return
    if horizontal:
        fig = px.bar(df, x=y, y=x, orientation="h", title=title, text=y)
        fig.update_traces(marker_color=AZUL_MEDIO, texttemplate="R$ %{x:,.0f}", textposition="outside")
        fig.update_layout(yaxis=dict(autorange="reversed"))
    else:
        fig = px.bar(df, x=x, y=y, title=title, text=y)
        fig.update_traces(marker_color=AZUL_MEDIO, texttemplate="R$ %{y:,.0f}", textposition="outside")
    fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f") if not horizontal else fig.update_xaxes(tickprefix="R$ ", tickformat=",.0f")
    fig = standard_layout(fig, height=height, legend=False)
    st.plotly_chart(fig, use_container_width=True)


def donut_chart(df: pd.DataFrame, names: str, values: str, title: str, height: int = 340):
    if df.empty or values not in df.columns or df[values].sum() == 0:
        st.info("Sem dados para exibir neste gráfico.")
        return
    fig = px.pie(df, names=names, values=values, hole=0.56, title=title)
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig = standard_layout(fig, height=height, legend=True)
    st.plotly_chart(fig, use_container_width=True)


def extract_fx_from_workbook(file_path: str) -> Dict[pd.Timestamp, float]:
    """Lê cotações USDBRL preenchidas na aba de análise, quando existirem."""
    fx_map: Dict[pd.Timestamp, float] = {}
    try:
        xl = pd.ExcelFile(file_path)
        sheet_name = next((s for s in xl.sheet_names if "ANÁLISE" in s.upper() or "ANALISE" in s.upper()), None)
        if not sheet_name:
            return fx_map
        tmp = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        month_map = {
            "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3, "ABRIL": 4,
            "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9,
            "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
        }
        year = None
        for val in tmp.values.flatten().tolist():
            m = re.search(r"20\d{2}", str(val))
            if m:
                year = int(m.group(0))
                break
        year = year or DATA_META.year

        header_row = None
        dolar_row = None
        for idx, row in tmp.iterrows():
            row_text = [str(x).strip().upper() for x in row.tolist()]
            if any(x in month_map for x in row_text):
                header_row = idx
            if any("DÓLAR PRICE" in x or "DOLAR PRICE" in x or "USDBRL" in x or x == "DOLAR" for x in row_text):
                dolar_row = idx
        if header_row is None or dolar_row is None:
            return fx_map

        for col_idx, header in tmp.loc[header_row].items():
            header_norm = str(header).strip().upper()
            if header_norm in month_map:
                fx = clean_numeric(pd.Series([tmp.iat[dolar_row, col_idx]])).iloc[0]
                if fx > 0:
                    fx_map[pd.Timestamp(year=year, month=month_map[header_norm], day=1)] = float(fx)
    except Exception:
        pass
    return fx_map


@st.cache_data(show_spinner=False)
def get_usdbrl_close(pl_date: pd.Timestamp, workbook_fx: Dict[pd.Timestamp, float]) -> float:
    """Prioridade: cotação manual da planilha > yfinance > última cotação manual anterior > 1."""
    month_key = pd.Timestamp(year=pl_date.year, month=pl_date.month, day=1)
    if month_key in workbook_fx and workbook_fx[month_key] > 0:
        return float(workbook_fx[month_key])
    try:
        import yfinance as yf
        start = (pl_date - pd.Timedelta(days=10)).strftime("%Y-%m-%d")
        end = (pl_date + pd.Timedelta(days=2)).strftime("%Y-%m-%d")
        hist = yf.download("USDBRL=X", start=start, end=end, progress=False, auto_adjust=False)
        if not hist.empty and "Close" in hist.columns:
            close = hist["Close"].dropna()
            close = close[close.index <= pl_date]
            if not close.empty:
                return float(close.iloc[-1])
    except Exception:
        pass
    valid = sorted((k, v) for k, v in workbook_fx.items() if v and v > 0 and k <= month_key)
    if valid:
        return float(valid[-1][1])
    return 1.0


def get_base_file() -> Path:
    DATA_DIR.mkdir(exist_ok=True)
    return DEFAULT_FILE


@st.cache_data(show_spinner="Carregando e tratando a base de clientes...")
def load_data(file_path: str, file_mtime: float) -> Tuple[pd.DataFrame, pd.DataFrame, List[str], Dict[str, pd.Timestamp], Dict[pd.Timestamp, float], str]:
    df = pd.read_excel(file_path, sheet_name=MAIN_SHEET)
    df.columns = make_unique_columns(df.columns)
    df["__row_id"] = np.arange(len(df))

    pl_cols = extract_pl_columns(df)
    if not pl_cols:
        raise ValueError("Não encontrei colunas de PL no padrão 'PL dd/mm/aaaa'.")

    text_cols = [
        "Corretora", "Grupo Geral", "Grupo Familiar", "Cliente", "PF/ PJ", "Canal",
        "Cliente - Corretora", "Conta", "UF", "Consultor", "Tipo de Marcação",
        "Perfil Carteira/ Renda", "Perfil de Investidor", "Consolidador", "Liquidez",
        "Observações", "Possui Crédito", "Possui Previdência", "Possui Previdência ", "Calculadora de IR", "E-mail",
    ]
    for col in text_cols:
        real_col = find_col(df, [col])
        if real_col is not None:
            df[real_col] = df[real_col].apply(normalize_text)

    conta_col = find_col(df, ["Conta"])
    if conta_col:
        df[conta_col] = df[conta_col].astype(str).str.replace(".0", "", regex=False).str.strip()

    corretora_col = find_col(df, ["Corretora"])
    if corretora_col:
        df["Região"] = df[corretora_col].apply(classify_region)
    else:
        df["Região"] = "Onshore"

    # Limpa PLs e converte Offshore em USD para BRL pelo fechamento do mês/data.
    workbook_fx = extract_fx_from_workbook(file_path)
    fx_used: Dict[pd.Timestamp, float] = {}
    offshore_mask = df["Região"].eq("Offshore")
    for col, dt in pl_cols.items():
        df[col] = clean_numeric(df[col])
        fx = get_usdbrl_close(dt, workbook_fx)
        fx_used[dt] = fx
        df.loc[offshore_mask, col] = df.loc[offshore_mask, col] * fx

    # Última coluna útil: evita escolher um mês D0 ainda vazio e zerar o dashboard.
    useful_cols = [col for col in pl_cols.keys() if df[col].abs().sum() > 0]
    latest_col = useful_cols[-1] if useful_cols else list(pl_cols.keys())[-1]
    df["PL Atual"] = df[latest_col].fillna(0)
    df["Data PL Atual"] = pl_cols[latest_col]

    id_cols = [c for c in df.columns if c not in pl_cols.keys()]
    long_df = df.melt(id_vars=id_cols, value_vars=list(pl_cols.keys()), var_name="Coluna PL", value_name="PL")
    long_df["Data"] = long_df["Coluna PL"].map(pl_cols)
    long_df["Mês"] = long_df["Data"].dt.to_period("M").astype(str)
    long_df["Ano"] = long_df["Data"].dt.year
    long_df["PL"] = long_df["PL"].fillna(0)

    return df, long_df, list(pl_cols.keys()), pl_cols, fx_used, latest_col


def apply_filters(base: pd.DataFrame, filtros: Dict[str, List[str]]) -> pd.DataFrame:
    out = base.copy()
    for col, selected in filtros.items():
        if col in out.columns and selected:
            out = out[out[col].isin(selected)]
    return out


def filtered_multiselect(label: str, df: pd.DataFrame, col: str):
    if col not in df.columns:
        return []
    options = sorted([x for x in df[col].dropna().unique().tolist() if str(x).strip() and str(x) != "Não informado"])
    return st.multiselect(label, options=options, default=[], placeholder="Todos")


def sort_money_df(df: pd.DataFrame, money_col: str, ascending: bool = False) -> pd.DataFrame:
    if money_col not in df.columns:
        return df
    return df.sort_values(money_col, ascending=ascending)

# ==============================
# SIDEBAR E CARGA DA BASE
# ==============================

st.sidebar.markdown("# M Wealth")
st.sidebar.markdown("Dashboard macro do escritório")

base_file = get_base_file()

with st.sidebar.expander("Atualização da base", expanded=False):
    st.caption("No GitHub, substitua o arquivo na pasta data. O upload abaixo é útil para teste local ou atualização manual no ambiente atual.")
    uploaded = st.file_uploader("Enviar nova planilha", type=["xlsx"])
    if uploaded is not None:
        DATA_DIR.mkdir(exist_ok=True)
        with open(DEFAULT_FILE, "wb") as f:
            f.write(uploaded.getbuffer())
        st.cache_data.clear()
        st.success("Base atualizada. Recarregue a página se necessário.")

if not base_file.exists():
    st.error("Arquivo base não encontrado. Inclua a planilha em data/Controle de Clientes MWealth 2026.xlsx ou envie pelo upload lateral.")
    st.stop()

try:
    raw_df, long_df, pl_cols, pl_col_dates, fx_used, latest_col = load_data(str(base_file), base_file.stat().st_mtime)
except Exception as exc:
    st.error(f"Erro ao carregar a base: {exc}")
    st.stop()

latest_date = pl_col_dates[latest_col]

pagina = st.sidebar.radio("Navegação", ["Dashboard Macro", "Análise Detalhada", "Base de Dados"], index=0)
st.sidebar.divider()
st.sidebar.caption(f"Base: {base_file.name}")
st.sidebar.caption(f"Última coluna útil de PL: {latest_col}")
if fx_used:
    st.sidebar.caption("Offshore convertido para BRL por USDBRL.")

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

    grupo_col = "Grupo Familiar" if "Grupo Familiar" in raw_df.columns else "Cliente"
    grupos = raw_df.groupby(grupo_col, dropna=False)["PL Atual"].sum().reset_index()
    grupos_validos = grupos[grupos[grupo_col] != "Não informado"]
    qtd_grupos = int(grupos_validos[grupo_col].nunique())
    pl_medio_grupo = grupos_validos["PL Atual"].sum() / qtd_grupos if qtd_grupos else 0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        kpi_card("PL total sob gestão", br_money(current_pl), f"Data-base: {latest_date.strftime('%d/%m/%Y')}")
    with col2:
        kpi_card("Meta patrimonial", br_money(META_PL_EMPRESA), f"Realizado: {br_percent(pct_meta)}")
    with col3:
        kpi_card("Desvio para a meta", br_money(gap_meta), f"Meses restantes: {meses_restantes} | Semanas: {semanas_restantes}")
    with col4:
        kpi_card("Grupos familiares", br_number(qtd_grupos), f"PL médio por grupo: {br_money(pl_medio_grupo)}")

    mensal = long_df.groupby("Data", as_index=False)["PL"].sum().sort_values("Data")
    mensal = mensal[mensal["PL"].abs() > 0]
    if not mensal.empty:
        mensal["Meta Linear"] = make_line_meta(mensal["Data"].tolist(), mensal["PL"].iloc[0], META_PL_EMPRESA).values
        mensal["Percentual da Meta"] = mensal["PL"] / META_PL_EMPRESA

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Evolução do PL vs. Meta Patrimonial")
    if mensal.empty:
        st.info("Sem dados de PL para exibir a evolução.")
    else:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=mensal["Data"], y=mensal["PL"], mode="lines+markers", name="PL Realizado", line=dict(color=AZUL_CLARO, width=3), marker=dict(size=7)))
        fig.add_trace(go.Scatter(x=mensal["Data"], y=mensal["Meta Linear"], mode="lines", name="Meta Linear", line=dict(color=AZUL_ESCURO, width=3)))
        fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
        fig = standard_layout(fig, height=420, legend=True)
        st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1.25, 1])
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("PL por Corretora")
        if "Corretora" in raw_df.columns:
            corretora = raw_df.groupby("Corretora", as_index=False)["PL Atual"].sum().pipe(sort_money_df, "PL Atual")
            bar_chart(corretora, "Corretora", "PL Atual", "Distribuição por Corretora", horizontal=False, height=390)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Onshore vs. Offshore")
        regiao = raw_df.groupby("Região", as_index=False)["PL Atual"].sum().pipe(sort_money_df, "PL Atual")
        donut_chart(regiao, "Região", "PL Atual", "Proporção de PL", height=390)
        st.markdown('</div>', unsafe_allow_html=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Proporção por Canal")
        if "Canal" in raw_df.columns:
            canal = raw_df.groupby("Canal", as_index=False)["PL Atual"].sum().pipe(sort_money_df, "PL Atual")
            donut_chart(canal, "Canal", "PL Atual", "Wealth, Tesouraria e Demais Canais", height=370)
        else:
            st.info("Coluna Canal não encontrada na base.")
        st.markdown('</div>', unsafe_allow_html=True)
    with c4:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Segmentação da Base por Faixa de PL")
        cliente_col = "Cliente" if "Cliente" in raw_df.columns else grupo_col
        cliente_pl = raw_df.groupby(cliente_col, as_index=False)["PL Atual"].sum()
        bins = [-0.01, 5_000_000, 10_000_000, 15_000_000, 30_000_000, 100_000_000, np.inf]
        labels = ["Até R$ 5 MM", "R$ 5 MM a R$ 10 MM", "R$ 10 MM a R$ 15 MM", "R$ 15 MM a R$ 30 MM", "R$ 30 MM a R$ 100 MM", "Acima de R$ 100 MM"]
        cliente_pl["Faixa"] = pd.cut(cliente_pl["PL Atual"], bins=bins, labels=labels)
        faixa = cliente_pl.groupby("Faixa", observed=False).agg(Clientes=(cliente_col, "count"), PL=("PL Atual", "sum")).reset_index()
        donut_chart(faixa, "Faixa", "Clientes", "Quantidade de Clientes por Faixa", height=370)
        st.markdown('</div>', unsafe_allow_html=True)

    c5, c6 = st.columns(2)
    with c5:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Top 10 Clientes")
        cliente_col = "Cliente" if "Cliente" in raw_df.columns else grupo_col
        top_clientes = raw_df.groupby(cliente_col, as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False).head(10)
        bar_chart(top_clientes.sort_values("PL Atual"), cliente_col, "PL Atual", "Maiores Clientes por PL", horizontal=True, height=450)
        st.markdown('</div>', unsafe_allow_html=True)
    with c6:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Top 10 Grupos Familiares")
        top_grupos = raw_df.groupby(grupo_col, as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False).head(10)
        bar_chart(top_grupos.sort_values("PL Atual"), grupo_col, "PL Atual", "Maiores Grupos por PL", horizontal=True, height=450)
        st.markdown('</div>', unsafe_allow_html=True)

# ==============================
# ANÁLISE DETALHADA
# ==============================

elif pagina == "Análise Detalhada":
    st.subheader("Análise Detalhada da Base")
    st.caption("Use os filtros para validar PL por consultor, cliente, grupo familiar, canal, corretora e perfil.")

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

    filtros = {"Consultor": consultores, "Cliente": clientes, "Grupo Familiar": grupos_familiares, "Corretora": corretoras, "Canal": canais, "PF/ PJ": tipo_pessoa}
    base_filtrada = apply_filters(raw_df, filtros)
    row_ids = base_filtrada["__row_id"].tolist() if "__row_id" in base_filtrada.columns else []

    contas_filtradas = base_filtrada["Conta"].nunique() if "Conta" in base_filtrada.columns else len(base_filtrada)
    clientes_filtrados = base_filtrada["Cliente"].nunique() if "Cliente" in base_filtrada.columns else len(base_filtrada)
    grupos_filtrados = base_filtrada["Grupo Familiar"].nunique() if "Grupo Familiar" in base_filtrada.columns else 0
    pl_filtrado = base_filtrada["PL Atual"].sum()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card("PL filtrado", br_money(pl_filtrado), f"{br_percent(pl_filtrado / META_PL_EMPRESA)} da meta total")
    with k2:
        kpi_card("Contas", br_number(contas_filtradas), "Quantidade de contas na seleção")
    with k3:
        kpi_card("Clientes", br_number(clientes_filtrados), "Clientes únicos na seleção")
    with k4:
        kpi_card("Grupos familiares", br_number(grupos_filtrados), "Grupos únicos na seleção")

    long_filtrado = long_df[long_df["__row_id"].isin(row_ids)] if row_ids else long_df.iloc[0:0]

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Evolução Histórica da Seleção")
    evo = long_filtrado.groupby("Data", as_index=False)["PL"].sum().sort_values("Data")
    evo = evo[evo["PL"].abs() > 0]
    if evo.empty:
        st.info("Nenhum dado encontrado para os filtros selecionados.")
    else:
        fig = px.line(evo, x="Data", y="PL", markers=True, title="PL Mensal da Seleção")
        fig.update_traces(line_color=AZUL_CLARO, line_width=3)
        fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
        fig = standard_layout(fig, height=360, legend=False)
        st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("PL por Consultor")
        if "Consultor" in base_filtrada.columns:
            por_consultor = base_filtrada.groupby("Consultor", as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False).head(20)
            bar_chart(por_consultor.sort_values("PL Atual"), "Consultor", "PL Atual", "Ranking de Consultores", horizontal=True, height=500)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("PL por Grupo Familiar")
        grupo_col = "Grupo Familiar" if "Grupo Familiar" in base_filtrada.columns else "Cliente"
        por_grupo = base_filtrada.groupby(grupo_col, as_index=False)["PL Atual"].sum().sort_values("PL Atual", ascending=False).head(20)
        bar_chart(por_grupo.sort_values("PL Atual"), grupo_col, "PL Atual", "Ranking de Grupos", horizontal=True, height=500)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Tabela Analítica")
    cols_show = ["Corretora", "Região", "Grupo Geral", "Grupo Familiar", "Cliente", "PF/ PJ", "Canal", "Conta", "UF", "Consultor", "Perfil Carteira/ Renda", "Perfil de Investidor", "PL Atual"]
    cols_show = [c for c in cols_show if c in base_filtrada.columns]
    tabela = base_filtrada[cols_show].sort_values("PL Atual", ascending=False).copy()
    st.dataframe(tabela, use_container_width=True, hide_index=True, column_config={"PL Atual": st.column_config.NumberColumn("PL Atual", format="R$ %.2f")})
    csv = tabela.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
    st.download_button("Baixar tabela filtrada em CSV", csv, "base_filtrada_mwealth.csv", "text/csv")
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================
# BASE DE DADOS
# ==============================

else:
    st.subheader("Base de Dados")
    st.caption("Visualização de controle para conferência da planilha carregada e da conversão de offshore.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Linhas", br_number(len(raw_df)))
    c2.metric("Colunas", br_number(len(raw_df.columns)))
    c3.metric("Colunas de PL", br_number(len(pl_cols)))
    c4.metric("Última data útil de PL", latest_date.strftime("%d/%m/%Y"))

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Colunas de PL Identificadas")
    mapa_pl = pd.DataFrame({"Coluna": list(pl_col_dates.keys()), "Data": list(pl_col_dates.values())})
    mapa_pl["Usada como PL Atual"] = mapa_pl["Coluna"].eq(latest_col)
    st.dataframe(mapa_pl, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Cotações USDBRL Utilizadas para Offshore")
    fx_df = pd.DataFrame({"Data PL": list(fx_used.keys()), "USDBRL usado": list(fx_used.values())}).sort_values("Data PL")
    st.dataframe(fx_df, use_container_width=True, hide_index=True, column_config={"USDBRL usado": st.column_config.NumberColumn("USDBRL usado", format="%.4f")})
    st.caption("Quando a cotação estiver preenchida na planilha, ela tem prioridade. Na ausência, o app tenta buscar via yfinance. Se nada estiver disponível, usa 1,00 como fallback para não quebrar a carga.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Prévia da Base Tratada")
    cols_preview = [c for c in ["Corretora", "Região", "Grupo Familiar", "Cliente", "Canal", "Conta", "Consultor", "PL Atual"] if c in raw_df.columns]
    st.dataframe(raw_df[cols_preview].head(300), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)
