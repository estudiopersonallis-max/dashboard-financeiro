# ================= IMPORTS =================
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import re
import unicodedata

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

# PPTX
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# ================= CONFIG =================
st.set_page_config(page_title="Dashboard Financeiro PRO", layout="wide")
st.title("📊 Dashboard Financeiro – Nível Consultoria")

# ================= NORMALIZAÇÃO =================
def normalizar(txt):
    if pd.isna(txt):
        return ""
    txt = str(txt).upper().strip()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    return txt

# ================= DETECTAR MÊS =================
mapa_meses = {
    "JAN":1, "JANEIRO":1,
    "FEV":2, "FEVEREIRO":2,
    "MAR":3, "MARCO":3,
    "ABR":4, "ABRIL":4,
    "MAI":5, "MAIO":5,
    "JUN":6, "JUNHO":6,
    "JUL":7, "JULHO":7,
    "AGO":8, "AGOSTO":8,
    "SET":9, "SETEMBRO":9,
    "OUT":10, "OUTUBRO":10,
    "NOV":11, "NOVEMBRO":11,
    "DEZ":12, "DEZEMBRO":12
}

def extrair_mes(nome):
    nome = normalizar(nome)

    match = re.search(r'\b(0?[1-9]|1[0-2])\b', nome)
    if match:
        return int(match.group())

    for k, v in mapa_meses.items():
        if k in nome:
            return v

    return 99

# ================= LEITURA =================
@st.cache_data(ttl=3600)
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        periodo_nome = f.name.split(".")[0]
        mes = extrair_mes(periodo_nome)

        df["Periodo"] = periodo_nome.upper()
        df["ordem_mes"] = mes

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)

        df["Nome do cliente"] = df.get("Nome do cliente", "").apply(normalizar)
        df = df[df["Nome do cliente"] != ""]

        df["Modalidade"] = df.get("Modalidade", "N/A").apply(normalizar)
        df["Modalidade"] = df["Modalidade"].replace("", "SEM MODALIDADE")

        df["Tipo"] = df.get("Tipo", "N/A")
        df["Professor"] = df.get("Professor", "N/A")
        df["Local"] = df.get("Local", "N/A")

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

@st.cache_data(ttl=3600)
def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        df = df.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df.empty:
            continue

        periodo_nome = f.name.split(".")[0]
        mes = extrair_mes(periodo_nome)

        df["Periodo"] = periodo_nome.upper()
        df["ordem_mes"] = mes

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Classe"] = df.get("Classe", "N/A").apply(normalizar)

        df["Local"] = df.get("Local", "N/A")

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= NOVO: LEITURA CLIENTES =================
@st.cache_data(ttl=3600)
def ler_clientes(file):
    if file is None:
        return pd.DataFrame()

    df = pd.read_excel(file)

    df["Nome do cliente"] = df.get("Nome do cliente", "").apply(normalizar)

    df["Data de início"] = pd.to_datetime(
        df.get("Data de início", None),
        errors="coerce"
    )

    df = df[df["Nome do cliente"] != ""]

    return df

# ================= FUNÇÕES AUX =================
def grafico_bar(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    return fig

def gerar_pivot(df, index):
    return df.pivot_table(
        index=index,
        columns="Periodo",
        values="Valor",
        aggfunc="sum",
        fill_value=0
    )

# ================= EXPORT HELPERS =================
figs_pdf = []

def capturar_grafico(fig, titulo, pivot):
    figs_pdf.append((titulo, fig, pivot))

# ================= UPLOAD =================
st.sidebar.header("📤 Upload")
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

# NOVO upload
uploaded_clientes = st.sidebar.file_uploader(
    "Base de Clientes (Nome + Data de Início)",
    type=["xlsx"],
    accept_multiple_files=False
)

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()
clientes_df = ler_clientes(uploaded_clientes) if uploaded_clientes else pd.DataFrame()

# ================= FILTROS =================
st.sidebar.header("🔎 Filtros")
periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))
periodo_sel = st.sidebar.multiselect("Períodos", periodos, default=periodos)

if not receitas.empty:
    receitas = receitas[receitas["Periodo"].isin(periodo_sel)]

if not despesas.empty:
    despesas = despesas[despesas["Periodo"].isin(periodo_sel)]
    despesas = despesas[despesas["Classe"] != "DEPOSITOS"]

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

receita_media = receitas.groupby("Periodo")["Valor"].sum().mean() if not receitas.empty else 0
despesa_media = despesas.groupby("Periodo")["Valor"].sum().mean() if not despesas.empty else 0
clientes_ativos_media = receitas.groupby("Periodo")["Nome do cliente"].nunique().mean() if not receitas.empty else 0

ticket_medio_receita = receita_media / clientes_ativos_media if clientes_ativos_media else 0
ticket_medio_despesa = abs(despesa_media) / clientes_ativos_media if clientes_ativos_media else 0

magic_number = abs(despesa_media)
cac = abs(despesa_media) / clientes_ativos_media if clientes_ativos_media else 0

# ================= LTV REAL =================
if not clientes_df.empty and not receitas.empty:
    hoje = pd.Timestamp.today()

    clientes_df["meses_vida"] = (
        (hoje - clientes_df["Data de início"]).dt.days / 30
    )

    media_meses = clientes_df["meses_vida"].dropna().mean()
    clientes_df["meses_vida"] = clientes_df["meses_vida"].fillna(media_meses)

    clientes_validos = clientes_df[
        clientes_df["Nome do cliente"].isin(receitas["Nome do cliente"])
    ]

    tempo_medio = clientes_validos["meses_vida"].mean() if not clientes_validos.empty else 0

    ltv = ticket_medio_receita * tempo_medio if tempo_medio else 0
else:
    ltv = ticket_medio_receita * 6 if ticket_medio_receita else 0

ltv_cac = (ltv / cac) if cac else 0

# ================= KPIs DISPLAY =================
col1, col2, col3, col4 = st.columns(4)
col1.metric("Receita", f"{receita_total:,.0f}€")
col2.metric("Despesa", f"{despesa_total:,.0f}€")
col3.metric("Lucro", f"{lucro_total:,.0f}€")
col4.metric("Margem", f"{margem:.1f}%")

col5, col6, col7, col8 = st.columns(4)
col5.metric("Ticket Receita", f"{ticket_medio_receita:,.0f}€")
col6.metric("Ticket Despesa", f"{ticket_medio_despesa:,.0f}€")
col7.metric("CAC", f"{cac:,.0f}€")
col8.metric("LTV/CAC", f"{ltv_cac:.2f}")

# ================= RESTO DO CÓDIGO (igual ao teu) =================
# (mantive tudo exatamente igual abaixo)

st.subheader("⚠️ Alertas Inteligentes")
if margem < 20:
    st.warning("Margem abaixo de 20%")
if despesa_total < -receita_total * 0.8:
    st.warning("Despesas muito altas")
if clientes_ativos_media < 10:
    st.warning("Poucos clientes ativos")

st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
