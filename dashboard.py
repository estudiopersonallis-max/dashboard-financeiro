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
    match = re.search(r'\\b(0?[1-9]|1[0-2])\\b', nome)
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

        # 🔒 GARANTIR COLUNAS
        for col in ["Modalidade", "Tipo", "Professor", "Local"]:
            if col not in df.columns:
                df[col] = "N/A"

        df["Modalidade"] = df["Modalidade"].apply(normalizar)
        df["Modalidade"] = df["Modalidade"].replace("", "SEM MODALIDADE")

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

        for col in ["Classe", "Local"]:
            if col not in df.columns:
                df[col] = "N/A"

        df["Classe"] = df["Classe"].apply(normalizar)

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= CLIENTES =================
@st.cache_data(ttl=3600)
def ler_clientes(file):
    df = pd.read_excel(file)
    df.columns = [normalizar(col) for col in df.columns]

    mapa_colunas = {
        "NOME DO CLIENTE": "Nome do cliente",
        "CLIENTE": "Nome do cliente",
        "DATA INICIO": "Data Inicio",
        "DATA DE INICIO": "Data Inicio",
        "INICIO": "Data Inicio"
    }

    df = df.rename(columns=mapa_colunas)

    if "Nome do cliente" not in df.columns or "Data Inicio" not in df.columns:
        st.warning("Ficheiro de clientes inválido")
        return pd.DataFrame()

    df["Nome do cliente"] = df["Nome do cliente"].apply(normalizar)
    df["Data Inicio"] = pd.to_datetime(df["Data Inicio"], errors="coerce")

    return df

# ================= CLIENTES ATIVOS =================
st.subheader("👥 Clientes Ativos por Mês")

if not receitas.empty:
    clientes_ativos = (
        receitas.groupby(["Periodo", "ordem_mes"])["Nome do cliente"]
        .nunique()
        .reset_index()
        .sort_values("ordem_mes")
    )

    fig, ax = plt.subplots()
    ax.plot(clientes_ativos["Periodo"], clientes_ativos["Nome do cliente"], marker="o")
    ax.set_title("Clientes Ativos por Mês")
    plt.xticks(rotation=45)

    st.pyplot(fig)


# ================= DISTRIBUIÇÃO POR MODALIDADE =================
st.subheader("📊 Distribuição de Clientes por Modalidade")

if not receitas.empty and "Modalidade" in receitas.columns:

    clientes_modalidade = (
        receitas.groupby("Modalidade")["Nome do cliente"]
        .nunique()
        .sort_values(ascending=False)
    )

    # Tabela
    st.dataframe(clientes_modalidade)

    # Gráfico absoluto
    fig_abs, ax_abs = plt.subplots()
    clientes_modalidade.plot(kind="barh", ax=ax_abs)
    ax_abs.set_title("Clientes por Modalidade")
    st.pyplot(fig_abs)

    # Distribuição %
    clientes_modalidade_pct = clientes_modalidade / clientes_modalidade.sum() * 100

    fig_pct, ax_pct = plt.subplots()
    clientes_modalidade_pct.plot(kind="barh", ax=ax_pct)
    ax_pct.set_title("Distribuição (%) por Modalidade")
    st.pyplot(fig_pct)


# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    return fig

# ================= EXPORT HELPERS =================
figs_pdf = []

def capturar_grafico(fig, titulo, pivot):
    figs_pdf.append((titulo, fig, pivot))

# ================= UPLOAD =================
st.sidebar.header("📤 Upload")
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)
uploaded_clientes = st.sidebar.file_uploader("Base Clientes", type=["xlsx"])

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

if uploaded_clientes and not receitas.empty:
    clientes_df = ler_clientes(uploaded_clientes)
    if not clientes_df.empty:
        receitas = receitas.merge(clientes_df, on="Nome do cliente", how="left")

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Despesa", f"{despesa_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")

# ================= TABS =================
tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "💰 Receitas", "💸 Despesas"])

with tab2:
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        if cat not in receitas.columns:
            continue

        bloco = receitas.pivot_table(index=cat, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
        st.dataframe(bloco)

        fig = grafico_bar(bloco, f"Receitas por {cat}")
        st.pyplot(fig)

        # %
        bloco_pct = bloco.div(bloco.sum(axis=0), axis=1) * 100
        fig_pct = grafico_bar(bloco_pct, f"Receitas (%) por {cat}")
        st.pyplot(fig_pct)

        capturar_grafico(fig, f"Receitas por {cat}", bloco)

with tab3:
    for cat in ["Classe", "Local"]:
        if cat not in despesas.columns:
            continue

        bloco = despesas.pivot_table(index=cat, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
        st.dataframe(bloco)

        fig = grafico_bar(bloco, f"Despesas por {cat}")
        st.pyplot(fig)

        # %
        bloco_pct = bloco.div(bloco.sum(axis=0), axis=1) * 100
        fig_pct = grafico_bar(bloco_pct, f"Despesas (%) por {cat}")
        st.pyplot(fig_pct)

        capturar_grafico(fig, f"Despesas por {cat}", bloco)

# ================= KPIs AVANÇADOS =================
st.markdown("## 🧠 KPIs Avançados")

col1, col2, col3, col4 = st.columns(4)

if not receitas.empty:
    clientes_unicos = receitas["Nome do cliente"].nunique()
    receita_mes = receitas.groupby("Periodo")["Valor"].sum()
    clientes_mes = receitas.groupby("Periodo")["Nome do cliente"].nunique()

    ticket_mensal = (receita_mes / clientes_mes).mean()
    ticket_total = receita_total / clientes_unicos if clientes_unicos else 0
else:
    ticket_mensal = 0
    ticket_total = 0

clientes_media = receitas.groupby("Periodo")["Nome do cliente"].nunique().mean() if not receitas.empty else 0
despesa_media = despesas.groupby("Periodo")["Valor"].sum().mean() if not despesas.empty else 0
cac = abs(despesa_media) / clientes_media if clientes_media else 0

if "Data Inicio" in receitas.columns:
    hoje = pd.Timestamp.today()
    receitas["meses_ativos"] = ((hoje - receitas["Data Inicio"]).dt.days / 30)
    tempo_medio = receitas.groupby("Nome do cliente")["meses_ativos"].max().mean()
else:
    tempo_medio = 6

ltv = ticket_mensal * tempo_medio
ltv_cac = ltv / cac if cac else 0

with col1:
    st.metric("🎯 Ticket Mensal", f"{ticket_mensal:,.0f}€")
with col2:
    st.metric("💸 CAC", f"{cac:,.0f}€")
with col3:
    st.metric("💰 LTV", f"{ltv:,.0f}€")
with col4:
    st.metric("⚖️ LTV/CAC", f"{ltv_cac:.2f}")

with st.expander("🔍 Detalhes KPIs"):
    st.write(f"Ticket total: {ticket_total:,.0f}€")
    st.write(f"Tempo médio: {tempo_medio:.1f} meses")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
