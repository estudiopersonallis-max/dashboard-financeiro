# ================= IMPORTS =================
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import re
import unicodedata

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
mapa_meses = {"JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,"JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12}

def extrair_mes(nome):
    nome = normalizar(nome)
    for k,v in mapa_meses.items():
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

        periodo = f.name.split(".")[0].upper()
        df["Periodo"] = periodo
        df["ordem_mes"] = extrair_mes(periodo)

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Nome do cliente"] = df.get("Nome do cliente", "").apply(normalizar)

        for col in ["Modalidade","Tipo","Professor","Local"]:
            if col not in df.columns:
                df[col] = "N/A"

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

@st.cache_data(ttl=3600)
def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        periodo = f.name.split(".")[0].upper()
        df["Periodo"] = periodo
        df["ordem_mes"] = extrair_mes(periodo)

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)

        for col in ["Classe","Local"]:
            if col not in df.columns:
                df[col] = "N/A"

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= UPLOAD =================
st.sidebar.header("📤 Upload")
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("Receita", f"{receita_total:,.0f}€")
col2.metric("Despesa", f"{despesa_total:,.0f}€")
col3.metric("Lucro", f"{lucro_total:,.0f}€")
col4.metric("Margem", f"{margem:.1f}%")

# ================= CLIENTES =================
st.subheader("👥 Evolução de Clientes")

if not receitas.empty:
    clientes_mes = receitas.groupby(["Periodo","ordem_mes"])["Nome do cliente"].nunique().reset_index().sort_values("ordem_mes")

    fig, ax = plt.subplots()
    ax.plot(clientes_mes["Periodo"], clientes_mes["Nome do cliente"], marker="o")
    plt.xticks(rotation=45)
    st.pyplot(fig)

# ================= NOVO =================
st.subheader("👥 Clientes Ativos por Mês")

if not receitas.empty:
    fig, ax = plt.subplots()
    ax.plot(clientes_mes["Periodo"], clientes_mes["Nome do cliente"], marker="o")
    st.pyplot(fig)

st.subheader("📊 Distribuição de Clientes por Modalidade")

if not receitas.empty:
    dist = receitas.groupby("Modalidade")["Nome do cliente"].nunique().sort_values()
    st.dataframe(dist)

    fig, ax = plt.subplots()
    dist.plot(kind="barh", ax=ax)
    st.pyplot(fig)

    fig2, ax2 = plt.subplots()
    (dist/dist.sum()*100).plot(kind="barh", ax=ax2)
    st.pyplot(fig2)

# ================= TABS =================
tab1, tab2, tab3 = st.tabs(["📊 Visão Geral","💰 Receitas","💸 Despesas"])

with tab2:
    for cat in ["Modalidade","Tipo","Professor","Local"]:
        if cat not in receitas.columns:
            continue
        bloco = receitas.pivot_table(index=cat, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
        st.dataframe(bloco)
        st.pyplot(bloco.plot(kind="barh").figure)
        st.pyplot((bloco.div(bloco.sum(axis=0), axis=1)*100).plot(kind="barh").figure)

with tab3:
    for cat in ["Classe","Local"]:
        if cat not in despesas.columns:
            continue
        bloco = despesas.pivot_table(index=cat, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
        st.dataframe(bloco)
        st.pyplot(bloco.plot(kind="barh").figure)
        st.pyplot((bloco.div(bloco.sum(axis=0), axis=1)*100).plot(kind="barh").figure)

# ================= KPIs AVANÇADOS =================
st.markdown("## 🧠 KPIs Avançados")

if not receitas.empty:
    clientes_unicos = receitas["Nome do cliente"].nunique()
    receita_mes = receitas.groupby("Periodo")["Valor"].sum()
    clientes_mes = receitas.groupby("Periodo")["Nome do cliente"].nunique()
    ticket = (receita_mes/clientes_mes).mean()
else:
    ticket = 0

cac = abs(despesa_total) / clientes_unicos if clientes_unicos else 0
ltv = ticket * 6

col1, col2, col3 = st.columns(3)
col1.metric("Ticket", f"{ticket:,.0f}€")
col2.metric("CAC", f"{cac:,.0f}€")
col3.metric("LTV", f"{ltv:,.0f}€")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
