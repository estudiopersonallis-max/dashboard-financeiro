import streamlit as st
import pandas as pd
import datetime

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro")

# ======================================================
# UPLOAD
# ======================================================
uploaded_files = st.file_uploader(
    "📤 Carregue um ficheiro Excel por mês",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("⬆️ Carregue pelo menos um ficheiro Excel para iniciar o dashboard")
    st.stop()

# ======================================================
# LEITURA DOS FICHEIROS
# ======================================================
dfs = []

for file in uploaded_files:
    df_temp = pd.read_excel(file)

    # Mês definido pelo nome do ficheiro
    mes_ficheiro = file.name.replace(".xlsx", "")
    df_temp["Mes"] = mes_ficheiro

    df_temp["Data"] = pd.to_datetime(df_temp["Data"])
    df_temp["Dia"] = df_temp["Data"].dt.day

    # Perdas
    df_temp["É Perda"] = df_temp["Perdas"].notna()

    dfs.append(df_temp)

df = pd.concat(dfs, ignore_index=True)

# ======================================================
# FILTRO POR MÊS (FICHEIRO)
# ======================================================
mes_selecionado = st.selectbox(
    "📅 Selecione o mês",
    sorted(df["Mes"].unique())
)

df_filtro = df[df["Mes"] == mes_selecionado]

# ======================================================
# KPIs
# ======================================================
clientes_ativos = df_filtro[~df_filtro["É Perda"]]["Nome do cliente"].nunique()
perdas = df_filtro["É Perda"].sum()

total_valor = df_filtro["Valor"].sum()
ticket_medio = df_filtro["Valor"].mean()

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()

# ======================================================
# TABELAS
# ======================================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("📌 Valor por Modalidade")
    valor_modalidade = df_filtro.groupby("Modalidade")["Valor"].sum()
    st.dataframe(valor_modalidade)

    st.subheader("📌 Valor por Tipo")
    valor_tipo = df_filtro.groupby("Tipo")["Valor"].sum()
    st.dataframe(valor_tipo)

with col2:
    st.subheader("📌 Valor por Professor")
    valor_professor = df_filtro.groupby("Professor")["Valor"].sum()
    st.dataframe(valor_professor)

    st.subheader("📌 Valor por Local")
    valor_local = df_filtro.groupby("Local")["Valor"].sum()
    st.dataframe(valor_local)

st.divider()

# ======================================================
# PERÍODOS DO MÊS
# ======================================================
st.subheader("📅 Valor por Período do Mês")

p1 = df_filtro[df_filtro["Dia"] <= 10]["Valor"].sum()
p2 = df_filtro[(df_filtro["Dia"] > 10) & (df_filtro["Dia"] <= 20)]["Valor"].sum()
p3 = df_filtro[df_filtro["Dia"] > 20]["Valor"].sum()

valor_periodo = pd.Series({
    "Dias 1–10": p1,
    "Dias 11–20": p2,
    "Dias 21–fim": p3
})

st.dataframe(valor_periodo)

st.divider()

# ======================================================
# CLIENTES
# ======================================================
st.subheader("👥 Clientes")

col1, col2 = st.columns(2)

with col1:
    clientes_local = df_filtro.groupby("Local")["Nome do cliente"].nunique()
    st.dataframe(clientes_local.rename("Clientes por Local"))

with col2:
    clientes_professor = df_filtro.groupby("Professor")["Nome do cliente"].nunique()
    st.dataframe(clientes_professor.rename("Clientes por Professor"))

st.divider()

st.subheader("🎟️ Ticket Médio por Tipo")
ticket_tipo = df_filtro.groupby("Tipo")["Valor"].mean()
st.dataframe(ticket_tipo)

# ======================================================
# GRÁFICOS
# ======================================================
st.divider()
st.header("📊 Gráficos")

st.bar_chart(valor_modalidade)
st.bar_chart(valor_tipo)
st.bar_chart(valor_professor)
st.bar_chart(valor_local)
st.bar_chart(valor_periodo)
st.bar_chart(clientes_local)
st.bar_chart(clientes_professor)
st.bar_chart(ticket_tipo)

# ======================================================
# RELATÓRIO PDF
# =======================================
