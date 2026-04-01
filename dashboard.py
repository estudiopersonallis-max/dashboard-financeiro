import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO
from datetime import datetime

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

# ================= CONFIG =================
st.set_page_config(page_title="Dashboard Financeiro PRO", layout="wide")
st.title("📊 Dashboard Financeiro – Nível Consultoria")

# ================= CACHE =================
@st.cache_data
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue
        df["Periodo"] = f.name.replace(".xlsx", "").upper()
        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Nome do cliente"] = df.get("Nome do cliente", "").astype(str).str.upper().str.strip()
        df["Modalidade"] = df.get("Modalidade", "N/A")
        df["Tipo"] = df.get("Tipo", "N/A")
        df["Professor"] = df.get("Professor", "N/A")
        df["Local"] = df.get("Local", "N/A")
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

@st.cache_data
def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        df = df.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df.empty:
            continue
        df["Periodo"] = f.name.replace(".xlsx", "").upper()
        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Classe"] = df.get("Classe", "N/A").astype(str).str.upper().str.strip()
        df["Local"] = df.get("Local", "N/A").astype(str).str.strip()
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= FUNÇÕES GRÁFICOS =================
def grafico_bar(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    return fig


def grafico_percentual(df, titulo):
    percent = df.div(df.sum(axis=0), axis=1) * 100
    fig, ax = plt.subplots()
    percent.plot(kind="barh", ax=ax)
    ax.set_title(titulo + " (%)")
    return fig


def grafico_pareto(series, titulo):
    series = series.sort_values(ascending=False)
    cum = series.cumsum() / series.sum() * 100
    fig, ax = plt.subplots()
    series.plot(kind="bar", ax=ax)
    cum.plot(ax=ax, secondary_y=True)
    ax.set_title(titulo + " (Pareto)")
    return fig


def grafico_heatmap(df, titulo):
    fig, ax = plt.subplots()
    cax = ax.imshow(df, aspect='auto')
    ax.set_title(titulo)
    fig.colorbar(cax)
    return fig


def grafico_waterfall(receita, despesa):
    fig, ax = plt.subplots()
    valores = [receita, despesa, receita+despesa]
    labels = ["Receita", "Custos", "Lucro"]
    ax.bar(labels, valores)
    ax.set_title("Waterfall Lucro")
    return fig

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

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Despesa", f"{despesa_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")

# ================= INSIGHTS =================
def gerar_insights():
    insights = []
    if margem < 10:
        insights.append("Margem pressionada indica necessidade de revisão de custos.")
    if lucro_total < 0:
        insights.append("Operação deficitária requer ação imediata.")
    return insights


def gerar_problemas():
    problemas = []
    if margem < 10:
        problemas.append("Baixa rentabilidade")
    if lucro_total < 0:
        problemas.append("Prejuízo operacional")
    return problemas[:3]


def gerar_oportunidades():
    oportunidades = []
    if margem > 20:
        oportunidades.append("Escalar operação atual")
    if not receitas.empty:
        oportunidades.append("Explorar clientes top")
    return oportunidades[:3]

st.subheader("🧠 Insights")
for i in gerar_insights():
    st.write("•", i)

st.subheader("🚨 Problemas")
for p in gerar_problemas():
    st.write("•", p)

st.subheader("🚀 Oportunidades")
for o in gerar_oportunidades():
    st.write("•", o)

# ================= GRÁFICOS EXECUTIVOS =================
st.subheader("📊 Análises Executivas")

if not receitas.empty:
    serie = receitas.groupby("Nome do cliente")["Valor"].sum()
    st.pyplot(grafico_pareto(serie, "Clientes"))

if not receitas.empty:
    pivot = receitas.pivot_table(index="Modalidade", columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
    st.pyplot(grafico_heatmap(pivot, "Heatmap Receitas"))

st.pyplot(grafico_waterfall(receita_total, despesa_total))

# ================= PDF =================
def gerar_pdf():
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("RELATÓRIO EXECUTIVO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    elementos.append(Paragraph("Sumário Executivo", styles["Heading2"]))
    elementos.append(Paragraph(f"Receita: {receita_total:,.0f}€, Lucro: {lucro_total:,.0f}€, Margem: {margem:.1f}%", styles["Normal"]))

    elementos.append(Spacer(1, 1*cm))

    elementos.append(Paragraph("Insights", styles["Heading2"]))
    for i in gerar_insights():
        elementos.append(Paragraph(i, styles["Normal"]))

    elementos.append(PageBreak())

    doc.build(elementos)
    buffer.seek(0)
    return buffer

st.subheader("📄 Exportar PDF Executivo")
if st.button("Gerar PDF"):
    pdf = gerar_pdf()
    st.download_button("Download", pdf, "relatorio_executivo.pdf")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
