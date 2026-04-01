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
mapa_meses = {"JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,"JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12}

def extrair_mes(nome):
    nome = normalizar(nome)
    match = re.search(r'\\b(0?[1-9]|1[0-2])\\b', nome)
    if match:
        return int(match.group())
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

        periodo_nome = f.name.split(".")[0]
        mes = extrair_mes(periodo_nome)

        df["Periodo"] = periodo_nome.upper()
        df["ordem_mes"] = mes

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Nome do cliente"] = df.get("Nome do cliente", "").apply(normalizar)
        df = df[df["Nome do cliente"] != ""]

        for col in ["Modalidade","Tipo","Professor","Local"]:
            if col not in df.columns:
                df[col] = "N/A"

        df["Modalidade"] = df["Modalidade"].apply(normalizar)
        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

@st.cache_data(ttl=3600)
def ler_despesas(files):
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

        for col in ["Classe","Local"]:
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

    mapa = {"NOME DO CLIENTE":"Nome do cliente","CLIENTE":"Nome do cliente","DATA INICIO":"Data Inicio","DATA DE INICIO":"Data Inicio"}
    df = df.rename(columns=mapa)

    if "Nome do cliente" not in df.columns or "Data Inicio" not in df.columns:
        return pd.DataFrame()

    df["Nome do cliente"] = df["Nome do cliente"].apply(normalizar)
    df["Data Inicio"] = pd.to_datetime(df["Data Inicio"], errors="coerce")
    return df

# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    return fig

def grafico_bar_percentual(df, titulo):
    df_pct = df.div(df.sum(axis=0), axis=1).fillna(0)*100
    fig, ax = plt.subplots()
    df_pct.plot(kind="barh", ax=ax)
    ax.set_title(titulo+" (%)")
    return fig

# ================= EXPORT HELPERS =================
figs_pdf = []

def capturar(fig, titulo):
    figs_pdf.append((titulo, fig))

# ================= UPLOAD =================
st.sidebar.header("📤 Upload")
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)
uploaded_clientes = st.sidebar.file_uploader("Clientes", type=["xlsx"])

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

col1,col2,col3 = st.columns(3)
col1.metric("Receita", f"{receita_total:,.0f}€")
col2.metric("Despesa", f"{despesa_total:,.0f}€")
col3.metric("Lucro", f"{lucro_total:,.0f}€")

# ================= EVOLUÇÃO =================
st.subheader("📈 Receita vs Despesa vs Lucro")
if not receitas.empty:
    receita_mes = receitas.groupby("Periodo")["Valor"].sum()
    despesa_mes = despesas.groupby("Periodo")["Valor"].sum() if not despesas.empty else receita_mes*0
    lucro_mes = receita_mes + despesa_mes

    df_ev = pd.DataFrame({"Receita":receita_mes,"Despesa":despesa_mes,"Lucro":lucro_mes})
    st.line_chart(df_ev)

# ================= TOP CLIENTES =================
st.subheader("🏆 Top Clientes")
if not receitas.empty:
    top = receitas.groupby("Nome do cliente")["Valor"].sum().sort_values(ascending=False).head(10)
    st.dataframe(top)

# ================= ALERTAS =================
st.subheader("🚨 Alertas")
if lucro_total < 0:
    st.error("Prejuízo detectado")
if receita_total > 0 and (lucro_total/receita_total)<0.2:
    st.warning("Margem baixa")

# ================= KPIs AVANÇADOS =================
st.subheader("🧠 KPIs Avançados")

if not receitas.empty:
    clientes = receitas["Nome do cliente"].nunique()
    ticket = receita_total/clientes if clientes else 0
else:
    ticket = 0

cac = abs(despesa_total)/clientes if clientes else 0
ltv = ticket*6

col1,col2,col3 = st.columns(3)
col1.metric("Ticket", f"{ticket:,.0f}€")
col2.metric("CAC", f"{cac:,.0f}€")
col3.metric("LTV", f"{ltv:,.0f}€")

# ================= EXPORT PDF =================
def gerar_pdf():
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elems = []

    elems.append(Paragraph("Dashboard Financeiro", styles['Title']))

    for titulo, fig in figs_pdf:
        img = BytesIO()
        fig.savefig(img, format='png')
        img.seek(0)
        elems.append(Paragraph(titulo, styles['Heading2']))
        elems.append(Image(img, width=16*cm, height=8*cm))
        elems.append(PageBreak())

    doc.build(elems)
    buffer.seek(0)
    return buffer

# ================= EXPORT PPT =================
def gerar_ppt():
    prs = Presentation()
    for titulo, fig in figs_pdf:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = titulo
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

st.download_button("📄 Download PDF", gerar_pdf(), "relatorio.pdf")
st.download_button("📊 Download PPT", gerar_ppt(), "relatorio.pptx")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
