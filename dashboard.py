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

st.set_page_config(layout="wide")

# ================= NORMALIZAR =================
def normalizar(txt):
    if pd.isna(txt):
        return ""
    txt = str(txt).upper().strip()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    return txt

# ================= MESES =================
mapa_meses = {"JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,"JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12}

def extrair_mes(nome):
    nome = normalizar(nome)
    for k,v in mapa_meses.items():
        if k in nome:
            return v
    return 99

# ================= LEITURA =================
@st.cache_data
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        periodo = f.name.split(".")[0]
        mes = extrair_mes(periodo)

        df["Periodo"] = periodo
        df["ordem_mes"] = mes
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Nome do cliente"] = df["Nome do cliente"].apply(normalizar)

        df = df[df["Nome do cliente"] != ""]

        df["Modalidade"] = df.get("Modalidade", "").apply(normalizar)
        df["Modalidade"] = df["Modalidade"].replace("", "SEM MODALIDADE")

        dfs.append(df)

    return pd.concat(dfs) if dfs else pd.DataFrame()

@st.cache_data
def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        periodo = f.name.split(".")[0]
        mes = extrair_mes(periodo)

        df["Periodo"] = periodo
        df["ordem_mes"] = mes
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Classe"] = df["Classe"].apply(normalizar)

        dfs.append(df)

    return pd.concat(dfs) if dfs else pd.DataFrame()

# ================= UPLOAD =================
rec_files = st.file_uploader("Receitas", accept_multiple_files=True)
desp_files = st.file_uploader("Despesas", accept_multiple_files=True)

receitas = ler_receitas(rec_files) if rec_files else pd.DataFrame()
despesas = ler_despesas(desp_files) if desp_files else pd.DataFrame()

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro = receita_total + despesa_total

receita_media = receitas.groupby("Periodo")["Valor"].sum().mean() if not receitas.empty else 0
despesa_media = despesas.groupby("Periodo")["Valor"].sum().mean() if not despesas.empty else 0
clientes_media = receitas.groupby("Periodo")["Nome do cliente"].nunique().mean() if not receitas.empty else 0

ticket_receita = receita_media / clientes_media if clientes_media else 0
ticket_despesa = abs(despesa_media) / clientes_media if clientes_media else 0

magic_number = abs(despesa_media)

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Despesa", f"{despesa_total:,.0f}€")
st.metric("Lucro", f"{lucro:,.0f}€")

st.metric("Ticket Receita", f"{ticket_receita:,.0f}€")
st.metric("Ticket Despesa", f"{ticket_despesa:,.0f}€")
st.metric("Break-even", f"{magic_number:,.0f}€")

# ================= CLIENTES (CHURN) =================
st.subheader("👥 Clientes: Novos vs Perdidos + Churn")

if not receitas.empty:
    base = receitas.groupby(["Periodo","ordem_mes"])["Nome do cliente"].apply(set).reset_index()
    base = base.sort_values("ordem_mes")

    novos, perdidos, churn = [], [], []

    for i in range(len(base)):
        if i == 0:
            novos.append(len(base.iloc[i]["Nome do cliente"]))
            perdidos.append(0)
            churn.append(0)
        else:
            atual = base.iloc[i]["Nome do cliente"]
            anterior = base.iloc[i-1]["Nome do cliente"]

            novos_mes = len(atual - anterior)
            perdidos_mes = len(anterior - atual)

            novos.append(novos_mes)
            perdidos.append(perdidos_mes)

            churn.append((perdidos_mes / len(anterior))*100 if len(anterior)>0 else 0)

    base["Novos"] = novos
    base["Perdidos"] = perdidos
    base["Churn %"] = churn

    st.dataframe(base[["Periodo","Novos","Perdidos","Churn %"]])

    fig, ax = plt.subplots()
    ax.plot(base["Periodo"], base["Novos"], label="Novos")
    ax.plot(base["Periodo"], base["Perdidos"], label="Perdidos")
    ax.legend()
    plt.xticks(rotation=45)
    st.pyplot(fig)

# ================= INSIGHTS =================
def gerar_insights():
    insights = []

    if churn and churn[-1] > 10:
        insights.append("⚠️ Churn elevado — risco de perda de clientes")

    if ticket_receita > ticket_despesa:
        insights.append("💰 Modelo saudável — receita por cliente cobre custos")

    if lucro < 0:
        insights.append("🔴 Negócio em prejuízo")

    if not insights:
        insights.append("✅ Negócio estável")

    return insights

st.subheader("🧠 Insights Automáticos")
for i in gerar_insights():
    st.write(i)

# ================= EXPORT =================
figs_pdf = []

def gerar_pdf():
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("Relatório Executivo", styles["Title"]))
    elementos.append(Spacer(1,1*cm))

    for i in gerar_insights():
        elementos.append(Paragraph(i, styles["Normal"]))

    elementos.append(PageBreak())

    doc.build(elementos)
    buffer.seek(0)
    return buffer

def gerar_ppt():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Resumo Executivo"

    tf = slide.placeholders[1].text = "\n".join(gerar_insights())

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

col1, col2 = st.columns(2)

with col1:
    st.download_button("📄 PDF", gerar_pdf(), "relatorio.pdf")

with col2:
    st.download_button("📊 PPT", gerar_ppt(), "relatorio.pptx")
