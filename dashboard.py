import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

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

# ================= FUNÇÕES AUX =================
def extrair_mes(nome):
    meses = {
        "JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,
        "JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12
    }
    nome = nome.upper()
    for k,v in meses.items():
        if k in nome:
            return v
    return 99

def normalizar(col):
    if isinstance(col, pd.Series):
        return col.astype(str).str.upper().str.strip()
    return ""

# ================= CACHE =================
@st.cache_data
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = f.name.replace(".xlsx", "").upper()
        df["Mes_num"] = extrair_mes(df["Periodo"].iloc[0])

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Nome do cliente"] = normalizar(df.get("Nome do cliente", ""))

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

@st.cache_data
def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = f.name.replace(".xlsx", "").upper()
        df["Mes_num"] = extrair_mes(df["Periodo"].iloc[0])

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= UPLOAD =================
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

receita_media = receitas.groupby("Periodo")["Valor"].sum().mean() if not receitas.empty else 0
clientes_media = receitas.groupby("Periodo")["Nome do cliente"].nunique().mean() if not receitas.empty else 0

ticket_medio_receita = receita_media / clientes_media if clientes_media else 0
ticket_medio_despesa = abs(despesa_total) / clientes_media if clientes_media else 0

magic_number = abs(despesa_total)

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Despesa", f"{despesa_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")
st.metric("Ticket Receita", f"{ticket_medio_receita:,.0f}€")
st.metric("Ticket Despesa", f"{ticket_medio_despesa:,.0f}€")
st.metric("Break-even", f"{magic_number:,.0f}€")

# ================= CLIENTES =================
st.subheader("👥 Clientes: Novos vs Perdidos + Churn")

if not receitas.empty:
    receitas = receitas.sort_values("Mes_num")

    clientes_mes = receitas.groupby("Periodo")["Nome do cliente"].apply(set)
    clientes_mes = clientes_mes.sort_index()

    novos = []
    perdidos = []
    churn_lista = []

    for i in range(len(clientes_mes)):
        if i == 0:
            novos.append(len(clientes_mes.iloc[i]))
            perdidos.append(0)
            churn_lista.append(0)
        else:
            atual = clientes_mes.iloc[i]
            anterior = clientes_mes.iloc[i-1]

            novos.append(len(atual - anterior))
            perdidos.append(len(anterior - atual))

            churn = (len(anterior - atual) / len(anterior))*100 if len(anterior)>0 else 0
            churn_lista.append(churn)

    df_clientes = pd.DataFrame({
        "Novos": novos,
        "Perdidos": perdidos,
        "Churn (%)": churn_lista
    }, index=clientes_mes.index)

    st.dataframe(df_clientes)

    fig, ax = plt.subplots()
    df_clientes[["Novos","Perdidos"]].plot(kind="bar", ax=ax)
    ax.set_title("Clientes Novos vs Perdidos")
    st.pyplot(fig)

    fig2, ax2 = plt.subplots()
    df_clientes["Churn (%)"].plot(ax=ax2, marker="o")
    ax2.set_title("Churn (%)")
    st.pyplot(fig2)

# ================= INSIGHTS =================
def gerar_insights():
    insights = []

    if receitas.empty:
        return insights

    clientes_por_mes = receitas.groupby("Periodo")["Nome do cliente"].nunique().sort_index()

    churn_lista = []

    for i in range(1, len(clientes_por_mes)):
        anterior = clientes_por_mes.iloc[i-1]
        atual = clientes_por_mes.iloc[i]

        if anterior > 0:
            churn = ((anterior - atual) / anterior) * 100
            churn_lista.append(churn)
        else:
            churn_lista.append(0)

    if len(churn_lista) > 0 and churn_lista[-1] > 10:
        insights.append(f"⚠️ Churn elevado no último mês: {churn_lista[-1]:.1f}%")

    if margem < 10:
        insights.append("⚠️ Margem baixa — risco operacional")

    if ticket_medio_receita < ticket_medio_despesa:
        insights.append("⚠️ Ticket médio abaixo do custo por cliente")

    if not insights:
        insights.append("✅ Negócio estável sem alertas críticos")

    return insights

st.subheader("🧠 Insights Automáticos")
for i in gerar_insights():
    st.write(i)

# ================= PDF =================
def gerar_pdf():
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("Relatório Financeiro", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    for i in gerar_insights():
        elementos.append(Paragraph(i, styles["Normal"]))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= PPT =================
def gerar_ppt():
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Resumo Executivo"
    slide.placeholders[1].text = f"Receita: {receita_total:,.0f}€ | Lucro: {lucro_total:,.0f}€"

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ================= EXPORT =================
st.subheader("📄 Exportações")

if st.button("Gerar PDF Completo"):
    st.download_button("Download PDF", gerar_pdf(), "relatorio.pdf")

if st.button("Gerar PPT Editável"):
    st.download_button("Download PPT", gerar_ppt(), "relatorio.pptx")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
