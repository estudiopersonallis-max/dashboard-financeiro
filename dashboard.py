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

clientes_unicos = receitas["Nome do cliente"].nunique() if not receitas.empty else 0
ticket_medio_receita = receita_total / clientes_unicos if clientes_unicos else 0

linhas_despesa = len(despesas) if not despesas.empty else 0
ticket_medio_despesa = abs(despesa_total) / linhas_despesa if linhas_despesa else 0

magic_number = abs(despesa_total)

# ================= SIMULADOR =================
clientes_para_break_even = (magic_number / ticket_medio_receita) if ticket_medio_receita else 0

aumento_ticket_10 = ticket_medio_receita * 1.1
novo_lucro_ticket = aumento_ticket_10 * clientes_unicos + despesa_total

reducao_custo_10 = despesa_total * 0.9
novo_lucro_custo = receita_total + reducao_custo_10

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")

st.metric("Ticket Médio", f"{ticket_medio_receita:,.0f}€")
st.metric("Break-even", f"{magic_number:,.0f}€")

st.subheader("🧠 Simulador")
st.write(f"Clientes necessários para break-even: {clientes_para_break_even:.0f}")
st.write(f"Lucro com +10% ticket: {novo_lucro_ticket:,.0f}€")
st.write(f"Lucro com -10% custo: {novo_lucro_custo:,.0f}€")

# ================= PPT EDITÁVEL =================
def gerar_ppt():
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Resumo Executivo"
    slide.placeholders[1].text = f"Receita: {receita_total:,.0f}€ | Lucro: {lucro_total:,.0f}€"

    if not receitas.empty:
        dados = receitas.groupby("Modalidade")["Valor"].sum()
        chart_data = CategoryChartData()
        chart_data.categories = list(dados.index)
        chart_data.add_series('Receita', list(dados.values))

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Receita por Modalidade"

        slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1), Inches(2), Inches(8), Inches(4),
            chart_data
        )

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ================= EXPORT =================
st.subheader("📄 Exportações")

if st.button("Gerar PPT Editável"):
    st.download_button("Download PPT", gerar_ppt(), "apresentacao_editavel.pptx")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
