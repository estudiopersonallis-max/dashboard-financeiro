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

# ================= DATA =================
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
        df["Local"] = df.get("Local", "N/A")
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
        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Local"] = df.get("Local", "N/A")
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

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
ticket_medio = receita_total / clientes_unicos if clientes_unicos else 0

magic_number = abs(despesa_total)

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")
st.metric("Ticket Médio", f"{ticket_medio:,.0f}€")
st.metric("Break-even", f"{magic_number:,.0f}€")

# ================= RENTABILIDADE REAL =================
st.subheader("💡 Rentabilidade por Cliente (com alocação por Local)")

if not receitas.empty and not despesas.empty:
    custo_local = despesas.groupby("Local")["Valor"].sum()

    receitas["Custo Alocado"] = receitas["Local"].map(custo_local) / receitas.groupby("Local")["Valor"].transform("count")

    cliente = receitas.groupby("Nome do cliente").agg({
        "Valor": "sum",
        "Custo Alocado": "sum"
    })

    cliente["Lucro"] = cliente["Valor"] + cliente["Custo Alocado"]
    cliente = cliente.sort_values("Lucro", ascending=False)

    st.dataframe(cliente)

# ================= RISCO (PARETO) =================
st.subheader("⚠️ Risco de Concentração")

if not receitas.empty:
    pareto = receitas.groupby("Nome do cliente")["Valor"].sum().sort_values(ascending=False)
    pareto_pct = pareto.cumsum() / pareto.sum() * 100

    top1 = pareto.iloc[0] / pareto.sum() * 100 if len(pareto) else 0
    st.write(f"Top 1 cliente representa {top1:.1f}% da receita")

    fig_pareto, ax = plt.subplots()
    pareto_pct.plot(ax=ax)
    ax.set_title("Pareto Clientes (%)")
    st.pyplot(fig_pareto)

# ================= SIMULAÇÕES =================
st.subheader("🧠 Simulações")

preco_slider = st.slider("Aumento de preço (%)", 0, 30, 10)
custo_slider = st.slider("Redução de custos (%)", 0, 30, 10)

nova_receita = receita_total * (1 + preco_slider/100)
novo_custo = despesa_total * (1 - custo_slider/100)
novo_lucro = nova_receita + novo_custo

st.write(f"Novo lucro: {novo_lucro:,.0f}€")

# ================= GRÁFICOS =================
figs_pdf = []

def grafico(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    return fig

if not receitas.empty:
    top = receitas.groupby("Nome do cliente")["Valor"].sum().sort_values(ascending=False).head(10)
    fig = grafico(top, "Top Clientes")
    st.pyplot(fig)
    figs_pdf.append(("Top Clientes", fig, top))

# ================= PDF =================
def gerar_pdf(figs_pdf):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("Relatório Executivo", styles["Title"]))

    for titulo, fig, _ in figs_pdf:
        elementos.append(Paragraph(f"{titulo} - análise de concentração", styles["Heading2"]))
        img = BytesIO()
        fig.savefig(img, format="png")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=8*cm))
        elementos.append(PageBreak())

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= PPT =================
def gerar_ppt():
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Resumo Executivo"
    slide.placeholders[1].text = f"Receita: {receita_total:,.0f}€ | Lucro: {lucro_total:,.0f}€"

    for titulo, _, data in figs_pdf:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"{titulo} – Receita concentrada"

        chart_data = CategoryChartData()
        chart_data.categories = list(data.index)
        chart_data.add_series("Receita", list(data.values))

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
if st.button("Gerar PDF Completo"):
    st.download_button("Download PDF", gerar_pdf(figs_pdf), "relatorio.pdf")

if st.button("Gerar PPT Editável"):
    st.download_button("Download PPT", gerar_ppt(), "apresentacao.pptx")

st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
