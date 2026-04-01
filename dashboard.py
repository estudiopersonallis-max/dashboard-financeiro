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
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= DATA =================
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro_total = receita_total + despesa_total

clientes = receitas.groupby("Nome do cliente")["Valor"].sum().reset_index()
clientes.columns = ["Cliente", "Receita"]

custo_medio = abs(despesa_total) / len(clientes) if len(clientes) else 0
clientes["Custo Estimado"] = custo_medio
clientes["Lucro"] = clientes["Receita"] - clientes["Custo Estimado"]

st.subheader("💡 Rentabilidade por Cliente")
st.dataframe(clientes.sort_values("Lucro", ascending=False))

# ================= SIMULAÇÕES =================
st.subheader("🧠 Simulações")

preco_up = receita_total * 1.1
custo_down = despesa_total * 0.9

st.write(f"Receita +10% → Lucro: {(preco_up + despesa_total):,.0f}€")
st.write(f"Custos -10% → Lucro: {(receita_total + custo_down):,.0f}€")

# ================= GRÁFICOS =================
figs_pdf = []

def grafico(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    return fig

if not clientes.empty:
    top = clientes.sort_values("Receita", ascending=False).head(10).set_index("Cliente")
    fig = grafico(top[["Receita"]], "Top Clientes")
    st.pyplot(fig)
    figs_pdf.append(("Top Clientes", fig, None, top))

# ================= PDF =================
def gerar_pdf(figs_pdf):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("Relatório Executivo", styles["Title"]))

    for titulo, fig1, _, _ in figs_pdf:
        elementos.append(Paragraph(titulo, styles["Heading2"]))
        img = BytesIO()
        fig1.savefig(img, format="png")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=8*cm))
        elementos.append(PageBreak())

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= PPT =================
def gerar_ppt():
    prs = Presentation()

    # Slide resumo
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Visão Geral"
    slide.placeholders[1].text = f"Receita: {receita_total:,.0f}€ | Lucro: {lucro_total:,.0f}€"

    for titulo, _, _, pivot in figs_pdf:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"{titulo} – Receita concentrada"

        chart_data = CategoryChartData()
        chart_data.categories = list(pivot.index)
        chart_data.add_series("Receita", list(pivot["Receita"]))

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
if st.button("Gerar PDF"):
    st.download_button("Download PDF", gerar_pdf(figs_pdf), "relatorio.pdf")

if st.button("Gerar PPT Estratégico"):
    st.download_button("Download PPT", gerar_ppt(), "apresentacao.pptx")

st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
