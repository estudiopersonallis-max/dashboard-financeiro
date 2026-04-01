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

# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    ax.set_xlabel("€")
    return fig


def grafico_percentual(df, titulo):
    percent = df.div(df.sum(axis=0), axis=1) * 100
    fig, ax = plt.subplots()
    percent.plot(kind="barh", ax=ax)
    ax.set_title(titulo + " (%)")
    ax.set_xlabel("%")
    return fig

# ================= UPLOAD =================
st.sidebar.header("📤 Upload")
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= FILTROS =================
st.sidebar.header("🔎 Filtros")
periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))
periodo_sel = st.sidebar.multiselect("Períodos", periodos, default=periodos)

if not receitas.empty:
    receitas = receitas[receitas["Periodo"].isin(periodo_sel)]

if not despesas.empty:
    despesas = despesas[despesas["Periodo"].isin(periodo_sel)]
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Despesa", f"{despesa_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")

# ================= NARRATIVA =================
def narrativa_geral():
    return f"A operação gerou {receita_total:,.0f}€ de receita, com lucro de {lucro_total:,.0f}€ e margem de {margem:.1f}%."

def narrativa_receita():
    if receitas.empty:
        return "Sem dados de receita."
    top = receitas.groupby("Modalidade")["Valor"].sum().idxmax()
    return f"A principal fonte de receita é {top}."

def narrativa_custos():
    if despesas.empty:
        return "Sem dados de custos."
    top = despesas.groupby("Classe")["Valor"].sum().idxmin()
    return f"O maior centro de custo é {top}."

# ================= BLOCO ANALISE =================
def bloco_analise(df, categoria, titulo, figs_pdf):
    if df.empty or categoria not in df.columns:
        return

    pivot = df.pivot_table(index=categoria, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)

    st.markdown(f"### {titulo} por {categoria}")
    st.dataframe(pivot)

    fig1 = grafico_bar(pivot, f"{titulo} por {categoria}")
    st.pyplot(fig1)

    fig2 = grafico_percentual(pivot, f"{titulo} por {categoria}")
    st.pyplot(fig2)

    figs_pdf.append((f"{titulo} - {categoria}", fig1, fig2))

# ================= TABS =================
figs_pdf = []

tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "💰 Receitas", "💸 Despesas"])

with tab1:
    st.write(narrativa_geral())

with tab2:
    st.write(narrativa_receita())
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        bloco_analise(receitas, cat, "Receitas", figs_pdf)

with tab3:
    st.write(narrativa_custos())
    for cat in ["Classe", "Local"]:
        bloco_analise(despesas, cat, "Despesas", figs_pdf)

# ================= PDF COMPLETO =================
def gerar_pdf():
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("RELATÓRIO EXECUTIVO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    # Capítulos
    elementos.append(Paragraph("1. Visão Geral", styles["Heading2"]))
    elementos.append(Paragraph(narrativa_geral(), styles["Normal"]))

    elementos.append(Paragraph("2. Receita", styles["Heading2"]))
    elementos.append(Paragraph(narrativa_receita(), styles["Normal"]))

    elementos.append(Paragraph("3. Custos", styles["Heading2"]))
    elementos.append(Paragraph(narrativa_custos(), styles["Normal"]))

    elementos.append(PageBreak())

    for titulo, fig1, fig2 in figs_pdf:
        elementos.append(Paragraph(titulo, styles["Heading3"]))

        if fig1:
            img = BytesIO()
            fig1.savefig(img, format="png", bbox_inches="tight")
            img.seek(0)
            elementos.append(Image(img, width=16*cm, height=8*cm))

        if fig2:
            img2 = BytesIO()
            fig2.savefig(img2, format="png", bbox_inches="tight")
            img2.seek(0)
            elementos.append(Image(img2, width=16*cm, height=8*cm))

        elementos.append(PageBreak())

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= PDF BIG4 =================
def gerar_pdf_big4():
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("RELATÓRIO ESTRATÉGICO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    elementos.append(Paragraph("Sumário Executivo", styles["Heading2"]))
    elementos.append(Paragraph(narrativa_geral(), styles["Normal"]))

    elementos.append(Paragraph("Diagnóstico", styles["Heading2"]))
    elementos.append(Paragraph(narrativa_receita() + " " + narrativa_custos(), styles["Normal"]))

    elementos.append(Paragraph("Recomendações", styles["Heading2"]))
    elementos.append(Paragraph("Recomenda-se otimização de custos e expansão de receita.", styles["Normal"]))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= PPT =================
def gerar_ppt():
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Resumo Executivo"
    slide.placeholders[1].text = narrativa_geral()

    for titulo, fig1, fig2 in figs_pdf:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = titulo

        if fig1:
            img = BytesIO()
            fig1.savefig(img, format="png")
            img.seek(0)
            slide.shapes.add_picture(img, Inches(1), Inches(1), width=Inches(8))

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ================= EXPORT =================
st.subheader("📄 Exportações")

if st.button("Gerar PDF Completo"):
    st.download_button("Download PDF", gerar_pdf(), "relatorio.pdf")

if st.button("Gerar PDF Estratégico (Big4)"):
    st.download_button("Download PDF Big4", gerar_pdf_big4(), "relatorio_big4.pdf")

if st.button("Gerar Apresentação (PPT)"):
    st.download_button("Download PPT", gerar_ppt(), "apresentacao.pptx")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
