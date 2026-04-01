import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
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
    if not receitas.empty or not despesas.empty:
        df_kpis = pd.DataFrame({
            "Receita": [receita_total],
            "Despesa": [despesa_total],
            "Lucro": [lucro_total]
        })
        fig = grafico_bar(df_kpis.T, "Resumo Financeiro")
        st.pyplot(fig)
        figs_pdf.append(("Resumo", fig, None))

with tab2:
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        bloco_analise(receitas, cat, "Receitas", figs_pdf)

with tab3:
    for cat in ["Classe", "Local"]:
        bloco_analise(despesas, cat, "Despesas", figs_pdf)

# ================= PDF COMPLETO =================
def gerar_pdf(figs_pdf):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("RELATÓRIO COMPLETO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    for titulo, fig1, fig2 in figs_pdf:
        elementos.append(Paragraph(titulo, styles["Heading2"]))

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

st.subheader("📄 Exportar PDF Completo")
if st.button("Gerar PDF Completo"):
    pdf = gerar_pdf(figs_pdf)
    st.download_button("Download PDF", pdf, "relatorio_completo.pdf")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
