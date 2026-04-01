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

# ================= ORDEM MESES =================
ordem_meses = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3,
    "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9,
    "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}

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

        # 🔥 NORMALIZAÇÃO CRÍTICA
        df["Nome do cliente"] = (
            df.get("Nome do cliente", "")
            .astype(str)
            .str.upper()
            .str.strip()
        )

        df = df[df["Nome do cliente"] != ""]

        df["Modalidade"] = df.get("Modalidade", "N/A").astype(str).str.upper().str.strip()
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

receita_media = receitas.groupby("Periodo")["Valor"].sum().mean() if not receitas.empty else 0
despesa_media = despesas.groupby("Periodo")["Valor"].sum().mean() if not despesas.empty else 0
clientes_ativos_media = receitas.groupby("Periodo")["Nome do cliente"].nunique().mean() if not receitas.empty else 0

ticket_medio_receita = receita_media / clientes_ativos_media if clientes_ativos_media else 0
ticket_medio_despesa = abs(despesa_media) / clientes_ativos_media if clientes_ativos_media else 0

magic_number = abs(despesa_media)

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Despesa", f"{despesa_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")

st.metric("Ticket Médio Receita", f"{ticket_medio_receita:,.0f}€")
st.metric("Ticket Médio Despesa", f"{ticket_medio_despesa:,.0f}€")
st.metric("Magic Number (Break-even)", f"{magic_number:,.0f}€")

# ================= CLIENTES (FIXED) =================
st.subheader("👥 Evolução de Clientes")

if not receitas.empty:

    clientes_por_mes = receitas.groupby("Periodo")["Nome do cliente"].nunique()

    # ordenar corretamente
    clientes_por_mes = clientes_por_mes.sort_index(key=lambda x: x.map(ordem_meses))

    fig, ax = plt.subplots()
    clientes_por_mes.plot(kind="line", marker="o", ax=ax)

    ax.set_title("Clientes Ativos por Mês")
    ax.set_xlabel("Período")
    ax.set_ylabel("Clientes")

    plt.xticks(rotation=45)

    st.pyplot(fig)

# ================= CLIENTES POR MODALIDADE =================
st.subheader("🏋️ Clientes por Modalidade")

if not receitas.empty:
    clientes_modalidade = (
        receitas.groupby("Modalidade")["Nome do cliente"]
        .nunique()
        .sort_values(ascending=False)
    )

    st.dataframe(clientes_modalidade)

    fig_mod, ax_mod = plt.subplots()
    clientes_modalidade.plot(kind="barh", ax=ax_mod)
    ax_mod.set_title("Distribuição de Clientes por Modalidade")
    st.pyplot(fig_mod)

# ================= BLOCO ANALISE =================
def bloco_analise(df, categoria, titulo, figs_pdf):
    if df.empty or categoria not in df.columns:
        return

    pivot = df.pivot_table(index=categoria, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
    pivot = pivot.reindex(sorted(pivot.columns, key=lambda x: ordem_meses.get(x, 99)), axis=1)

    st.markdown(f"### {titulo} por {categoria}")
    st.dataframe(pivot)

    fig1 = grafico_bar(pivot, f"{titulo} por {categoria}")
    st.pyplot(fig1)

    fig2 = grafico_percentual(pivot, f"{titulo} por {categoria}")
    st.pyplot(fig2)

    figs_pdf.append((f"{titulo} - {categoria}", fig1, fig2, pivot))

# ================= GRÁFICOS AUX =================
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

# ================= TABS =================
figs_pdf = []

tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "💰 Receitas", "💸 Despesas"])

with tab1:
    st.write(f"Receita: {receita_total:,.0f}€, Lucro: {lucro_total:,.0f}€, Margem: {margem:.1f}%, Ticket Médio: {ticket_medio_receita:,.0f}€, Break-even: {magic_number:,.0f}€")

with tab2:
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        bloco_analise(receitas, cat, "Receitas", figs_pdf)

with tab3:
    for cat in ["Classe", "Local"]:
        bloco_analise(despesas, cat, "Despesas", figs_pdf)

# ================= PDF =================
def gerar_pdf(figs_pdf):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("RELATÓRIO COMPLETO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    for titulo, fig1, fig2, _ in figs_pdf:
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

# ================= PPT =================
def gerar_ppt():
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Resumo Executivo"
    slide.placeholders[1].text = f"Receita: {receita_total:,.0f}€ | Lucro: {lucro_total:,.0f}€"

    for titulo, fig1, fig2, pivot in figs_pdf:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = titulo

        chart_data = CategoryChartData()
        chart_data.categories = list(pivot.index)

        for col in pivot.columns:
            chart_data.add_series(str(col), list(pivot[col].values))

        slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            Inches(1), Inches(2), Inches(8), Inches(4),
            chart_data
        )

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ================= EXPORT =================
st.subheader("📄 Exportações")

if st.button("Gerar PDF Completo"):
    pdf = gerar_pdf(figs_pdf)
    st.download_button("Download PDF", pdf, "relatorio_completo.pdf")

if st.button("Gerar PPT Editável"):
    st.download_button("Download PPT", gerar_ppt(), "apresentacao_editavel.pptx")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
