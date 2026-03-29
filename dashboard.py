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
st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

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
st.subheader("📌 KPIs")

receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0

# CORREÇÃO DO LUCRO
lucro_total = receita_total - despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("Receita", f"{receita_total:,.0f}€")
col2.metric("Despesa", f"{despesa_total:,.0f}€")
col3.metric("Lucro", f"{lucro_total:,.0f}€")
col4.metric("Margem", f"{margem:.1f}%")

# ================= KPIs POR PERÍODO =================
kpis = []
for p in periodos:
    r = receitas[receitas["Periodo"] == p]
    d = despesas[despesas["Periodo"] == p]

    receita = r["Valor"].sum()
    despesa = d["Valor"].sum()
    lucro = receita - despesa

    kpis.append({
        "Período": p,
        "Receita": receita,
        "Despesa": despesa,
        "Lucro": lucro
    })

df_kpis = pd.DataFrame(kpis)

if not df_kpis.empty:
    df_kpis["Margem (%)"] = (df_kpis["Lucro"] / df_kpis["Receita"]) * 100
    df_kpis["Δ Receita (%)"] = df_kpis["Receita"].pct_change() * 100
    df_kpis["Δ Lucro (%)"] = df_kpis["Lucro"].pct_change() * 100

    st.dataframe(df_kpis.round(2), use_container_width=True)

# ================= GRÁFICO RESUMO =================
def gerar_fig_resumo(df):
    fig, ax = plt.subplots()
    df.set_index("Período")[["Receita", "Despesa", "Lucro"]].plot(kind="bar", ax=ax)
    ax.set_title("Resumo Financeiro")
    return fig

fig_resumo = gerar_fig_resumo(df_kpis) if not df_kpis.empty else None

if fig_resumo:
    st.pyplot(fig_resumo)

# ================= PDF =================
st.subheader("📄 Relatório Executivo PDF")

def gerar_pdf(df_kpis, fig_resumo):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    elementos = []

    # CAPA
    elementos.append(Spacer(1, 6*cm))
    elementos.append(Paragraph("RELATÓRIO FINANCEIRO", styles["Title"]))
    elementos.append(Paragraph("Resumo Executivo", styles["Heading2"]))
    elementos.append(Spacer(1, 1*cm))
    elementos.append(Paragraph(datetime.now().strftime("%d/%m/%Y"), styles["Normal"]))
    elementos.append(PageBreak())

    # KPIs
    elementos.append(Paragraph("Resumo Financeiro", styles["Heading1"]))

    tabela = [["Período", "Receita", "Despesa", "Lucro", "Margem"]]
    for _, row in df_kpis.iterrows():
        tabela.append([
            row["Período"],
            f"{row['Receita']:,.2f}€",
            f"{row['Despesa']:,.2f}€",
            f"{row['Lucro']:,.2f}€",
            f"{row['Margem (%)']:.1f}%"
        ])

    elementos.append(Table(tabela))
    elementos.append(Spacer(1, 1*cm))

    # INSIGHT AUTOMÁTICO
    lucro_total = df_kpis["Lucro"].sum()
    texto = "Resultado positivo." if lucro_total > 0 else "Atenção: prejuízo no período."
    elementos.append(Paragraph(texto, styles["Normal"]))
    elementos.append(Spacer(1, 1*cm))

    # GRÁFICO
    if fig_resumo:
        img = BytesIO()
        fig_resumo.savefig(img, format="png", bbox_inches="tight")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=9*cm))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

if st.button("Gerar PDF Executivo"):
    pdf = gerar_pdf(df_kpis, fig_resumo)

    st.download_button(
        label="📥 Download PDF",
        data=pdf,
        file_name="relatorio_executivo.pdf",
        mime="application/pdf"
    )

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
