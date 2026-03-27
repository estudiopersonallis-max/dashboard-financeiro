import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import matplotlib

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from datetime import datetime

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= UPLOAD =================
st.subheader("📤 Upload de Ficheiros (cada ficheiro = um período)")
uploaded_receitas = st.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

# ================= FUNÇÕES =================
def nome_periodo(nome):
    return nome.replace(".xlsx", "").upper()

def limpar_valor(valor):
    if isinstance(valor, str):
        valor = valor.replace("€", "").replace(" ", "").replace(",", ".")
    return pd.to_numeric(valor, errors="coerce")

def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)

        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)

        dfs.append(df)

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame(columns=["Periodo","Valor"])

def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)

        if df.empty:
            continue

        df.columns = df.columns.str.strip()

        if "Valor" not in df.columns:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = df["Valor"].apply(limpar_valor)

        # manter só despesas reais (negativas)
        df = df[df["Valor"] < 0]

        dfs.append(df)

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame(columns=["Periodo","Valor"])

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else ler_receitas([])
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else ler_despesas([])

# ================= KPIs =================
st.subheader("📌 KPIs Comparativos")

periodos = sorted(set(receitas["Periodo"]).union(set(despesas["Periodo"])))
kpis = []

for p in periodos:
    r = receitas[receitas["Periodo"] == p]
    d = despesas[despesas["Periodo"] == p]

    receita = r["Valor"].sum()
    despesa = d["Valor"].sum()
    lucro = receita + despesa  # despesas já negativas

    kpis.append({
        "Período": p,
        "Receita (€)": round(receita, 2),
        "Despesa (€)": round(despesa, 2),
        "Lucro (€)": round(lucro, 2)
    })

df_kpis = pd.DataFrame(kpis)
st.dataframe(df_kpis, use_container_width=True)

# ================= GRÁFICO =================
def grafico_resumo(df):
    if df.empty:
        return None

    fig, ax = plt.subplots()
    df.set_index("Período")[["Receita (€)", "Despesa (€)", "Lucro (€)"]].plot(kind="bar", ax=ax)

    ax.set_title("Resumo Financeiro")
    ax.set_ylabel("€")

    return fig

fig_resumo = grafico_resumo(df_kpis)

if fig_resumo:
    st.pyplot(fig_resumo)

# ================= PDF EXECUTIVO =================
st.subheader("📄 Exportar PDF Executivo")

def gerar_pdf_executivo(df_kpis, fig):

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    elementos = []

    # ===== CAPA =====
    elementos.append(Spacer(1, 6*cm))
    elementos.append(Paragraph("RELATÓRIO FINANCEIRO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))
    elementos.append(Paragraph("Dashboard Financeiro", styles["Heading2"]))
    elementos.append(Spacer(1, 1*cm))

    data_hoje = datetime.now().strftime("%d/%m/%Y")
    elementos.append(Paragraph(f"Data: {data_hoje}", styles["Normal"]))

    elementos.append(PageBreak())

    # ===== RESUMO =====
    elementos.append(Paragraph("Resumo Executivo", styles["Heading1"]))
    elementos.append(Spacer(1, 0.5*cm))

    tabela_data = [["Período", "Receita (€)", "Despesa (€)", "Lucro (€)"]]

    for _, row in df_kpis.iterrows():
        tabela_data.append([
            row["Período"],
            f"{row['Receita (€)']:,.2f}",
            f"{row['Despesa (€)']:,.2f}",
            f"{row['Lucro (€)']:,.2f}"
        ])

    tabela = Table(tabela_data, hAlign='LEFT')

    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ]))

    elementos.append(tabela)
    elementos.append(PageBreak())

    # ===== GRÁFICO =====
    elementos.append(Paragraph("Análise Gráfica", styles["Heading1"]))
    elementos.append(Spacer(1, 0.5*cm))

    if fig:
        img_buffer = BytesIO()
        fig.savefig(img_buffer, format="png", bbox_inches="tight")
        img_buffer.seek(0)

        elementos.append(Image(img_buffer, width=16*cm, height=9*cm))

    # ===== BUILD =====
    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= BOTÃO =================
if st.button("📄 Gerar PDF Executivo"):
    pdf = gerar_pdf_executivo(df_kpis, fig_resumo)

    st.download_button(
        label="📥 Download PDF Executivo",
        data=pdf,
        file_name="relatorio_financeiro_executivo.pdf",
        mime="application/pdf"
    )
