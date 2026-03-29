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
lucro_total = receita_total - despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

# ================= KPIs POR PERÍODO =================
periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))

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

# ================= GRÁFICO =================
def gerar_fig_resumo(df):
    fig, ax = plt.subplots()
    df.set_index("Período")[["Receita", "Despesa", "Lucro"]].plot(kind="bar", ax=ax)
    return fig

fig_resumo = gerar_fig_resumo(df_kpis) if not df_kpis.empty else None

# ================= INSIGHTS =================
def gerar_insights(df):
    textos = []
    if df.empty:
        return textos

    if len(df) > 1:
        ult = df.iloc[-1]
        prev = df.iloc[-2]

        var_receita = ((ult["Receita"] - prev["Receita"]) / prev["Receita"] * 100) if prev["Receita"] else 0
        var_lucro = ((ult["Lucro"] - prev["Lucro"]) / prev["Lucro"] * 100) if prev["Lucro"] else 0

        textos.append(f"A receita variou {var_receita:.1f}% no último período.")
        textos.append(f"O lucro variou {var_lucro:.1f}% no último período.")

    if df["Lucro"].sum() > 0:
        textos.append("A operação apresenta resultado positivo no período analisado.")
    else:
        textos.append("A operação apresenta prejuízo e requer atenção.")

    return textos

insights = gerar_insights(df_kpis)

# ================= TOP CLIENTES =================
top_clientes = None
if not receitas.empty:
    top_clientes = receitas.groupby("Nome do cliente")["Valor"].sum().nlargest(5)

# ================= PDF CONSULTORIA =================
st.subheader("📄 Relatório Executivo (Consultoria)")

def gerar_pdf_consultoria(df_kpis, fig_resumo, insights, top_clientes):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    elementos = []

    # CAPA
    elementos.append(Spacer(1, 6*cm))
    elementos.append(Paragraph("RELATÓRIO EXECUTIVO", styles["Title"]))
    elementos.append(Paragraph("Análise Financeira", styles["Heading2"]))
    elementos.append(Spacer(1, 1*cm))
    elementos.append(Paragraph(datetime.now().strftime("%d/%m/%Y"), styles["Normal"]))
    elementos.append(PageBreak())

    # RESUMO
    elementos.append(Paragraph("Resumo Executivo", styles["Heading1"]))

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

    # INSIGHTS
    elementos.append(Paragraph("Principais Insights", styles["Heading2"]))
    for i in insights:
        elementos.append(Paragraph(f"• {i}", styles["Normal"]))

    elementos.append(Spacer(1, 1*cm))

    # TOP CLIENTES
    if top_clientes is not None:
        elementos.append(Paragraph("Top Clientes", styles["Heading2"]))
        tabela_top = [["Cliente", "Receita"]]
        for idx, val in top_clientes.items():
            tabela_top.append([idx, f"{val:,.2f}€"])
        elementos.append(Table(tabela_top))
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

if st.button("Gerar PDF Consultoria"):
    pdf = gerar_pdf_consultoria(df_kpis, fig_resumo, insights, top_clientes)

    st.download_button(
        label="📥 Download PDF Consultoria",
        data=pdf,
        file_name="relatorio_consultoria.pdf",
        mime="application/pdf"
    )
