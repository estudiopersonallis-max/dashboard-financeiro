import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import matplotlib
import re

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= LIMPEZA FORTE DE VALORES =================
def limpar_valor(x):
    try:
        if pd.isna(x):
            return 0.0

        x = str(x)

        # remove tudo exceto números, vírgula, ponto e sinal
        x = re.sub(r"[^\d,.\-]", "", x)

        # padrão europeu → converter
        if "," in x and "." in x:
            x = x.replace(".", "").replace(",", ".")
        else:
            x = x.replace(",", ".")

        return float(x)

    except:
        return 0.0

# ================= PERÍODO =================
def nome_periodo(nome):
    nome = nome.replace(".xlsx", "").strip().upper()
    nome = nome.replace(".R", "").replace(".D", "")
    return nome

# ================= UPLOAD =================
uploaded_receitas = st.file_uploader("Receitas (Excel)", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.file_uploader("Despesas (Excel)", type=["xlsx"], accept_multiple_files=True)

# ================= LEITURA =================
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = df["Valor"].apply(limpar_valor)

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = df["Valor"].apply(limpar_valor)

        # 🔥 CORREÇÃO DEFINITIVA
        df["Valor"] = df["Valor"].abs() * -1

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= FILTROS =================
st.sidebar.header("🔎 Filtros")

periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))

periodo_sel = st.sidebar.multiselect("Período", periodos, default=periodos)

# aplicar filtros
if not receitas.empty:
    receitas = receitas[receitas["Periodo"].isin(periodo_sel)]

if not despesas.empty:
    despesas = despesas[despesas["Periodo"].isin(periodo_sel)]

# ================= KPIs =================
st.subheader("📊 Visão Geral")

kpis = []

for p in periodos:
    r = receitas[receitas["Periodo"] == p]
    d = despesas[despesas["Periodo"] == p]

    receita = r["Valor"].sum() if not r.empty else 0
    despesa = d["Valor"].sum() if not d.empty else 0
    lucro = receita + despesa

    kpis.append({
        "Período": p,
        "Receita (€)": receita,
        "Despesa (€)": despesa,
        "Lucro (€)": lucro
    })

df_kpis = pd.DataFrame(
    kpis,
    columns=["Período", "Receita (€)", "Despesa (€)", "Lucro (€)"]
)

# KPIs
col1, col2, col3 = st.columns(3)

col1.metric("💰 Receita Total", f"{df_kpis['Receita (€)'].sum():,.2f} €")
col2.metric("💸 Despesa Total", f"{df_kpis['Despesa (€)'].sum():,.2f} €")
col3.metric("📈 Lucro Total", f"{df_kpis['Lucro (€)'].sum():,.2f} €")

st.dataframe(df_kpis, use_container_width=True)

# ================= DEBUG (REMOVE DEPOIS) =================
st.write("🔍 Debug Despesas por período:")
st.write(despesas.groupby("Periodo")["Valor"].sum())

# ================= GRÁFICO =================
def grafico_resumo(df):
    if df.empty:
        return None
    fig, ax = plt.subplots()
    df.set_index("Período").plot(kind="bar", ax=ax)
    return fig

fig = grafico_resumo(df_kpis)

if fig:
    st.pyplot(fig)

# ================= PDF =================
def gerar_pdf(df_kpis, fig):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    elementos = []
    elementos.append(Paragraph("Relatório Financeiro", styles["Title"]))

    tabela_data = [df_kpis.columns.tolist()] + df_kpis.values.tolist()
    elementos.append(Table(tabela_data))

    if fig:
        img = BytesIO()
        fig.savefig(img, format="png")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=9*cm))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

if st.button("📄 Gerar PDF"):
    pdf = gerar_pdf(df_kpis, fig)
    st.download_button("📥 Download PDF", pdf, "relatorio.pdf")
