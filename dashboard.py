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

# ================= NORMALIZAR COLUNAS =================
def normalizar_colunas(df):
    df.columns = df.columns.str.strip().str.upper()
    return df

# ================= LIMPAR VALORES =================
def limpar_valor(x):
    try:
        if pd.isna(x):
            return 0.0
        x = str(x)
        x = re.sub(r"[^\d,.\-]", "", x)

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
uploaded_receitas = st.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

# ================= LEITURA =================
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        df = normalizar_colunas(df)

        if "VALOR" not in df.columns:
            continue

        df["PERIODO"] = nome_periodo(f.name)
        df["VALOR"] = df["VALOR"].apply(limpar_valor)

        df = df.drop_duplicates()

        dfs.append(df)

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame(columns=["PERIODO", "VALOR"])

def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        df = normalizar_colunas(df)

        if "VALOR" not in df.columns:
            continue

        df["PERIODO"] = nome_periodo(f.name)
        df["VALOR"] = df["VALOR"].apply(limpar_valor)

        # 🔥 GARANTE NEGATIVO CORRETO
        df["VALOR"] = -df["VALOR"].abs()

        df = df.drop_duplicates()

        dfs.append(df)

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame(columns=["PERIODO", "VALOR"])

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame(columns=["PERIODO","VALOR"])
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame(columns=["PERIODO","VALOR"])

# ================= FILTROS =================
st.sidebar.header("Filtros")

periodos = sorted(set(receitas["PERIODO"]).union(set(despesas["PERIODO"])))

periodo_sel = st.sidebar.multiselect("Período", periodos, default=periodos)

receitas = receitas[receitas["PERIODO"].isin(periodo_sel)]
despesas = despesas[despesas["PERIODO"].isin(periodo_sel)]

# ================= KPIs =================
kpis = []

for p in periodos:
    r = receitas[receitas["PERIODO"] == p]
    d = despesas[despesas["PERIODO"] == p]

    receita = r["VALOR"].sum()
    despesa = d["VALOR"].sum()
    lucro = receita + despesa

    kpis.append({
        "Período": p,
        "Receita (€)": receita,
        "Despesa (€)": despesa,
        "Lucro (€)": lucro
    })

df_kpis = pd.DataFrame(kpis)

# KPIs
col1, col2, col3 = st.columns(3)

col1.metric("Receita", f"{df_kpis['Receita (€)'].sum():,.2f} €")
col2.metric("Despesa", f"{df_kpis['Despesa (€)'].sum():,.2f} €")
col3.metric("Lucro", f"{df_kpis['Lucro (€)'].sum():,.2f} €")

st.dataframe(df_kpis)

# ================= DEBUG REAL =================
st.subheader("DEBUG")

st.write("Receitas somadas:")
st.write(receitas.groupby("PERIODO")["VALOR"].sum())

st.write("Despesas somadas:")
st.write(despesas.groupby("PERIODO")["VALOR"].sum())

# ================= GRÁFICO =================
fig, ax = plt.subplots()
df_kpis.set_index("Período").plot(kind="bar", ax=ax)
st.pyplot(fig)

# ================= PDF =================
def gerar_pdf(df):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    elementos = []
    elementos.append(Paragraph("Relatório Financeiro", styles["Title"]))

    tabela_data = [df.columns.tolist()] + df.values.tolist()
    elementos.append(Table(tabela_data))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

if st.button("Gerar PDF"):
    pdf = gerar_pdf(df_kpis)
    st.download_button("Download PDF", pdf, "relatorio.pdf")
