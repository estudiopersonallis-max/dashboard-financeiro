import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import matplotlib
import re

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, PageBreak, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= FUNÇÃO LIMPEZA € =================
def limpar_valor(x):
    if isinstance(x, str):
        x = re.sub(r"[^\d,.-]", "", x)
        x = x.replace(",", ".")
        try:
            return float(x)
        except:
            return 0.0
    return float(x) if pd.notnull(x) else 0.0

# ================= UPLOAD =================
uploaded_receitas = st.file_uploader("Receitas (Excel)", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.file_uploader("Despesas (Excel)", type=["xlsx"], accept_multiple_files=True)

# ================= FUNÇÕES =================
def nome_periodo(nome):
    return nome.replace(".xlsx", "").strip().upper()

def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = df["Valor"].apply(limpar_valor)

        df["Professor"] = df.get("Professor", "N/A")
        df["Local"] = df.get("Local", "N/A")
        df["Modalidade"] = df.get("Modalidade", "N/A")
        df["Tipo"] = df.get("Tipo", "N/A")

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

        # 🔥 FORÇAR DESPESAS NEGATIVAS
        df["Valor"] = -abs(df["Valor"])

        df["Classe"] = df.get("Classe", "N/A")
        df["Local"] = df.get("Local", "N/A")

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= FILTROS =================
st.sidebar.header("🔎 Filtros")

periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))

periodo_sel = st.sidebar.multiselect("Período", periodos, default=periodos)

prof_sel = st.sidebar.multiselect(
    "Professor",
    receitas["Professor"].dropna().unique() if not receitas.empty else [],
    default=receitas["Professor"].dropna().unique() if not receitas.empty else []
)

local_sel = st.sidebar.multiselect(
    "Local",
    receitas["Local"].dropna().unique() if not receitas.empty else [],
    default=receitas["Local"].dropna().unique() if not receitas.empty else []
)

# ================= APLICAR FILTROS =================
if not receitas.empty:
    receitas = receitas[
        receitas["Periodo"].isin(periodo_sel) &
        receitas["Professor"].isin(prof_sel) &
        receitas["Local"].isin(local_sel)
    ]

if not despesas.empty:
    despesas = despesas[
        despesas["Periodo"].isin(periodo_sel)
    ]

# ================= KPIs =================
st.subheader("📊 Visão Geral")

kpis = []

for p in periodos:
    r = receitas[receitas["Periodo"] == p] if not receitas.empty else pd.DataFrame()
    d = despesas[despesas["Periodo"] == p] if not despesas.empty else pd.DataFrame()

    receita = r["Valor"].sum() if not r.empty else 0
    despesa = d["Valor"].sum() if not d.empty else 0
    lucro = receita + despesa

    kpis.append({
        "Período": p,
        "Receita (€)": receita,
        "Despesa (€)": despesa,
        "Lucro (€)": lucro
    })

df_kpis = pd.DataFrame(kpis, columns=["Período", "Receita (€)", "Despesa (€)", "Lucro (€)"])

# KPIs cards
col1, col2, col3 = st.columns(3)

col1.metric("💰 Receita Total", f"{df_kpis['Receita (€)'].sum():,.2f} €")
col2.metric("💸 Despesa Total", f"{df_kpis['Despesa (€)'].sum():,.2f} €")
col3.metric("📈 Lucro Total", f"{df_kpis['Lucro (€)'].sum():,.2f} €")

st.dataframe(df_kpis, use_container_width=True)

# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    if df.empty:
        return None
    fig, ax = plt.subplots()
    df.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    return fig

def bloco_analise(df, categoria, titulo):
    if df.empty or categoria not in df.columns:
        return

    pivot = df.pivot_table(
        index=categoria,
        columns="Periodo",
        values="Valor",
        aggfunc="sum",
        fill_value=0
    )

    st.markdown(f"### {titulo} por {categoria}")
    st.dataframe(pivot.round(2))

    fig = grafico_bar(pivot, titulo)
    if fig:
        st.pyplot(fig)

# ================= RECEITAS =================
st.subheader("📌 Receitas")

col1, col2 = st.columns(2)

with col1:
    bloco_analise(receitas, "Modalidade", "Receitas")
    bloco_analise(receitas, "Tipo", "Receitas")

with col2:
    bloco_analise(receitas, "Professor", "Receitas")
    bloco_analise(receitas, "Local", "Receitas")

# ================= DESPESAS =================
st.subheader("📌 Despesas")

col1, col2 = st.columns(2)

with col1:
    bloco_analise(despesas, "Classe", "Despesas")

with col2:
    bloco_analise(despesas, "Local", "Despesas")

# ================= RESUMO =================
def grafico_resumo(df):
    if df.empty:
        return None
    fig, ax = plt.subplots()
    df.set_index("Período")[["Receita (€)", "Despesa (€)", "Lucro (€)"]].plot(kind="bar", ax=ax)
    ax.set_title("Resumo Financeiro")
    return fig

fig_resumo = grafico_resumo(df_kpis)

if fig_resumo:
    st.pyplot(fig_resumo)

# ================= DEBUG (opcional) =================
# st.write("DEBUG Despesas:", despesas.groupby("Periodo")["Valor"].sum())

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
    pdf = gerar_pdf(df_kpis, fig_resumo)

    st.download_button("📥 Download PDF", pdf, "relatorio.pdf")
