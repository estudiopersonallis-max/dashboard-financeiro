import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import matplotlib

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from datetime import datetime

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= UPLOAD =================
st.subheader("📤 Upload de Ficheiros (cada ficheiro = um período)")
uploaded_receitas = st.file_uploader("Receitas (Excel)", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.file_uploader("Despesas (Excel)", type=["xlsx"], accept_multiple_files=True)

# ================= FUNÇÕES =================
def nome_periodo(nome):
    return nome.replace(".xlsx", "").upper()

def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Nome do cliente"] = df["Nome do cliente"].astype(str).str.upper().str.strip()
        df["Modalidade"] = df.get("Modalidade", "N/A")
        df["Tipo"] = df.get("Tipo", "N/A")
        df["Professor"] = df.get("Professor", "N/A")
        df["Local"] = df.get("Local", "N/A")

        coluna_status = df.columns[2]
        df["Ativo"] = df[coluna_status].astype(str).str.upper().eq("ATIVO")
        df["É Perda"] = df["Perdas"].notna() if "Perdas" in df.columns else False

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        df = df.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Classe"] = df["Classe"].astype(str).str.upper().str.strip()
        df["Local"] = df["Local"].astype(str).str.strip()

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= FILTRO DEPÓSITOS =================
if not despesas.empty:
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= FILTROS (POWER BI STYLE) =================
st.sidebar.header("🔎 Filtros")

periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))

periodo_sel = st.sidebar.multiselect(
    "Período",
    options=periodos,
    default=periodos
)

prof_sel = st.sidebar.multiselect(
    "Professor",
    options=receitas["Professor"].dropna().unique() if not receitas.empty else [],
    default=receitas["Professor"].dropna().unique() if not receitas.empty else []
)

local_sel = st.sidebar.multiselect(
    "Local",
    options=receitas["Local"].dropna().unique() if not receitas.empty else [],
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

periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))
kpis = []

for p in periodos:
    r = receitas[receitas["Periodo"] == p] if not receitas.empty else pd.DataFrame()
    d = despesas[despesas["Periodo"] == p] if not despesas.empty else pd.DataFrame()

    receita = r["Valor"].sum() if not r.empty else 0
    despesa = d["Valor"].sum() if not d.empty else 0
    lucro = receita + despesa

    kpis.append({
        "Período": p,
        "Receita (€)": round(receita, 2),
        "Despesa (€)": round(despesa, 2),
        "Lucro (€)": round(lucro, 2)
    })

df_kpis = pd.DataFrame(kpis)

col1, col2, col3 = st.columns(3)

total_receita = df_kpis["Receita (€)"].sum()
total_despesa = df_kpis["Despesa (€)"].sum()
total_lucro = df_kpis["Lucro (€)"].sum()

col1.metric("💰 Receita Total", f"{total_receita:,.2f} €")
col2.metric("💸 Despesa Total", f"{total_despesa:,.2f} €")
col3.metric("📈 Lucro Total", f"{total_lucro:,.2f} €")

st.dataframe(df_kpis, use_container_width=True)

# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    if df.empty:
        return None
    fig, ax = plt.subplots()
    df.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    ax.set_ylabel("€")
    return fig

def bloco_analise(df, categoria, titulo):
    if df.empty or categoria not in df.columns:
        return

    pivot = df.pivot_table(index=categoria, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
    percent = pivot.div(pivot.sum(axis=0), axis=1) * 100

    st.markdown(f"### {titulo} por {categoria}")
    st.dataframe(pivot.round(2))

    fig = grafico_bar(pivot, f"{titulo} por {categoria}")
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
def grafico_resumo(df_kpis):
    if df_kpis.empty:
        return None

    fig, ax = plt.subplots()
    df_kpis.set_index("Período")[["Receita (€)", "Despesa (€)", "Lucro (€)"]].plot(kind="bar", ax=ax)
    ax.set_title("Resumo Financeiro")
    return fig

fig_resumo = grafico_resumo(df_kpis)

if fig_resumo:
    st.pyplot(fig_resumo)

st.divider()

# ================= PDF =================
st.subheader("📄 Exportar PDF Executivo")

def gerar_pdf_executivo(df_kpis, receitas, despesas, fig_resumo):

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    elementos = []

    elementos.append(Paragraph("RELATÓRIO FINANCEIRO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    tabela_data = [["Período", "Receita (€)", "Despesa (€)", "Lucro (€)"]]
    for _, row in df_kpis.iterrows():
        tabela_data.append(list(row))

    elementos.append(Table(tabela_data))
    elementos.append(PageBreak())

    if fig_resumo:
        img = BytesIO()
        fig_resumo.savefig(img, format="png")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=9*cm))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

if st.button("📄 Gerar PDF Executivo"):
    pdf = gerar_pdf_executivo(df_kpis, receitas, despesas, fig_resumo)

    st.download_button(
        label="📥 Download PDF",
        data=pdf,
        file_name="relatorio.pdf",
        mime="application/pdf"
    )
