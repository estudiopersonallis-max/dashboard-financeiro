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
lucro_total = receita_total + despesa_total  # despesas já negativas
margem = (lucro_total / receita_total * 100) if receita_total else 0

# estratégicos
ticket_medio = receitas["Valor"].mean() if not receitas.empty else 0
custo_ratio = abs(despesa_total) / receita_total * 100 if receita_total else 0

concentracao_top5 = 0
if not receitas.empty:
    top = receitas.groupby("Nome do cliente")["Valor"].sum()
    concentracao_top5 = top.nlargest(5).sum() / top.sum() * 100 if top.sum() else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Receita", f"{receita_total:,.0f}€")
col2.metric("Despesa", f"{despesa_total:,.0f}€")
col3.metric("Lucro", f"{lucro_total:,.0f}€")
col4.metric("Margem", f"{margem:.1f}%")
col5.metric("Ticket Médio", f"{ticket_medio:,.0f}€")

col6, col7 = st.columns(2)
col6.metric("Custo/Receita", f"{custo_ratio:.1f}%")
col7.metric("Concentração Top 5", f"{concentracao_top5:.1f}%")

# ================= KPIs POR PERÍODO =================
kpis = []
for p in periodos:
    r = receitas[receitas["Periodo"] == p]
    d = despesas[despesas["Periodo"] == p]
    receita = r["Valor"].sum()
    despesa = d["Valor"].sum()
    lucro = receita + despesa
    kpis.append({"Período": p, "Receita": receita, "Despesa": despesa, "Lucro": lucro})

df_kpis = pd.DataFrame(kpis)

if not df_kpis.empty:
    df_kpis["Margem (%)"] = (df_kpis["Lucro"] / df_kpis["Receita"]) * 100
    df_kpis["Δ Receita (%)"] = df_kpis["Receita"].pct_change() * 100
    df_kpis["Δ Lucro (%)"] = df_kpis["Lucro"].pct_change() * 100
    st.dataframe(df_kpis.round(2), use_container_width=True)

# ================= FUNÇÕES =================
def grafico_bar(df, titulo):
    if df.empty:
        return None
    df = df.loc[df.sum(axis=1).sort_values(ascending=False).index]
    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    ax.set_xlabel("€")
    return fig


def bloco_analise(df, categoria, titulo):
    if df.empty or categoria not in df.columns:
        return None
    pivot = df.pivot_table(index=categoria, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
    percent = pivot.div(pivot.sum(axis=0), axis=1) * 100

    st.markdown(f"### {titulo} por {categoria}")
    tabela = pivot.round(2).astype(str) + " € | " + percent.round(1).astype(str) + " %"
    st.dataframe(tabela, use_container_width=True)

    fig = grafico_bar(pivot, f"{titulo} por {categoria}")
    if fig:
        st.pyplot(fig)
    return fig

# ================= TABS =================
tab1, tab2, tab3, tab4 = st.tabs(["📊 Visão Geral", "💰 Receitas", "💸 Despesas", "🧠 Drivers"])

with tab1:
    st.subheader("Resumo Financeiro")
    fig_resumo = None
    if not df_kpis.empty:
        fig_resumo, ax = plt.subplots()
        df_kpis.set_index("Período")[["Receita", "Despesa", "Lucro"]].plot(kind="bar", ax=ax)
        st.pyplot(fig_resumo)

with tab2:
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        bloco_analise(receitas, cat, "Receitas")

with tab3:
    for cat in ["Classe", "Local"]:
        bloco_analise(despesas, cat, "Despesas")

with tab4:
    st.subheader("Drivers do Negócio")
    if not receitas.empty:
        st.write("Top Professores")
        st.bar_chart(receitas.groupby("Professor")["Valor"].sum().sort_values(ascending=False).head(10))
    if not receitas.empty:
        st.write("Top Modalidades")
        st.bar_chart(receitas.groupby("Modalidade")["Valor"].sum().sort_values(ascending=False))

# ================= ALERTAS =================
alertas = []
if margem < 10:
    alertas.append("Margem baixa (<10%)")
if not df_kpis.empty and (df_kpis["Lucro"] < 0).any():
    alertas.append("Períodos com prejuízo")
if concentracao_top5 > 50:
    alertas.append("Alta dependência de poucos clientes")
if custo_ratio > 80:
    alertas.append("Estrutura de custos elevada")
if receita_total > 0 and lucro_total < 0:
    alertas.append("Crescimento sem lucro")

if alertas:
    st.warning("\n".join([f"⚠️ {a}" for a in alertas]))

# ================= SCORE =================
score = 0
if margem > 20: score += 1
if lucro_total > 0: score += 1
if custo_ratio < 70: score += 1
if concentracao_top5 < 50: score += 1

st.subheader("🏥 Saúde do Negócio")
st.metric("Score", f"{score}/4")

# ================= PDF =================
def gerar_pdf(df_kpis, receitas, despesas, fig_resumo):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("RELATÓRIO FINANCEIRO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))

    if not df_kpis.empty:
        tabela = [["Período", "Receita", "Despesa", "Lucro", "Margem"]]
        for _, row in df_kpis.iterrows():
            tabela.append([
                row["Período"],
                f"{row['Receita']:,.0f}€",
                f"{row['Despesa']:,.0f}€",
                f"{row['Lucro']:,.0f}€",
                f"{row['Margem (%)']:.1f}%"
            ])
        elementos.append(Table(tabela))

    if fig_resumo:
        img = BytesIO()
        fig_resumo.savefig(img, format="png", bbox_inches="tight")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=9*cm))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

st.subheader("📄 Exportar PDF")
if st.button("Gerar Relatório PDF"):
    pdf = gerar_pdf(df_kpis, receitas, despesas, fig_resumo)
    st.download_button(
        label="📥 Download PDF",
        data=pdf,
        file_name="relatorio_financeiro.pdf",
        mime="application/pdf"
    )

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
