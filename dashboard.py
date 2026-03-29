import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

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

# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    if df.empty:
        return

    df = df.loc[df.sum(axis=1).sort_values(ascending=False).index]

    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    ax.set_xlabel("€")
    st.pyplot(fig)

# ================= ANÁLISE =================
def bloco_analise(df, categoria, titulo):
    if df.empty or categoria not in df.columns:
        return

    pivot = df.pivot_table(index=categoria, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
    percent = pivot.div(pivot.sum(axis=0), axis=1) * 100

    st.markdown(f"### {titulo} por {categoria}")

    tabela = pivot.round(2).astype(str) + " € | " + percent.round(1).astype(str) + " %"
    st.dataframe(tabela, use_container_width=True)

    grafico_bar(pivot, f"{titulo} por {categoria}")

# ================= TABS =================
tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "💰 Receitas", "💸 Despesas"])

with tab1:
    st.subheader("Resumo Financeiro")
    if not df_kpis.empty:
        fig, ax = plt.subplots()
        df_kpis.set_index("Período")[["Receita", "Despesa", "Lucro"]].plot(kind="bar", ax=ax)
        st.pyplot(fig)

with tab2:
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        bloco_analise(receitas, cat, "Receitas")

with tab3:
    for cat in ["Classe", "Local"]:
        bloco_analise(despesas, cat, "Despesas")

# ================= TOP CLIENTES =================
st.subheader("🏆 Top Clientes")
if not receitas.empty:
    top = receitas.groupby("Nome do cliente")["Valor"].sum().nlargest(10)
    st.bar_chart(top)

# ================= EXPORT CSV =================
st.subheader("📥 Exportar Dados")
if not df_kpis.empty:
    csv = df_kpis.to_csv(index=False).encode("utf-8")
    st.download_button("Download KPIs", csv, "kpis.csv", "text/csv")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
