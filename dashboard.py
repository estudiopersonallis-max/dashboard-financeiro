import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

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
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

# KPIs estratégicos
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

# ================= TENDÊNCIA =================
def classificar_tendencia(series):
    if len(series) < 2:
        return "Neutra"
    if series.iloc[-1] > series.iloc[0]:
        return "Alta 📈"
    elif series.iloc[-1] < series.iloc[0]:
        return "Queda 📉"
    return "Estável"

if not df_kpis.empty:
    st.info(f"Tendência Receita: {classificar_tendencia(df_kpis['Receita'])} | Tendência Lucro: {classificar_tendencia(df_kpis['Lucro'])}")

# ================= ALERTAS =================
alertas = []

if margem < 10:
    alertas.append("Margem baixa (<10%)")

if (df_kpis["Lucro"] < 0).any():
    alertas.append("Períodos com prejuízo")

if concentracao_top5 > 50:
    alertas.append("Alta dependência de poucos clientes")

if custo_ratio > 80:
    alertas.append("Estrutura de custos elevada")

if receita_total > 0 and lucro_total < 0:
    alertas.append("Crescimento sem lucro")

if alertas:
    st.warning("\n".join([f"⚠️ {a}" for a in alertas]))

# ================= GRÁFICOS =================
if not df_kpis.empty:
    fig, ax = plt.subplots()
    df_kpis.set_index("Período")[["Receita", "Despesa", "Lucro"]].plot(kind="bar", ax=ax)
    st.pyplot(fig)

# ================= DRIVERS =================
st.subheader("🧠 Drivers do Negócio")

if not receitas.empty:
    st.write("Top Professores")
    st.bar_chart(receitas.groupby("Professor")["Valor"].sum().sort_values(ascending=False).head(10))

if not receitas.empty:
    st.write("Top Modalidades")
    st.bar_chart(receitas.groupby("Modalidade")["Valor"].sum().sort_values(ascending=False))

# ================= SCORE =================
score = 0
if margem > 20: score += 1
if lucro_total > 0: score += 1
if custo_ratio < 70: score += 1
if concentracao_top5 < 50: score += 1

st.subheader("🏥 Saúde do Negócio")
st.metric("Score", f"{score}/4")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
