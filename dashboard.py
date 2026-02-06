import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro")

uploaded_files = st.file_uploader(
    "📤 Carregue um ou mais arquivos Excel (1 por mês)",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("⬆️ Carregue pelo menos um arquivo Excel")
    st.stop()

# ---------- LEITURA DOS ARQUIVOS ----------
dfs = []

for file in uploaded_files:
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df.dropna(subset=["Data"])

    df["Mes"] = df["Data"].dt.strftime("%Y-%m")
    df["Dia"] = df["Data"].dt.day
    df["É Perda"] = df["Perdas"].notna()

    dfs.append(df)

df = pd.concat(dfs, ignore_index=True)

# ---------- FILTRO DE MÊS ----------
meses = sorted(df["Mes"].unique())
mes_selecionado = st.selectbox("📅 Selecione o mês", meses)

df_mes = df[df["Mes"] == mes_selecionado]

# ---------- KPIs ----------
total_valor = df_mes["Valor"].sum()
ticket_medio = df_mes["Valor"].mean()
perdas = df_mes["É Perda"].sum()
clientes_ativos = df_mes[~df_mes["É Perda"]]["Nome do cliente"].nunique()

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()

# ---------- TIPO (A–D FORÇADO) ----------
tipos = ["A", "B", "C", "D"]

valor_tipo = (
    df_mes.groupby("Tipo")["Valor"]
    .sum()
    .reindex(tipos, fill_value=0)
)

ticket_tipo = (
    df_mes.groupby("Tipo")["Valor"]
    .mean()
    .reindex(tipos, fill_value=0)
)

col1, col2 = st.columns(2)

with col1:
    st.subheader("💰 Valor por Tipo")
    st.dataframe(valor_tipo)
    st.bar_chart(valor_tipo)

with col2:
    st.subheader("🎟️ Ticket Médio por Tipo")
    st.dataframe(ticket_tipo)
    st.bar_chart(ticket_tipo)

st.divider()

# ---------- OUTRAS DIMENSÕES ----------
def bloco(titulo, grupo):
    st.subheader(titulo)
    tabela = df_mes.groupby(grupo)["Valor"].sum()
    st.dataframe(tabela)
    st.bar_chart(tabela)

col1, col2 = st.columns(2)

with col1:
    bloco("Valor por Professor", "Professor")
    bloco("Valor por Modalidade", "Modalidade")

with col2:
    bloco("Valor por Local", "Local")

st.divider()

# ---------- PERÍODO DO MÊS ----------
periodos = pd.Series({
    "Dias 1–10": df_mes[df_mes["Dia"] <= 10]["Valor"].sum(),
    "Dias 11–20": df_mes[(df_mes["Dia"] > 10) & (df_mes["Dia"] <= 20)]["Valor"].sum(),
    "Dias 21–fim": df_mes[df_mes["Dia"] > 20]["Valor"].sum(),
})

st.subheader("📅 Valor por Período do Mês")
st.dataframe(periodos)
st.bar_chart(periodos)

st.divider()

# ---------- COMPARAÇÃO ENTRE MESES ----------
st.subheader("📈 Comparação entre Meses")

comparativo = (
    df.groupby("Mes")["Valor"]
    .sum()
    .sort_index()
)

st.dataframe(comparativo)
st.line_chart(comparativo)
