import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro")

# ================= UPLOAD =================
uploaded_files = st.file_uploader(
    "📤 Carregue um ficheiro Excel por mês",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("⬆️ Carregue pelo menos um ficheiro Excel para iniciar o dashboard")
    st.stop()

# ================= LEITURA =================
dfs = []

for file in uploaded_files:
    df_temp = pd.read_excel(file)

    mes_ficheiro = file.name.replace(".xlsx", "")
    df_temp["Mes"] = mes_ficheiro

    df_temp["Data"] = pd.to_datetime(df_temp["Data"])
    df_temp["Dia"] = df_temp["Data"].dt.day
    df_temp["Ano"] = df_temp["Data"].dt.year
    df_temp["Trimestre"] = df_temp["Data"].dt.to_period("Q").astype(str)

    df_temp["Nome do cliente"] = (
        df_temp["Nome do cliente"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    coluna_status = df_temp.columns[2]
    df_temp["Ativo"] = (
        df_temp[coluna_status]
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("ATIVO")
    )

    df_temp["É Perda"] = df_temp["Perdas"].notna()

    dfs.append(df_temp)

df = pd.concat(dfs, ignore_index=True)

# ================= FILTRO =================
tipo_periodo = st.selectbox(
    "📅 Tipo de análise",
    ["Mês (ficheiro)", "Trimestre", "Ano"]
)

if tipo_periodo == "Mês (ficheiro)":
    periodo = st.selectbox("Selecione o mês", sorted(df["Mes"].unique()))
    df_filtro = df[df["Mes"] == periodo]
elif tipo_periodo == "Trimestre":
    periodo = st.selectbox("Selecione o trimestre", sorted(df["Trimestre"].unique()))
    df_filtro = df[df["Trimestre"] == periodo]
else:
    periodo = st.selectbox("Selecione o ano", sorted(df["Ano"].unique()))
    df_filtro = df[df["Ano"] == periodo]

st.caption(f"📌 Período selecionado: **{periodo}**")

# ================= KPIs =================
clientes_ativos = df_filtro.loc[df_filtro["Ativo"], "Nome do cliente"].nunique()
perdas = int(df_filtro["É Perda"].sum())
total_valor = df_filtro["Valor"].sum()
ticket_medio = total_valor / clientes_ativos if clientes_ativos > 0 else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()

# ================= MODALIDADE =================
st.subheader("📌 Valor por Modalidade")
valor_modalidade = df_filtro.groupby("Modalidade")["Valor"].sum()
st.dataframe(valor_modalidade)
st.bar_chart(valor_modalidade)

st.subheader("📊 % do Valor por Modalidade")
st.bar_chart((valor_modalidade / total_valor) * 100)

# ================= TIPO =================
st.subheader("📌 Valor por Tipo")
valor_tipo = df_filtro.groupby("Tipo")["Valor"].sum()
st.dataframe(valor_tipo)
st.bar_chart(valor_tipo)

st.subheader("📊 % do Valor por Tipo")
st.bar_chart((valor_tipo / total_valor) * 100)

# ================= PROFESSOR =================
st.subheader("📌 Valor por Professor")
valor_professor = df_filtro.groupby("Professor")["Valor"].sum()
st.dataframe(valor_professor)
st.bar_chart(valor_professor)

st.subheader("📊 % do Valor por Professor")
st.bar_chart((valor_professor / total_valor) * 100)

# ================= LOCAL =================
st.subheader("📌 Valor por Local")
valor_local = df_filtro.groupby("Local")["Valor"].sum()
st.dataframe(valor_local)
st.bar_chart(valor_local)

st.subheader("📊 % do Valor por Local")
st.bar_chart((valor_local / total_valor) * 100)

st.divider()

# ================= PERÍODO DO MÊS =================
st.subheader("📅 Valor por Período do Mês")

valor_periodo = pd.Series({
    "Dias 1–10": df_filtro[df_filtro["Dia"] <= 10]["Valor"].sum(),
    "Dias 11–20": df_filtro[(df_filtro["Dia"] > 10) & (df_filtro["Dia"] <= 20)]["Valor"].sum(),
    "Dias 21–fim": df_filtro[df_filtro["Dia"] > 20]["Valor"].sum()
})

st.dataframe(valor_periodo)
st.bar_chart(valor_periodo)

st.subheader("📊 % do Valor por Período do Mês")
st.bar_chart((valor_periodo / total_valor) * 100)

st.divider()

# ================= CLIENTES =================
st.subheader("👥 Clientes por Local")
clientes_local = df_filtro[df_filtro["Ativo"]].groupby("Local")["Nome do cliente"].nunique()
st.dataframe(clientes_local)
st.bar_chart(clientes_local)

st.subheader("📊 % Clientes por Local")
st.bar_chart((clientes_local / clientes_ativos) * 100)

st.subheader("👥 Clientes por Professor")
clientes_professor = df_filtro[df_filtro["Ativo"]].groupby("Professor")["Nome do cliente"].nunique()
st.dataframe(clientes_professor)
st.bar_chart(clientes_professor)

st.subheader("📊 % Clientes por Professor")
st.bar_chart((clientes_professor / clientes_ativos) * 100)

st.divider()

# ================= TICKET =================
st.subheader("🎟️ Ticket Médio por Tipo")
ticket_tipo = (
    df_filtro.groupby("Tipo")["Valor"].sum()
    / df_filtro[df_filtro["Ativo"]].groupby("T_]()_
