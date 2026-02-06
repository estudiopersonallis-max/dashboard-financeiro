import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro")

uploaded_files = st.file_uploader(
    "📤 Carregue um ficheiro Excel por mês",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("⬆️ Carregue pelo menos um ficheiro Excel")
    st.stop()

# ================= LEITURA DOS FICHEIROS =================
dfs = []

for file in uploaded_files:
    df_temp = pd.read_excel(file)

    # Normalizar colunas
    df_temp.columns = df_temp.columns.str.strip()

    # Datas (mantido como no código original)
    df_temp["Data"] = pd.to_datetime(df_temp["Data"], errors="coerce")
    df_temp = df_temp.dropna(subset=["Data"])
    df_temp["Dia"] = df_temp["Data"].dt.day

    # Coluna mês
    df_temp["Mes"] = df_temp["Data"].dt.strftime("%Y-%m")

    # Perdas
    df_temp["É Perda"] = df_temp["Perdas"].notna()

    dfs.append(df_temp)

df = pd.concat(dfs, ignore_index=True)

# ================= SELETOR DE MÊS =================
mes_selecionado = st.selectbox(
    "📅 Selecione o mês",
    sorted(df["Mes"].unique())
)

df = df[df["Mes"] == mes_selecionado]

# ================= KPIs =================
total_valor = df["Valor"].sum()
ticket_medio = df["Valor"].mean()
perdas = df["É Perda"].sum()
clientes_ativos = df[~df["É Perda"]]["Nome do cliente"].nunique()

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()

# ================= TABELAS =================
col1, col2 = st.columns(2)

with col1:
    st.subheader("Valor por Modalidade")
    valor_modalidade = df.groupby("Modalidade")["Valor"].sum()
    st.dataframe(valor_modalidade)

    st.subheader("Valor por Tipo")
    valor_tipo = df.groupby("Tipo")["Valor"].sum()
    st.dataframe(valor_tipo)

with col2:
    st.subheader("Valor por Professor")
    valor_professor = df.groupby("Professor")["Valor"].sum()
    st.dataframe(valor_professor)

    st.subheader("Valor por Local")
    valor_local = df.groupby("Local")["Valor"].sum()
    st.dataframe(valor_local)

st.divider()

st.subheader("Valor por Período do Mês")

p1 = df[df["Dia"] <= 10]["Valor"].sum()
p2 = df[(df["Dia"] > 10) & (df["Dia"] <= 20)]["Valor"].sum()
p3 = df[df["Dia"] > 20]["Valor"].sum()

valor_periodo = pd.Series(
    {
        "Dias 1–10": p1,
        "Dias 11–20": p2,
        "Dias 21–fim": p3,
    }
)

st.dataframe(valor_periodo)

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("Clientes por Local")
    clientes_local = df.groupby("Local")["Nome do cliente"].nunique()
    st.dataframe(clientes_local)

with col2:
    st.subheader("Clientes por Professor")
    clientes_professor = df.groupby("Professor")["Nome do cliente"].nunique()
    st.dataframe(clientes_professor)

st.divider()

st.subheader("Ticket Médio por Tipo")
ticket_tipo = df.groupby("Tipo")["Valor"].mean()
st.dataframe(ticket_tipo)

# ================= GRÁFICOS =================
st.divider()
st.header("📊 Gráficos")

st.subheader("Valor por Modalidade")
st.bar_chart(valor_modalidade)

st.subheader("Valor por Tipo")
st.bar_chart(valor_tipo)

st.subheader("Valor por Professor")
st.bar_chart(valor_professor)

st.subheader("Valor por Local")
st.bar_chart(valor_local)

st.subheader("Valor por Período do Mês")
st.bar_chart(valor_periodo)

st.subheader("Clientes por Local")
st.bar_chart(clientes_local)

st.subheader("Clientes por Professor")
st.bar_chart(clientes_professor)

st.subheader("Ticket Médio por Tipo")
st.bar_chart(ticket_tipo)
