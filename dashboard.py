import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")

st.title("📊 Dashboard Financeiro")

uploaded_file = st.file_uploader("📤 Carregue o arquivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    df["Data"] = pd.to_datetime(df["Data"])
    df["Dia"] = df["Data"].dt.day

    # Perdas
    df["É Perda"] = df["Perdas"].notna()

    # Clientes ativos
    clientes_ativos = df[~df["É Perda"]]["Nome do cliente"].nunique()
    perdas = df["É Perda"].sum()

    # KPIs
    total_valor = df["Valor"].sum()
    ticket_medio = df["Valor"].mean()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
    col2.metric("👥 Clientes Ativos", clientes_ativos)
    col3.metric("❌ Perdas", perdas)
    col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📌 Valor por Modalidade")
        st.dataframe(df.groupby("Modalidade")["Valor"].sum())

        st.subheader("📌 Valor por Tipo")
        st.dataframe(df.groupby("Tipo")["Valor"].sum())

    with col2:
        st.subheader("📌 Valor por Professor")
        st.dataframe(df.groupby("Professor")["Valor"].sum())

        st.subheader("📌 Valor por Local")
        st.dataframe(df.groupby("Local")["Valor"].sum())

    st.divider()

    st.subheader("📅 Valor por Período do Mês")

    periodo_1 = df[df["Dia"] <= 10]["Valor"].sum()
    periodo_2 = df[(df["Dia"] > 10) & (df["Dia"] <= 20)]["Valor"].sum()
    periodo_3 = df[df["Dia"] > 20]["Valor"].sum()

    st.write(f"🟢 Dias 1–10: € {periodo_1:,.2f}")
    st.write(f"🟡 Dias 11–20: € {periodo_2:,.2f}")
    st.write(f"🔵 Dias 21–fim: € {periodo_3:,.2f}")

    st.divider()

    st.subheader("👥 Clientes")

    col1, col2 = st.columns(2)
    with col1:
        st.dataframe(df.groupby("Local")["Nome do cliente"].nunique().rename("Clientes por Local"))

    with col2:
        st.dataframe(df.groupby("Professor")["Nome do cliente"].nunique().rename("Clientes por Professor"))

    st.divider()

    st.subheader("🎟️ Ticket Médio por Tipo")
    st.dataframe(df.groupby("Tipo")["Valor"].mean())

else:
    st.info("⬆️ Carregue um arquivo Excel para iniciar o dashboard")
