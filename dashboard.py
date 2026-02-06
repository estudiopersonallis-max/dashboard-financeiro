import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro")

uploaded_file = st.file_uploader("📤 Carregue o arquivo Excel", type=["xlsx"])

if uploaded_file is None:
    st.info("⬆️ Carregue um arquivo Excel para iniciar o dashboard")
else:
    try:
        df = pd.read_excel(uploaded_file)

        # Normalizar nomes das colunas
        df.columns = df.columns.str.strip()

        # Converter data com segurança
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        df = df.dropna(subset=["Data"])
        df["Dia"] = df["Data"].dt.day

        # Perdas
        df["É Perda"] = df["Perdas"].notna()

        # KPIs
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

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Valor por Modalidade")
            st.dataframe(df.groupby("Modalidade")["Valor"].sum())

            st.subheader("Valor por Tipo")
            st.dataframe(df.groupby("Tipo")["Valor"].sum())

        with col2:
            st.subheader("Valor por Professor")
            st.dataframe(df.groupby("Professor")["Valor"].sum())

            st.subheader("Valor por Local")
            st.dataframe(df.groupby("Local")["Valor"].sum())

        st.divider()

        st.subheader("Valor por Período do Mês")

        p1 = df[df["Dia"] <= 10]["Valor"].sum()
        p2 = df[(df["Dia"] > 10) & (df["Dia"] <= 20)]["Valor"].sum()
        p3 = df[df["Dia"] > 20]["Valor"].sum()

        st.write(f"🟢 Dias 1–10: € {p1:,.2f}")
        st.write(f"🟡 Dias 11–20: € {p2:,.2f}")
        st.write(f"🔵 Dias 21–fim: € {p3:,.2f}")

        st.divider()

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Clientes por Local")
            st.dataframe(df.groupby("Local")["Nome do cliente"].nunique())

        with col2:
            st.subheader("Clientes por Professor")
            st.dataframe(df.groupby("Professor")["Nome do cliente"].nunique())

        st.divider()

        st.subheader("Ticket Médio por Tipo")
        st.dataframe(df.groupby("Tipo")["Valor"].mean())

    except Exception as e:
        st.error("❌ Erro ao processar o arquivo")
        st.exception(e)
