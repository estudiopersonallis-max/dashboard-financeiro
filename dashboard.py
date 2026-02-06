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

dfs = []

for file in uploaded_files:
    xls = pd.ExcelFile(file)

    sheet_name = st.selectbox(
        f"Selecione a aba do ficheiro {file.name}",
        xls.sheet_names
    )

    df = pd.read_excel(file, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df.dropna(subset=["Data"])

    df["Valor_Correto"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)

    df["Dia"] = df["Data"].dt.day
    df["Mes"] = df["Data"].dt.strftime("%Y-%m")
    df["É Perda"] = df["Perdas"].notna()

    dfs.append(df)

df = pd.concat(dfs, ignore_index=True)

meses = sorted(df["Mes"].unique())
mes_sel = st.selectbox("📅 Selecione o mês", meses)
df_mes = df[df["Mes"] == mes_sel]

clientes_por_cliente = (
    df_mes.groupby("Nome do cliente")["É Perda"]
    .any()
)

clientes_ativos = (~clientes_por_cliente).sum()
perdas = clientes_por_cliente.sum()

total_valor = df_mes["Valor_Correto"].sum()
ticket_medio = df_mes["Valor_Correto"].mean()

st.subheader("📌 Indicadores")
c1, c2, c3, c4 = st.columns(4)
c1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
c2.metric("👥 Clientes Ativos", int(clientes_ativos))
c3.metric("❌ Perdas", int(perdas))
c4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()
st.header("📋 Tabelas")

tipos = ["A", "B", "C", "D"]

valor_tipo = (
    df_mes.groupby("Tipo")["Valor_Correto"]
    .sum()
    .reindex(tipos, fill_value=0)
)

ticket_tipo = (
    df_mes.groupby("Tipo")["Valor_Correto"]
    .mean()
    .reindex(tipos, fill_value=0)
)

st.subheader("Valor por Tipo")
st.dataframe(valor_tipo)

st.subheader("Ticket Médio por Tipo")
st.dataframe(ticket_tipo)

st.subheader("Valor por Professor")
st.dataframe(df_mes.groupby("Professor")["Valor_Correto"].sum())

st.subheader("Valor por Local")
st.dataframe(df_mes.groupby("Local")["Valor_Correto"].sum())

st.subheader("Valor por Modalidade")
st.dataframe(df_mes.groupby("Modalidade")["Valor_Correto"].sum())

st.subheader("Valor por Período do Mês")
periodos = pd.Series({
    "Dias 1–10": df_mes[df_mes["Dia"] <= 10]["Valor_Correto"].sum(),
    "Dias 11–20": df_mes[(df_mes["Dia"] > 10) & (df_mes["Dia"] <= 20)]["Valor_Correto"].sum(),
    "Dias 21–fim": df_mes[df_mes["Dia"] > 20]["Valor_Correto"].sum(),
})
st.dataframe(periodos)

st.divider()
st.header("📊 Gráficos")

st.subheader("Valor por Tipo")
st.bar_chart(valor_tipo)

st.subheader("Ticket Médio por Tipo")
st.bar_chart(ticket_tipo)

st.subheader("Valor por Professor")
st.bar_chart(df_mes.groupby("Professor")["Valor_Correto"].sum())

st.subheader("Valor por Local")
st.bar_chart(df_mes.groupby("Local")["Valor_Correto"].sum())

st.subheader("Valor por Modalidade")
st.bar_chart(df_mes.groupby("Modalidade")["Valor_Correto"].sum())

st.subheader("Valor por Período do Mês")
st.bar_chart(periodos)

st.divider()
st.header("📈 Comparação entre Meses")

comparativo = df.groupby("Mes")["Valor_Correto"].sum().sort_index()
st.dataframe(comparativo)
st.line_chart(comparativo)

