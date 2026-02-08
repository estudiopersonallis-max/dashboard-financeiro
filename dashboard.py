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
    try:
        df_temp = pd.read_excel(file)

        mes_ficheiro = file.name.replace(".xlsx", "")
        df_temp["Mes"] = mes_ficheiro

        df_temp["Data"] = pd.to_datetime(df_temp["Data"], errors="coerce")
        df_temp = df_temp.dropna(subset=["Data"])

        df_temp["Dia"] = df_temp["Data"].dt.day
        df_temp["Ano"] = df_temp["Data"].dt.year
        df_temp["Trimestre"] = df_temp["Data"].dt.to_period("Q").astype(str)

        df_temp["Nome do cliente"] = (
            df_temp["Nome do cliente"].astype(str).str.strip().str.upper()
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

    except Exception as e:
        st.error(f"Erro ao ler o ficheiro {file.name}")
        st.exception(e)

# 🔒 PROTEÇÃO CRÍTICA
if not dfs:
    st.error("❌ Nenhum ficheiro Excel válido foi processado.")
    st.stop()

# ================= CONCATENAÇÃO SEGURA =================
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
total_valor = df_filtro["Valor"].sum()
perdas = int(df_filtro["É Perda"].sum())
ticket_medio = total_valor / clientes_ativos if clientes_ativos > 0 else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.success("✅ Dashboard carregado com sucesso")
