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

    # Mês pelo nome do ficheiro
    mes_ficheiro = file.name.replace(".xlsx", "")
    df_temp["Mes"] = mes_ficheiro

    # Datas (apenas para dia / ano / trimestre)
    df_temp["Data"] = pd.to_datetime(df_temp["Data"])
    df_temp["Dia"] = df_temp["Data"].dt.day
    df_temp["Ano"] = df_temp["Data"].dt.year
    df_temp["Trimestre"] = df_temp["Data"].dt.to_period("Q").astype(str)

    # Normalizar nome do cliente
    df_temp["Nome do cliente"] = (
        df_temp["Nome do cliente"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ================= ATIVOS (COLUNA C) =================
    coluna_status = df_temp.columns[2]  # coluna C

    df_temp["Ativo"] = (
        df_temp[coluna_status]
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("ATIVO")
    )

    # Perdas
    df_temp["É Perda"] = df_temp["Perdas"].notna()

    dfs.append(df_temp)

df = pd.concat(dfs, ignore_index=True)

# ================= FILTRO DE PERÍODO ========
