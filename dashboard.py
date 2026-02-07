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
    st.info("⬆️ Carregue pelo menos um ficheiro Excel")
    st.stop()

# ================= LEITURA DOS FICHEIROS =================
dfs = []

for file in uploaded_files:
    df_temp = pd.read_excel(file)

    # Normalizar colunas
    df_temp.columns = df_temp.columns.str.strip()

    # Converter datas (mantido como no código original que funcionava)
    df_temp["Data"] = pd.to_datetime(df_temp["Data"], errors="coerce")
    df_temp = df_temp.dropna(subset=["Data"])
    df_temp["Dia"] = df_temp["Data"].dt.day

    # Criar coluna Mês
