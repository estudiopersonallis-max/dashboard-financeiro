import streamlit as st
import pandas as pd
import datetime

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

    # Normalizar nomes dos clientes
    df_temp["Nome do cliente"] = (
        df_temp["Nome do cliente"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # 🔹 REGRA CORRETA DE CLIENTE ATIVO (COLUNA C)
    df_temp["Ativo"] = (
        df_temp["C"]
        .astype(str)
        .str.strip()
        .str.lower()
        .ne("")
        & ~df_temp["C"].astype(str).str.lower().str.contains("inativo")
    )

    # Perdas continuam separadas
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
perdas = df_filtro["É Perda"].sum()

total_valor = df_filtro["Valor"].sum()
ticket_medio = df_filtro["Valor"].mean()

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()

# ================= RELATÓRIO HTML =================
def gerar_relatorio_html():
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial; }}
            h1 {{ text-align: center; }}
        </style>
    </head>
    <body>
        <h1>Relatório Financeiro</h1>
        <p><b>Período:</b> {periodo}</p>
        <p><b>Valor Total:</b> € {total_valor:,.2f}</p>
        <p><b>Clientes Ativos:</b> {clientes_ativos}</p>
        <p><b>Perdas:</b> {perdas}</p>
        <p><b>Ticket Médio:</b> € {ticket_medio:,.2f}</p>
    </body>
    </html>
    """
    return html

st.divider()
st.header("📄 Relatório")

st.download_button(
    "⬇️ Download do relatório (HTML → PDF)",
    data=gerar_relatorio_html(),
    file_name=f"Relatorio_{periodo}.html",
    mime="text/html"
)
