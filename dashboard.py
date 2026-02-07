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

    # 🔹 NORMALIZAÇÃO DOS NOMES (CORRIGE CLIENTES ATIVOS)
    df_temp["Nome do cliente"] = (
        df_temp["Nome do cliente"]
        .astype(str)
        .str.strip()
        .str.upper()
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
clientes_ativos = df_filtro.loc[~df_filtro["É Perda"], "Nome do cliente"].nunique()
perdas = df_filtro["É Perda"].sum()

total_valor = df_filtro["Valor"].sum()
ticket_medio = df_filtro["Valor"].mean()

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()

# ================= TABELAS =================
col1, col2 = st.columns(2)

with col1:
    valor_modalidade = df_filtro.groupby("Modalidade")["Valor"].sum()
    st.subheader("📌 Valor por Modalidade")
    st.dataframe(valor_modalidade)

    valor_tipo = df_filtro.groupby("Tipo")["Valor"].sum()
    st.subheader("📌 Valor por Tipo")
    st.dataframe(valor_tipo)

with col2:
    valor_professor = df_filtro.groupby("Professor")["Valor"].sum()
    st.subheader("📌 Valor por Professor")
    st.dataframe(valor_professor)

    valor_local = df_filtro.groupby("Local")["Valor"].sum()
    st.subheader("📌 Valor por Local")
    st.dataframe(valor_local)

st.divider()

# ================= CLIENTES =================
clientes_local = df_filtro.groupby("Local")["Nome do cliente"].nunique()
clientes_professor = df_filtro.groupby("Professor")["Nome do cliente"].nunique()

# ================= RELATÓRIO HTML =================
def gerar_relatorio_html():
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial; }}
            h1 {{ text-align: center; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
            th, td {{ border: 1px solid #ccc; padding: 8px; }}
            th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>

    <h1>Relatório Financeiro</h1>
    <p><b>Período:</b> {periodo}</p>
    <p><b>Gerado em:</b> {datetime.date.today().strftime('%d/%m/%Y')}</p>

    <h2>Resumo Executivo</h2>
    <ul>
        <li><b>Valor Total:</b> € {total_valor:,.2f}</li>
        <li><b>Clientes Ativos:</b> {clientes_ativos}</li>
        <li><b>Perdas:</b> {perdas}</li>
        <li><b>Ticket Médio:</b> € {ticket_medio:,.2f}</li>
    </ul>

    <h2>Valor por Modalidade</h2>
    {valor_modalidade.to_frame("Valor").to_html()}

    <h2>Valor por Tipo</h2>
    {valor_tipo.to_frame("Valor").to_html()}

    <h2>Valor por Professor</h2>
    {valor_professor.to_frame("Valor").to_html()}

    <h2>Valor por Local</h2>
    {valor_local.to_frame("Valor").to_html()}

    </body>
    </html>
    """
    return html

st.divider()
st.header("📄 Relatório")

html_relatorio = gerar_relatorio_html()

st.download_button(
    "⬇️ Download do relatório (HTML → PDF)",
    data=html_relatorio,
    file_name=f"Relatorio_{periodo}.html",
    mime="text/html"
)
