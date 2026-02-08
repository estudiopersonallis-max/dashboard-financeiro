import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from pathlib import Path

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
        df_temp["Nome do cliente"].astype(str).str.strip().str.upper()
    )

    coluna_status = df_temp.columns[2]
    df_temp["Ativo"] = (
        df_temp[coluna_status].astype(str).str.strip().str.upper().eq("ATIVO")
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
total_valor = df_filtro["Valor"].sum()
perdas = int(df_filtro["É Perda"].sum())
ticket_medio = total_valor / clientes_ativos if clientes_ativos > 0 else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Total", f"€ {total_valor:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("🎟️ Ticket Médio", f"€ {ticket_medio:,.2f}")

st.divider()

# ================= DADOS PARA RELATÓRIO =================
valor_modalidade = df_filtro.groupby("Modalidade")["Valor"].sum()
valor_tipo = df_filtro.groupby("Tipo")["Valor"].sum()
valor_professor = df_filtro.groupby("Professor")["Valor"].sum()
valor_local = df_filtro.groupby("Local")["Valor"].sum()

# ================= RELATÓRIO PDF (HTML LEVE) =================
st.header("📄 Relatório Mensal (PDF)")

st.info("👉 Clique para gerar o relatório e depois use **Ctrl+P → Salvar como PDF**")

if st.button("🧾 Gerar relatório em HTML (leve)"):
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>Relatório Financeiro - {periodo}</title>
        <style>
            body {{ font-family: Arial; margin: 30px; }}
            h1, h2 {{ border-bottom: 1px solid #ccc; padding-bottom: 4px; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
            th, td {{ border: 1px solid #ccc; padding: 6px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>

        <h1>Relatório Financeiro</h1>
        <p><b>Período:</b> {periodo}</p>

        <h2>Resumo</h2>
        <ul>
            <li><b>Valor Total:</b> € {total_valor:,.2f}</li>
            <li><b>Clientes Ativos:</b> {clientes_ativos}</li>
            <li><b>Perdas:</b> {perdas}</li>
            <li><b>Ticket Médio:</b> € {ticket_medio:,.2f}</li>
        </ul>

        <h2>Valor por Modalidade</h2>
        {valor_modalidade.to_frame("Valor (€)").to_html()}

        <h2>Valor por Tipo</h2>
        {valor_tipo.to_frame("Valor (€)").to_html()}

        <h2>Valor por Professor</h2>
        {valor_professor.to_frame("Valor (€)").to_html()}

        <h2>Valor por Local</h2>
        {valor_local.to_frame("Valor (€)").to_html()}

    </body>
    </html>
    """

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    Path(tmp.name).write_text(html, encoding="utf-8")

    st.success("Relatório gerado com sucesso")
    st.markdown(f"👉 [Abrir relatório para imprimir em PDF]({tmp.name})", unsafe_allow_html=True)
