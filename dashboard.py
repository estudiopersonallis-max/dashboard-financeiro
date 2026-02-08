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

# ================= FUNÇÕES DE GRÁFICO =================
def bar_chart_com_valor(series, titulo, prefixo=""):
    fig, ax = plt.subplots()
    bars = ax.bar(series.index.astype(str), series.values)
    ax.set_title(titulo)
    ax.set_xticklabels(series.index.astype(str), rotation=45, ha="right")

    for bar in bars:
        altura = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            altura,
            f"{prefixo}{altura:,.2f}" if prefixo else f"{int(altura)}",
            ha="center",
            va="bottom",
            fontsize=8
        )

    st.pyplot(fig)


def pizza_percentual(series, titulo):
    fig, ax = plt.subplots(figsize=(5, 5))

    ax.pie(
        series,
        startangle=90,
        autopct="%1.1f%%",
        textprops={"fontsize": 8}
    )

    ax.legend(
        series.index,
        title="Legenda",
        loc="center left",
        bbox_to_anchor=(1, 0.5),
        fontsize=8
    )

    ax.set_title(titulo)
    ax.axis("equal")
    st.pyplot(fig)

# ================= MODALIDADE =================
st.subheader("📌 Valor por Modalidade")
valor_modalidade = df_filtro.groupby("Modalidade")["Valor"].sum()
st.dataframe(valor_modalidade)
bar_chart_com_valor(valor_modalidade, "Valor por Modalidade (€)", "€ ")
pizza_percentual(valor_modalidade, "% Valor por Modalidade")

# ================= TIPO =================
st.subheader("📌 Valor por Tipo")
valor_tipo = df_filtro.groupby("Tipo")["Valor"].sum()
st.dataframe(valor_tipo)
bar_chart_com_valor(valor_tipo, "Valor por Tipo (€)", "€ ")
pizza_percentual(valor_tipo, "% Valor por Tipo")

# ================= PROFESSOR =================
st.subheader("📌 Valor por Professor")
valor_professor = df_filtro.groupby("Professor")["Valor"].sum()
st.dataframe(valor_professor)
bar_chart_com_valor(valor_professor, "Valor por Professor (€)", "€ ")
pizza_percentual(valor_professor, "% Valor por Professor")

# ================= LOCAL =================
st.subheader("📌 Valor por Local")
valor_local = df_filtro.groupby("Local")["Valor"].sum()
st.dataframe(valor_local)
bar_chart_com_valor(valor_local, "Valor por Local (€)", "€ ")
pizza_percentual(valor_local, "% Valor por Local")

st.divider()

# ================= CLIENTES =================
st.subheader("👥 Clientes por Local")
clientes_local = df_filtro[df_filtro["Ativo"]].groupby("Local")["Nome do cliente"].nunique()
st.dataframe(clientes_local)
bar_chart_com_valor(clientes_local, "Clientes por Local")
pizza_percentual(clientes_local, "% Clientes por Local")

st.subheader("👥 Clientes por Professor")
clientes_professor = df_filtro[df_filtro["Ativo"]].groupby("Professor")["Nome do cliente"].nunique()
st.dataframe(clientes_professor)
bar_chart_com_valor(clientes_professor, "Clientes por Professor")
pizza_percentual(clientes_professor, "% Clientes por Professor")

st.divider()

# ================= RELATÓRIO HTML =================
st.header("📄 Relatório")

if st.button("🧾 Gerar relatório em HTML"):
    html = f"""
    <html>
    <head>
        <title>Relatório Financeiro - {periodo}</title>
        <style>
            body {{ font-family: Arial; margin: 40px; }}
            table {{ border-collapse: collapse; width: 100%; }}
            th, td {{ border: 1px solid #ccc; padding: 8px; }}
            th {{ background: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h1>Relatório Financeiro – {periodo}</h1>
        <p><b>Valor Total:</b> € {total_valor:,.2f}</p>
        <p><b>Clientes Ativos:</b> {clientes_ativos}</p>
        <p><b>Ticket Médio:</b> € {ticket_medio:,.2f}</p>

        <h2>Valor por Modalidade</h2>
        {valor_modalidade.to_frame("Valor").to_html()}
    </body>
    </html>
    """

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    Path(tmp.name).write_text(html, encoding="utf-8")

    st.success("Relatório HTML gerado com sucesso")
    st.markdown(f"[👉 Abrir relatório]({tmp.name})", unsafe_allow_html=True)
