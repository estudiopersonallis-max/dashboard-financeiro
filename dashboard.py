import streamlit as st
import pandas as pd
import datetime

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER

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

    # Datas (usadas só para dia / trimestre / ano)
    df_temp["Data"] = pd.to_datetime(df_temp["Data"])
    df_temp["Dia"] = df_temp["Data"].dt.day
    df_temp["Ano"] = df_temp["Data"].dt.year
    df_temp["Trimestre"] = df_temp["Data"].dt.to_period("Q").astype(str)

    # 🔹 NORMALIZAÇÃO DOS NOMES (CORREÇÃO DOS CLIENTES ATIVOS)
    df_temp["Nome do cliente"] = (
        df_temp["Nome do cliente"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Perdas
    df_temp["É Perda"] = df_temp["Perdas"].notna()

    dfs.append(df_temp)

df = pd.concat(dfs, ignore_index=True)

# ================= FILTRO DE PERÍODO =================
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
    st.subheader("📌 Valor por Modalidade")
    valor_modalidade = df_filtro.groupby("Modalidade")["Valor"].sum()
    st.dataframe(valor_modalidade)

    st.subheader("📌 Valor por Tipo")
    valor_tipo = df_filtro.groupby("Tipo")["Valor"].sum()
    st.dataframe(valor_tipo)

with col2:
    st.subheader("📌 Valor por Professor")
    valor_professor = df_filtro.groupby("Professor")["Valor"].sum()
    st.dataframe(valor_professor)

    st.subheader("📌 Valor por Local")
    valor_local = df_filtro.groupby("Local")["Valor"].sum()
    st.dataframe(valor_local)

st.divider()

# ================= PERÍODOS DO MÊS =================
st.subheader("📅 Valor por Período do Mês")

p1 = df_filtro[df_filtro["Dia"] <= 10]["Valor"].sum()
p2 = df_filtro[(df_filtro["Dia"] > 10) & (df_filtro["Dia"] <= 20)]["Valor"].sum()
p3 = df_filtro[df_filtro["Dia"] > 20]["Valor"].sum()

valor_periodo = pd.Series({
    "Dias 1–10": p1,
    "Dias 11–20": p2,
    "Dias 21–fim": p3
})

st.dataframe(valor_periodo)

st.divider()

# ================= CLIENTES =================
st.subheader("👥 Clientes")

col1, col2 = st.columns(2)

with col1:
    clientes_local = df_filtro.groupby("Local")["Nome do cliente"].nunique()
    st.dataframe(clientes_local.rename("Clientes por Local"))

with col2:
    clientes_professor = df_filtro.groupby("Professor")["Nome do cliente"].nunique()
    st.dataframe(clientes_professor.rename("Clientes por Professor"))

st.divider()

st.subheader("🎟️ Ticket Médio por Tipo")
ticket_tipo = df_filtro.groupby("Tipo")["Valor"].mean()
st.dataframe(ticket_tipo)

# ================= GRÁFICOS =================
st.divider()
st.header("📊 Gráficos")

st.bar_chart(valor_modalidade)
st.bar_chart(valor_tipo)
st.bar_chart(valor_professor)
st.bar_chart(valor_local)
st.bar_chart(valor_periodo)
st.bar_chart(clientes_local)
st.bar_chart(clientes_professor)
st.bar_chart(ticket_tipo)

# ================= COMPARATIVO ANUAL =================
st.divider()
st.header("📈 Comparativo Anual / Global")

valor_por_mes = df.groupby("Mes")["Valor"].sum()
st.line_chart(valor_por_mes)

# ================= RELATÓRIO PDF =================
def gerar_relatorio_pdf():
    nome = f"Relatorio_{periodo}.pdf"

    doc = SimpleDocTemplate(
        nome,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    story = []

    titulo = ParagraphStyle(
        "Titulo",
        parent=styles["Heading1"],
        alignment=TA_CENTER
    )

    story.append(Paragraph("Relatório Financeiro", titulo))
    story.append(Spacer(1, 12))

    story.append(Paragraph(f"<b>Período:</b> {periodo}", styles["Normal"]))
    story.append(Paragraph(
        f"<b>Gerado em:</b> {datetime.date.today().strftime('%d/%m/%Y')}",
        styles["Normal"]
    ))

    story.append(Spacer(1, 20))

    story.append(Paragraph("<b>Resumo Executivo</b>", styles["Heading2"]))
    story.append(Paragraph(f"Valor Total: € {total_valor:,.2f}", styles["Normal"]))
    story.append(Paragraph(f"Clientes Ativos: {clientes_ativos}", styles["Normal"]))
    story.append(Paragraph(f"Perdas: {perdas}", styles["Normal"]))
    story.append(Paragraph(f"Ticket Médio: € {ticket_medio:,.2f}", styles["Normal"]))

    story.append(Spacer(1, 20))

    def tabela(titulo, serie):
        story.append(Paragraph(titulo, styles["Heading3"]))
        data = [["Categoria", "Valor"]]
        for k, v in serie.items():
            data.append([str(k), f"€ {v:,.2f}"])
        story.append(Table(data))
        story.append(Spacer(1, 15))

    tabela("Valor por Modalidade", valor_modalidade)
    tabela("Valor por Tipo", valor_tipo)
    tabela("Valor por Professor", valor_professor)
    tabela("Valor por Local", valor_local)

    doc.build(story)
    return nome

st.divider()
st.header("📄 Relatório Mensal")

if st.button("Gerar relatório PDF"):
    pdf = gerar_relatorio_pdf()
    with open(pdf, "rb") as f:
        st.download_button(
            "⬇️ Download do relatório",
            f,
            file_name=pdf,
            mime="application/pdf"
        )
