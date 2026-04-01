import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

# ================= CONFIG =================
st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= CACHE =================
@st.cache_data
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = f.name.replace(".xlsx", "").upper()
        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Nome do cliente"] = df.get("Nome do cliente", "").astype(str).str.upper().str.strip()
        df["Modalidade"] = df.get("Modalidade", "N/A")
        df["Tipo"] = df.get("Tipo", "N/A")
        df["Professor"] = df.get("Professor", "N/A")
        df["Local"] = df.get("Local", "N/A")

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

@st.cache_data
def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        df = df.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df.empty:
            continue

        df["Periodo"] = f.name.replace(".xlsx", "").upper()
        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)
        df["Classe"] = df.get("Classe", "N/A").astype(str).str.upper().str.strip()
        df["Local"] = df.get("Local", "N/A").astype(str).str.strip()

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= UPLOAD =================
st.sidebar.header("📤 Upload")
uploaded_receitas = st.sidebar.file_uploader("Receitas", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.sidebar.file_uploader("Despesas", type=["xlsx"], accept_multiple_files=True)

receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= FILTROS =================
st.sidebar.header("🔎 Filtros")
periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))

periodo_sel = st.sidebar.multiselect("Períodos", periodos, default=periodos)

if not receitas.empty:
    receitas = receitas[receitas["Periodo"].isin(periodo_sel)]

if not despesas.empty:
    despesas = despesas[despesas["Periodo"].isin(periodo_sel)]
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= KPIs =================
st.subheader("📌 KPIs")

receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0

# CORREÇÃO DO LUCRO
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("Receita", f"{receita_total:,.0f}€")
col2.metric("Despesa", f"{despesa_total:,.0f}€")
col3.metric("Lucro", f"{lucro_total:,.0f}€")
col4.metric("Margem", f"{margem:.1f}%")

# ================= KPIs POR PERÍODO =================
kpis = []
for p in periodos:
    r = receitas[receitas["Periodo"] == p]
    d = despesas[despesas["Periodo"] == p]

    receita = r["Valor"].sum()
    despesa = d["Valor"].sum()
    lucro = receita - despesa

    kpis.append({
        "Período": p,
        "Receita": receita,
        "Despesa": despesa,
        "Lucro": lucro
    })

df_kpis = pd.DataFrame(kpis)

if not df_kpis.empty:
    df_kpis["Margem (%)"] = (df_kpis["Lucro"] / df_kpis["Receita"]) * 100
    df_kpis["Δ Receita (%)"] = df_kpis["Receita"].pct_change() * 100
    df_kpis["Δ Lucro (%)"] = df_kpis["Lucro"].pct_change() * 100

    st.dataframe(df_kpis.round(2), use_container_width=True)

# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    if df.empty:
        return

    df = df.loc[df.sum(axis=1).sort_values(ascending=False).index]

    fig, ax = plt.subplots()
    df.plot(kind="barh", ax=ax)
    ax.set_title(titulo)
    ax.set_xlabel("€")
    st.pyplot(fig)

# ================= ANÁLISE =================
def bloco_analise(df, categoria, titulo):
    if df.empty or categoria not in df.columns:
        return

    pivot = df.pivot_table(index=categoria, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
    percent = pivot.div(pivot.sum(axis=0), axis=1) * 100

    st.markdown(f"### {titulo} por {categoria}")

    tabela = pivot.round(2).astype(str) + " € | " + percent.round(1).astype(str) + " %"
    st.dataframe(tabela, use_container_width=True)

    grafico_bar(pivot, f"{titulo} por {categoria}")

# ================= TABS =================
tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "💰 Receitas", "💸 Despesas"])

with tab1:
    st.subheader("Resumo Financeiro")
    if not df_kpis.empty:
        fig_resumo, ax = plt.subplots()
        df_kpis.set_index("Período")[["Receita", "Despesa", "Lucro"]].plot(kind="bar", ax=ax)
        st.pyplot(fig_resumo)
    else:
        fig_resumo = None

with tab2:
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        bloco_analise(receitas, cat, "Receitas")

with tab3:
    for cat in ["Classe", "Local"]:
        bloco_analise(despesas, cat, "Despesas")

# ================= TOP CLIENTES =================
st.subheader("🏆 Top Clientes")
if not receitas.empty:
    top = receitas.groupby("Nome do cliente")["Valor"].sum().nlargest(10)
    st.bar_chart(top)

# ================= PDF COMPLETO (CONSULTORIA) =================
st.subheader("📄 Exportar Relatório Completo (Consultoria)")

def gerar_insights(df_kpis):
    insights = []
    if df_kpis.empty:
        return insights

    total_lucro = df_kpis["Lucro"].sum()
    if total_lucro > 0:
        insights.append("A operação apresenta resultado positivo no período analisado.")
    else:
        insights.append("A operação apresenta prejuízo e requer revisão de custos.")

    if len(df_kpis) > 1:
        ult = df_kpis.iloc[-1]
        prev = df_kpis.iloc[-2]

        if prev["Receita"] != 0:
            var_r = (ult["Receita"] - prev["Receita"]) / prev["Receita"] * 100
            insights.append(f"A receita variou {var_r:.1f}% no último período.")

        if prev["Lucro"] != 0:
            var_l = (ult["Lucro"] - prev["Lucro"]) / prev["Lucro"] * 100
            insights.append(f"O lucro variou {var_l:.1f}% no último período.")

    return insights


def tabela_categoria(df, categoria):
    if df.empty or categoria not in df.columns:
        return None
    pivot = df.pivot_table(index=categoria, columns="Periodo", values="Valor", aggfunc="sum", fill_value=0)
    percent = pivot.div(pivot.sum(axis=0), axis=1) * 100
    data = [[categoria] + list(pivot.columns)]
    for idx in pivot.index:
        linha = [str(idx)]
        for col in pivot.columns:
            linha.append(f"{pivot.loc[idx,col]:,.2f}€ ({percent.loc[idx,col]:.1f}%)")
        data.append(linha)
    return data


def gerar_diagnostico_avancado(df_kpis, receitas, despesas):
    textos = []

    if df_kpis.empty:
        return textos

    # Margem média
    margem_media = df_kpis["Margem (%)"].mean()
    textos.append(f"A margem média do período foi de {margem_media:.1f}%.")

    # Tendência
    if len(df_kpis) > 1:
        if df_kpis["Lucro"].iloc[-1] > df_kpis["Lucro"].iloc[0]:
            textos.append("Observa-se uma tendência de crescimento do lucro ao longo do período.")
        else:
            textos.append("Observa-se uma tendência de queda do lucro ao longo do período.")

    # Concentração clientes
    if not receitas.empty:
        top = receitas.groupby("Nome do cliente")["Valor"].sum()
        share_top = top.nlargest(3).sum() / top.sum() * 100 if top.sum() else 0
        textos.append(f"Os 3 principais clientes representam {share_top:.1f}% da receita total.")

    # Peso despesas
    if not despesas.empty and not receitas.empty:
        ratio = despesas["Valor"].sum() / receitas["Valor"].sum() * 100 if receitas["Valor"].sum() else 0
        textos.append(f"As despesas representam {ratio:.1f}% da receita.")

    return textos


def gerar_pdf_completo(df_kpis, receitas, despesas, fig_resumo):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    insights = gerar_insights(df_kpis)
    diagnostico = gerar_diagnostico_avancado(df_kpis, receitas, despesas)

    # CAPA
    elementos.append(Spacer(1, 6*cm))
    elementos.append(Paragraph("RELATÓRIO EXECUTIVO FINANCEIRO", styles["Title"]))
    elementos.append(Paragraph("Análise Premium", styles["Heading2"]))
    elementos.append(Spacer(1, 1*cm))
    elementos.append(Paragraph(datetime.now().strftime("%d/%m/%Y"), styles["Normal"]))
    elementos.append(PageBreak())

    # RESUMO
    elementos.append(Paragraph("Resumo Executivo", styles["Heading1"]))

    tabela = [["Período", "Receita", "Despesa", "Lucro", "Margem"]]
    for _, row in df_kpis.iterrows():
        tabela.append([
            row["Período"],
            f"{row['Receita']:,.2f}€",
            f"{row['Despesa']:,.2f}€",
            f"{row['Lucro']:,.2f}€",
            f"{row['Margem (%)']:.1f}%"
        ])

    elementos.append(Table(tabela))
    elementos.append(Spacer(1, 1*cm))

    # INSIGHTS
    elementos.append(Paragraph("Principais Insights", styles["Heading2"]))
    for i in insights:
        elementos.append(Paragraph(f"• {i}", styles["Normal"]))

    elementos.append(Spacer(1, 1*cm))

    # DIAGNÓSTICO AVANÇADO
    elementos.append(Paragraph("Diagnóstico Financeiro", styles["Heading2"]))
    for d in diagnostico:
        elementos.append(Paragraph(f"• {d}", styles["Normal"]))

    elementos.append(PageBreak())

    # GRÁFICO PRINCIPAL
    if fig_resumo:
        img = BytesIO()
        fig_resumo.savefig(img, format="png", bbox_inches="tight")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=9*cm))
        elementos.append(PageBreak())

    # RECEITAS
    elementos.append(Paragraph("Análise de Receitas", styles["Heading1"]))
    for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
        data = tabela_categoria(receitas, cat)
        if data:
            elementos.append(Paragraph(f"Receitas por {cat}", styles["Heading2"]))
            elementos.append(Table(data))
            elementos.append(Spacer(1, 1*cm))

    elementos.append(PageBreak())

    # DESPESAS
    elementos.append(Paragraph("Análise de Despesas", styles["Heading1"]))
    for cat in ["Classe", "Local"]:
        data = tabela_categoria(despesas, cat)
        if data:
            elementos.append(Paragraph(f"Despesas por {cat}", styles["Heading2"]))
            elementos.append(Table(data))
            elementos.append(Spacer(1, 1*cm))

    # TOP CLIENTES
    if not receitas.empty:
        top = receitas.groupby("Nome do cliente")["Valor"].sum().nlargest(10)
        elementos.append(PageBreak())
        elementos.append(Paragraph("Top Clientes", styles["Heading1"]))
        data = [["Cliente", "Receita"]]
        for idx, val in top.items():
            data.append([idx, f"{val:,.2f}€"])
        elementos.append(Table(data))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

if st.button("Gerar PDF Premium"):
    pdf = gerar_pdf_completo(df_kpis, receitas, despesas, fig_resumo)

    st.download_button(
        label="📥 Download PDF Premium",
        data=pdf,
        file_name="relatorio_financeiro_premium.pdf",
        mime="application/pdf"
    )
# ================= BIG4 - TEXTO EXECUTIVO =================
def gerar_texto_executivo(df_kpis):
    if df_kpis.empty:
        return "Dados insuficientes para análise."

    receita_total = df_kpis["Receita"].sum()
    lucro_total = df_kpis["Lucro"].sum()
    margem_media = df_kpis["Margem (%)"].mean()

    tendencia = ""
    if len(df_kpis) > 1:
        if df_kpis["Lucro"].iloc[-1] > df_kpis["Lucro"].iloc[0]:
            tendencia = "Observa-se uma trajetória de crescimento consistente ao longo do período analisado."
        else:
            tendencia = "Verifica-se uma tendência de deterioração do resultado ao longo do período."

    texto = f"""
    O desempenho financeiro demonstra uma receita acumulada de {receita_total:,.0f}€, 
    com resultado líquido de {lucro_total:,.0f}€, refletindo uma margem média de {margem_media:.1f}%.

    {tendencia}
    """

    return texto


# ================= ALERTAS INTELIGENTES =================
def gerar_alertas(receitas, despesas, df_kpis):
    alertas = []

    if df_kpis.empty:
        return alertas

    # Margem baixa
    if df_kpis["Margem (%)"].mean() < 10:
        alertas.append("Margem média inferior a 10% — possível pressão de custos.")

    # Prejuízo
    if (df_kpis["Lucro"] < 0).any():
        alertas.append("Existem períodos com prejuízo operacional.")

    # Concentração clientes
    if not receitas.empty:
        top = receitas.groupby("Nome do cliente")["Valor"].sum()
        if not top.empty:
            share = top.max() / top.sum() * 100
            if share > 40:
                alertas.append(f"Elevada dependência de um único cliente ({share:.1f}% da receita).")

    # Despesas elevadas
    if not receitas.empty and not despesas.empty:
        ratio = despesas["Valor"].sum() / receitas["Valor"].sum() * 100
        if ratio > 80:
            alertas.append("Estrutura de custos elevada (>80% da receita).")

    return alertas


# ================= PDF BIG4 FINAL =================
def gerar_pdf_big4(df_kpis, receitas, despesas, fig_resumo):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, PageBreak, Image
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []

    texto_exec = gerar_texto_executivo(df_kpis)
    alertas = gerar_alertas(receitas, despesas, df_kpis)

    # CAPA
    elementos.append(Spacer(1, 6*cm))
    elementos.append(Paragraph("RELATÓRIO FINANCEIRO", styles["Title"]))
    elementos.append(Paragraph("Análise Executiva - Nível Consultoria", styles["Heading2"]))
    elementos.append(Spacer(1, 1*cm))
    elementos.append(Paragraph(datetime.now().strftime("%d/%m/%Y"), styles["Normal"]))
    elementos.append(PageBreak())

    # TEXTO EXECUTIVO
    elementos.append(Paragraph("Sumário Executivo", styles["Heading1"]))
    elementos.append(Paragraph(texto_exec, styles["Normal"]))
    elementos.append(Spacer(1, 1*cm))

    # ALERTAS
    if alertas:
        elementos.append(Paragraph("Principais Alertas", styles["Heading2"]))
        for a in alertas:
            elementos.append(Paragraph(f"• {a}", styles["Normal"]))
        elementos.append(Spacer(1, 1*cm))

    # TABELA KPI
    tabela = [["Período", "Receita", "Despesa", "Lucro", "Margem"]]
    for _, row in df_kpis.iterrows():
        tabela.append([
            row["Período"],
            f"{row['Receita']:,.0f}€",
            f"{row['Despesa']:,.0f}€",
            f"{row['Lucro']:,.0f}€",
            f"{row['Margem (%)']:.1f}%"
        ])
    elementos.append(Table(tabela))

    elementos.append(PageBreak())

    # GRÁFICO
    if fig_resumo:
        img = BytesIO()
        fig_resumo.savefig(img, format="png", bbox_inches="tight")
        img.seek(0)
        elementos.append(Image(img, width=16*cm, height=9*cm))

    elementos.append(PageBreak())

    # TOP CLIENTES
    if not receitas.empty:
        top = receitas.groupby("Nome do cliente")["Valor"].sum().nlargest(10)
        elementos.append(Paragraph("Top Clientes", styles["Heading1"]))

        data = [["Cliente", "Receita"]]
        for c, v in top.items():
            data.append([c, f"{v:,.0f}€"])

        elementos.append(Table(data))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
