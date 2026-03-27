import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import tempfile
import matplotlib

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from datetime import datetime

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= UPLOAD =================
st.subheader("📤 Upload de Ficheiros (cada ficheiro = um período)")
uploaded_receitas = st.file_uploader("Receitas (Excel)", type=["xlsx"], accept_multiple_files=True)
uploaded_despesas = st.file_uploader("Despesas (Excel)", type=["xlsx"], accept_multiple_files=True)

# ================= FUNÇÕES =================
def nome_periodo(nome):
    return nome.replace(".xlsx", "").upper()

def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Nome do cliente"] = df["Nome do cliente"].astype(str).str.upper().str.strip()
        df["Modalidade"] = df.get("Modalidade", "N/A")
        df["Tipo"] = df.get("Tipo", "N/A")
        df["Professor"] = df.get("Professor", "N/A")
        df["Local"] = df.get("Local", "N/A")

        coluna_status = df.columns[2]
        df["Ativo"] = df[coluna_status].astype(str).str.upper().eq("ATIVO")
        df["É Perda"] = df["Perdas"].notna() if "Perdas" in df.columns else False

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def ler_despesas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        df = df.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df.empty:
            continue

        df["Periodo"] = nome_periodo(f.name)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Classe"] = df["Classe"].astype(str).str.upper().str.strip()
        df["Local"] = df["Local"].astype(str).str.strip()

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= FILTRO =================
if not despesas.empty:
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= KPIs =================
st.subheader("📌 KPIs Comparativos")

periodos = sorted(set(receitas.get("Periodo", [])).union(set(despesas.get("Periodo", []))))
kpis = []

for p in periodos:
    r = receitas[receitas["Periodo"] == p] if not receitas.empty else pd.DataFrame()
    d = despesas[despesas["Periodo"] == p] if not despesas.empty else pd.DataFrame()

    receita = r["Valor"].sum() if not r.empty else 0
    despesa = d["Valor"].sum() if not d.empty else 0
    lucro = receita + despesa

    kpis.append({
        "Período": p,
        "Receita (€)": round(receita, 2),
        "Despesa (€)": round(despesa, 2),
        "Lucro (€)": round(lucro, 2)
    })

df_kpis = pd.DataFrame(kpis)
st.dataframe(df_kpis, use_container_width=True)

# ================= GRÁFICOS =================
def grafico_bar(df, titulo):
    if df.empty:
        return None
    fig, ax = plt.subplots()
    df.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    ax.set_ylabel("€")
    ax.legend(title="Período")
    return fig

def grafico_pizza(series, titulo):
    if series.sum() == 0:
        return None

    valores = series.abs()
    fig, ax = plt.subplots(figsize=(5, 5))
    ax.pie(
        valores,
        autopct=lambda p: f"{p:.1f}%",
        startangle=90,
        pctdistance=1.15,
        labeldistance=1.3,
        textprops={"fontsize": 8}
    )
    ax.legend(
        valores.index,
        loc="center left",
        bbox_to_anchor=(1, 0.5),
        fontsize=8
    )
    ax.set_title(titulo)
    ax.axis("equal")
    return fig

def bloco_analise(df, categoria, titulo_base):
    if df.empty or categoria not in df.columns:
        return

    pivot = df.pivot_table(
        index=categoria,
        columns="Periodo",
        values="Valor",
        aggfunc="sum",
        fill_value=0
    )

    percent = pivot.div(pivot.sum(axis=0), axis=1) * 100
    tabela = pivot.round(2).astype(str) + " € | " + percent.round(1).astype(str) + " %"

    st.markdown(f"### {titulo_base} por {categoria}")
    st.dataframe(tabela, use_container_width=True)

    fig_bar = grafico_bar(pivot, f"{titulo_base} por {categoria} (€)")
    if fig_bar:
        st.pyplot(fig_bar)

    for p in pivot.columns:
        fig = grafico_pizza(pivot[p], f"{titulo_base} – {categoria} (%) | {p}")
        if fig:
            st.pyplot(fig)

# ================= RECEITAS =================
st.subheader("📌 Receitas – Distribuição Percentual e Valor")
for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
    bloco_analise(receitas, cat, "Receitas")

# ================= DESPESAS =================
st.subheader("📌 Despesas – Distribuição Percentual e Valor")
for cat in ["Classe", "Local"]:
    bloco_analise(despesas, cat, "Despesas")


# ================= FIGURA RESUMO (FIX DO ERRO) =================
def grafico_resumo(df_kpis):
    if df_kpis.empty:
        return None

    fig, ax = plt.subplots()
    df_kpis.set_index("Período")[["Receita (€)", "Despesa (€)", "Lucro (€)"]].plot(kind="bar", ax=ax)
    ax.set_title("Resumo Financeiro")
    ax.set_ylabel("€")
    return fig

fig_resumo = grafico_resumo(df_kpis)

if fig_resumo:
    st.pyplot(fig_resumo)

st.divider()

# ================= PDF =================
st.subheader("📄 Exportar PDF Executivo")

def gerar_pdf_executivo(df_kpis, fig):

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    elementos = []

    elementos.append(Spacer(1, 6*cm))
    elementos.append(Paragraph("RELATÓRIO FINANCEIRO", styles["Title"]))
    elementos.append(Spacer(1, 1*cm))
    elementos.append(Paragraph("Dashboard Financeiro", styles["Heading2"]))
    elementos.append(Spacer(1, 1*cm))

    data_hoje = datetime.now().strftime("%d/%m/%Y")
    elementos.append(Paragraph(f"Data: {data_hoje}", styles["Normal"]))

    elementos.append(PageBreak())

    elementos.append(Paragraph("Resumo Executivo", styles["Heading1"]))
    elementos.append(Spacer(1, 0.5*cm))

    tabela_data = [["Período", "Receita (€)", "Despesa (€)", "Lucro (€)"]]

    for _, row in df_kpis.iterrows():
        tabela_data.append([
            row["Período"],
            f"{row['Receita (€)']:,.2f}",
            f"{row['Despesa (€)']:,.2f}",
            f"{row['Lucro (€)']:,.2f}"
        ])

    tabela = Table(tabela_data)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
    ]))

    elementos.append(tabela)
    elementos.append(PageBreak())

    if fig:
        img_buffer = BytesIO()
        fig.savefig(img_buffer, format="png", bbox_inches="tight")
        img_buffer.seek(0)
        elementos.append(Image(img_buffer, width=16*cm, height=9*cm))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================= BOTÃO =================
if st.button("📄 Gerar PDF Executivo"):
    pdf = gerar_pdf_executivo(df_kpis, fig_resumo)

    st.download_button(
        label="📥 Download PDF Executivo",
        data=pdf,
        file_name="relatorio_financeiro_executivo.pdf",
        mime="application/pdf"
    )
