import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import tempfile
import matplotlib

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= UPLOAD =================
st.subheader("📤 Upload de Ficheiros (cada ficheiro = um período)")
uploaded_receitas = st.file_uploader(
    "Receitas (Excel)",
    type=["xlsx"],
    accept_multiple_files=True
)
uploaded_despesas = st.file_uploader(
    "Despesas (Excel)",
    type=["xlsx"],
    accept_multiple_files=True
)

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

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame(columns=[
        "Periodo","Valor","Nome do cliente","Modalidade",
        "Tipo","Professor","Local","Ativo","É Perda"
    ])

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

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame(columns=["Periodo","Valor","Classe","Local"])

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else ler_receitas([])
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else ler_despesas([])

# ================= FILTRO DEPÓSITOS =================
if not despesas.empty:
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= KPIs =================
st.subheader("📌 KPIs Comparativos")

periodos = sorted(set(receitas["Periodo"]).union(set(despesas["Periodo"])))
kpis = []

for p in periodos:
    r = receitas[receitas["Periodo"] == p]
    d = despesas[despesas["Periodo"] == p]

    receita = r["Valor"].sum()
    despesa = d["Valor"].sum()
    lucro = receita + despesa  # despesas já negativas

    kpis.append({
        "Período": p,
        "Receita (€)": round(receita, 2),
        "Despesa (€)": round(despesa, 2),
        "Lucro (€)": round(lucro, 2)
    })

df_kpis = pd.DataFrame(kpis)
st.dataframe(df_kpis, use_container_width=True)
st.divider()

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

# ================= POWERPOINT =================
st.subheader("💾 Exportar PowerPoint")

def slide_fig(prs, fig, titulo):
    if fig is None:
        return
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = titulo
    img = BytesIO()
    fig.savefig(img, format="png", bbox_inches="tight")
    img.seek(0)
    slide.shapes.add_picture(img, Inches(1), Inches(1.5), width=Inches(8))

if st.button("🖇️ Gerar PowerPoint"):
    prs = Presentation()
    slide_fig(
        prs,
        grafico_bar(
            df_kpis.set_index("Período")[["Receita (€)", "Despesa (€)", "Lucro (€)"]],
            "Resumo Financeiro"
        ),
        "Resumo Financeiro"
    )

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)
    st.success("PowerPoint gerado com sucesso")
    st.markdown(f"[👉 Abrir PowerPoint]({tmp.name})", unsafe_allow_html=True)
