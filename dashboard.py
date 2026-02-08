import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import matplotlib

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro – Comparativo por Período")

# ================= UPLOAD =================
st.subheader("📤 Upload de Ficheiros (cada ficheiro = um período)")
uploaded_receitas = st.file_uploader(
    "Carregue ficheiros de RECEITAS (Excel)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="receitas"
)
uploaded_despesas = st.file_uploader(
    "Carregue ficheiros de DESPESAS (Excel)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="despesas"
)

# ================= FUNÇÕES =================
def extrair_periodo(nome):
    return nome.replace(".xlsx", "").upper()

def ler_receitas(ficheiros):
    dfs = []
    for f in ficheiros:
        df = pd.read_excel(f)
        if df.empty:
            continue

        df["Periodo"] = extrair_periodo(f.name)
        df["Nome do cliente"] = df["Nome do cliente"].astype(str).str.upper().str.strip()
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Modalidade"] = df.get("Modalidade", "N/A")
        df["Local"] = df.get("Local", "N/A")
        df["Tipo"] = df.get("Tipo", "N/A")
        df["Professor"] = df.get("Professor", "N/A")

        coluna_status = df.columns[2]
        df["Ativo"] = df[coluna_status].astype(str).str.upper().eq("ATIVO")
        df["É Perda"] = df["Perdas"].notna() if "Perdas" in df.columns else False

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(columns=["Periodo"])

def ler_despesas(ficheiros):
    dfs = []
    for f in ficheiros:
        df = pd.read_excel(f)
        df = df.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df.empty:
            continue

        df["Periodo"] = extrair_periodo(f.name)
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Classe"] = df["Classe"].astype(str).str.upper().str.strip()
        df["Local"] = df["Local"].astype(str).str.strip()

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(columns=["Periodo"])

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame(columns=["Periodo"])
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame(columns=["Periodo"])

# ================= FILTRO DEPÓSITOS =================
if not despesas.empty and "Classe" in despesas.columns:
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= REDISTRIBUIÇÃO GERAL (POR PERÍODO) =================
if not despesas.empty and not receitas.empty:
    novas = []

    for periodo in despesas["Periodo"].unique():
        desp_p = despesas[despesas["Periodo"] == periodo]
        rec_p = receitas[(receitas["Periodo"] == periodo) & (receitas["Ativo"])]

        ativos_local = rec_p.groupby("Local")["Nome do cliente"].nunique()
        total_ativos = ativos_local.sum()

        for _, row in desp_p.iterrows():
            if row["Local"].upper() == "GERAL" and total_ativos > 0:
                for loc, qtd in ativos_local.items():
                    nova = row.copy()
                    nova["Valor"] = row["Valor"] * qtd / total_ativos
                    nova["Local"] = loc
                    novas.append(nova)
            else:
                novas.append(row)

    despesas = pd.DataFrame(novas)

# ================= KPIs =================
st.subheader("📌 KPIs por Período")

periodos = sorted(
    set(receitas["Periodo"].unique()).union(set(despesas["Periodo"].unique()))
)

kpis = []
for p in periodos:
    r = receitas[receitas["Periodo"] == p]
    d = despesas[despesas["Periodo"] == p]

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

st.divider()

# ================= FUNÇÃO GRÁFICO =================
def grafico_bar(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    ax.set_ylabel("€")
    ax.legend(title="Período")
    return fig

# ================= RECEITAS =================
st.subheader("📌 Receitas – Comparativo")
for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
    if cat in receitas.columns and not receitas.empty:
        pivot = receitas.pivot_table(
            index=cat,
            columns="Periodo",
            values="Valor",
            aggfunc="sum",
            fill_value=0
        )
        st.markdown(f"**Receitas por {cat}**")
        st.dataframe(pivot)
        st.pyplot(grafico_bar(pivot, f"Receitas por {cat}"))

# ================= DESPESAS =================
st.subheader("📌 Despesas – Comparativo")
for cat in ["Classe", "Local"]:
    if cat in despesas.columns and not despesas.empty:
        pivot = despesas.pivot_table(
            index=cat,
            columns="Periodo",
            values="Valor",
            aggfunc="sum",
            fill_value=0
        )
        st.markdown(f"**Despesas por {cat}**")
        st.dataframe(pivot)
        st.pyplot(grafico_bar(pivot, f"Despesas por {cat}"))

# ================= EXPORTAR PPT =================
st.subheader("💾 Exportar PowerPoint (leve)")

def slide_fig(prs, fig, titulo):
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
        grafico_bar(df_kpis.set_index("Período")[["Receita (€)", "Despesa (€)", "Lucro (€)"]],
                    "Resumo Financeiro"),
        "Resumo Financeiro"
    )

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)

    st.success("PowerPoint gerado com sucesso")
    st.markdown(f"[👉 Abrir PowerPoint]({tmp.name})", unsafe_allow_html=True)
