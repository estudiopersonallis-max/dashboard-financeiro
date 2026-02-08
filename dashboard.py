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
def extrair_periodo(nome_ficheiro):
    return nome_ficheiro.replace(".xlsx", "").upper()

def ler_receitas(ficheiros):
    dfs = []
    for file in ficheiros:
        periodo = extrair_periodo(file.name)
        df = pd.read_excel(file)
        if df.empty:
            continue

        df["Periodo"] = periodo
        df["Nome do cliente"] = df["Nome do cliente"].astype(str).str.strip().str.upper()
        coluna_status = df.columns[2]
        df["Ativo"] = df[coluna_status].astype(str).str.upper().eq("ATIVO")
        df["É Perda"] = df["Perdas"].notna() if "Perdas" in df.columns else False
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Modalidade"] = df.get("Modalidade", "N/A")
        df["Local"] = df.get("Local", "N/A")
        df["Tipo"] = df.get("Tipo", "N/A")
        df["Professor"] = df.get("Professor", "N/A")
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def ler_despesas(ficheiros):
    dfs = []
    for file in ficheiros:
        periodo = extrair_periodo(file.name)
        df = pd.read_excel(file)
        df = df.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df.empty:
            continue

        df["Periodo"] = periodo
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
        df["Classe"] = df["Classe"].astype(str).str.upper().str.strip()
        df["Local"] = df["Local"].astype(str).str.strip()
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= FILTRO DEPÓSITOS =================
if not despesas.empty:
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= REDISTRIBUIÇÃO GERAL (POR PERÍODO) =================
if not despesas.empty and not receitas.empty:
    novas_despesas = []

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
                    novas_despesas.append(nova)
            else:
                novas_despesas.append(row)

    despesas = pd.DataFrame(novas_despesas)

# ================= KPIs COMPARATIVOS =================
st.subheader("📌 KPIs por Período")

kpis = []
for periodo in sorted(set(receitas["Periodo"]).union(set(despesas["Periodo"]))):
    rec = receitas[receitas["Periodo"] == periodo]
    desp = despesas[despesas["Periodo"] == periodo]

    total_receita = rec["Valor"].sum()
    total_despesa = desp["Valor"].sum()
    lucro = total_receita + total_despesa

    kpis.append({
        "Período": periodo,
        "Receita (€)": round(total_receita, 2),
        "Despesa (€)": round(total_despesa, 2),
        "Lucro (€)": round(lucro, 2)
    })

df_kpis = pd.DataFrame(kpis)
st.dataframe(df_kpis, use_container_width=True)

st.divider()

# ================= FUNÇÕES DE GRÁFICO =================
def grafico_bar_comparativo(df, titulo):
    fig, ax = plt.subplots()
    df.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    ax.set_ylabel("€")
    ax.legend(title="Período")
    return fig

def grafico_pizza_periodo(df, titulo):
    figs = {}
    for periodo in df.columns:
        fig, ax = plt.subplots(figsize=(4,4))
        valores = df[periodo].abs()
        ax.pie(
            valores,
            labels=valores.index,
            autopct="%1.1f%%",
            pctdistance=1.15,
            labeldistance=1.3,
            textprops={"fontsize": 7}
        )
        ax.set_title(f"{titulo} – {periodo}")
        figs[periodo] = fig
    return figs

# ================= RECEITAS =================
st.subheader("📌 Receitas – Comparativo")
for cat in ["Modalidade", "Tipo", "Professor", "Local"]:
    if cat in receitas.columns:
        pivot = receitas.pivot_table(
            index=cat,
            columns="Periodo",
            values="Valor",
            aggfunc="sum",
            fill_value=0
        )
        st.markdown(f"**Receitas por {cat}**")
        st.dataframe(pivot)
        st.pyplot(grafico_bar_comparativo(pivot, f"Receitas por {cat}"))

st.divider()

# ================= DESPESAS =================
st.subheader("📌 Despesas – Comparativo")
for cat in ["Classe", "Local"]:
    if cat in despesas.columns:
        pivot = despesas.pivot_table(
            index=cat,
            columns="Periodo",
            values="Valor",
            aggfunc="sum",
            fill_value=0
        )
        st.markdown(f"**Despesas por {cat}**")
        st.dataframe(pivot)
        st.pyplot(grafico_bar_comparativo(pivot, f"Despesas por {cat}"))

# ================= EXPORTAR PPT (LEVE) =================
st.subheader("💾 Exportar PowerPoint Comparativo")

def slide_fig(prs, fig, titulo):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = titulo
    img = BytesIO()
    fig.savefig(img, format="png", bbox_inches="tight")
    img.seek(0)
    slide.shapes.add_picture(img, Inches(1), Inches(1.5), width=Inches(8))

if st.button("🖇️ Gerar PowerPoint Comparativo"):
    prs = Presentation()

    slide_fig(prs, grafico_bar_comparativo(df_kpis.set_index("Período")[["Receita (€)", "Despesa (€)", "Lucro (€)"]],
                                          "Receita x Despesa x Lucro"),
              "Resumo Financeiro")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)

    st.success("PowerPoint gerado com sucesso (leve e comparativo)")
    st.markdown(f"[👉 Abrir PowerPoint]({tmp.name})", unsafe_allow_html=True)
