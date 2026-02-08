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
st.title("📊 Dashboard Financeiro – Análise Executiva")

# ================= UPLOAD =================
st.subheader("📤 Upload de Ficheiros (1 ficheiro = 1 período)")
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

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(
        columns=["Periodo","Valor","Nome do cliente","Modalidade","Tipo","Professor","Local","Ativo"]
    )

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

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(
        columns=["Periodo","Valor","Classe","Local"]
    )

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else ler_receitas([])
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else ler_despesas([])

# ================= FILTRO DEPÓSITOS =================
despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]

# ================= RESULTADO POR PERÍODO =================
periodos = sorted(set(receitas["Periodo"]).union(set(despesas["Periodo"])))
resumo = []

for p in periodos:
    r = receitas[receitas["Periodo"] == p]["Valor"].sum()
    d = despesas[despesas["Periodo"] == p]["Valor"].sum()
    lucro = r + d
    margem = (lucro / r * 100) if r != 0 else 0

    resumo.append({
        "Período": p,
        "Receita (€)": round(r, 2),
        "Despesa (€)": round(d, 2),
        "Lucro (€)": round(lucro, 2),
        "Margem (%)": round(margem, 1)
    })

df_resumo = pd.DataFrame(resumo)

# ================= KPIs =================
st.subheader("📌 KPIs por Período")
st.dataframe(df_resumo, use_container_width=True)

col1, col2, col3, col4 = st.columns(4)
col1.metric("📈 Melhor Mês", df_resumo.loc[df_resumo["Lucro (€)"].idxmax()]["Período"])
col2.metric("📉 Pior Mês", df_resumo.loc[df_resumo["Lucro (€)"].idxmin()]["Período"])
col3.metric("💰 Receita Média", f"€ {df_resumo['Receita (€)'].mean():,.2f}")
col4.metric("🎯 Margem Média", f"{df_resumo['Margem (%)'].mean():.1f} %")

st.divider()

# ================= GRÁFICOS EXECUTIVOS =================
st.subheader("📊 Análise Executiva")

def grafico_linha(df, col, titulo):
    fig, ax = plt.subplots()
    ax.plot(df["Período"], df[col], marker="o")
    ax.set_title(titulo)
    ax.grid(True)
    return fig

def grafico_bar_duplo(df):
    fig, ax = plt.subplots()
    df.set_index("Período")[["Receita (€)", "Despesa (€)"]].plot(kind="bar", ax=ax)
    ax.set_title("Receita vs Despesa por Período")
    ax.set_ylabel("€")
    return fig

fig_lucro = grafico_linha(df_resumo, "Lucro (€)", "Evolução do Lucro")
fig_margem = grafico_linha(df_resumo, "Margem (%)", "Evolução da Margem (%)")
fig_receita_despesa = grafico_bar_duplo(df_resumo)

st.pyplot(fig_receita_despesa)
st.pyplot(fig_lucro)
st.pyplot(fig_margem)

# ================= POWERPOINT =================
st.subheader("💾 Exportar PowerPoint Executivo")

def slide_fig(prs, fig, titulo):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = titulo
    img = BytesIO()
    fig.savefig(img, format="png", bbox_inches="tight")
    img.seek(0)
    slide.shapes.add_picture(img, Inches(1), Inches(1.5), width=Inches(8))

if st.button("🖇️ Gerar PowerPoint"):
    prs = Presentation()
    slide_fig(prs, fig_receita_despesa, "Receita vs Despesa")
    slide_fig(prs, fig_lucro, "Evolução do Lucro")
    slide_fig(prs, fig_margem, "Evolução da Margem")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)

    st.success("PowerPoint executivo gerado com sucesso")
    st.markdown(f"[👉 Abrir PowerPoint]({tmp.name})", unsafe_allow_html=True)
