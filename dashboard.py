import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Dashboard Financeiro Comparativo", layout="wide")
st.title("📊 Dashboard Financeiro – Comparação por Período")

# ======================================================
# UPLOAD
# ======================================================
col1, col2 = st.columns(2)

with col1:
    uploaded_receitas = st.file_uploader(
        "📤 Ficheiros de RECEITAS (um ficheiro = um período)",
        type=["xlsx"],
        accept_multiple_files=True
    )

with col2:
    uploaded_despesas = st.file_uploader(
        "📤 Ficheiros de DESPESAS (um ficheiro = um período)",
        type=["xlsx"],
        accept_multiple_files=True
    )

if not uploaded_receitas and not uploaded_despesas:
    st.info("⬆️ Carregue pelo menos um ficheiro de receitas ou despesas")
    st.stop()

# ======================================================
# FUNÇÕES AUXILIARES
# ======================================================
def nome_periodo(file):
    return file.name.replace(".xlsx", "").upper()

def grafico_barra(df_pivot, titulo, prefixo="€ "):
    fig, ax = plt.subplots()
    df_pivot.plot(kind="bar", ax=ax)
    ax.set_title(titulo)
    ax.set_ylabel("Valor")
    ax.tick_params(axis="x", rotation=45)

    for container in ax.containers:
        ax.bar_label(container, fmt=f"{prefixo}%.2f", fontsize=8)

    return fig


def grafico_pizza(serie, titulo):
    serie = serie[serie > 0]
    if serie.empty:
        return None

    n = len(serie)
    fontsize = 10 if n <= 5 else 8 if n <= 10 else 6

    fig, ax = plt.subplots(figsize=(5,5))
    ax.pie(
        serie,
        autopct="%1.1f%%",
        startangle=90,
        pctdistance=1.15,
        textprops={"fontsize": fontsize}
    )
    ax.legend(
        serie.index,
        loc="center left",
        bbox_to_anchor=(1, 0.5),
        fontsize=8
    )
    ax.set_title(titulo)
    ax.axis("equal")
    return fig

# ======================================================
# LEITURA RECEITAS
# ======================================================
receitas = []

for file in uploaded_receitas or []:
    df = pd.read_excel(file)
    df["Periodo"] = nome_periodo(file)
    df["Valor"] = df["Valor"].astype(float)
    df["Nome do cliente"] = df["Nome do cliente"].astype(str).str.upper().str.strip()
    df["Local"] = df["Local"].astype(str).str.upper().str.strip()
    df["Modalidade"] = df["Modalidade"].astype(str).str.upper().str.strip()
    df["Ativo"] = df.iloc[:,2].astype(str).str.upper().eq("ATIVO")
    receitas.append(df)

receitas = pd.concat(receitas, ignore_index=True) if receitas else pd.DataFrame()

# ======================================================
# LEITURA DESPESAS
# ======================================================
despesas = []

for file in uploaded_despesas or []:
    df = pd.read_excel(file)
    df["Periodo"] = nome_periodo(file)
    df["Valor"] = df["Valor"].abs().astype(float)
    df["Classe"] = df["Classe"].astype(str).str.upper().str.strip()
    df["Local"] = df["Local"].astype(str).str.upper().str.strip()
    despesas.append(df)

despesas = pd.concat(despesas, ignore_index=True) if despesas else pd.DataFrame()

# ======================================================
# LIMPEZA DESPESAS
# ======================================================
if not despesas.empty:
    despesas = despesas[despesas["Classe"] != "DEPÓSITOS"]
    despesas = despesas.dropna(subset=["Classe", "Local", "Valor"])

# ======================================================
# REDISTRIBUIR DESPESAS GERAIS
# ======================================================
if not despesas.empty and not receitas.empty:
    ativos_local = receitas[receitas["Ativo"]].groupby("Local")["Nome do cliente"].nunique()
    total_ativos = ativos_local.sum()

    novas = []
    mask_geral = despesas["Local"] == "GERAL"

    for _, row in despesas[mask_geral].iterrows():
        for local, n in ativos_local.items():
            nova = row.copy()
            nova["Local"] = local
            nova["Valor"] = row["Valor"] * n / total_ativos
            novas.append(nova)

    despesas = pd.concat([despesas[~mask_geral], pd.DataFrame(novas)], ignore_index=True)

# ======================================================
# KPIs COMPARATIVOS
# ======================================================
st.subheader("📌 KPIs por Período")

kpis = []

for periodo in sorted(set(receitas["Periodo"]).union(set(despesas["Periodo"]))):
    r = receitas[receitas["Periodo"] == periodo]
    d = despesas[despesas["Periodo"] == periodo]

    total_r = r["Valor"].sum()
    total_d = d["Valor"].sum()
    lucro = total_r - total_d

    kpis.append({
        "Período": periodo,
        "Receitas (€)": total_r,
        "Despesas (€)": total_d,
        "Lucro (€)": lucro
    })

df_kpi = pd.DataFrame(kpis).set_index("Período")
st.dataframe(df_kpi)

# ======================================================
# RECEITAS – COMPARAÇÃO
# ======================================================
if not receitas.empty:
    st.divider()
    st.header("💰 Receitas – Comparativo")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Receita por Modalidade")
        pivot = receitas.pivot_table(values="Valor", index="Modalidade", columns="Periodo", aggfunc="sum").fillna(0)
        st.dataframe(pivot)
        st.pyplot(grafico_barra(pivot, "Receita por Modalidade"))

    with col2:
        for periodo in pivot.columns:
            fig = grafico_pizza(pivot[periodo], f"% Receita por Modalidade – {periodo}")
            if fig:
                st.pyplot(fig)

# ======================================================
# DESPESAS – COMPARAÇÃO
# ======================================================
if not despesas.empty:
    st.divider()
    st.header("💸 Despesas – Comparativo")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Despesa por Classe")
        pivot = despesas.pivot_table(values="Valor", index="Classe", columns="Periodo", aggfunc="sum").fillna(0)
        st.dataframe(pivot)
        st.pyplot(grafico_barra(pivot, "Despesa por Classe"))

    with col2:
        for periodo in pivot.columns:
            fig = grafico_pizza(pivot[periodo], f"% Despesa por Classe – {periodo}")
            if fig:
                st.pyplot(fig)

    st.divider()

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("Despesa por Local")
        pivot = despesas.pivot_table(values="Valor", index="Local", columns="Periodo", aggfunc="sum").fillna(0)
        st.dataframe(pivot)
        st.pyplot(grafico_barra(pivot, "Despesa por Local"))

    with col4:
        for periodo in pivot.columns:
            fig = grafico_pizza(pivot[periodo], f"% Despesa por Local – {periodo}")
            if fig:
                st.pyplot(fig)
