import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import matplotlib

matplotlib.use("Agg")

st.set_page_config(page_title="Dashboard Financeiro", layout="wide")
st.title("📊 Dashboard Financeiro")

# ================= UPLOAD =================
st.subheader("📤 Upload de Ficheiros")
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
def ler_receitas(ficheiros):
    dfs = []
    for file in ficheiros:
        df_temp = pd.read_excel(file)
        if df_temp.empty:
            continue
        df_temp["Nome do cliente"] = df_temp["Nome do cliente"].astype(str).str.strip().str.upper()
        coluna_status = df_temp.columns[2]
        df_temp["Ativo"] = df_temp[coluna_status].astype(str).str.strip().str.upper().eq("ATIVO")
        df_temp["É Perda"] = df_temp["Perdas"].notna() if "Perdas" in df_temp.columns else False
        df_temp["Valor"] = pd.to_numeric(df_temp["Valor"], errors='coerce').fillna(0)
        df_temp["Modalidade"] = df_temp.get("Modalidade", "N/A")
        df_temp["Local"] = df_temp.get("Local", "N/A")
        df_temp["Tipo"] = df_temp.get("Tipo", "N/A")
        df_temp["Professor"] = df_temp.get("Professor", "N/A")
        dfs.append(df_temp)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def ler_despesas(ficheiros):
    dfs = []
    for file in ficheiros:
        df_temp = pd.read_excel(file)
        df_temp = df_temp.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df_temp.empty:
            continue
        df_temp["Nome do cliente"] = df_temp["Descrição da Despesa"].astype(str).str.strip().str.upper()
        df_temp["Valor"] = pd.to_numeric(df_temp["Valor"], errors='coerce').fillna(0)
        df_temp["Classe"] = df_temp["Classe"].astype(str).str.strip().str.upper()
        df_temp["Local"] = df_temp["Local"].astype(str).str.strip()
        df_temp["Ativo"] = True
        df_temp["É Perda"] = False
        dfs.append(df_temp)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= REDISTRIBUIÇÃO =================
if not despesas.empty and not receitas.empty:
    ativos_local = receitas[receitas["Ativo"]].groupby("Local")["Nome do cliente"].nunique()
    geral_mask = despesas["Local"].str.upper() == "GERAL"
    despesas_geral = despesas[geral_mask]
    despesas_nao_geral = despesas[~geral_mask]
    redistribuidas = []
    for _, row in despesas_geral.iterrows():
        total_ativos = ativos_local.sum()
        for loc, n_ativos in ativos_local.items():
            nova_linha = row.copy()
            nova_linha["Valor"] = row["Valor"] * n_ativos / total_ativos
            nova_linha["Local"] = loc
            redistribuidas.append(nova_linha)
    despesas = pd.concat([despesas_nao_geral, pd.DataFrame(redistribuidas)], ignore_index=True) if redistribuidas else despesas_nao_geral

# ================= KPIs =================
clientes_ativos = receitas["Nome do cliente"].nunique() if not receitas.empty else 0
total_receita = receitas["Valor"].sum() if not receitas.empty else 0
perdas = int(receitas["É Perda"].sum()) if not receitas.empty else 0
ticket_medio = total_receita / clientes_ativos if clientes_ativos else 0
total_despesa = despesas["Valor"].sum() if not despesas.empty else 0
lucro_liquido = total_receita + total_despesa  # soma porque despesas são negativas

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("💰 Total Receita", f"€ {total_receita:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("💸 Total Despesa", f"€ {total_despesa:,.2f}")
col5.metric("📈 Lucro Líquido", f"€ {lucro_liquido:,.2f}")

st.divider()

# ================= FUNÇÕES DE GRÁFICO =================
def gerar_grafico_bar(df_grupo, titulo):
    df_grupo = df_grupo[df_grupo > 0]
    if df_grupo.empty:
        return None
    fig, ax = plt.subplots()
    bars = ax.bar(df_grupo.index.astype(str), df_grupo.values)
    ax.set_title(titulo)
    ax.set_xticklabels(df_grupo.index.astype(str), rotation=45, ha="right")
    for bar in bars:
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height(), f"{bar.get_height():,.2f}", ha="center", va="bottom", fontsize=8)
    return fig

def gerar_grafico_pizza(df_grupo, titulo):
    df_grupo = df_grupo[df_grupo > 0]
    if df_grupo.empty:
        return None
    fig, ax = plt.subplots(figsize=(5,5))
    ax.pie(df_grupo, startangle=90, autopct="%1.1f%%", textprops={"fontsize": 8})
    ax.legend(df_grupo.index, title="Legenda", loc="center left", bbox_to_anchor=(1,0.5), fontsize=8)
    ax.set_title(titulo)
    ax.axis("equal")
    return fig

# ================= DASHBOARD =================
st.subheader("📌 Receitas x Despesas")
categorias_receita = ["Modalidade", "Tipo", "Professor", "Local"]
categorias_despesa = ["Classe", "Local"]  # Classe é equivalente à Modalidade

for cat in categorias_receita:
    col_receita, col_despesa = st.columns(2)
    with col_receita:
        st.markdown(f"**Receitas – {cat}**")
        if cat in receitas.columns:
            receita_grupo = receitas.groupby(cat)["Valor"].sum()
            st.dataframe(receita_grupo)
            fig_bar = gerar_grafico_bar(receita_grupo, f"Receitas por {cat}")
            fig_pizza = gerar_grafico_pizza(receita_grupo, f"% Receitas por {cat}")
            if fig_bar: st.pyplot(fig_bar)
            if fig_pizza: st.pyplot(fig_pizza)
    with col_despesa:
        if cat in categorias_despesa and cat in despesas.columns:
            st.markdown(f"**Despesas – {cat}**")
            despesa_grupo = despesas.groupby(cat)["Valor"].sum()
            st.dataframe(despesa_grupo)
            fig_bar = gerar_grafico_bar(despesa_grupo, f"Despesas por {cat}")
            fig_pizza = gerar_grafico_pizza(despesa_grupo, f"% Despesas por {cat}")
            if fig_bar: st.pyplot(fig_bar)
            if fig_pizza: st.pyplot(fig_pizza)

# ================= COMPARATIVO =================
st.subheader("📌 Comparativo Receita x Despesa por Classe/Modalidade")
receita_modalidade = receitas.groupby("Modalidade")["Valor"].sum() if not receitas.empty else pd.Series(dtype=float)
despesa_classe = despesas.groupby("Classe")["Valor"].sum() if not despesas.empty else pd.Series(dtype=float)
comparativo = pd.concat([receita_modalidade, despesa_classe], axis=1).fillna(0)
comparativo.columns = ["Receita", "Despesa"]
comparativo = comparativo.astype(float)
st.dataframe(comparativo)

if not comparativo.empty:
    fig_comparativo, ax = plt.subplots()
    comparativo.plot(kind="bar", ax=ax)
    ax.set_title("Comparativo Receita x Despesa por Classe/Modalidade")
    ax.set_ylabel("€")
    st.pyplot(fig_comparativo)
