st.write("VERSAO 2 TESTE")


import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import re
import unicodedata

# ================= CONFIG =================
st.set_page_config(page_title="Dashboard Financeiro PRO", layout="wide")
st.title("📊 Dashboard Financeiro – Nível Consultoria")

# ================= NORMALIZAÇÃO =================
def normalizar(txt):
    if pd.isna(txt):
        return ""
    txt = str(txt).upper().strip()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    return txt

# ================= DETECTAR MÊS =================
mapa_meses = {
    "JAN":1, "JANEIRO":1,
    "FEV":2, "FEVEREIRO":2,
    "MAR":3, "MARCO":3,
    "ABR":4, "ABRIL":4,
    "MAI":5, "MAIO":5,
    "JUN":6, "JUNHO":6,
    "JUL":7, "JULHO":7,
    "AGO":8, "AGOSTO":8,
    "SET":9, "SETEMBRO":9,
    "OUT":10, "OUTUBRO":10,
    "NOV":11, "NOVEMBRO":11,
    "DEZ":12, "DEZEMBRO":12
}

def extrair_mes(nome):
    nome = normalizar(nome)

    match = re.search(r'\b(0?[1-9]|1[0-2])\b', nome)
    if match:
        return int(match.group())

    for k, v in mapa_meses.items():
        if k in nome:
            return v

    return 99

# ================= LEITURA =================
@st.cache_data
def ler_receitas(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty:
            continue

        periodo_nome = f.name.split(".")[0]
        mes = extrair_mes(periodo_nome)

        df["Periodo"] = periodo_nome.upper()
        df["ordem_mes"] = mes

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)

        # ✅ CORRIGIDO (.apply)
        df["Nome do cliente"] = df.get("Nome do cliente", "").apply(normalizar)
        df = df[df["Nome do cliente"] != ""]

        df["Modalidade"] = df.get("Modalidade", "N/A").apply(normalizar)
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

        periodo_nome = f.name.split(".")[0]
        mes = extrair_mes(periodo_nome)

        df["Periodo"] = periodo_nome.upper()
        df["ordem_mes"] = mes

        df["Valor"] = pd.to_numeric(df.get("Valor", 0), errors="coerce").fillna(0)

        # ✅ CORRIGIDO (.apply)
        df["Classe"] = df.get("Classe", "N/A").apply(normalizar)

        df["Local"] = df.get("Local", "N/A")

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
    despesas = despesas[despesas["Classe"] != "DEPOSITOS"]

# ================= KPIs =================
receita_total = receitas["Valor"].sum() if not receitas.empty else 0
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0
lucro_total = receita_total + despesa_total
margem = (lucro_total / receita_total * 100) if receita_total else 0

receita_media = receitas.groupby("Periodo")["Valor"].sum().mean() if not receitas.empty else 0
despesa_media = despesas.groupby("Periodo")["Valor"].sum().mean() if not despesas.empty else 0
clientes_ativos_media = receitas.groupby("Periodo")["Nome do cliente"].nunique().mean() if not receitas.empty else 0

ticket_medio_receita = receita_media / clientes_ativos_media if clientes_ativos_media else 0
ticket_medio_despesa = abs(despesa_media) / clientes_ativos_media if clientes_ativos_media else 0

# ✅ MAGIC NUMBER CORRIGIDO
magic_number = abs(despesa_media)

st.metric("Receita", f"{receita_total:,.0f}€")
st.metric("Despesa", f"{despesa_total:,.0f}€")
st.metric("Lucro", f"{lucro_total:,.0f}€")
st.metric("Margem", f"{margem:.1f}%")

st.metric("Ticket Médio Receita", f"{ticket_medio_receita:,.0f}€")
st.metric("Ticket Médio Despesa", f"{ticket_medio_despesa:,.0f}€")
st.metric("Magic Number (Break-even mensal)", f"{magic_number:,.0f}€")

# ================= CLIENTES =================
st.subheader("👥 Evolução de Clientes")

if not receitas.empty:
    clientes_por_mes = (
        receitas.groupby(["Periodo", "ordem_mes"])["Nome do cliente"]
        .nunique()
        .reset_index()
        .sort_values("ordem_mes")
    )

    fig, ax = plt.subplots()
    ax.plot(clientes_por_mes["Periodo"], clientes_por_mes["Nome do cliente"], marker="o")

    ax.set_title("Clientes Ativos por Mês")
    ax.set_xlabel("Período")
    ax.set_ylabel("Clientes")

    plt.xticks(rotation=45)
    st.pyplot(fig)

# ================= MODALIDADE =================
st.subheader("🏋️ Clientes por Modalidade")

if not receitas.empty:
    clientes_modalidade = receitas.groupby("Modalidade")["Nome do cliente"].nunique().sort_values(ascending=False)

    st.dataframe(clientes_modalidade)

    fig_mod, ax_mod = plt.subplots()
    clientes_modalidade.plot(kind="barh", ax=ax_mod)
    ax_mod.set_title("Distribuição de Clientes por Modalidade")
    st.pyplot(fig_mod)

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
