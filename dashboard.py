import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import matplotlib

matplotlib.use("Agg")  # Necessário para exportar gráficos

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

if not uploaded_receitas and not uploaded_despesas:
    st.info("⬆️ Carregue pelo menos um ficheiro de receitas ou despesas para iniciar o dashboard")
    st.stop()

# ================= FUNÇÃO DE LEITURA =================
def ler_receitas(ficheiros):
    dfs = []
    for file in ficheiros:
        df_temp = pd.read_excel(file)
        mes_ficheiro = file.name.replace(".xlsx", "")
        df_temp["Mes"] = mes_ficheiro
        df_temp["Data"] = pd.to_datetime(df_temp["Data"])
        df_temp["Dia"] = df_temp["Data"].dt.day
        df_temp["Ano"] = df_temp["Data"].dt.year
        df_temp["Trimestre"] = df_temp["Data"].dt.to_period("Q").astype(str)
        df_temp["Nome do cliente"] = df_temp["Nome do cliente"].astype(str).str.strip().str.upper()
        coluna_status = df_temp.columns[2]
        df_temp["Ativo"] = df_temp[coluna_status].astype(str).str.strip().str.upper().eq("ATIVO")
        df_temp["É Perda"] = df_temp["Perdas"].notna()
        dfs.append(df_temp)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def ler_despesas(ficheiros):
    dfs = []
    for file in ficheiros:
        df_temp = pd.read_excel(file)
        mes_ficheiro = file.name.replace(".xlsx", "")
        df_temp["Mes"] = mes_ficheiro
        df_temp["Data"] = pd.to_datetime(df_temp["Data"])
        df_temp["Dia"] = df_temp["Data"].dt.day
        df_temp["Ano"] = df_temp["Data"].dt.year
        df_temp["Trimestre"] = df_temp["Data"].dt.to_period("Q").astype(str)
        df_temp["Nome do cliente"] = df_temp["Descrição"].astype(str).str.strip().str.upper()
        df_temp["Valor"] = df_temp["Valor"].astype(float)
        df_temp["Modalidade"] = df_temp["Classe"].astype(str).str.strip().str.upper()
        df_temp["Ativo"] = True
        df_temp["É Perda"] = False
        dfs.append(df_temp)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= LEITURA =================
df_receitas = ler_receitas(uploaded_receitas)
df_despesas = ler_despesas(uploaded_despesas)

# Concatenar para análise conjunta
df = pd.concat([df_receitas, df_despesas], ignore_index=True)
df["Tipo Registro"] = ["Receita"]*len(df_receitas) + ["Despesa"]*len(df_despesas)

# ================= FILTRO =================
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
receitas = df_filtro[df_filtro["Tipo Registro"]=="Receita"]
despesas = df_filtro[df_filtro["Tipo Registro"]=="Despesa"]

clientes_ativos = receitas.loc[receitas["Ativo"], "Nome do cliente"].nunique()
total_receita = receitas["Valor"].sum()
perdas = int(receitas["É Perda"].sum())
ticket_medio = total_receita / clientes_ativos if clientes_ativos > 0 else 0
total_despesa = despesas["Valor"].sum()
lucro_liquido = total_receita - total_despesa

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("💰 Total Receita", f"€ {total_receita:,.2f}")
col2.metric("👥 Clientes Ativos", clientes_ativos)
col3.metric("❌ Perdas", perdas)
col4.metric("💸 Total Despesa", f"€ {total_despesa:,.2f}")
col5.metric("📈 Lucro Líquido", f"€ {lucro_liquido:,.2f}")

st.divider()

# ================= FUNÇÕES DE GRÁFICOS =================
def gerar_grafico_bar(df_grupo, titulo):
    fig, ax = plt.subplots()
    bars = ax.bar(df_grupo.index.astype(str), df_grupo.values)
    ax.set_title(titulo)
    ax.set_xticklabels(df_grupo.index.astype(str), rotation=45, ha="right")
    for bar in bars:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f"{bar.get_height():,.2f}", ha="center", va="bottom", fontsize=8)
    st.pyplot(fig)
    return fig

def gerar_grafico_pizza(df_grupo, titulo):
    fig, ax = plt.subplots(figsize=(5,5))
    ax.pie(df_grupo, startangle=90, autopct="%1.1f%%", textprops={"fontsize": 8})
    ax.legend(df_grupo.index, title="Legenda", loc="center left", bbox_to_anchor=(1,0.5), fontsize=8)
    ax.set_title(titulo)
    ax.axis("equal")
    st.pyplot(fig)
    return fig

# ================= DASHBOARD LADO A LADO =================
st.subheader("📌 Receitas x Despesas por Modalidade")
col_receita, col_despesa = st.columns(2)

with col_receita:
    st.markdown("**Receitas**")
    if "Modalidade" in receitas.columns:
        receita_modalidade = receitas.groupby("Modalidade")["Valor"].sum()
        st.dataframe(receita_modalidade)
        fig_receita_modalidade = gerar_grafico_bar(receita_modalidade, "Receitas por Modalidade")
        fig_receita_pizza = gerar_grafico_pizza(receita_modalidade, "% Receitas por Modalidade")

with col_despesa:
    st.markdown("**Despesas**")
    if "Modalidade" in despesas.columns:
        despesa_modalidade = despesas.groupby("Modalidade")["Valor"].sum()
        st.dataframe(despesa_modalidade)
        fig_despesa_modalidade = gerar_grafico_bar(despesa_modalidade, "Despesas por Modalidade")
        fig_despesa_pizza = gerar_grafico_pizza(despesa_modalidade, "% Despesas por Modalidade")

# ================= COMPARATIVO =================
st.subheader("📌 Comparativo Receita x Despesa por Modalidade")
comparativo = pd.concat([receita_modalidade, despesa_modalidade], axis=1).fillna(0)
comparativo.columns = ["Receita", "Despesa"]
st.dataframe(comparativo)

fig, ax = plt.subplots()
comparativo.plot(kind="bar", ax=ax)
ax.set_title("Comparativo Receita x Despesa por Modalidade")
ax.set_ylabel("€")
st.pyplot(fig)
fig_comparativo = fig

# ================= EXPORTAR POWERPOINT =================
st.subheader("💾 Exportar para PowerPoint")

def adicionar_figura_slide(prs, fig, titulo):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = titulo
    img_stream = BytesIO()
    fig.savefig(img_stream, format='png', bbox_inches='tight')
    img_stream.seek(0)
    slide.shapes.add_picture(img_stream, Inches(1), Inches(1.5), width=Inches(8), height=Inches(4.5))

def adicionar_tabela_slide(prs, df, titulo):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = titulo
    rows, cols = df.shape
    table = slide.shapes.add_table(rows+1, cols, Inches(1), Inches(1.5), Inches(8), Inches(4.5)).table
    for j, col_name in enumerate(df.columns):
        table.cell(0, j).text = str(col_name)
    for i in range(rows):
        for j in range(cols):
            table.cell(i+1, j).text = str(df.iloc[i, j])

if st.button("🖇️ Gerar PowerPoint"):
    prs = Presentation()
    # Slides gráficos
    adicionar_figura_slide(prs, fig_receita_modalidade, "Receitas por Modalidade")
    adicionar_figura_slide(prs, fig_receita_pizza, "% Receitas por Modalidade")
    adicionar_figura_slide(prs, fig_despesa_modalidade, "Despesas por Modalidade")
    adicionar_figura_slide(prs, fig_despesa_pizza, "% Despesas por Modalidade")
    adicionar_figura_slide(prs, fig_comparativo, "Comparativo Receita x Despesa")
    # Slides tabelas
    adicionar_tabela_slide(prs, receita_modalidade.to_frame("Valor"), "Receitas por Modalidade")
    adicionar_tabela_slide(prs, despesa_modalidade.to_frame("Valor"), "Despesas por Modalidade")
    adicionar_tabela_slide(prs, comparativo, "Comparativo Receita x Despesa")
    
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp_file.name)
    st.success("PowerPoint gerado com sucesso")
    st.markdown(f"[👉 Abrir PowerPoint]({tmp_file.name})", unsafe_allow_html=True)
