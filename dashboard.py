import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import matplotlib

matplotlib.use("Agg")  # Para exportar gráficos

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

# ================= FUNÇÕES DE LEITURA =================
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
        df_temp["Valor"] = df_temp["Valor"].astype(float)
        df_temp["Modalidade"] = df_temp["Modalidade"] if "Modalidade" in df_temp.columns else "N/A"
        df_temp["Local"] = df_temp["Local"] if "Local" in df_temp.columns else "N/A"
        df_temp["Tipo"] = df_temp["Tipo"] if "Tipo" in df_temp.columns else "N/A"
        df_temp["Professor"] = df_temp["Professor"] if "Professor" in df_temp.columns else "N/A"
        dfs.append(df_temp)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def ler_despesas(ficheiros):
    dfs = []
    for file in ficheiros:
        df_temp = pd.read_excel(file)
        # Remove linhas sem Valor ou sem descrição/modalidade
        df_temp = df_temp.dropna(subset=["Valor", "Descrição da Despesa", "Classe"])
        if df_temp.empty:
            continue

        # Mapear colunas
        df_temp["Nome do cliente"] = df_temp["Descrição da Despesa"].astype(str).str.strip().str.upper()
        df_temp["Valor"] = df_temp["Valor"].astype(float)
        df_temp["Modalidade"] = df_temp["Classe"].astype(str).str.strip().str.upper()
        df_temp["Local"] = df_temp["Local"].astype(str).str.strip()
        df_temp["Ativo"] = True
        df_temp["É Perda"] = False

        dfs.append(df_temp)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ================= LEITURA =================
receitas = ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas = ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= REDISTRIBUIR DESPESAS COM LOCAL = 'GERAL' =================
if not despesas.empty and not receitas.empty:
    # Contar ativos por Local a partir das receitas
    ativos_local = receitas[receitas["Ativo"]].groupby("Local")["Nome do cliente"].nunique()
    # Selecionar despesas "Geral"
    geral_mask = despesas["Local"].str.upper() == "GERAL"
    despesas_geral = despesas[geral_mask]
    despesas_nao_geral = despesas[~geral_mask]

    # Redistribuir proporcionalmente
    redistribuidas = []
    for idx, row in despesas_geral.iterrows():
        total_ativos = ativos_local.sum()
        for loc, n_ativos in ativos_local.items():
            nova_linha = row.copy()
            nova_linha["Valor"] = row["Valor"] * n_ativos / total_ativos
            nova_linha["Local"] = loc
            redistribuidas.append(nova_linha)
    if redistribuidas:
        despesas_redistribuidas = pd.DataFrame(redistribuidas)
        despesas = pd.concat([despesas_nao_geral, despesas_redistribuidas], ignore_index=True)
    else:
        despesas = despesas_nao_geral

# ================= KPIs =================
clientes_ativos = receitas.loc[receitas["Ativo"], "Nome do cliente"].nunique() if not receitas.empty else 0
total_receita = receitas["Valor"].sum() if not receitas.empty else 0
perdas = int(receitas["É Perda"].sum()) if not receitas.empty else 0
ticket_medio = total_receita / clientes_ativos if clientes_ativos > 0 else 0
total_despesa = despesas["Valor"].sum() if not despesas.empty else 0
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
    df_grupo = df_grupo.dropna()
    df_grupo = df_grupo[df_grupo > 0]
    if df_grupo.empty:
        return None
    fig, ax = plt.subplots()
    bars = ax.bar(df_grupo.index.astype(str), df_grupo.values)
    ax.set_title(titulo)
    ax.set_xticklabels(df_grupo.index.astype(str), rotation=45, ha="right")
    for bar in bars:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f"{bar.get_height():,.2f}", ha="center", va="bottom", fontsize=8)
    return fig

def gerar_grafico_pizza(df_grupo, titulo):
    df_grupo = df_grupo.dropna()
    df_grupo = df_grupo[df_grupo > 0]
    if df_grupo.empty:
        return None
    fig, ax = plt.subplots(figsize=(5,5))
    ax.pie(df_grupo, startangle=90, autopct="%1.1f%%", textprops={"fontsize": 8})
    ax.legend(df_grupo.index, title="Legenda", loc="center left", bbox_to_anchor=(1,0.5), fontsize=8)
    ax.set_title(titulo)
    ax.axis("equal")
    return fig

# ================= DASHBOARD LADO A LADO =================
st.subheader("📌 Receitas x Despesas")

categorias_receita = ["Modalidade", "Tipo", "Professor", "Local"]
categorias_despesa = ["Modalidade", "Local"]

for cat in categorias_receita:
    col_receita, col_despesa = st.columns(2)
    with col_receita:
        st.markdown(f"**Receitas – {cat}**")
        if cat in receitas.columns:
            receita_grupo = receitas.groupby(cat)["Valor"].sum()
            st.dataframe(receita_grupo)
            fig_receita_bar = gerar_grafico_bar(receita_grupo, f"Receitas por {cat}")
            fig_receita_pizza = gerar_grafico_pizza(receita_grupo, f"% Receitas por {cat}")
            if fig_receita_bar: st.pyplot(fig_receita_bar)
            if fig_receita_pizza: st.pyplot(fig_receita_pizza)
    with col_despesa:
        if cat in categorias_despesa and cat in despesas.columns:
            st.markdown(f"**Despesas – {cat}**")
            despesa_grupo = despesas.groupby(cat)["Valor"].sum()
            st.dataframe(despesa_grupo)
            fig_despesa_bar = gerar_grafico_bar(despesa_grupo, f"Despesas por {cat}")
            fig_despesa_pizza = gerar_grafico_pizza(despesa_grupo, f"% Despesas por {cat}")
            if fig_despesa_bar: st.pyplot(fig_despesa_bar)
            if fig_despesa_pizza: st.pyplot(fig_despesa_pizza)

# ================= COMPARATIVO =================
st.subheader("📌 Comparativo Receita x Despesa por Modalidade")
receita_modalidade = receitas.groupby("Modalidade")["Valor"].sum() if not receitas.empty else pd.Series()
despesa_modalidade = despesas.groupby("Modalidade")["Valor"].sum() if not despesas.empty else pd.Series()
comparativo = pd.concat([receita_modalidade, despesa_modalidade], axis=1).fillna(0)
comparativo.columns = ["Receita", "Despesa"]
st.dataframe(comparativo)

fig_comparativo, ax = plt.subplots()
comparativo.plot(kind="bar", ax=ax)
ax.set_title("Comparativo Receita x Despesa por Modalidade")
ax.set_ylabel("€")
st.pyplot(fig_comparativo)

# ================= EXPORTAR POWERPOINT =================
st.subheader("💾 Exportar para PowerPoint")

def adicionar_figura_slide(prs, fig, titulo):
    if fig is None:
        return
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

if st.button("🖇️ Gerar PowerPoint Automático"):
    prs = Presentation()
    for cat in categorias_receita:
        if cat in receitas.columns:
            receita_grupo = receitas.groupby(cat)["Valor"].sum()
            if cat in categorias_despesa and cat in despesas.columns:
                despesa_grupo = despesas.groupby(cat)["Valor"].sum()
            else:
                despesa_grupo = None
            adicionar_figura_slide(prs, gerar_grafico_bar(receita_grupo, f"Receitas por {cat}"), f"Receitas por {cat}")
            adicionar_figura_slide(prs, gerar_grafico_pizza(receita_grupo, f"% Receitas por {cat}"), f"% Receitas por {cat}")
            adicionar_figura_slide(prs, gerar_grafico_bar(despesa_grupo, f"Despesas por {cat}"), f"Despesas por {cat}")
            adicionar_figura_slide(prs, gerar_grafico_pizza(despesa_grupo, f"% Despesas por {cat}"), f"% Despesas por {cat}")
            adicionar_tabela_slide(prs, receita_grupo.to_frame("Valor"), f"Receitas por {cat}")
            if despesa_grupo is not None:
                adicionar_tabela_slide(prs, despesa_grupo.to_frame("Valor"), f"Despesas por {cat}")

    adicionar_figura_slide(prs, fig_comparativo, "Comparativo Receita x Despesa")
    adicionar_tabela_slide(prs, comparativo, "Comparativo Receita x Despesa")

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp_file.name)
    st.success("PowerPoint gerado com sucesso")
    st.markdown(f"[👉 Abrir PowerPoint]({tmp_file.name})", unsafe_allow_html=True)
