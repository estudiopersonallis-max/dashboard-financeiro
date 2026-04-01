# ================= IMPORTS =================
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import re
import unicodedata

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

# PPTX
from pptx import Presentation

# ================= CONFIG =================
st.set_page_config(page_title="Dashboard Financeiro PRO", layout="wide")
st.title("📊 Dashboard Financeiro – Nível Consultoria")

# ================= NORMALIZAÇÃO =================
def normalizar(txt):
    if pd.isna(txt): return ""
    txt = str(txt).upper().strip()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII','ignore').decode('ASCII')
    return txt

# ================= DETECTAR MÊS =================
mapa_meses = {"JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,"JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12}

def extrair_mes(nome):
    nome = normalizar(nome)
    match = re.search(r'\\b(0?[1-9]|1[0-2])\\b', nome)
    if match: return int(match.group())
    for k,v in mapa_meses.items():
        if k in nome: return v
    return 99

# ================= LEITURA =================
@st.cache_data(ttl=3600)
def ler_receitas(files):
    dfs=[]
    for f in files:
        df=pd.read_excel(f)
        if df.empty: continue

        periodo=f.name.split(".")[0]
        df["Periodo"]=periodo.upper()
        df["ordem_mes"]=extrair_mes(periodo)

        df["Valor"]=pd.to_numeric(df.get("Valor",0),errors="coerce").fillna(0)
        df["Nome do cliente"]=df.get("Nome do cliente","").apply(normalizar)
        df=df[df["Nome do cliente"]!=""]

        df["Modalidade"]=df.get("Modalidade","N/A").apply(normalizar)
        df["Tipo"]=df.get("Tipo","N/A")
        df["Professor"]=df.get("Professor","N/A")
        df["Local"]=df.get("Local","N/A")

        dfs.append(df)
    return pd.concat(dfs,ignore_index=True) if dfs else pd.DataFrame()

@st.cache_data(ttl=3600)
def ler_despesas(files):
    dfs=[]
    for f in files:
        df=pd.read_excel(f)
        if df.empty: continue

        periodo=f.name.split(".")[0]
        df["Periodo"]=periodo.upper()
        df["ordem_mes"]=extrair_mes(periodo)

        df["Valor"]=pd.to_numeric(df.get("Valor",0),errors="coerce").fillna(0)
        df["Classe"]=df.get("Classe","N/A").apply(normalizar)
        df["Local"]=df.get("Local","N/A")

        dfs.append(df)
    return pd.concat(dfs,ignore_index=True) if dfs else pd.DataFrame()

# ================= UPLOAD =================
st.sidebar.header("📤 Upload")
uploaded_receitas=st.sidebar.file_uploader("Receitas",type=["xlsx"],accept_multiple_files=True)
uploaded_despesas=st.sidebar.file_uploader("Despesas",type=["xlsx"],accept_multiple_files=True)

receitas=ler_receitas(uploaded_receitas) if uploaded_receitas else pd.DataFrame()
despesas=ler_despesas(uploaded_despesas) if uploaded_despesas else pd.DataFrame()

# ================= KPIs =================
receita_total=receitas["Valor"].sum() if not receitas.empty else 0
despesa_total=despesas["Valor"].sum() if not despesas.empty else 0
lucro_total=receita_total+despesa_total

# ================= KPIs EM COLUNAS =================
col1,col2,col3,col4=st.columns(4)
col1.metric("Receita",f"{receita_total:,.0f}€")
col2.metric("Despesa",f"{despesa_total:,.0f}€")
col3.metric("Lucro",f"{lucro_total:,.0f}€")
col4.metric("Margem",f"{(lucro_total/receita_total*100 if receita_total else 0):.1f}%")

# ================= TABS (MANTIDAS) =================
tab1,tab2,tab3=st.tabs(["📊 Visão Geral","💰 Receitas","💸 Despesas"])

# ================= VISÃO GERAL =================
with tab1:
    st.subheader("📈 Receita vs Despesa vs Lucro")
    if not receitas.empty:
        r=receitas.groupby("Periodo")["Valor"].sum()
        d=despesas.groupby("Periodo")["Valor"].sum() if not despesas.empty else r*0
        l=r+d
        st.line_chart(pd.DataFrame({"Receita":r,"Despesa":d,"Lucro":l}))

    st.subheader("🏆 Top Clientes")
    if not receitas.empty:
        top=receitas.groupby("Nome do cliente")["Valor"].sum().sort_values(ascending=False).head(10)
        st.dataframe(top)

    st.subheader("🚨 Alertas")
    if lucro_total<0: st.error("Prejuízo")
    if receita_total>0 and (lucro_total/receita_total)<0.2: st.warning("Margem baixa")

# ================= RECEITAS =================
with tab2:
    for cat in ["Modalidade","Tipo","Professor","Local"]:
        if cat in receitas.columns:
            bloco=receitas.pivot_table(index=cat,columns="Periodo",values="Valor",aggfunc="sum",fill_value=0)
            st.dataframe(bloco)
            fig,ax=plt.subplots()
            bloco.plot(kind="barh",ax=ax)
            st.pyplot(fig)

# ================= DESPESAS =================
with tab3:
    for cat in ["Classe","Local"]:
        if cat in despesas.columns:
            bloco=despesas.pivot_table(index=cat,columns="Periodo",values="Valor",aggfunc="sum",fill_value=0)
            st.dataframe(bloco)
            fig,ax=plt.subplots()
            bloco.plot(kind="barh",ax=ax)
            st.pyplot(fig)

# ================= KPIs AVANÇADOS (CORRIGIDO) =================
st.subheader("🧠 KPIs Avançados")

if not receitas.empty:
    receita_mes=receitas.groupby("Periodo")["Valor"].sum()
    clientes_mes=receitas.groupby("Periodo")["Nome do cliente"].nunique()
    ticket_mensal=(receita_mes/clientes_mes).mean()
else:
    ticket_mensal=0

clientes_unicos=receitas["Nome do cliente"].nunique() if not receitas.empty else 0

cac=abs(despesa_total)/clientes_unicos if clientes_unicos else 0
ltv=ticket_mensal*6

col1,col2,col3,col4=st.columns(4)
col1.metric("🎯 Ticket Mensal",f"{ticket_mensal:,.0f}€")
col2.metric("💸 CAC",f"{cac:,.0f}€")
col3.metric("💰 LTV",f"{ltv:,.0f}€")
col4.metric("⚖️ LTV/CAC",f"{(ltv/cac if cac else 0):.2f}")

# ================= EXPORT =================
def gerar_pdf():
    buffer=BytesIO()
    doc=SimpleDocTemplate(buffer,pagesize=A4)
    styles=getSampleStyleSheet()
    elems=[Paragraph("Dashboard Financeiro",styles['Title'])]
    doc.build(elems)
    buffer.seek(0)
    return buffer

st.download_button("📄 PDF",gerar_pdf(),"relatorio.pdf")

# ================= FOOTER =================
st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
