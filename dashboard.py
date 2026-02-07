uploaded_files = st.file_uploader(
    "📤 Carregue um ficheiro Excel por mês",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files is None or len(uploaded_files) == 0:
    st.info("⬆️ Carregue pelo menos um ficheiro Excel")
    st.stop()

# Feedback visual imediato
st.success(f"✅ {len(uploaded_files)} ficheiro(s) carregado(s):")
for f in uploaded_files:
    st.write("•", f.name)
# ================= LEITURA DOS FICHEIROS =================
dfs = []

for file in uploaded_files:
    df_temp = pd.read_excel(file)

    # Normalizar colunas
    df_temp.columns = df_temp.columns.str.strip()

    # Converter datas (mantido como no código original que funcionava)
    df_temp["Data"] = pd.to_datetime(df_temp["Data"], errors="coerce")
    df_temp = df_temp.dropna(subset=["Data"])
    df_temp["Dia"] = df_temp["Data"].dt.day

    # Criar coluna Mês
