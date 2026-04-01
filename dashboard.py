# ================= KPIs AVANÇADOS =================
st.markdown("## 🧠 KPIs Avançados")

col1, col2, col3, col4 = st.columns(4)

# Ticket médio
clientes_unicos = receitas["Nome do cliente"].nunique() if not receitas.empty else 0
ticket_medio = receita_total / clientes_unicos if clientes_unicos else 0

# CAC (proxy)
clientes_media = receitas.groupby("Periodo")["Nome do cliente"].nunique().mean() if not receitas.empty else 0
despesa_media = despesas.groupby("Periodo")["Valor"].sum().mean() if not despesas.empty else 0
cac = abs(despesa_media) / clientes_media if clientes_media else 0

# LTV (real se tiver data)
if "Data Inicio" in receitas.columns:
    hoje = pd.Timestamp.today()
    receitas["meses_ativos"] = ((hoje - receitas["Data Inicio"]).dt.days / 30).fillna(0)
    tempo_medio = receitas.groupby("Nome do cliente")["meses_ativos"].max().mean()
else:
    tempo_medio = 6  # fallback

ltv = ticket_medio * tempo_medio
ltv_cac = ltv / cac if cac else 0

with col1:
    st.metric("🎯 Ticket Médio", f"{ticket_medio:,.0f}€")

with col2:
    st.metric("💸 CAC", f"{cac:,.0f}€")

with col3:
    st.metric("💰 LTV", f"{ltv:,.0f}€")

with col4:
    st.metric("⚖️ LTV/CAC", f"{ltv_cac:.2f}")


# ================= EVOLUÇÃO FINANCEIRA =================
st.markdown("## 📈 Evolução Financeira")

if not receitas.empty and not despesas.empty:
    receita_mes = receitas.groupby(["Periodo", "ordem_mes"])["Valor"].sum().reset_index()
    despesa_mes = despesas.groupby(["Periodo", "ordem_mes"])["Valor"].sum().reset_index()

    df_merge = receita_mes.merge(despesa_mes, on=["Periodo", "ordem_mes"], how="left", suffixes=("_rec", "_des"))
    df_merge["Lucro"] = df_merge["Valor_rec"] + df_merge["Valor_des"]
    df_merge = df_merge.sort_values("ordem_mes")

    fig, ax = plt.subplots()
    ax.plot(df_merge["Periodo"], df_merge["Valor_rec"], marker="o", label="Receita")
    ax.plot(df_merge["Periodo"], df_merge["Valor_des"], marker="o", label="Despesa")
    ax.plot(df_merge["Periodo"], df_merge["Lucro"], marker="o", label="Lucro")
    ax.legend()
    plt.xticks(rotation=45)

    st.pyplot(fig)


# ================= TOP CLIENTES =================
st.markdown("## 🏆 Top Clientes")

if not receitas.empty:
    top_clientes = (
        receitas.groupby("Nome do cliente")["Valor"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )

    st.dataframe(top_clientes)

    fig_top, ax_top = plt.subplots()
    top_clientes.sort_values().plot(kind="barh", ax=ax_top)
    st.pyplot(fig_top)


# ================= ALERTAS AUTOMÁTICOS =================
st.markdown("## 🚨 Alertas Automáticos")

alertas = []

if margem < 10:
    alertas.append("⚠️ Margem baixa (<10%)")

if ltv_cac < 3:
    alertas.append("⚠️ LTV/CAC abaixo do ideal (<3)")

if not receitas.empty:
    crescimento = receita_mes["Valor"].pct_change().mean()
    if crescimento < 0:
        alertas.append("📉 Receita em queda")

if clientes_unicos < 10:
    alertas.append("⚠️ Base de clientes pequena")

if alertas:
    for a in alertas:
        st.warning(a)
else:
    st.success("✅ Nenhum alerta crítico")
