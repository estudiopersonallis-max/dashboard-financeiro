import pandas as pd
from pathlib import Path

# ================= CONFIGURAÇÕES =================
PASTA_EXCEL = Path(".")          # onde estão os Excel
PASTA_SAIDA = Path("relatorios")
PASTA_SAIDA.mkdir(exist_ok=True)

# ================= LEITURA DOS FICHEIROS =================
dfs = []

for ficheiro in PASTA_EXCEL.glob("*.xlsx"):
    df = pd.read_excel(ficheiro)

    mes = ficheiro.stem
    df["Mes"] = mes

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df["Dia"] = df["Data"].dt.day

    df["Nome do cliente"] = (
        df["Nome do cliente"].astype(str).str.strip().str.upper()
    )

    coluna_status = df.columns[2]
    df["Ativo"] = (
        df[coluna_status].astype(str).str.strip().str.upper().eq("ATIVO")
    )

    df["É Perda"] = df["Perdas"].notna()

    dfs.append(df)

df = pd.concat(dfs, ignore_index=True)

# ================= AGRUPAMENTO POR MÊS =================
for mes, df_mes in df.groupby("Mes"):

    total_valor = df_mes["Valor"].sum()
    clientes_ativos = df_mes.loc[df_mes["Ativo"], "Nome do cliente"].nunique()
    perdas = int(df_mes["É Perda"].sum())
    ticket_medio = total_valor / clientes_ativos if clientes_ativos > 0 else 0

    valor_modalidade = df_mes.groupby("Modalidade")["Valor"].sum()
    valor_tipo = df_mes.groupby("Tipo")["Valor"].sum()
    valor_professor = df_mes.groupby("Professor")["Valor"].sum()
    valor_local = df_mes.groupby("Local")["Valor"].sum()

    # ================= HTML =================
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>Relatório Financeiro - {mes}</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 40px;
            }}
            h1 {{
                border-bottom: 3px solid #333;
                padding-bottom: 10px;
            }}
            h2 {{
                margin-top: 40px;
                border-bottom: 1px solid #ccc;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin-top: 10px;
            }}
            th, td {{
                border: 1px solid #ccc;
                padding: 8px;
            }}
            th {{
                background-color: #f2f2f2;
            }}
            ul {{
                line-height: 1.8;
            }}
        </style>
    </head>
    <body>

        <h1>Relatório Financeiro</h1>
        <p><b>Mês:</b> {mes}</p>

        <h2>Resumo Executivo</h2>
        <ul>
            <li><b>Valor Total:</b> € {total_valor:,.2f}</li>
            <li><b>Clientes Ativos:</b> {clientes_ativos}</li>
            <li><b>Perdas:</b> {perdas}</li>
            <li><b>Ticket Médio:</b> € {ticket_medio:,.2f}</li>
        </ul>

        <h2>Valor por Modalidade</h2>
        {valor_modalidade.to_frame("Valor (€)").to_html()}

        <h2>Valor por Tipo</h2>
        {valor_tipo.to_frame("Valor (€)").to_html()}

        <h2>Valor por Professor</h2>
        {valor_professor.to_frame("Valor (€)").to_html()}

        <h2>Valor por Local</h2>
        {valor_local.to_frame("Valor (€)").to_html()}

        <p style="margin-top:50px;font-size:12px;color:#666;">
            Relatório gerado automaticamente.
        </p>

    </body>
    </html>
    """

    ficheiro_saida = PASTA_SAIDA / f"Relatorio_{mes}.html"
    ficheiro_saida.write_text(html, encoding="utf-8")

    print(f"✔ Relatório gerado: {ficheiro_saida}")
