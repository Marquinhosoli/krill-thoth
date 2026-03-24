def build_prices(frutas: pd.DataFrame, legumes: pd.DataFrame, order_df: pd.DataFrame) -> bytes:
    # Pega apenas os dados do pedido original
    base_precos = (
        order_df[["Descrição do Produto", "CodigoPedido", "PrecoPedido"]]
        .copy()
        .drop_duplicates(subset=["Descrição do Produto"])
    )

    base_precos["PRODUTO_KEY"] = base_precos["Descrição do Produto"].map(norm_key)

    def make_df(df):
        if df.empty:
            return pd.DataFrame(columns=["CÓDIGO", "PRODUTO", "PREÇO"])

        # Garante a ordem alfabética de A-Z
        produtos = sorted(df.index.tolist(), key=lambda x: str(x).strip().upper())
        linhas = []

        for prod in produtos:
            key = norm_key(prod)
            achou = base_precos[base_precos["PRODUTO_KEY"] == key]

            if not achou.empty:
                row = achou.iloc[0]
                linhas.append(
                    {
                        "CÓDIGO": row["CodigoPedido"],
                        "PRODUTO": prod,
                        "PREÇO": row["PrecoPedido"],
                    }
                )
            else:
                linhas.append(
                    {
                        "CÓDIGO": "",
                        "PRODUTO": prod,
                        "PREÇO": "",
                    }
                )

        return pd.DataFrame(linhas)

    frutas_df = make_df(frutas)
    legumes_df = make_df(legumes)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        frutas_df.to_excel(writer, sheet_name="FRUTAS", index=False)
        legumes_df.to_excel(writer, sheet_name="LEGUMES", index=False)
        
        # Deixa a planilha de preços bonita e formatada (largura das colunas)
        for sheet_name in ["FRUTAS", "LEGUMES"]:
            worksheet = writer.sheets[sheet_name]
            worksheet.column_dimensions['A'].width = 12  # Coluna CÓDIGO
            worksheet.column_dimensions['B'].width = 45  # Coluna PRODUTO
            worksheet.column_dimensions['C'].width = 15  # Coluna PREÇO

    out.seek(0)
    return out.getvalue()
