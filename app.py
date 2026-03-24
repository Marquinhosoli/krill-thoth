import pandas as pd
from io import BytesIO
import re
import unicodedata

def normalizar_texto(txt):
    txt = "" if txt is None else str(txt)
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"\s+", " ", txt).strip().upper()
    return txt

def para_numero(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d\.-]", "", s)
    try:
        return float(s)
    except:
        return 0.0

def extrair_numero_loja(loja):
    m = re.search(r"(\d+)", str(loja))
    return int(m.group(1)) if m else None

def nome_coluna_krill(numero):
    return f"KRILL {numero}"

def detectar_coluna(df, candidatos):
    mapa = {normalizar_texto(c): c for c in df.columns}
    for cand in candidatos:
        n = normalizar_texto(cand)
        if n in mapa:
            return mapa[n]
    return None

def classificar_categoria(nome_produto, frutas_modelo, legumes_modelo):
    nome_norm = normalizar_texto(nome_produto)

    if nome_norm in frutas_modelo:
        return "FRUTAS"
    if nome_norm in legumes_modelo:
        return "LEGUMES"

    chaves_legumes = [
        "ABOBRINHA","ABOBORA","AIPIM","ALFACE","ALHO","BATATA","BATATA DOCE",
        "BERINJELA","BETERRABA","BROCOLIS","CEBOLA","CENOURA","CHUCHU","COUVE",
        "COUVE FLOR","ERVILHA","ESPINAFRE","INHAME","JILO","MANDIOCA",
        "MANDIOQUINHA","MAXIXE","MILHO","MILHO VERDE","PEPINO","PIMENTA",
        "PIMENTAO","QUIABO","REPOLHO","RUCULA","SALSINHA","TOMATE","VAGEM"
    ]

    chaves_frutas = [
        "ABACATE","ABACAXI","ACEROLA","AMEIXA","AMORA","BANANA","CAJU","CAQUI",
        "CARAMBOLA","COCO","FIGO","FRAMBOESA","GOIABA","KIWI","LARANJA","LIMAO",
        "MACA","MAMAO","MANGA","MELANCIA","MELAO","MEXERICA","MIRTILO","MORANGO",
        "NECTARINA","PERA","PESSEGO","PITAYA","ROMA","TANGERINA","UVA"
    ]

    if any(k in nome_norm for k in chaves_legumes):
        return "LEGUMES"
    if any(k in nome_norm for k in chaves_frutas):
        return "FRUTAS"

    return "LEGUMES"

def gerar_planilha_precos(pedido_df, modelo_frutas_df, modelo_legumes_df):
    col_produto = detectar_coluna(pedido_df, ["Produto", "Descrição do Produto", "Descricao do Produto", "Item"])
    col_codigo = detectar_coluna(pedido_df, ["Código", "Codigo", "Cod"])
    col_preco = detectar_coluna(pedido_df, ["Preço", "Preco", "GTIN/PLU Unitário", "Valor Unitário", "Valor"])
    col_loja = detectar_coluna(pedido_df, ["Loja", "Loja Destino", "Destino"])

    if not col_produto:
        raise ValueError("Não encontrei a coluna de produto no pedido.")
    if not col_loja:
        raise ValueError("Não encontrei a coluna de loja no pedido.")

    frutas_modelo = {}
    legumes_modelo = {}

    for _, row in modelo_frutas_df.iterrows():
        nome = str(row.iloc[0]).strip()
        if nome:
            frutas_modelo[normalizar_texto(nome)] = nome

    for _, row in modelo_legumes_df.iterrows():
        nome = str(row.iloc[0]).strip()
        if nome:
            legumes_modelo[normalizar_texto(nome)] = nome

    frutas = {}
    legumes = {}

    for _, row in pedido_df.iterrows():
        produto_original = str(row[col_produto]).strip() if pd.notna(row[col_produto]) else ""
        if not produto_original:
            continue

        loja = row[col_loja]
        numero_loja = extrair_numero_loja(loja)
        if numero_loja is None:
            continue

        codigo = ""
        if col_codigo and pd.notna(row[col_codigo]):
            codigo = str(row[col_codigo]).strip()

        preco = 0.0
        if col_preco and pd.notna(row[col_preco]):
            preco = para_numero(row[col_preco])

        categoria = classificar_categoria(produto_original, frutas_modelo, legumes_modelo)
        nome_norm = normalizar_texto(produto_original)

        if categoria == "FRUTAS":
            nome_final = frutas_modelo.get(nome_norm, produto_original)
            if nome_final not in frutas:
                frutas[nome_final] = {
                    "CODIGO": codigo,
                    "PRODUTO": nome_final,
                    "PRECO": preco
                }
            else:
                if not frutas[nome_final]["CODIGO"] and codigo:
                    frutas[nome_final]["CODIGO"] = codigo
                if frutas[nome_final]["PRECO"] == 0 and preco:
                    frutas[nome_final]["PRECO"] = preco
        else:
            nome_final = legumes_modelo.get(nome_norm, produto_original)
            if nome_final not in legumes:
                legumes[nome_final] = {
                    "CODIGO": codigo,
                    "PRODUTO": nome_final,
                    "PRECO": preco
                }
            else:
                if not legumes[nome_final]["CODIGO"] and codigo:
                    legumes[nome_final]["CODIGO"] = codigo
                if legumes[nome_final]["PRECO"] == 0 and preco:
                    legumes[nome_final]["PRECO"] = preco

    df_frutas = pd.DataFrame(list(frutas.values())).sort_values("PRODUTO")
    df_legumes = pd.DataFrame(list(legumes.values())).sort_values("PRODUTO")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_frutas.to_excel(writer, index=False, sheet_name="FRUTAS")
        df_legumes.to_excel(writer, index=False, sheet_name="LEGUMES")

        for aba in ["FRUTAS", "LEGUMES"]:
            ws = writer.book[aba]
            for row in ws.iter_rows(min_row=2):
                row[2].number_format = 'R$ #,##0.00'

    output.seek(0)
    return output
