import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import re
import unicodedata

st.set_page_config(page_title="KRILL → THOTH", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_MODELO_FRUTAS = BASE_DIR / "KRILL_FRUTAS_Branco.xlsx"
ARQUIVO_MODELO_LEGUMES = BASE_DIR / "KRILL_LEGUMES_Branco.xlsx"

st.title("KRILL → THOTH")
st.caption("Modelos fixos no sistema. Mapeamento por número da loja. Código e preço saem do pedido original.")

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

def ajustar_largura(ws, larguras):
    from openpyxl.utils import get_column_letter
    for i, largura in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura

def carregar_modelos_fixos():
    if not ARQUIVO_MODELO_FRUTAS.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_MODELO_FRUTAS.name}")
    if not ARQUIVO_MODELO_LEGUMES.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_MODELO_LEGUMES.name}")

    modelo_frutas_df = pd.read_excel(ARQUIVO_MODELO_FRUTAS)
    modelo_legumes_df = pd.read_excel(ARQUIVO_MODELO_LEGUMES)
    return modelo_frutas_df, modelo_legumes_df

def primeira_coluna_nome(df):
    return df.columns[0]

def mapear_modelo_por_nome(df_modelo):
    col = primeira_coluna_nome(df_modelo)
    mapa = {}
    ordem = []
    for _, row in df_modelo.iterrows():
        nome = str(row[col]).strip() if pd.notna(row[col]) else ""
        if nome:
            mapa[normalizar_texto(nome)] = nome
            ordem.append(nome)
    return mapa, ordem

def classificar_categoria(nome_produto, frutas_modelo, legumes_modelo):
    nome_norm = normalizar_texto(nome_produto)

    if nome_norm in frutas_modelo:
        return "FRUTAS"
    if nome_norm in legumes_modelo:
        return "LEGUMES"

    chaves_legumes = [
        "ABOBRINHA","ABOBORA","AIPIM","ALFACE","ALHO","BATATA","BATATA DOCE",
        "BERINJELA","BETERRABA","BROCOLIS","BROCOLIS NINJA","BROCOLIS RAMOSO",
        "CEBOLA","CENOURA","CHUCHU","COUVE","COUVE FLOR","ERVILHA","ESPINAFRE",
        "INHAME","JILO","MANDIOCA","MANDIOQUINHA","MAXIXE","MILHO","MILHO VERDE",
        "PEPINO","PEPINO JAPONES","PIMENTA","PIMENTAO","QUIABO","REPOLHO",
        "RUCULA","SALSINHA","TOMATE","TOMATE CEREJA","TOMATE GRAPE","VAGEM"
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

def pontuar_linha_cabecalho(linha):
    score = 0
    textos = [normalizar_texto(v) for v in linha if pd.notna(v)]
    for t in textos:
        if "PRODUTO" in t or "DESCRICAO" in t:
            score += 3
        if "QTDE" in t or "QTD" in t or "QUANTIDADE" in t:
            score += 2
        if "LOJA" in t or "DESTINO" in t or "FILIAL" in t:
            score += 2
        if "COD" in t or "CODIGO" in t:
            score += 1
        if "PRECO" in t or "VALOR" in t or "GTIN" in t or "PLU" in t:
            score += 1
    return score

def ler_pedido_inteligente(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    melhor_df = None
    melhor_sheet = None
    melhor_header = None
    melhor_score = -1

    for sheet in xls.sheet_names:
        bruto = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
        limite = min(15, len(bruto))

        for header_row in range(limite):
            linha = bruto.iloc[header_row].tolist()
            score = pontuar_linha_cabecalho(linha)

            if score > melhor_score:
                try:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row)
                    df = df.dropna(axis=1, how="all")
                    df = df.dropna(axis=0, how="all")
                    melhor_df = df
                    melhor_sheet = sheet
                    melhor_header = header_row
                    melhor_score = score
                except:
                    pass

    if melhor_df is None:
        raise ValueError("Não foi possível ler o pedido.")

    return melhor_df, melhor_sheet, melhor_header

def detectar_campos_pedido(df):
    col_produto = detectar_coluna(df, [
        "Descrição do Produto", "Descricao do Produto", "Produto", "Item", "Descrição", "Descricao"
    ])
    col_qtd = detectar_coluna(df, [
        "Qtde.", "Qtde", "Quantidade", "Qtd", "Caixas", "Qtd. Pedido"
    ])
    col_loja = detectar_coluna(df, [
        "Loja", "Loja Destino", "Destino", "Filial", "Numero Loja"
    ])
    col_codigo = detectar_coluna(df, [
        "Código", "Codigo", "Cod", "Código do Produto", "Codigo do Produto"
    ])
    col_preco = detectar_coluna(df, [
        "Preço", "Preco", "GTIN/PLU Unitário", "GTIN/PLU Unitario", "Valor Unitário",
        "Valor Unitario", "Valor", "Preço Unitário", "Preco Unitario"
    ])

    return {
        "produto": col_produto,
        "qtd": col_qtd,
        "loja": col_loja,
        "codigo": col_codigo,
        "preco": col_preco,
    }

def gerar_planilhas_principais(pedido_df, modelo_frutas_df, modelo_legumes_df):
    campos = detectar_campos_pedido(pedido_df)

    if not campos["produto"]:
        raise ValueError("Não encontrei a coluna de produto no pedido.")
    if not campos["qtd"]:
        raise ValueError("Não encontrei a coluna de quantidade no pedido.")
    if not campos["loja"]:
        raise ValueError("Não encontrei a coluna de loja no pedido.")

    mapa_frutas, ordem_frutas = mapear_modelo_por_nome(modelo_frutas_df)
    mapa_legumes, ordem_legumes = mapear_modelo_por_nome(modelo_legumes_df)

    lojas_encontradas = set()
    frutas_agrupadas = {}
    legumes_agrupadas = {}

    for _, row in pedido_df.iterrows():
        produto_original = str(row[campos["produto"]]).strip() if pd.notna(row[campos["produto"]]) else ""
        if not produto_original:
            continue

        qtd = para_numero(row[campos["qtd"]])
        if qtd == 0:
            continue

        loja = row[campos["loja"]]
        numero_loja = extrair_numero_loja(loja)
        if numero_loja is None:
            continue

        col_loja = nome_coluna_krill(numero_loja)
        lojas_encontradas.add(col_loja)

        categoria = classificar_categoria(produto_original, mapa_frutas, mapa_legumes)
        nome_norm = normalizar_texto(produto_original)

        if categoria == "FRUTAS":
            nome_final = mapa_frutas.get(nome_norm, produto_original)
            if nome_final not in frutas_agrupadas:
                frutas_agrupadas[nome_final] = {}
            frutas_agrupadas[nome_final][col_loja] = frutas_agrupadas[nome_final].get(col_loja, 0) + qtd
        else:
            nome_final = mapa_legumes.get(nome_norm, produto_original)
            if nome_final not in legumes_agrupadas:
                legumes_agrupadas[nome_final] = {}
            legumes_agrupadas[nome_final][col_loja] = legumes_agrupadas[nome_final].get(col_loja, 0) + qtd

    lojas_ordenadas = sorted(
        list(lojas_encontradas),
        key=lambda x: int(re.search(r"(\d+)", x).group(1))
    )

    def montar_df(agrupado, ordem_modelo):
        nomes_existentes = set(agrupado.keys())
        nomes_fora_modelo = [n for n in agrupado.keys() if n not in ordem_modelo]
        ordem_final = list(ordem_modelo) + sorted(nomes_fora_modelo, key=lambda x: normalizar_texto(x))

        linhas = []
        for nome in ordem_final:
            if nome in nomes_existentes or nome in ordem_modelo:
                linha = {"PRODUTO": nome}
                for loja in lojas_ordenadas:
                    linha[loja] = agrupado.get(nome, {}).get(loja, 0)
                linhas.append(linha)

        return pd.DataFrame(linhas)

    df_frutas = montar_df(frutas_agrupadas, ordem_frutas)
    df_legumes = montar_df(legumes_agrupadas, ordem_legumes)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_frutas.to_excel(writer, index=False, sheet_name="FRUTAS")
        df_legumes.to_excel(writer, index=False, sheet_name="LEGUMES")

        wsf = writer.book["FRUTAS"]
        wsl = writer.book["LEGUMES"]

        ajustar_largura(wsf, [42] + [12] * (len(df_frutas.columns) - 1))
        ajustar_largura(wsl, [42] + [12] * (len(df_legumes.columns) - 1))

    output.seek(0)
    return output, df_frutas, df_legumes, lojas_ordenadas

def gerar_planilha_precos(pedido_df, modelo_frutas_df, modelo_legumes_df):
    campos = detectar_campos_pedido(pedido_df)

    if not campos["produto"]:
        raise ValueError("Não encontrei a coluna de produto no pedido.")

    mapa_frutas, _ = mapear_modelo_por_nome(modelo_frutas_df)
    mapa_legumes, _ = mapear_modelo_por_nome(modelo_legumes_df)

    frutas = {}
    legumes = {}

    for _, row in pedido_df.iterrows():
        produto_original = str(row[campos["produto"]]).strip() if pd.notna(row[campos["produto"]]) else ""
        if not produto_original:
            continue

        codigo = ""
        if campos["codigo"] and pd.notna(row[campos["codigo"]]):
            codigo = str(row[campos["codigo"]]).strip()

        preco = 0.0
        if campos["preco"] and pd.notna(row[campos["preco"]]):
            preco = para_numero(row[campos["preco"]])

        categoria = classificar_categoria(produto_original, mapa_frutas, mapa_legumes)
        nome_norm = normalizar_texto(produto_original)

        if categoria == "FRUTAS":
            nome_final = mapa_frutas.get(nome_norm, produto_original)
            if nome_final not in frutas:
                frutas[nome_final] = {"CODIGO": codigo, "PRODUTO": nome_final, "PRECO": preco}
            else:
                if not frutas[nome_final]["CODIGO"] and codigo:
                    frutas[nome_final]["CODIGO"] = codigo
                if frutas[nome_final]["PRECO"] == 0 and preco:
                    frutas[nome_final]["PRECO"] = preco
        else:
            nome_final = mapa_legumes.get(nome_norm, produto_original)
            if nome_final not in legumes:
                legumes[nome_final] = {"CODIGO": codigo, "PRODUTO": nome_final, "PRECO": preco}
            else:
                if not legumes[nome_final]["CODIGO"] and codigo:
                    legumes[nome_final]["CODIGO"] = codigo
                if legumes[nome_final]["PRECO"] == 0 and preco:
                    legumes[nome_final]["PRECO"] = preco

    df_frutas = pd.DataFrame(list(frutas.values()))
    df_legumes = pd.DataFrame(list(legumes.values()))

    if df_frutas.empty:
        df_frutas = pd.DataFrame(columns=["CODIGO", "PRODUTO", "PRECO"])
    else:
        df_frutas = df_frutas.sort_values("PRODUTO")

    if df_legumes.empty:
        df_legumes = pd.DataFrame(columns=["CODIGO", "PRODUTO", "PRECO"])
    else:
        df_legumes = df_legumes.sort_values("PRODUTO")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_frutas.to_excel(writer, index=False, sheet_name="FRUTAS")
        df_legumes.to_excel(writer, index=False, sheet_name="LEGUMES")

        wsf = writer.book["FRUTAS"]
        wsl = writer.book["LEGUMES"]

        ajustar_largura(wsf, [14, 42, 14])
        ajustar_largura(wsl, [14, 42, 14])

        for ws in [wsf, wsl]:
            for row in ws.iter_rows(min_row=2):
                if len(row) >= 3:
                    row[2].number_format = 'R$ #,##0.00'

    output.seek(0)
    return output, df_frutas, df_legumes

pedido_file = st.file_uploader("Pedido Krill", type=["xlsx", "xls"], key="pedido")

if pedido_file:
    try:
        modelo_frutas_df, modelo_legumes_df = carregar_modelos_fixos()
        pedido_df, aba_escolhida, header_escolhido = ler_pedido_inteligente(pedido_file)

        arquivo_principal, df_frutas_main, df_legumes_main, lojas = gerar_planilhas_principais(
            pedido_df, modelo_frutas_df, modelo_legumes_df
        )

        arquivo_precos, df_frutas_precos, df_legumes_precos = gerar_planilha_precos(
            pedido_df, modelo_frutas_df, modelo_legumes_df
        )

        campos = detectar_campos_pedido(pedido_df)

        st.success("Processamento concluído com sucesso.")
        st.caption(
            f"Aba lida: {aba_escolhida} | Linha do cabeçalho: {header_escolhido + 1}"
        )
        st.caption(
            f"Colunas encontradas no pedido → "
            f"Loja: {campos['loja'] or 'NÃO ENCONTRADA'} | "
            f"Produto: {campos['produto'] or 'NÃO ENCONTRADA'} | "
            f"Qtde: {campos['qtd'] or 'NÃO ENCONTRADA'} | "
            f"Código: {campos['codigo'] or 'NÃO ENCONTRADA'} | "
            f"Preço: {campos['preco'] or 'NÃO ENCONTRADA'}"
        )

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Itens FRUTAS", 0 if df_frutas_precos.empty else len(df_frutas_precos))
        c2.metric("Itens LEGUMES", 0 if df_legumes_precos.empty else len(df_legumes_precos))
        c3.metric("Lojas processadas", len(lojas))
        c4.metric("Não encontrados", 0)

        b1, b2, b3 = st.columns(3)

        with b1:
            st.download_button(
                "Baixar FRUTAS + LEGUMES",
                data=arquivo_principal,
                file_name="KRILL_THOTH_PRINCIPAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with b2:
            st.download_button(
                "Baixar PREÇOS",
                data=arquivo_precos,
                file_name="KRILL_PRECOS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with b3:
            resumo_output = BytesIO()
            with pd.ExcelWriter(resumo_output, engine="openpyxl") as writer:
                df_frutas_precos.to_excel(writer, index=False, sheet_name="PRECOS_FRUTAS")
                df_legumes_precos.to_excel(writer, index=False, sheet_name="PRECOS_LEGUMES")
                df_frutas_main.to_excel(writer, index=False, sheet_name="THOTH_FRUTAS")
                df_legumes_main.to_excel(writer, index=False, sheet_name="THOTH_LEGUMES")
            resumo_output.seek(0)

            st.download_button(
                "Baixar TUDO",
                data=resumo_output,
                file_name="KRILL_TUDO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with st.expander("Pré-visualização do pedido lido"):
            st.dataframe(pedido_df.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
else:
    st.info("Suba apenas o pedido original da Krill.")
