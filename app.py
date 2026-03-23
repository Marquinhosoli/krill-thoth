
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="KRILL → THOTH PRO", layout="wide")

st.title("📦 KRILL → THOTH (PRO)")

pedido = st.file_uploader("Upload Pedido Krill", type=["xlsx","xls"])

def processar(pedido_file):
    raw = pd.read_excel(pedido_file, header=None)

    header_row = None
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).tolist()
        if "Loja" in row and "Descrição do Produto" in row:
            header_row = i
            break

    df = raw.iloc[header_row+1:].copy()
    df.columns = raw.iloc[header_row]

    df = df[["Loja","Descrição do Produto","Qtde."]]
    df = df.dropna(subset=["Descrição do Produto"])

    df["Loja"] = df["Loja"].astype(str)
    df["Qtde."] = pd.to_numeric(df["Qtde."], errors="coerce").fillna(0)

    df = df[df["Loja"].str.isnumeric()]
    df = df[df["Loja"] != "0"]

    pivot = df.pivot_table(index="Descrição do Produto", columns="Loja", values="Qtde.", aggfunc="sum").fillna(0)

    return pivot

def gerar_saida(modelo_path, pivot):
    wb = load_workbook(modelo_path)
    ws = wb.active

    lojas = {}
    for col in range(2, 40):
        loja = ws.cell(2, col).value
        if isinstance(loja, (int,float)):
            lojas[str(int(loja))] = col

    for row in range(3, ws.max_row+1):
        for col in lojas.values():
            ws.cell(row,col).value = None

    for row in range(3, ws.max_row+1):
        prod = ws.cell(row,1).value
        if prod in pivot.index:
            for loja,col in lojas.items():
                if loja in pivot.columns:
                    val = pivot.loc[prod, loja]
                    if val:
                        ws.cell(row,col).value = val

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

if st.button("PROCESSAR", use_container_width=True):
    if not pedido:
        st.error("Envie o pedido")
    else:
        pivot = processar(pedido)

        frutas = gerar_saida("modelo_frutas.xlsx", pivot)
        legumes = gerar_saida("modelo_legumes.xlsx", pivot)

        st.success("Processado com sucesso")

        st.download_button("Baixar FRUTAS", frutas, "KRILL_FRUTAS.xlsx")
        st.download_button("Baixar LEGUMES", legumes, "KRILL_LEGUMES.xlsx")
