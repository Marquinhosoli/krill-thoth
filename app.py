import io
from copy import copy
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title='Krill → Thoth ERP', page_icon='📦', layout='wide')

VALID_STORES = {
    '1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30'
}
IGNORE_PRODUCT_NAMES = {'TOTAIS:', 'TOTAL', 'SUBTOTAL'}


def normalize_text(value) -> str:
    if value is None:
        return ''
    text = str(value).strip()
    if text.lower() == 'nan':
        return ''
    return ' '.join(text.split())


def find_header_row(df_raw: pd.DataFrame) -> int:
    for i in range(len(df_raw)):
        row = [normalize_text(v) for v in df_raw.iloc[i].tolist()]
        joined = ' | '.join(row)
        if 'Loja' in row and 'Descrição do Produto' in row and 'Qtde.' in row:
            return i
        if 'Loja' in joined and 'Descrição do Produto' in joined and 'Qtde.' in joined:
            return i
    raise ValueError('Não encontrei a linha de cabeçalho do pedido. Verifique o arquivo da Krill.')


def extract_order_table(order_file) -> pd.DataFrame:
    raw = pd.read_excel(order_file, header=None)
    header_row = find_header_row(raw)
    df = raw.iloc[header_row + 1:].copy()
    df.columns = raw.iloc[header_row]

    needed = ['Loja', 'Descrição do Produto', 'Qtde.']
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f'Colunas obrigatórias não encontradas: {missing}')

    df = df[needed].copy()
    df['Loja'] = df['Loja'].map(normalize_text)
    df['Descrição do Produto'] = df['Descrição do Produto'].map(normalize_text)
    df['Qtde.'] = pd.to_numeric(df['Qtde.'], errors='coerce').fillna(0)

    df = df[df['Descrição do Produto'] != '']
    df = df[~df['Descrição do Produto'].str.upper().isin(IGNORE_PRODUCT_NAMES)]
    df = df[df['Loja'].isin(VALID_STORES)]
    df = df[df['Loja'] != '0']  # CD nunca entra
    df = df[df['Qtde.'] > 0]

    if df.empty:
        raise ValueError('Nenhum item válido foi encontrado no pedido.')

    return df


def build_pivot(df: pd.DataFrame) -> pd.DataFrame:
    pivot = df.pivot_table(
        index='Descrição do Produto',
        columns='Loja',
        values='Qtde.',
        aggfunc='sum',
        fill_value=0,
    )
    pivot = pivot.reindex(sorted(pivot.columns, key=lambda x: int(x)), axis=1)
    return pivot


def model_products(model_file) -> List[str]:
    df = pd.read_excel(model_file)
    products: List[str] = []
    for value in df.iloc[:, 0].tolist():
        txt = normalize_text(value)
        if not txt:
            continue
        if txt.upper() in {'PRODUTO', 'PRODUTOS'}:
            continue
        products.append(txt)
    return products


def split_products(pivot: pd.DataFrame, frutas_list: List[str], legumes_list: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    frutas_set = set(frutas_list)
    legumes_set = set(legumes_list)

    frutas_idx = [p for p in pivot.index if p in frutas_set]
    legumes_idx = [p for p in pivot.index if p in legumes_set]
    extra_idx = [p for p in pivot.index if p not in frutas_set and p not in legumes_set]

    frutas = pivot.loc[frutas_idx].copy() if frutas_idx else pivot.iloc[0:0].copy()
    legumes = pivot.loc[legumes_idx + extra_idx].copy() if (legumes_idx or extra_idx) else pivot.iloc[0:0].copy()
    return frutas, legumes, extra_idx


STORE_LABELS = {
    '1': 'KRILL 1', '2': 'KRILL 2', '3': 'LOJA 3', '4': 'LOJA 4', '5': 'LOJA 5',
    '6': 'LOJA 6', '7': 'LOJA 7', '8': 'LOJA 8', '9': 'LOJA 9', '10': 'LOJA 10',
    '11': 'LOJA 11', '12': 'LOJA 12', '13': 'LOJA 13', '14': 'LOJA 14', '15': 'LOJA 15',
    '16': 'LOJA 16', '17': 'LOJA 17', '18': 'LOJA 18', '19': 'LOJA 19', '20': 'LOJA 20',
    '21': 'LOJA 21', '22': 'LOJA 22', '23': 'LOJA 23', '24': 'LOJA 24', '25': 'LOJA 25',
    '26': 'LOJA 26', '27': 'LOJA 27', '28': 'LOJA 28', '29': 'LOJA 29', '30': 'LOJA 30',
}


def detect_store_columns(ws) -> Tuple[Dict[str, int], Optional[int]]:
    store_to_col: Dict[str, int] = {}
    total_col: Optional[int] = None

    for row in range(1, min(ws.max_row, 5) + 1):
        for col in range(1, ws.max_column + 1):
            val = normalize_text(ws.cell(row, col).value).upper()
            if not val:
                continue
            if 'TOTAL' == val or val.endswith(' TOTAL'):
                total_col = col
            if 'CD' in val:
                continue
            digits = ''.join(ch for ch in val if ch.isdigit())
            if digits and ('KRILL' in val or 'LOJA' in val):
                store_to_col[digits] = col

    return store_to_col, total_col


def detect_product_rows(ws) -> Dict[str, int]:
    rows: Dict[str, int] = {}
    for row in range(1, ws.max_row + 1):
        txt = normalize_text(ws.cell(row, 1).value)
        if not txt:
            continue
        if txt.upper() in {'PRODUTO', 'PRODUTOS'}:
            continue
        rows[txt] = row
    return rows


def clear_existing_values(ws, product_rows: Dict[str, int], store_cols: List[int], total_col: Optional[int]):
    for row in product_rows.values():
        for col in store_cols:
            ws.cell(row, col).value = None
        if total_col:
            ws.cell(row, total_col).value = None


def copy_row_style(ws, source_row: int, target_row: int):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        dst._style = copy(src._style)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def write_model_output(model_file, data: pd.DataFrame) -> bytes:
    wb = load_workbook(model_file)
    ws = wb.active

    store_to_col, total_col = detect_store_columns(ws)
    product_rows = detect_product_rows(ws)
    clear_existing_values(ws, product_rows, list(store_to_col.values()), total_col)

    used: Set[str] = set()

    for product, row in product_rows.items():
        if product not in data.index:
            continue
        used.add(product)
        row_total = 0
        for store in data.columns:
            if store not in store_to_col:
                continue
            qty = float(data.loc[product, store])
            if qty:
                ws.cell(row, store_to_col[store]).value = qty
                row_total += qty
        if total_col and row_total:
            ws.cell(row, total_col).value = row_total

    missing = [p for p in data.index if p not in used]
    if missing:
        last_product_row = max(product_rows.values()) if product_rows else 1
        current_row = last_product_row + 1
        style_source = last_product_row
        for product in missing:
            copy_row_style(ws, style_source, current_row)
            ws.cell(current_row, 1).value = product
            row_total = 0
            for store in data.columns:
                if store not in store_to_col:
                    continue
                qty = float(data.loc[product, store])
                if qty:
                    ws.cell(current_row, store_to_col[store]).value = qty
                    row_total += qty
            if total_col and row_total:
                ws.cell(current_row, total_col).value = row_total
            current_row += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def build_price_sheet(frutas: pd.DataFrame, legumes: pd.DataFrame) -> bytes:
    def to_df(data: pd.DataFrame) -> pd.DataFrame:
        if data.empty:
            return pd.DataFrame(columns=['CODIGO', 'PRODUTO', 'PRECO'])
        items = sorted(list(data.index))
        return pd.DataFrame({'CODIGO': [''] * len(items), 'PRODUTO': items, 'PRECO': [''] * len(items)})

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        to_df(frutas).to_excel(writer, sheet_name='FRUTAS', index=False)
        to_df(legumes).to_excel(writer, sheet_name='LEGUMES', index=False)
    out.seek(0)
    return out.getvalue()


def main():
    st.title('📦 Krill → Thoth ERP')
    st.caption('Suba o pedido bruto e os dois modelos. O sistema gera FRUTAS, LEGUMES e a planilha auxiliar de preços.')

    with st.sidebar:
        st.subheader('Como usar')
        st.write('1. Envie o pedido bruto da Krill.')
        st.write('2. Envie o modelo de FRUTAS.')
        st.write('3. Envie o modelo de LEGUMES.')
        st.write('4. Clique em Processar.')
        st.write('5. Baixe os arquivos prontos para o Thoth.')
        st.info('KRILL CD é ignorado automaticamente.')

    col1, col2, col3 = st.columns(3)
    with col1:
        order_file = st.file_uploader('Pedido Krill', type=['xlsx', 'xls'])
    with col2:
        frutas_model = st.file_uploader('Modelo FRUTAS', type=['xlsx', 'xls'])
    with col3:
        legumes_model = st.file_uploader('Modelo LEGUMES', type=['xlsx', 'xls'])

    process = st.button('PROCESSAR PEDIDO KRILL', type='primary', use_container_width=True)

    if process:
        if not order_file or not frutas_model or not legumes_model:
            st.error('Envie os 3 arquivos para continuar.')
            return

        try:
            order_df = extract_order_table(order_file)
            pivot = build_pivot(order_df)
            frutas_list = model_products(frutas_model)
            legumes_list = model_products(legumes_model)
            frutas_df, legumes_df, extras = split_products(pivot, frutas_list, legumes_list)

            frutas_bytes = write_model_output(frutas_model, frutas_df)
            legumes_bytes = write_model_output(legumes_model, legumes_df)
            precos_bytes = build_price_sheet(frutas_df, legumes_df)

            st.success('Arquivos gerados com sucesso.')

            info1, info2, info3 = st.columns(3)
            info1.metric('Itens FRUTAS', len(frutas_df.index))
            info2.metric('Itens LEGUMES', len(legumes_df.index))
            info3.metric('Lojas válidas', len(pivot.columns))

            if extras:
                st.warning('Itens fora dos modelos foram incluídos no final de LEGUMES: ' + ', '.join(extras))

            down1, down2, down3 = st.columns(3)
            with down1:
                st.download_button(
                    'Baixar FRUTAS',
                    data=frutas_bytes,
                    file_name='KRILL_FRUTAS_Thoth.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                )
            with down2:
                st.download_button(
                    'Baixar LEGUMES',
                    data=legumes_bytes,
                    file_name='KRILL_LEGUMES_Thoth.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                )
            with down3:
                st.download_button(
                    'Baixar PREÇOS',
                    data=precos_bytes,
                    file_name='KRILL_PRECOS.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                )

            with st.expander('Prévia dos itens encontrados'):
                preview = order_df.groupby('Descrição do Produto', as_index=False)['Qtde.'].sum().sort_values('Descrição do Produto')
                st.dataframe(preview, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f'Erro ao processar: {e}')


if __name__ == '__main__':
    main()
