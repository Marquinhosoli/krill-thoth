/* =========================
   KRILL PRO - EXPORTAÇÃO COMPLETA
   Gera:
   1) planilha principal FRUTAS
   2) planilha principal LEGUMES
   3) planilha CODIGOS_E_PRECOS_FRUTAS
   4) planilha CODIGOS_E_PRECOS_LEGUMES
   ========================= */

function normalizarTexto(txt) {
  return String(txt || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function paraNumero(valor) {
  if (typeof valor === "number") return valor;
  if (valor == null) return 0;
  const limpo = String(valor)
    .replace(/\./g, "")
    .replace(",", ".")
    .replace(/[^\d.-]/g, "");
  const n = parseFloat(limpo);
  return isNaN(n) ? 0 : n;
}

function moedaBR(valor) {
  return Number(paraNumero(valor) || 0);
}

function extrairNumeroLoja(loja) {
  const m = String(loja || "").match(/(\d+)/);
  return m ? parseInt(m[1], 10) : null;
}

function ordenarColunasKrill(colunas) {
  return [...new Set(colunas)]
    .filter(v => v != null)
    .sort((a, b) => extrairNumeroLoja(a) - extrairNumeroLoja(b));
}

function nomeColunaKrill(numero) {
  return `KRILL ${numero}`;
}

function obterNomeProdutoPedido(item) {
  return (
    item.produto ||
    item.PRODUTO ||
    item.descricao ||
    item.DESCRICAO ||
    item.item ||
    item.ITEM ||
    ""
  );
}

function obterCodigoPedido(item) {
  return (
    item.codigo ||
    item.CODIGO ||
    item.cod ||
    item.COD ||
    item["Código"] ||
    item["CODIGO"] ||
    ""
  );
}

function obterPrecoPedido(item) {
  return (
    item.preco ||
    item.PRECO ||
    item.valor ||
    item.VALOR ||
    item["Preço"] ||
    item["PREÇO"] ||
    0
  );
}

function obterQuantidadePedido(item) {
  return (
    paraNumero(item.quantidade) ||
    paraNumero(item.QUANTIDADE) ||
    paraNumero(item.qtd) ||
    paraNumero(item.QTD) ||
    paraNumero(item.caixas) ||
    paraNumero(item.CAIXAS) ||
    0
  );
}

function obterLojaPedido(item) {
  return (
    item.loja ||
    item.LOJA ||
    item.destino ||
    item.DESTINO ||
    item["Loja"] ||
    ""
  );
}

function mapearModeloPorNome(modeloLinhas) {
  const mapa = new Map();
  for (const linha of modeloLinhas || []) {
    const nome = linha.PRODUTO || linha.produto || linha.DESCRICAO || linha.descricao || "";
    if (nome) {
      mapa.set(normalizarTexto(nome), nome);
    }
  }
  return mapa;
}

function listaProdutosModelo(modeloLinhas) {
  return (modeloLinhas || [])
    .map(l => l.PRODUTO || l.produto || l.DESCRICAO || l.descricao || "")
    .filter(Boolean);
}

function detectarCategoriaProduto(nomeProduto, mapaFrutas, mapaLegumes) {
  const nomeNorm = normalizarTexto(nomeProduto);

  if (mapaFrutas.has(nomeNorm)) return "FRUTAS";
  if (mapaLegumes.has(nomeNorm)) return "LEGUMES";

  const chavesLegumes = [
    "ABOBRINHA","ABOBORA","AIPIM","ALFACE","ALHO","BATATA","BATATA DOCE","BATATA LAVADA",
    "BERINJELA","BETERRABA","BROCOLIS","BROCOLIS NINJA","BROCOLIS RAMOSO","CARA","CEBOLA",
    "CENOURA","CHUCHU","COUVE","COUVE FLOR","ERVILHA","ERVILHA TORTA","ESPINAFRE","INHAME",
    "JILO","MANDIOCA","MANDIOQUINHA","MAXIXE","MILHO","MILHO VERDE","PEPINO","PEPINO JAPONES",
    "PIMENTA","PIMENTAO","QUIABO","REPOLHO","RUCULA","SALSINHA","TOMATE","TOMATE CEREJA",
    "TOMATE GRAPE","VAGEM"
  ];

  const chavesFrutas = [
    "ABACATE","ABACAXI","ACEROLA","AMEIXA","AMORA","BANANA","CAJU","CAQUI","CARAMBOLA",
    "CEREJA","COCO","FIGO","FRAMBOESA","GOIABA","JABUTICABA","KIWI","LARANJA","LIMAO",
    "MACA","MAMAO","MANGA","MELANCIA","MELAO","MEXERICA","MIRTILO","MORANGO","NECTARINA",
    "PERA","PESSEGO","PITAYA","PITANGA","ROMA","TANGERINA","UVA"
  ];

  if (chavesLegumes.some(k => nomeNorm.includes(k))) return "LEGUMES";
  if (chavesFrutas.some(k => nomeNorm.includes(k))) return "FRUTAS";

  return "LEGUMES";
}

function agruparPedidoHorizontal(pedidoItens, modeloFrutas, modeloLegumes) {
  const mapaFrutas = mapearModeloPorNome(modeloFrutas);
  const mapaLegumes = mapearModeloPorNome(modeloLegumes);

  const lojasSet = new Set();
  const agrupadoFrutas = new Map();
  const agrupadoLegumes = new Map();

  for (const item of pedidoItens || []) {
    const produtoOriginal = obterNomeProdutoPedido(item);
    const codigo = obterCodigoPedido(item);
    const preco = moedaBR(obterPrecoPedido(item));
    const qtd = obterQuantidadePedido(item);
    const lojaBruta = obterLojaPedido(item);
    const numeroLoja = extrairNumeroLoja(lojaBruta);

    if (!produtoOriginal || !numeroLoja || !qtd) continue;

    const loja = nomeColunaKrill(numeroLoja);
    lojasSet.add(loja);

    const categoria = detectarCategoriaProduto(produtoOriginal, mapaFrutas, mapaLegumes);
    const mapaCategoria = categoria === "FRUTAS" ? agrupadoFrutas : agrupadoLegumes;

    const nomeNorm = normalizarTexto(produtoOriginal);
    let nomeFinal = produtoOriginal;

    if (categoria === "FRUTAS" && mapaFrutas.has(nomeNorm)) {
      nomeFinal = mapaFrutas.get(nomeNorm);
    } else if (categoria === "LEGUMES" && mapaLegumes.has(nomeNorm)) {
      nomeFinal = mapaLegumes.get(nomeNorm);
    }

    if (!mapaCategoria.has(nomeFinal)) {
      mapaCategoria.set(nomeFinal, {
        PRODUTO: nomeFinal,
        CODIGO: codigo,
        PRECO: preco
      });
    }

    const linha = mapaCategoria.get(nomeFinal);
    linha[loja] = (linha[loja] || 0) + qtd;

    if (!linha.CODIGO && codigo) linha.CODIGO = codigo;
    if ((!linha.PRECO || linha.PRECO === 0) && preco) linha.PRECO = preco;
  }

  const colunasLojas = ordenarColunasKrill([...lojasSet]);

  function montarLinhasFinais(mapaCategoria, modeloLinhas) {
    const linhasModelo = listaProdutosModelo(modeloLinhas);
    const nomesPedido = [...mapaCategoria.keys()];
    const nomesForaDoModelo = nomesPedido.filter(n => !linhasModelo.some(m => normalizarTexto(m) === normalizarTexto(n)));
    const ordemFinal = [...linhasModelo, ...nomesForaDoModelo];

    return ordemFinal.map(nome => {
      const base = { PRODUTO: nome };
      for (const loja of colunasLojas) base[loja] = 0;

      const registro = [...mapaCategoria.entries()]
        .find(([k]) => normalizarTexto(k) === normalizarTexto(nome));

      if (registro) {
        const dados = registro[1];
        for (const loja of colunasLojas) {
          base[loja] = dados[loja] || 0;
        }
      }
      return base;
    });
  }

  function montarCodigosPrecos(mapaCategoria) {
    return [...mapaCategoria.values()]
      .map(item => ({
        CODIGO: item.CODIGO || "",
        PRODUTO: item.PRODUTO || "",
        PRECO: moedaBR(item.PRECO || 0)
      }))
      .sort((a, b) => normalizarTexto(a.PRODUTO).localeCompare(normalizarTexto(b.PRODUTO), "pt-BR"));
  }

  return {
    colunasLojas,
    frutas: montarLinhasFinais(agrupadoFrutas, modeloFrutas),
    legumes: montarLinhasFinais(agrupadoLegumes, modeloLegumes),
    codigosFrutas: montarCodigosPrecos(agrupadoFrutas),
    codigosLegumes: montarCodigosPrecos(agrupadoLegumes)
  };
}

function ajustarLarguraColunas(ws, larguras) {
  ws["!cols"] = larguras.map(w => ({ wch: w }));
}

function aplicarFormatoPreco(ws, dados) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  let colunaPreco = -1;

  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    const cel = ws[addr];
    if (cel && String(cel.v).toUpperCase() === "PRECO") {
      colunaPreco = c;
      break;
    }
  }

  if (colunaPreco >= 0) {
    for (let r = 1; r <= range.e.r; r++) {
      const addr = XLSX.utils.encode_cell({ r, c: colunaPreco });
      if (ws[addr]) ws[addr].z = 'R$ #,##0.00';
    }
  }
}

function baixarWorkbook(workbook, nomeArquivo) {
  XLSX.writeFile(workbook, nomeArquivo);
}

function exportarKrillProCompleto({
  pedidoItens,
  modeloFrutas,
  modeloLegumes,
  nomeBase = "KRILL_PRO"
}) {
  if (!Array.isArray(pedidoItens) || pedidoItens.length === 0) {
    throw new Error("Pedido vazio.");
  }

  const resultado = agruparPedidoHorizontal(pedidoItens, modeloFrutas, modeloLegumes);

  // Workbook principal
  const wbPrincipal = XLSX.utils.book_new();

  const wsFrutas = XLSX.utils.json_to_sheet(resultado.frutas);
  const wsLegumes = XLSX.utils.json_to_sheet(resultado.legumes);

  ajustarLarguraColunas(wsFrutas, [42, ...resultado.colunasLojas.map(() => 12)]);
  ajustarLarguraColunas(wsLegumes, [42, ...resultado.colunasLojas.map(() => 12)]);

  XLSX.utils.book_append_sheet(wbPrincipal, wsFrutas, "FRUTAS");
  XLSX.utils.book_append_sheet(wbPrincipal, wsLegumes, "LEGUMES");

  baixarWorkbook(wbPrincipal, `${nomeBase}.xlsx`);

  // Workbook códigos e preços
  const wbCodigos = XLSX.utils.book_new();

  const wsCodFrutas = XLSX.utils.json_to_sheet(resultado.codigosFrutas);
  const wsCodLegumes = XLSX.utils.json_to_sheet(resultado.codigosLegumes);

  ajustarLarguraColunas(wsCodFrutas, [14, 42, 14]);
  ajustarLarguraColunas(wsCodLegumes, [14, 42, 14]);

  aplicarFormatoPreco(wsCodFrutas, resultado.codigosFrutas);
  aplicarFormatoPreco(wsCodLegumes, resultado.codigosLegumes);

  XLSX.utils.book_append_sheet(wbCodigos, wsCodFrutas, "FRUTAS");
  XLSX.utils.book_append_sheet(wbCodigos, wsCodLegumes, "LEGUMES");

  baixarWorkbook(wbCodigos, `${nomeBase}_CODIGOS_E_PRECOS.xlsx`);

  return resultado;
}
