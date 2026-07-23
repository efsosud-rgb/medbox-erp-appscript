/************************************************************
 * MEDBOX ERP | MOTOR DE DADOS MVP (dados reais das bases BD_)
 *
 * Le exclusivamente:
 * - BD_SKU MASTER
 * - BD_PRECOS
 * - abas BD_<CLIENTE> descobertas dinamicamente
 *
 * NUNCA le MAPA_ERP nem AUX_BUSCA.
 * Este arquivo e somente leitura: nenhuma funcao aqui grava na planilha.
 ************************************************************/

const MVP_ABA_SKU_MASTER = 'BD_SKU MASTER';
const MVP_ABA_PRECOS = 'BD_PRECOS';
const MVP_MAX_RESULTADOS_BUSCA = 300;
const MVP_MAX_RESULTADOS_MAPA = 500;

/************************************************************
 * NORMALIZACAO / CONVERSAO
 ************************************************************/

function mvpNormalizarTexto_(valor) {
  if (valor === null || valor === undefined) return '';

  const MVP_DIACRITICOS_RE = new RegExp('[' + String.fromCharCode(0x0300) + '-' + String.fromCharCode(0x036f) + ']', 'g');

  return String(valor)
    .normalize('NFD')
    .replace(MVP_DIACRITICOS_RE, '')
    .toUpperCase()
    .replace(/,/g, '.')
    .replace(/([0-9])\s+(MG|ML|G|L|MCG|UI|UN)\b/g, '$1$2')
    .replace(/([A-Z])([0-9])/g, '$1 $2')
    .replace(/([0-9])([A-Z])/g, '$1$2')
    .replace(/[^A-Z0-9.%/+ X\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function mvpConverterNumero_(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;

  if (typeof valor === 'number') {
    return isNaN(valor) ? 0 : valor;
  }

  let txt = String(valor).trim();
  if (!txt) return 0;
  if (txt === '-' || txt === 'R$ -' || txt === 'R$-' || txt === '--') return 0;

  txt = txt
    .replace(/R\$/gi, '')
    .replace(/\$/g, '')
    .replace(/%/g, '')
    .replace(/\s+/g, '');

  if (!txt || txt === '-') return 0;

  if (txt.indexOf(',') !== -1) {
    txt = txt.replace(/\./g, '').replace(',', '.');
  }

  txt = txt.replace(/[^0-9.\-]/g, '');
  if (!txt || txt === '-') return 0;

  const n = Number(txt);
  return isNaN(n) ? 0 : n;
}

function mvpCriarChaveProduto_(skuMasterNormalizado, descricaoNormalizada) {
  if (skuMasterNormalizado) return 'SKU:' + skuMasterNormalizado;
  if (descricaoNormalizada) return 'PRODUTO:' + descricaoNormalizada;
  return null;
}

function mvpEncontrarColuna_(headers, candidatos) {
  const headersNorm = headers.map(function(h) {
    return mvpNormalizarTexto_(h).replace(/[^A-Z0-9]/g, '');
  });

  for (let c = 0; c < candidatos.length; c++) {
    const candNorm = mvpNormalizarTexto_(candidatos[c]).replace(/[^A-Z0-9]/g, '');
    const idxExato = headersNorm.indexOf(candNorm);
    if (idxExato !== -1) return idxExato;
  }

  for (let c = 0; c < candidatos.length; c++) {
    const candNorm = mvpNormalizarTexto_(candidatos[c]).replace(/[^A-Z0-9]/g, '');
    for (let i = 0; i < headersNorm.length; i++) {
      if (headersNorm[i] !== '' &&
          (headersNorm[i].indexOf(candNorm) !== -1 || candNorm.indexOf(headersNorm[i]) !== -1)) {
        return i;
      }
    }
  }

  return -1;
}

/************************************************************
 * DESCOBERTA DE ABAS DE CLIENTES
 ************************************************************/

function mvpDescobrirAbasClientes_(ss) {
  const sheets = ss.getSheets();

  const longTerms = [
    'COTACOES', 'COTAÇÕES',
    'SKU_MASTER', 'SKU MASTER',
    'PRECOS', 'PREÇOS',
    'SEFAZ',
    'ADMIN', 'BACKUP',
    'REVISAO', 'REVISÃO',
    'RELATORIO', 'RELATÓRIO',
    'IMPORTACOES', 'IMPORTAÇÕES'
  ];
  const shortTerms = ['LOG', 'AUX', 'MAPA'];

  const abas = [];

  sheets.forEach(function(sheet) {
    const nome = sheet.getName();
    const upperNome = nome.trim().toUpperCase();

    if (upperNome.indexOf('BD_') !== 0) return;

    for (let i = 0; i < longTerms.length; i++) {
      if (upperNome.indexOf(longTerms[i]) !== -1) return;
    }

    const tokenNome = upperNome.replace(/[_\-.\/\\]+/g, ' ').replace(/\s+/g, ' ').trim();
    const tokens = tokenNome.split(' ');
    for (let i = 0; i < shortTerms.length; i++) {
      if (tokens.indexOf(shortTerms[i]) !== -1) return;
    }

    const deteccao = mvpDetectarCabecalhoCliente_(sheet);
    if (deteccao) {
      abas.push({ sheet: sheet, deteccao: deteccao, nomeCliente: mvpNomeClienteDaAba_(nome) });
    }
  });

  return abas;
}

function mvpNomeClienteDaAba_(nomeAba) {
  return String(nomeAba).substring(3).trim().toUpperCase();
}

function mvpDetectarCabecalhoCliente_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return null;

  const colLimit = Math.min(lastCol, 8);
  const range = sheet.getRange(1, 1, Math.min(lastRow, 5), colLimit);
  const values = range.getValues();

  for (let r = 0; r < values.length; r++) {
    const headers = values[r].map(String);

    const colProdutos = mvpEncontrarColuna_(headers, ['PRODUTOS', 'PRODUTO']);
    const colQuant = mvpEncontrarColuna_(headers, ['QUANT', 'QUANTIDADE', 'QTD']);
    const colVUnit = mvpEncontrarColuna_(headers, ['V. UNIT', 'V UNIT', 'VALOR UNITARIO', 'VALOR_UNITARIO']);
    const colVTotal = mvpEncontrarColuna_(headers, ['V. TOTAL', 'V TOTAL', 'VALOR TOTAL', 'VALOR_TOTAL']);

    if (colProdutos !== -1 && colQuant !== -1 && (colVUnit !== -1 || colVTotal !== -1)) {
      return {
        headerRowIndex: r,
        colCategoria: mvpEncontrarColuna_(headers, ['CATEGORIA']),
        colCodigo: mvpEncontrarColuna_(headers, ['CODIGO', 'CÓDIGO']),
        colProdutos: colProdutos,
        colSkuMaster: mvpEncontrarColuna_(headers, ['SKU_MASTER', 'SKU MASTER']),
        colQuant: colQuant,
        colUnd: mvpEncontrarColuna_(headers, ['UND', 'UNIDADE']),
        colVUnit: colVUnit,
        colVTotal: colVTotal
      };
    }
  }

  return null;
}

/************************************************************
 * BD_SKU MASTER
 ************************************************************/

function mvpCarregarSkuMaster_(ss) {
  const resultado = { porDescricao: {}, porSku: {} };

  const sheet = ss.getSheetByName(MVP_ABA_SKU_MASTER);
  if (!sheet) return resultado;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return resultado;

  const lastCol = Math.min(sheet.getLastColumn(), 3);
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(String);

  const colCategoria = mvpEncontrarColuna_(headers, ['CATEGORIA']);
  const colProdutos = mvpEncontrarColuna_(headers, ['PRODUTOS', 'PRODUTO']);
  const colSku = mvpEncontrarColuna_(headers, ['SKU MASTER', 'SKU_MASTER']);

  if (colProdutos === -1 || colSku === -1) return resultado;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const descricaoRaw = String(row[colProdutos] || '').trim();
    const skuRaw = String(row[colSku] || '').trim();

    if (!descricaoRaw && !skuRaw) continue;

    const categoriaRaw = colCategoria !== -1 ? String(row[colCategoria] || '').trim() : '';
    const descNorm = mvpNormalizarTexto_(descricaoRaw);
    const skuNorm = mvpNormalizarTexto_(skuRaw);

    const registro = { categoria: categoriaRaw, descricao: descricaoRaw, skuMaster: skuRaw };

    if (descNorm && !resultado.porDescricao[descNorm]) {
      resultado.porDescricao[descNorm] = registro;
    }
    if (skuNorm && !resultado.porSku[skuNorm]) {
      resultado.porSku[skuNorm] = registro;
    }
  }

  return resultado;
}

/************************************************************
 * BD_PRECOS
 ************************************************************/

function mvpCarregarPrecos_(ss) {
  const resultado = { porSku: {}, porDescricao: {} };

  const sheet = ss.getSheetByName(MVP_ABA_PRECOS);
  if (!sheet) return resultado;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return resultado;

  const lastCol = Math.min(sheet.getLastColumn(), 7);
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(String);

  const colFornecedor = mvpEncontrarColuna_(headers, ['FORNECEDOR']);
  const colMarca = mvpEncontrarColuna_(headers, ['MARCA']);
  const colCategoria = mvpEncontrarColuna_(headers, ['CATEGORIA']);
  const colDescricao = mvpEncontrarColuna_(headers, ['DESCRICAO DO PRODUTO', 'DESCRICAO_DO_PRODUTO', 'DESCRICAO']);
  const colSku = mvpEncontrarColuna_(headers, ['SKU MASTER', 'SKU_MASTER']);
  const colUnidade = mvpEncontrarColuna_(headers, ['UNIDADE DE COMPRA', 'UNIDADE_DE_COMPRA', 'UNIDADE']);
  const colValor = mvpEncontrarColuna_(headers, ['VALOR UNITARIO', 'VALOR_UNITARIO', 'VALOR UNITÁRIO']);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const skuRaw = colSku !== -1 ? String(row[colSku] || '').trim() : '';
    const descricaoRaw = colDescricao !== -1 ? String(row[colDescricao] || '').trim() : '';

    if (!skuRaw && !descricaoRaw) continue;

    const valor = colValor !== -1 ? mvpConverterNumero_(row[colValor]) : 0;
    if (valor <= 0) continue;

    const registro = {
      valorUnitario: valor,
      fornecedor: colFornecedor !== -1 ? String(row[colFornecedor] || '').trim() : '',
      marca: colMarca !== -1 ? String(row[colMarca] || '').trim() : '',
      categoria: colCategoria !== -1 ? String(row[colCategoria] || '').trim() : '',
      unidadeCompra: colUnidade !== -1 ? String(row[colUnidade] || '').trim() : '',
      descricaoProduto: descricaoRaw
    };

    const skuNorm = mvpNormalizarTexto_(skuRaw);
    const descNorm = mvpNormalizarTexto_(descricaoRaw);

    if (skuNorm) {
      const atual = resultado.porSku[skuNorm];
      if (!atual || valor < atual.valorUnitario) resultado.porSku[skuNorm] = registro;
    } else if (descNorm) {
      const atual = resultado.porDescricao[descNorm];
      if (!atual || valor < atual.valorUnitario) resultado.porDescricao[descNorm] = registro;
    }
  }

  return resultado;
}

/************************************************************
 * ABAS BD_<CLIENTE>
 ************************************************************/

function mvpCarregarAbaCliente_(sheetInfo, skuMasterIndex) {
  const sheet = sheetInfo.sheet;
  const deteccao = sheetInfo.deteccao;
  const itens = [];

  const lastRow = sheet.getLastRow();
  const colLimit = Math.min(sheet.getLastColumn(), 8);
  const values = sheet.getRange(1, 1, lastRow, colLimit).getValues();
  const startRow = deteccao.headerRowIndex + 1;

  for (let i = startRow; i < values.length; i++) {
    const row = values[i];

    const categoriaRaw = deteccao.colCategoria !== -1 ? String(row[deteccao.colCategoria] || '').trim() : '';
    const produtoRaw = deteccao.colProdutos !== -1 ? String(row[deteccao.colProdutos] || '').trim() : '';
    const skuRawLinha = deteccao.colSkuMaster !== -1 ? String(row[deteccao.colSkuMaster] || '').trim() : '';
    const undRaw = deteccao.colUnd !== -1 ? String(row[deteccao.colUnd] || '').trim() : '';

    if (!produtoRaw && !skuRawLinha) continue;

    const quantidade = deteccao.colQuant !== -1 ? mvpConverterNumero_(row[deteccao.colQuant]) : 0;
    const valorUnit = deteccao.colVUnit !== -1 ? mvpConverterNumero_(row[deteccao.colVUnit]) : 0;
    let valorTotal = deteccao.colVTotal !== -1 ? mvpConverterNumero_(row[deteccao.colVTotal]) : 0;

    if (valorTotal <= 0 && quantidade > 0 && valorUnit > 0) {
      valorTotal = quantidade * valorUnit;
    }

    if (quantidade <= 0 && valorTotal <= 0) continue;

    // Regra de SKU: 1) linha do cliente; 2) BD_SKU MASTER por descricao; 3) vazio. Nunca usar Codigo.
    let skuMasterFinal = skuRawLinha;
    if (!skuMasterFinal && produtoRaw) {
      const descNorm = mvpNormalizarTexto_(produtoRaw);
      const registro = skuMasterIndex.porDescricao[descNorm];
      if (registro && registro.skuMaster) {
        skuMasterFinal = registro.skuMaster;
      }
    }

    const skuNorm = mvpNormalizarTexto_(skuMasterFinal);
    const descNormItem = mvpNormalizarTexto_(produtoRaw);
    const chave = mvpCriarChaveProduto_(skuNorm, descNormItem);
    if (!chave) continue;

    itens.push({
      chave: chave,
      categoria: categoriaRaw,
      produto: produtoRaw,
      skuMaster: skuMasterFinal,
      unidade: undRaw,
      quantidade: quantidade > 0 ? quantidade : 0,
      valorUnit: valorUnit > 0 ? valorUnit : 0,
      valorTotal: valorTotal > 0 ? valorTotal : 0
    });
  }

  return itens;
}

/************************************************************
 * CONSOLIDACAO
 ************************************************************/

function mvpCriarProdutoBase_(item) {
  return {
    categoria: item.categoria,
    produto: item.produto,
    skuMaster: item.skuMaster,
    descricaoOrigem: item.produto,
    unidade: item.unidade,
    quantidadeTotalDemandada: 0,
    valorTotalDemandado: 0,
    clientesComDemandaSet: {},
    clientesValores: {}
  };
}

function mvpAcumularProduto_(produto, item, nomeCliente) {
  if (!produto.categoria && item.categoria) produto.categoria = item.categoria;
  if (!produto.skuMaster && item.skuMaster) produto.skuMaster = item.skuMaster;
  if (!produto.descricaoOrigem && item.produto) produto.descricaoOrigem = item.produto;
  if (!produto.unidade && item.unidade) produto.unidade = item.unidade;

  produto.quantidadeTotalDemandada += item.quantidade;
  produto.valorTotalDemandado += item.valorTotal;

  if (!produto.clientesValores[nomeCliente]) {
    produto.clientesValores[nomeCliente] = {
      quantidadeAtual: 0,
      valorUnitarioAtual: null,
      demandaAtual: 0,
      dataUltimaCotacao: null,
      mediaMensalQuantidade: null,
      mediaMensalValorUnitario: null
    };
  }

  const cv = produto.clientesValores[nomeCliente];
  cv.quantidadeAtual += item.quantidade;
  cv.demandaAtual += item.valorTotal;

  if (item.valorUnit > 0 && (cv.valorUnitarioAtual === null || item.valorUnit < cv.valorUnitarioAtual)) {
    cv.valorUnitarioAtual = item.valorUnit;
  }

  if (item.quantidade > 0 || item.valorTotal > 0) {
    produto.clientesComDemandaSet[nomeCliente] = true;
  }
}

function mvpAplicarPreco_(produto, precosIndex, markup) {
  const skuNorm = mvpNormalizarTexto_(produto.skuMaster);
  const descNorm = mvpNormalizarTexto_(produto.descricaoOrigem);

  let registro = null;
  if (skuNorm && precosIndex.porSku[skuNorm]) {
    registro = precosIndex.porSku[skuNorm];
  } else if (descNorm && precosIndex.porDescricao[descNorm]) {
    registro = precosIndex.porDescricao[descNorm];
  }

  const precoCusto = registro ? registro.valorUnitario : null;

  produto.precoCustoAtual = precoCusto;
  produto.precoCusto = precoCusto;
  produto.fornecedor = registro ? registro.fornecedor : '';
  produto.fabricante = registro ? registro.marca : '';
  if (registro && registro.unidadeCompra) produto.unidade = registro.unidadeCompra;
  if (!produto.categoria && registro && registro.categoria) produto.categoria = registro.categoria;

  produto.precoMedbox = calcularPrecoMedBox(produto.skuMaster, precoCusto, markup) || null;

  produto.statusCadastro = produto.skuMaster ? 'CADASTRADO' : 'SEM SKU MASTER';
  produto.statusPreco = (precoCusto && precoCusto > 0) ? 'COM PREÇO' : 'SEM PREÇO';
}

function mvpFinalizarProduto_(produto) {
  const nomeExibicao = produto.skuMaster ? produto.skuMaster : produto.produto;

  return {
    categoria: upperText(produto.categoria),
    produto: upperText(nomeExibicao),
    skuMaster: upperText(produto.skuMaster),
    skuOperacional: '',
    produtoNormalizado: upperText(produto.descricaoOrigem),
    descricaoOrigem: upperText(produto.descricaoOrigem),
    unidade: upperText(produto.unidade),
    precoCustoAtual: produto.precoCustoAtual,
    precoCusto: produto.precoCusto,
    precoMedbox: produto.precoMedbox,
    fornecedor: upperText(produto.fornecedor),
    fabricante: upperText(produto.fabricante),
    dataUltimoPreco: null,
    ultimaQuantidadeCotada: null,
    ultimoValorUnitarioCotado: null,
    quantTotal: arredondar(produto.quantidadeTotalDemandada, 4),
    quantidadeTotalDemandada: arredondar(produto.quantidadeTotalDemandada, 4),
    demandaR: arredondar(produto.valorTotalDemandado, 4),
    valorTotalDemandado: arredondar(produto.valorTotalDemandado, 4),
    clientesComDemanda: Object.keys(produto.clientesComDemandaSet).sort().join('; '),
    clientesValores: produto.clientesValores,
    statusCadastro: produto.statusCadastro,
    statusPreco: produto.statusPreco
  };
}

function mvpConsolidarDados_(markup) {
  const ss = obterPlanilhaMedbox_();

  const skuMasterIndex = mvpCarregarSkuMaster_(ss);
  const precosIndex = mvpCarregarPrecos_(ss);
  const abasClientes = mvpDescobrirAbasClientes_(ss);

  const produtosMap = {};
  const clientesDetectados = [];
  const linhasValidasPorCliente = {};
  const abasDetectadas = [];

  abasClientes.forEach(function(sheetInfo) {
    clientesDetectados.push(sheetInfo.nomeCliente);
    abasDetectadas.push(sheetInfo.sheet.getName());

    const itens = mvpCarregarAbaCliente_(sheetInfo, skuMasterIndex);
    linhasValidasPorCliente[sheetInfo.nomeCliente] = itens.length;

    itens.forEach(function(item) {
      if (!produtosMap[item.chave]) {
        produtosMap[item.chave] = mvpCriarProdutoBase_(item);
      }
      mvpAcumularProduto_(produtosMap[item.chave], item, sheetInfo.nomeCliente);
    });
  });

  const chaves = Object.keys(produtosMap);
  const produtosFinais = [];

  for (let i = 0; i < chaves.length; i++) {
    const produto = produtosMap[chaves[i]];
    mvpAplicarPreco_(produto, precosIndex, markup);
    produtosFinais.push(mvpFinalizarProduto_(produto));
  }

  clientesDetectados.sort();

  const regSku = skuMasterIndex.porDescricao ? Object.keys(skuMasterIndex.porDescricao).length : 0;
  const regPrecos = precosIndex.porSku ? Object.keys(precosIndex.porSku).length : 0;
  const regPrecosDesc = precosIndex.porDescricao ? Object.keys(precosIndex.porDescricao).length : 0;

  return {
    clientes: clientesDetectados,
    produtos: produtosFinais,
    atualizadoEm: formatarDataHora(new Date()),
    _diagnostico: {
      abasClientes: abasDetectadas,
      linhasValidasPorCliente: linhasValidasPorCliente,
      registrosSkuMaster: regSku,
      registrosPrecos: regPrecos + regPrecosDesc
    }
  };
}

function diagnosticarMvpSomenteLeitura() {
  const dados = mvpConsolidarDados_(0.8);
  const produtos = dados.produtos;
  const diag = dados._diagnostico;

  var totalComSku = 0;
  var totalSemSku = 0;
  var totalComPreco = 0;
  var totalSemPreco = 0;
  var quantidadeTotal = 0;
  var valorTotal = 0;

  produtos.forEach(function(p) {
    quantidadeTotal += p.quantTotal || 0;
    valorTotal += p.demandaR || 0;
    if (p.skuMaster) { totalComSku++; } else { totalSemSku++; }
    if (p.precoCustoAtual && p.precoCustoAtual > 0) { totalComPreco++; } else { totalSemPreco++; }
  });

  var exemplosProdutos = [];
  for (var i = 0; i < Math.min(10, produtos.length); i++) {
    var p = produtos[i];
    exemplosProdutos.push({
      produto: p.produto,
      skuMaster: p.skuMaster,
      categoria: p.categoria,
      fornecedor: p.fornecedor,
      fabricante: p.fabricante,
      precoCustoAtual: p.precoCustoAtual,
      quantidadeTotalDemandada: p.quantTotal,
      valorTotalDemandado: p.demandaR,
      clientesComDemanda: p.clientesComDemanda
    });
  }

  var exemplosComDoisClientes = [];
  for (var j = 0; j < produtos.length && exemplosComDoisClientes.length < 5; j++) {
    var pj = produtos[j];
    var clientesArr = pj.clientesComDemanda ? pj.clientesComDemanda.split('; ').filter(Boolean) : [];
    if (clientesArr.length >= 2) {
      exemplosComDoisClientes.push({ produto: pj.produto, skuMaster: pj.skuMaster, clientes: clientesArr });
    }
  }

  var exemplosCorrelacionadosSku = [];
  for (var k = 0; k < produtos.length && exemplosCorrelacionadosSku.length < 5; k++) {
    var pk = produtos[k];
    if (pk.skuMaster) {
      exemplosCorrelacionadosSku.push({ produto: pk.produto, skuMaster: pk.skuMaster });
    }
  }

  var exemplosComPreco = [];
  for (var m = 0; m < produtos.length && exemplosComPreco.length < 5; m++) {
    var pm = produtos[m];
    if (pm.precoCustoAtual && pm.precoCustoAtual > 0) {
      exemplosComPreco.push({
        produto: pm.produto,
        skuMaster: pm.skuMaster,
        precoCusto: pm.precoCustoAtual,
        fornecedor: pm.fornecedor,
        fabricante: pm.fabricante
      });
    }
  }

  var resultado = {
    executadoEm: formatarDataHora(new Date()),
    planilhaId: '1QLurrnKCXorjXUoUtPl6MoSP0Gk32Yz0AWBHxebifCw',
    abasClientes: diag.abasClientes,
    totalAbasClientes: diag.abasClientes.length,
    linhasValidasPorCliente: diag.linhasValidasPorCliente,
    registrosSkuMaster: diag.registrosSkuMaster,
    registrosPrecos: diag.registrosPrecos,
    totalProdutos: produtos.length,
    totalComSku: totalComSku,
    totalSemSku: totalSemSku,
    totalComPreco: totalComPreco,
    totalSemPreco: totalSemPreco,
    quantidadeTotal: quantidadeTotal,
    valorTotal: valorTotal,
    exemplosProdutos: exemplosProdutos,
    exemplosComDoisClientes: exemplosComDoisClientes,
    exemplosCorrelacionadosSku: exemplosCorrelacionadosSku,
    exemplosComPreco: exemplosComPreco
  };

  var json = JSON.stringify(resultado, null, 2);
  Logger.log(json);
  return resultado;
}

/************************************************************
 * FILTROS COMPARTILHADOS
 ************************************************************/

function mvpTextoCombinadoProduto_(produto) {
  return mvpNormalizarTexto_([
    produto.produto,
    produto.skuMaster,
    produto.descricaoOrigem,
    produto.categoria,
    produto.fornecedor,
    produto.fabricante
  ].join(' '));
}

function mvpProdutoAtendeFiltros_(produto, filtrosTratados, termos) {
  const categoria = mvpNormalizarTexto_(filtrosTratados.categoria || '');
  const sku = mvpNormalizarTexto_(filtrosTratados.sku || '');
  const fabricante = mvpNormalizarTexto_(filtrosTratados.fabricante || '');
  const fornecedor = mvpNormalizarTexto_(filtrosTratados.fornecedor || '');

  const textoCompleto = mvpTextoCombinadoProduto_(produto);

  if (termos.length > 0) {
    for (let t = 0; t < termos.length; t++) {
      if (textoCompleto.indexOf(termos[t]) === -1) return false;
    }
  }

  if (categoria && mvpNormalizarTexto_(produto.categoria).indexOf(categoria) === -1) return false;

  if (sku) {
    const skuProduto = mvpNormalizarTexto_(produto.skuMaster);
    if (skuProduto.indexOf(sku) === -1) return false;
  }

  if (fabricante && mvpNormalizarTexto_(produto.fabricante).indexOf(fabricante) === -1) return false;

  if (fornecedor && mvpNormalizarTexto_(produto.fornecedor).indexOf(fornecedor) === -1) return false;

  return true;
}

function mvpColetarTermos_(filtrosTratados) {
  return [filtrosTratados.termo1, filtrosTratados.termo2, filtrosTratados.termo3, filtrosTratados.termo4]
    .map(function(t) { return mvpNormalizarTexto_(t || ''); })
    .filter(Boolean);
}

/************************************************************
 * FUNCOES PUBLICAS DO MVP (consumidas pelos wrappers em Codigo.js / Fase2_Dashboard.js)
 ************************************************************/

function mvpObterOpcoesSistema() {
  const dados = mvpConsolidarDados_(0.8);

  const categoriasSet = {};
  const fabricantesSet = {};
  const fornecedoresSet = {};

  dados.produtos.forEach(function(p) {
    if (p.categoria) categoriasSet[p.categoria] = true;
    if (p.fabricante) fabricantesSet[p.fabricante] = true;
    if (p.fornecedor) fornecedoresSet[p.fornecedor] = true;
  });

  return {
    atualizadoEm: dados.atualizadoEm,
    totalProdutosMapa: dados.produtos.length,
    totalAuxBusca: 0,
    clientes: dados.clientes,
    categorias: Object.keys(categoriasSet).sort(),
    fabricantes: Object.keys(fabricantesSet).sort().slice(0, 500),
    fornecedores: Object.keys(fornecedoresSet).sort().slice(0, 500)
  };
}

function mvpPesquisarBuscaAvancada(filtros) {
  const filtrosTratados = filtros || {};
  const markup = normalizarMarkup(filtrosTratados.markup);
  const clienteFiltro = mvpNormalizarTexto_(filtrosTratados.cliente || '');
  const termos = mvpColetarTermos_(filtrosTratados);

  const dados = mvpConsolidarDados_(markup);

  let encontrados = dados.produtos.filter(function(p) {
    if (!mvpProdutoAtendeFiltros_(p, filtrosTratados, termos)) return false;
    if (clienteFiltro) {
      const clienteReal = dados.clientes.filter(function(c) { return mvpNormalizarTexto_(c) === clienteFiltro; })[0];
      const cv = clienteReal ? p.clientesValores[clienteReal] : null;
      if (!cv || cv.quantidadeAtual <= 0) return false;
    }
    return true;
  });

  encontrados.sort(function(a, b) {
    if (b.quantTotal !== a.quantTotal) return b.quantTotal - a.quantTotal;
    return mvpNormalizarTexto_(a.produto).localeCompare(mvpNormalizarTexto_(b.produto));
  });

  const totalEncontrado = encontrados.length;
  const limitados = encontrados.slice(0, MVP_MAX_RESULTADOS_BUSCA);

  let quantidadeTotalDemandada = 0;
  let valorTotalDemandado = 0;
  limitados.forEach(function(p) {
    quantidadeTotalDemandada += p.quantTotal || 0;
    valorTotalDemandado += p.demandaR || 0;
  });

  return {
    tela: 'Busca_Avançada',
    atualizadoEm: dados.atualizadoEm,
    markupUsado: markup,
    totalEncontrado: totalEncontrado,
    exibidos: limitados.length,
    quantidadeTotalDemandada: arredondar(quantidadeTotalDemandada, 4),
    valorTotalDemandado: arredondar(valorTotalDemandado, 4),
    sugestoes: [],
    resultados: limitados
  };
}

function mvpObterMapaDeProdutos(filtros) {
  const filtrosTratados = filtros || {};
  const markup = normalizarMarkup(filtrosTratados.markup);
  const limite = normalizarLimite(filtrosTratados.limite, MVP_MAX_RESULTADOS_MAPA);
  const clienteFiltro = mvpNormalizarTexto_(filtrosTratados.cliente || '');
  const termos = mvpColetarTermos_(filtrosTratados);

  const dados = mvpConsolidarDados_(markup);

  const clienteReal = clienteFiltro
    ? dados.clientes.filter(function(c) { return mvpNormalizarTexto_(c) === clienteFiltro; })[0]
    : null;

  let produtosFiltrados = dados.produtos.filter(function(p) {
    if (!mvpProdutoAtendeFiltros_(p, filtrosTratados, termos)) return false;
    if (clienteReal) {
      const cv = p.clientesValores[clienteReal];
      if (!cv || cv.quantidadeAtual <= 0) return false;
    }
    return true;
  });

  produtosFiltrados = produtosFiltrados.map(function(p) {
    if (!clienteReal) return p;

    const cv = p.clientesValores[clienteReal] || { quantidadeAtual: 0, demandaAtual: 0 };
    const clone = Object.assign({}, p);
    clone.quantTotal = arredondar(cv.quantidadeAtual, 4);
    clone.quantidadeTotalDemandada = arredondar(cv.quantidadeAtual, 4);
    clone.demandaR = arredondar(cv.demandaAtual, 4);
    clone.valorTotalDemandado = arredondar(cv.demandaAtual, 4);
    return clone;
  });

  produtosFiltrados.sort(function(a, b) {
    if ((b.quantTotal || 0) !== (a.quantTotal || 0)) return (b.quantTotal || 0) - (a.quantTotal || 0);
    return mvpNormalizarTexto_(a.produto).localeCompare(mvpNormalizarTexto_(b.produto));
  });

  const totalEncontrado = produtosFiltrados.length;
  const limitados = produtosFiltrados.slice(0, limite);

  let quantidadeTotalDemandada = 0;
  let valorTotalDemandado = 0;
  limitados.forEach(function(p) {
    quantidadeTotalDemandada += p.quantTotal || 0;
    valorTotalDemandado += p.demandaR || 0;
  });

  return {
    tela: 'Mapa_de_Produtos',
    atualizadoEm: dados.atualizadoEm,
    markupUsado: markup,
    totalEncontrado: totalEncontrado,
    exibidos: limitados.length,
    quantidadeTotalDemandada: arredondar(quantidadeTotalDemandada, 4),
    valorTotalDemandado: arredondar(valorTotalDemandado, 4),
    clientes: dados.clientes,
    produtos: limitados
  };
}

function mvpCalcularDashboard() {
  const dados = mvpConsolidarDados_(0.8);
  const produtos = dados.produtos;

  let totalSkuMasterSet = {};
  let quantidadeTotalDemandada = 0;
  let valorTotalDemandado = 0;
  let itensComPrecoCusto = 0;
  let itensSemPreco = 0;
  let itensSemSkuMaster = 0;

  const listaComSku = [];
  const listaSemSku = [];
  const listaSemPreco = [];

  produtos.forEach(function(p) {
    quantidadeTotalDemandada += p.quantTotal || 0;
    valorTotalDemandado += p.demandaR || 0;

    if (p.skuMaster) {
      if (!totalSkuMasterSet[p.skuMaster]) totalSkuMasterSet[p.skuMaster] = true;
      listaComSku.push(p);
    } else {
      itensSemSkuMaster++;
      listaSemSku.push(p);
    }

    if (p.precoCustoAtual && p.precoCustoAtual > 0) {
      itensComPrecoCusto++;
    } else {
      itensSemPreco++;
      listaSemPreco.push(p);
    }
  });

  const topQuantidade = produtos.slice()
    .sort(function(a, b) { return (b.quantTotal || 0) - (a.quantTotal || 0); })
    .slice(0, 30)
    .map(function(p) { return { skuMaster: p.skuMaster, produto: p.produto, quantidade: p.quantTotal || 0 }; });

  const topValor = produtos.slice()
    .sort(function(a, b) { return (b.demandaR || 0) - (a.demandaR || 0); })
    .slice(0, 30)
    .map(function(p) { return { skuMaster: p.skuMaster, produto: p.produto, valor: p.demandaR || 0 }; });

  const demandaPorCliente = {};
  produtos.forEach(function(p) {
    Object.keys(p.clientesValores || {}).forEach(function(cliente) {
      const cv = p.clientesValores[cliente];
      demandaPorCliente[cliente] = (demandaPorCliente[cliente] || 0) + (cv.demandaAtual || 0);
    });
  });

  const topClientes = Object.keys(demandaPorCliente)
    .map(function(c) { return { cliente: c, demanda: arredondar(demandaPorCliente[c], 4) }; })
    .sort(function(a, b) { return b.demanda - a.demanda; })
    .slice(0, 30);

  const semSku = listaSemSku.slice(0, 50).map(function(p) { return { produto: p.produto || p.skuMaster }; });
  const semPreco = listaSemPreco.slice(0, 50).map(function(p) { return { produto: p.produto, skuMaster: p.skuMaster }; });

  return {
    resumo: {
      totalSkuMaster: Object.keys(totalSkuMasterSet).length,
      totalProdutos: produtos.length,
      totalClientes: dados.clientes.length,
      quantidadeTotalDemandada: arredondar(quantidadeTotalDemandada, 4),
      valorTotalDemandado: arredondar(valorTotalDemandado, 2),
      itensComPrecoCusto: itensComPrecoCusto,
      itensSemPreco: itensSemPreco,
      itensSemSkuMaster: itensSemSkuMaster,
      itensCotadosNoMes: 0,
      valorCotadoNoMes: 0
    },
    topQuantidade: topQuantidade,
    topValor: topValor,
    topClientes: topClientes,
    semSku: semSku,
    semPreco: semPreco,
    ultimosPrecos: [],
    ultimasCotacoes: []
  };
}
