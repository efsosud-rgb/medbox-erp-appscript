/************************************************************
 * MEDBOX ERP | FASE 2D - DASHBOARD
 *
 * Calcula automaticamente:
 * - Top 30 SKU Master por Quantidade
 * - Top 30 SKU Master por Valor Demandado
 * - Produtos sem SKU Master
 * - Produtos sem Preço Custo
 * - Total de quantidade demandada
 * - Total de valor demandado
 *
 * Fonte principal:
 * - MAPA_ERP
 *
 * Não apaga BD_ antigas.
 ************************************************************/

const F2D_ABA_MAPA = 'MAPA_ERP';
const F2D_ABA_DASHBOARD = 'DASHBOARD';

function calcularDashboardFase2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapa = ss.getSheetByName(F2D_ABA_MAPA);

  if (!mapa) {
    throw new Error('Aba MAPA_ERP não encontrada.');
  }

  const values = mapa.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error('MAPA_ERP está vazio.');
  }

  const headers = values[0].map(String);
  const idx = f2dCriarIndice_(headers);

  const clientes = f2dObterClientesDoMapa_(headers);

  const porSku = new Map();
  const categoriasSet = new Set();
  const clientesSet = new Set();

  let totalProdutos = 0;
  let totalQuantidadeDemandada = 0;
  let totalValorDemandado = 0;
  let produtosSemSku = 0;
  let produtosSemCusto = 0;
  let produtosComCusto = 0;
  let produtosComDemanda = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const categoria = f2dUpper_(f2dGetCell_(row, idx, 'Categoria'));
    const skuMaster = f2dUpper_(f2dGetCell_(row, idx, 'SKU Master'));
    const skuOperacional = f2dUpper_(f2dGetCell_(row, idx, 'SKU Operacional'));
    const produto = f2dUpper_(f2dGetCell_(row, idx, 'Produto'));
    const produtoNormalizado = f2dUpper_(f2dGetCell_(row, idx, 'Produto Normalizado'));
    const clientesConsumo = f2dUpper_(f2dGetCell_(row, idx, 'Clientes com Consumo'));

    const quantTotal = f2dToNumber_(f2dGetCell_(row, idx, 'Quant. total')) || 0;
    const menorValorUnit = f2dToNumber_(f2dGetCell_(row, idx, 'Menor Valor Unit'));
    const precoMedioClientes = f2dToNumber_(f2dGetCell_(row, idx, 'Preço Médio Clientes'));
    const precoCusto = f2dToNumber_(f2dGetCell_(row, idx, 'Custo Compra 2026'));

    const valorUnitBase = menorValorUnit !== null
      ? menorValorUnit
      : precoMedioClientes;

    const valorDemandado = valorUnitBase !== null
      ? quantTotal * valorUnitBase
      : 0;

    totalProdutos++;

    if (categoria) {
      categoriasSet.add(categoria);
    }

    if (clientesConsumo) {
      clientesConsumo.split(/[;,]/).forEach(c => {
        const cliente = f2dUpper_(c);
        if (cliente) clientesSet.add(cliente);
      });
    }

    clientes.forEach(cliente => {
      const q = f2dToNumber_(f2dGetCell_(row, idx, `${cliente} Quant`)) || 0;
      if (q > 0) clientesSet.add(cliente);
    });

    if (!skuMaster) {
      produtosSemSku++;
    }

    if (precoCusto === null) {
      produtosSemCusto++;
    } else {
      produtosComCusto++;
    }

    if (quantTotal > 0) {
      produtosComDemanda++;
    }

    totalQuantidadeDemandada += quantTotal;
    totalValorDemandado += valorDemandado;

    const chaveSku = skuMaster || skuOperacional || produtoNormalizado || produto || 'SEM SKU';

    if (!porSku.has(chaveSku)) {
      porSku.set(chaveSku, {
        skuMaster: chaveSku,
        skuMasterOriginal: skuMaster,
        categoria: categoria,
        produtoExemplo: produto,
        quantidade: 0,
        valorDemandado: 0,
        menorPrecoCliente: null,
        precoMedioCliente: null,
        precoCusto: null,
        clientesSet: new Set(),
        linhas: 0,
        semSku: !skuMaster,
        semCusto: precoCusto === null
      });
    }

    const item = porSku.get(chaveSku);

    item.quantidade += quantTotal;
    item.valorDemandado += valorDemandado;
    item.linhas += 1;

    if (!item.categoria && categoria) item.categoria = categoria;
    if (!item.produtoExemplo && produto) item.produtoExemplo = produto;

    if (menorValorUnit !== null) {
      if (item.menorPrecoCliente === null || menorValorUnit < item.menorPrecoCliente) {
        item.menorPrecoCliente = menorValorUnit;
      }
    }

    if (precoMedioClientes !== null) {
      if (item.precoMedioCliente === null) {
        item.precoMedioCliente = precoMedioClientes;
      } else {
        item.precoMedioCliente = (item.precoMedioCliente + precoMedioClientes) / 2;
      }
    }

    if (precoCusto !== null) {
      if (item.precoCusto === null || precoCusto < item.precoCusto) {
        item.precoCusto = precoCusto;
      }
      item.semCusto = false;
    }

    if (clientesConsumo) {
      clientesConsumo.split(/[;,]/).forEach(c => {
        const cliente = f2dUpper_(c);
        if (cliente) item.clientesSet.add(cliente);
      });
    }

    clientes.forEach(cliente => {
      const q = f2dToNumber_(f2dGetCell_(row, idx, `${cliente} Quant`)) || 0;
      if (q > 0) item.clientesSet.add(cliente);
    });
  }

  const listaSku = Array.from(porSku.values()).map(item => {
    const precoMedbox50 = item.precoCusto !== null
      ? f2dArredondar_(item.precoCusto * 1.5, 4)
      : '';

    let status50 = 'SEM CUSTO';

    if (item.precoCusto !== null && item.menorPrecoCliente !== null) {
      status50 = precoMedbox50 <= item.menorPrecoCliente
        ? 'GANHANDO ITEM'
        : 'ACIMA DO CLIENTE';
    } else if (item.precoCusto !== null && item.menorPrecoCliente === null) {
      status50 = 'SEM PREÇO CLIENTE';
    }

    return {
      skuMaster: item.skuMaster,
      skuMasterOriginal: item.skuMasterOriginal,
      categoria: item.categoria,
      produtoExemplo: item.produtoExemplo,
      quantidade: f2dArredondar_(item.quantidade, 4),
      valorDemandado: f2dArredondar_(item.valorDemandado, 4),
      menorPrecoCliente: item.menorPrecoCliente,
      precoMedioCliente: item.precoMedioCliente,
      precoCusto: item.precoCusto,
      precoMedbox50: precoMedbox50,
      status50: status50,
      clientes: Array.from(item.clientesSet).sort().join('; '),
      qtdClientes: item.clientesSet.size,
      linhas: item.linhas,
      semSku: item.semSku,
      semCusto: item.semCusto
    };
  });

  const topQuantidade = listaSku
    .slice()
    .sort((a, b) => {
      if ((b.quantidade || 0) !== (a.quantidade || 0)) return (b.quantidade || 0) - (a.quantidade || 0);
      return (b.valorDemandado || 0) - (a.valorDemandado || 0);
    })
    .slice(0, 30);

  const topValor = listaSku
    .slice()
    .sort((a, b) => {
      if ((b.valorDemandado || 0) !== (a.valorDemandado || 0)) return (b.valorDemandado || 0) - (a.valorDemandado || 0);
      return (b.quantidade || 0) - (a.quantidade || 0);
    })
    .slice(0, 30);

  const semSku = listaSku
    .filter(x => x.semSku)
    .sort((a, b) => (b.valorDemandado || 0) - (a.valorDemandado || 0))
    .slice(0, 50);

  const semCusto = listaSku
    .filter(x => x.semCusto)
    .sort((a, b) => (b.valorDemandado || 0) - (a.valorDemandado || 0))
    .slice(0, 50);

  const competitivos50 = listaSku
    .filter(x => x.status50 === 'GANHANDO ITEM')
    .sort((a, b) => (b.valorDemandado || 0) - (a.valorDemandado || 0))
    .slice(0, 50);

  const acimaCliente50 = listaSku
    .filter(x => x.status50 === 'ACIMA DO CLIENTE')
    .sort((a, b) => (b.valorDemandado || 0) - (a.valorDemandado || 0))
    .slice(0, 50);

  return {
    atualizadoEm: f2dFormatarDataHora_(new Date()),
    resumo: {
      totalProdutos: totalProdutos,
      totalSkuAgrupados: listaSku.length,
      totalQuantidadeDemandada: f2dArredondar_(totalQuantidadeDemandada, 4),
      totalValorDemandado: f2dArredondar_(totalValorDemandado, 4),
      produtosSemSku: produtosSemSku,
      produtosSemCusto: produtosSemCusto,
      produtosComCusto: produtosComCusto,
      produtosComDemanda: produtosComDemanda,
      clientesAtivos: clientesSet.size,
      categoriasAtivas: categoriasSet.size,
      skusCompetitivos50: competitivos50.length,
      skusAcimaCliente50: acimaCliente50.length
    },
    topQuantidade: topQuantidade,
    topValor: topValor,
    semSku: semSku,
    semCusto: semCusto,
    competitivos50: competitivos50,
    acimaCliente50: acimaCliente50
  };
}

function atualizarDashboardFase2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(F2D_ABA_DASHBOARD) || ss.insertSheet(F2D_ABA_DASHBOARD);

  const data = calcularDashboardFase2();

  ws.clear();

  ws.getRange('A1:J1').merge();
  ws.getRange('A1').setValue('MEDBOX ERP | DASHBOARD COMERCIAL');
  ws.getRange('A1')
    .setBackground('#002E56')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  ws.getRange('A2:J2').merge();
  ws.getRange('A2').setValue('ATUALIZADO EM: ' + data.atualizadoEm);
  ws.getRange('A2')
    .setBackground('#EEF1F4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  escreverResumoDashboard_(ws, data);
  escreverTopQuantidade_(ws, data.topQuantidade);
  escreverTopValor_(ws, data.topValor);
  escreverSemCusto_(ws, data.semCusto);
  escreverSemSku_(ws, data.semSku);

  ws.setFrozenRows(2);

  ws.setColumnWidth(1, 260);
  ws.setColumnWidth(2, 180);
  ws.setColumnWidth(3, 40);
  ws.setColumnWidth(4, 360);
  ws.setColumnWidth(5, 130);
  ws.setColumnWidth(6, 150);
  ws.setColumnWidth(7, 40);
  ws.setColumnWidth(8, 360);
  ws.setColumnWidth(9, 130);
  ws.setColumnWidth(10, 150);

  f2dRegistrarLog_('DASHBOARD FASE 2D ATUALIZADO.');

  SpreadsheetApp.getUi().alert('Dashboard Fase 2D atualizado com sucesso.');
}

function escreverResumoDashboard_(ws, data) {
  const resumo = data.resumo;

  const rows = [
    ['INDICADOR', 'VALOR'],
    ['TOTAL DE PRODUTOS NO MAPA', resumo.totalProdutos],
    ['TOTAL DE SKU AGRUPADOS', resumo.totalSkuAgrupados],
    ['QUANTIDADE TOTAL DEMANDADA', resumo.totalQuantidadeDemandada],
    ['VALOR TOTAL DEMANDADO', resumo.totalValorDemandado],
    ['PRODUTOS SEM SKU MASTER', resumo.produtosSemSku],
    ['PRODUTOS SEM PREÇO CUSTO', resumo.produtosSemCusto],
    ['PRODUTOS COM PREÇO CUSTO', resumo.produtosComCusto],
    ['PRODUTOS COM DEMANDA', resumo.produtosComDemanda],
    ['CLIENTES ATIVOS', resumo.clientesAtivos],
    ['CATEGORIAS ATIVAS', resumo.categoriasAtivas],
    ['SKU COMPETITIVOS COM 50%', resumo.skusCompetitivos50],
    ['SKU ACIMA DO CLIENTE COM 50%', resumo.skusAcimaCliente50]
  ];

  ws.getRange(4, 1, rows.length, 2).setValues(rows);

  f2dFormatarCabecalho_(ws.getRange(4, 1, 1, 2), '#D80024', '#FFFFFF');
  ws.getRange(5, 1, rows.length - 1, 1).setFontWeight('bold');
  ws.getRange(5, 2, rows.length - 1, 1).setBackground('#FFF5C7');

  ws.getRange(8, 2).setNumberFormat('#,##0.00');
  ws.getRange(9, 2).setNumberFormat('R$ #,##0.00');
}

function escreverTopQuantidade_(ws, lista) {
  const startRow = 4;
  const startCol = 4;

  const headers = [
    'TOP 30 SKU MASTER POR QUANTIDADE',
    'QUANTIDADE',
    'VALOR DEMANDADO'
  ];

  ws.getRange(startRow, startCol, 1, headers.length).setValues([headers]);
  f2dFormatarCabecalho_(ws.getRange(startRow, startCol, 1, headers.length), '#002E56', '#FFFFFF');

  const rows = lista.map(item => [
    item.skuMaster,
    item.quantidade,
    item.valorDemandado
  ]);

  if (rows.length > 0) {
    ws.getRange(startRow + 1, startCol, rows.length, headers.length).setValues(rows);
    ws.getRange(startRow + 1, startCol + 1, rows.length, 1).setNumberFormat('#,##0.00');
    ws.getRange(startRow + 1, startCol + 2, rows.length, 1).setNumberFormat('R$ #,##0.00');
  }
}

function escreverTopValor_(ws, lista) {
  const startRow = 4;
  const startCol = 8;

  const headers = [
    'TOP 30 SKU MASTER POR VALOR',
    'VALOR',
    'QUANTIDADE'
  ];

  ws.getRange(startRow, startCol, 1, headers.length).setValues([headers]);
  f2dFormatarCabecalho_(ws.getRange(startRow, startCol, 1, headers.length), '#002E56', '#FFFFFF');

  const rows = lista.map(item => [
    item.skuMaster,
    item.valorDemandado,
    item.quantidade
  ]);

  if (rows.length > 0) {
    ws.getRange(startRow + 1, startCol, rows.length, headers.length).setValues(rows);
    ws.getRange(startRow + 1, startCol + 1, rows.length, 1).setNumberFormat('R$ #,##0.00');
    ws.getRange(startRow + 1, startCol + 2, rows.length, 1).setNumberFormat('#,##0.00');
  }
}

function escreverSemCusto_(ws, lista) {
  const startRow = 38;
  const startCol = 4;

  const headers = [
    'TOP PRODUTOS SEM PREÇO CUSTO',
    'VALOR DEMANDADO',
    'QUANTIDADE'
  ];

  ws.getRange(startRow, startCol, 1, headers.length).setValues([headers]);
  f2dFormatarCabecalho_(ws.getRange(startRow, startCol, 1, headers.length), '#D80024', '#FFFFFF');

  const rows = lista.slice(0, 30).map(item => [
    item.skuMaster,
    item.valorDemandado,
    item.quantidade
  ]);

  if (rows.length > 0) {
    ws.getRange(startRow + 1, startCol, rows.length, headers.length).setValues(rows);
    ws.getRange(startRow + 1, startCol + 1, rows.length, 1).setNumberFormat('R$ #,##0.00');
    ws.getRange(startRow + 1, startCol + 2, rows.length, 1).setNumberFormat('#,##0.00');
  }
}

function escreverSemSku_(ws, lista) {
  const startRow = 38;
  const startCol = 8;

  const headers = [
    'TOP PRODUTOS SEM SKU MASTER',
    'VALOR DEMANDADO',
    'QUANTIDADE'
  ];

  ws.getRange(startRow, startCol, 1, headers.length).setValues([headers]);
  f2dFormatarCabecalho_(ws.getRange(startRow, startCol, 1, headers.length), '#D80024', '#FFFFFF');

  const rows = lista.slice(0, 30).map(item => [
    item.produtoExemplo || item.skuMaster,
    item.valorDemandado,
    item.quantidade
  ]);

  if (rows.length > 0) {
    ws.getRange(startRow + 1, startCol, rows.length, headers.length).setValues(rows);
    ws.getRange(startRow + 1, startCol + 1, rows.length, 1).setNumberFormat('R$ #,##0.00');
    ws.getRange(startRow + 1, startCol + 2, rows.length, 1).setNumberFormat('#,##0.00');
  }
}

/************************************************************
 * HELPERS
 ************************************************************/

function f2dObterClientesDoMapa_(headers) {
  return headers
    .filter(h => String(h).trim().endsWith(' Quant'))
    .map(h => f2dUpper_(String(h).replace(' Quant', '').trim()))
    .filter(Boolean)
    .sort();
}

function f2dCriarIndice_(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    idx[String(h).trim()] = i;
  });

  return idx;
}

function f2dGetCell_(row, idx, nomeColuna) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) return '';

  return row[pos];
}

function f2dLimpar_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .trim()
    .replace(/\s+/g, ' ');
}

function f2dUpper_(valor) {
  return f2dLimpar_(valor).toUpperCase();
}

function f2dToNumber_(valor) {
  if (valor === null || valor === undefined || valor === '') return null;

  if (typeof valor === 'number') {
    if (isNaN(valor)) return null;
    return valor;
  }

  let txt = String(valor).trim();

  if (!txt) return null;

  txt = txt
    .replace(/R\$/gi, '')
    .replace(/%/g, '')
    .replace(/\s+/g, '');

  if (txt.includes(',')) {
    txt = txt.replace(/\./g, '').replace(',', '.');
  }

  txt = txt.replace(/[^0-9.\-]/g, '');

  if (!txt) return null;

  const n = Number(txt);

  return isNaN(n) ? null : n;
}

function f2dArredondar_(n, casas) {
  if (n === null || n === undefined || isNaN(n)) return '';

  const f = Math.pow(10, casas || 2);

  return Math.round(n * f) / f;
}

function f2dFormatarDataHora_(data) {
  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );
}

function f2dFormatarCabecalho_(range, corFundo, corFonte) {
  range
    .setBackground(corFundo)
    .setFontColor(corFonte)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function f2dRegistrarLog_(mensagem) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('LOG_PROCESSAMENTO');

  if (!ws) {
    ws = ss.insertSheet('LOG_PROCESSAMENTO');
    ws.appendRow(['DATA_HORA', 'EVENTO']);
  }

  if (ws.getLastRow() === 0) {
    ws.appendRow(['DATA_HORA', 'EVENTO']);
  }

  ws.appendRow([
    f2dFormatarDataHora_(new Date()),
    mensagem
  ]);
}