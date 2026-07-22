/************************************************************
 * MEDBOX ERP | FASE 2D - DASHBOARD
 *
 * Calcula automaticamente de acordo com o contrato do Web App:
 * - Resumo Comercial (SKU, Quantidade, Valor, Com/Sem Preço etc.)
 * - Rankings e Listas Comerciais
 *
 * Fonte principal:
 * - MAPA_ERP
 ************************************************************/

const F2D_ABA_MAPA = 'MAPA_ERP';
const F2D_ABA_DASHBOARD = 'DASHBOARD';

function calcularDashboardFase2() {
  const ss = obterPlanilhaMedbox_();
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

  const porSku = new Map();

  let totalSkuMasterSet = new Set();
  let totalQuantidadeDemandada = 0;
  let totalValorDemandado = 0;
  let produtosSemSku = 0;
  let produtosSemCusto = 0;
  let produtosComCusto = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const categoria = f2dUpper_(f2dGetCell_(row, idx, 'Categoria'));
    const skuMaster = f2dUpper_(f2dGetCell_(row, idx, 'SKU Master'));
    const skuOperacional = f2dUpper_(f2dGetCell_(row, idx, 'SKU Operacional'));
    const produto = f2dUpper_(f2dGetCell_(row, idx, 'Produto'));
    const produtoNormalizado = f2dUpper_(f2dGetCell_(row, idx, 'Produto Normalizado'));

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

    if (skuMaster) {
      totalSkuMasterSet.add(skuMaster);
    } else {
      produtosSemSku++;
    }

    if (precoCusto === null) {
      produtosSemCusto++;
    } else {
      produtosComCusto++;
    }

    totalQuantidadeDemandada += quantTotal;
    totalValorDemandado += valorDemandado;

    const chaveSku = skuMaster || skuOperacional || produtoNormalizado || produto || 'SEM SKU';

    if (!porSku.has(chaveSku)) {
      porSku.set(chaveSku, {
        skuMaster: chaveSku,
        produto: produto,
        quantidade: 0,
        valorDemandado: 0,
        semSku: !skuMaster,
        semCusto: precoCusto === null
      });
    }

    const item = porSku.get(chaveSku);
    item.quantidade += quantTotal;
    item.valorDemandado += valorDemandado;
  }

  const listaSku = Array.from(porSku.values());

  const topQuantidade = listaSku
    .slice()
    .sort((a, b) => b.quantidade - a.quantidade)
    .slice(0, 30)
    .map(x => ({ skuMaster: x.skuMaster, produto: x.produto, quantidade: f2dArredondar_(x.quantidade, 4) }));

  const topValor = listaSku
    .slice()
    .sort((a, b) => b.valorDemandado - a.valorDemandado)
    .slice(0, 30)
    .map(x => ({ skuMaster: x.skuMaster, produto: x.produto, valor: f2dArredondar_(x.valorDemandado, 4) }));

  const semSku = listaSku
    .filter(x => x.semSku)
    .sort((a, b) => b.quantidade - a.quantidade)
    .slice(0, 50)
    .map(x => ({ produto: x.produto || x.skuMaster }));

  const semPreco = listaSku
    .filter(x => x.semCusto)
    .sort((a, b) => b.quantidade - a.quantidade)
    .slice(0, 50)
    .map(x => ({ produto: x.produto, skuMaster: x.skuMaster }));

  return {
    resumo: {
      totalSkuMaster: totalSkuMasterSet.size,
      quantidadeTotalDemandada: f2dArredondar_(totalQuantidadeDemandada, 4),
      valorTotalDemandado: f2dArredondar_(totalValorDemandado, 2),
      itensComPrecoCusto: produtosComCusto,
      itensSemPreco: produtosSemCusto,
      itensSemSkuMaster: produtosSemSku,
      itensCotadosNoMes: 0, // Placeholder
      valorCotadoNoMes: 0   // Placeholder
    },
    topQuantidade: topQuantidade,
    topValor: topValor,
    topClientes: [],       // Placeholder
    semSku: semSku,
    semPreco: semPreco,
    ultimosPrecos: [],     // Placeholder
    ultimasCotacoes: []    // Placeholder
  };
}

function atualizarDashboardFase2() {
  const ss = obterPlanilhaMedbox_();
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
  ws.getRange('A2').setValue('ATUALIZADO EM: ' + f2dFormatarDataHora_(new Date()));
  ws.getRange('A2')
    .setBackground('#EEF1F4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Re-write summary and charts for sheet dashboard fallback
  const resumo = data.resumo;
  const rows = [
    ['INDICADOR', 'VALOR'],
    ['TOTAL DE PRODUTOS NO MAPA', resumo.totalSkuMaster + resumo.itensSemSkuMaster],
    ['TOTAL DE SKU MASTER', resumo.totalSkuMaster],
    ['QUANTIDADE TOTAL DEMANDADA', resumo.quantidadeTotalDemandada],
    ['VALOR TOTAL DEMANDADO', resumo.valorTotalDemandado],
    ['PRODUTOS SEM SKU MASTER', resumo.itensSemSkuMaster],
    ['PRODUTOS SEM PREÇO CUSTO', resumo.itensSemPreco],
    ['PRODUTOS COM PREÇO CUSTO', resumo.itensComPrecoCusto]
  ];

  ws.getRange(4, 1, rows.length, 2).setValues(rows);
  f2dFormatarCabecalho_(ws.getRange(4, 1, 1, 2), '#D80024', '#FFFFFF');
  ws.getRange(5, 1, rows.length - 1, 1).setFontWeight('bold');
  ws.getRange(5, 2, rows.length - 1, 1).setBackground('#FFF5C7');

  f2dRegistrarLog_('DASHBOARD FASE 2D ATUALIZADO.');

  SpreadsheetApp.getUi().alert('Dashboard Fase 2D atualizado com sucesso.');
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
  const ss = obterPlanilhaMedbox_();
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