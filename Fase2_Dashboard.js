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
  // MVP emergencial: consome exclusivamente DadosMVP.js (BD_SKU MASTER, BD_PRECOS, BD_<CLIENTE>).
  // Nao le MAPA_ERP nem AUX_BUSCA diretamente.
  return mvpCalcularDashboard();
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