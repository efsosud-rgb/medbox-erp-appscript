/**
 * MEDBOX ERP | MÓDULO EXTRAIR COTAÇÃO
 *
 * Lógica para listagem de clientes e extração de cotações com cruzamento de custos.
 */

/**
 * Lista todos os clientes que possuem uma aba de banco de dados (BD_NomeCliente).
 * Filtra abas técnicas conhecidas.
 */
function listarClientesComCotacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const abasTecnicas = [
    'BD_PRECOS', 'BD_COTACOES', 'BD_SKU_MASTER_SEFAZ',
    'BD_IMPORTACOES_ARQUIVOS', 'BD_CONSOLIDADO', 'BD_INTS',
    'MAPA_ERP', 'AUX_BUSCA', 'LOG_PROCESSAMENTO', 'DASHBOARD'
  ];

  const clientes = sheets
    .map(s => s.getName())
    .filter(nome => nome.startsWith("BD_") && !abasTecnicas.includes(nome))
    .map(nome => nome.replace("BD_", ""));

  return clientes.sort();
}

/**
 * Extrai os itens da aba do cliente e cruza com BD_PRECOS para obter o custo.
 *
 * @param {string} nomeCliente
 * @return {Array} Lista de objetos com os dados da cotação e cálculos.
 */
function extrairCotacaoCliente(nomeCliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "BD_" + nomeCliente;
  const sheetCliente = ss.getSheetByName(nomeAba);
  const sheetPrecos = ss.getSheetByName("BD_PRECOS");

  if (!sheetCliente) throw new Error("Aba do cliente não encontrada: " + nomeAba);
  if (!sheetPrecos) throw new Error("Aba BD_PRECOS não encontrada.");

  const dataCliente = sheetCliente.getDataRange().getValues();
  const headersCliente = dataCliente[0];
  const idxSkuCli = headersCliente.indexOf("SKU_MASTER");
  const idxDescCli = headersCliente.indexOf("DESCRICAO_CLIENTE");
  const idxValorCli = headersCliente.indexOf("PRECO_REFERENCIA_CLIENTE");

  // Mapa de Custos (Menor preço por SKU)
  const dataPrecos = sheetPrecos.getDataRange().getValues();
  const headersPrecos = dataPrecos[0];
  const idxSkuPre = headersPrecos.indexOf("SKU_MASTER");
  const idxCustoPre = headersPrecos.indexOf("PRECO_CUSTO");

  const mapaCustos = {};
  for (let i = 1; i < dataPrecos.length; i++) {
    const sku = String(dataPrecos[i][idxSkuPre]).trim();
    const custo = parseFloat(dataPrecos[i][idxCustoPre]);
    if (sku && !isNaN(custo)) {
      if (!mapaCustos[sku] || custo < mapaCustos[sku]) {
        mapaCustos[sku] = custo;
      }
    }
  }

  const resultados = [];
  const markupPadrao = 80;

  for (let i = 1; i < dataCliente.length; i++) {
    const sku = String(dataCliente[i][idxSkuCli]).trim();
    const desc = dataCliente[i][idxDescCli] || "";
    const precoOsinfo = parseFloat(dataCliente[i][idxValorCli]) || 0;
    const precoCusto = mapaCustos[sku] || 0;

    // Cálculos
    const precoMedbox = precoCusto * (1 + markupPadrao / 100);
    const custoMaxLance80 = precoOsinfo > 0 ? (precoOsinfo - 0.05) / 1.80 : 0;
    const custoMaxLance50 = precoOsinfo > 0 ? (precoOsinfo - 0.05) / 1.50 : 0;

    resultados.push({
      skuMaster: sku,
      descricao: desc,
      precoOsinfo: precoOsinfo,
      markupPercent: markupPadrao,
      precoCusto: precoCusto,
      precoMedbox: Math.round(precoMedbox * 100) / 100,
      custoMax80: Math.round(custoMaxLance80 * 100) / 100,
      custoMax50: Math.round(custoMaxLance50 * 100) / 100
    });
  }

  return resultados;
}

/**
 * Helper para recalcular o preço MedBox baseado em novo markup.
 */
function recalcularPrecoMedbox(precoCusto, markupPercent) {
  const preco = precoCusto * (1 + markupPercent / 100);
  return Math.round(preco * 100) / 100;
}
