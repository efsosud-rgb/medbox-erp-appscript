/**
 * MEDBOX ERP | MÓDULO EXTRAIR COTAÇÃO
 *
 * Lógica para listagem de clientes e extração de cotações com cruzamento de custos.
 * Arquitetura: Centralizada na aba BD_COTACOES.
 */

/**
 * Lista todos os nomes únicos de clientes presentes na aba BD_COTACOES.
 */
function listarClientesComCotacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("BD_COTACOES");

  if (!ws) return [];

  const data = ws.getDataRange().getValues();
  const headers = data[0];
  const idxCliente = headers.indexOf("CLIENTE");

  if (idxCliente === -1) return [];

  const clientesSet = new Set();
  for (let i = 1; i < data.length; i++) {
    const nome = String(data[i][idxCliente]).trim();
    if (nome) clientesSet.add(nome.toUpperCase());
  }

  return Array.from(clientesSet).sort();
}

/**
 * Extrai os itens de um cliente da aba BD_COTACOES e cruza com BD_PRECOS para obter o custo.
 *
 * @param {string} nomeCliente
 * @return {Array} Lista de objetos com os dados da cotação e cálculos.
 */
function extrairCotacaoCliente(nomeCliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsCot = ss.getSheetByName("BD_COTACOES");
  const sheetPrecos = ss.getSheetByName("BD_PRECOS");

  if (!wsCot) throw new Error("Aba BD_COTACOES não encontrada.");
  if (!sheetPrecos) throw new Error("Aba BD_PRECOS não encontrada.");

  const dataCot = wsCot.getDataRange().getValues();
  const headersCot = dataCot[0];
  const idxCli = headersCot.indexOf("CLIENTE");
  const idxSkuCot = headersCot.indexOf("SKU_MASTER");
  const idxDescCot = headersCot.indexOf("DESCRICAO_CLIENTE");
  const idxValorCot = headersCot.indexOf("PRECO_REFERENCIA_CLIENTE");

  const filtroCliente = nomeCliente.trim().toUpperCase();

  // Mapa de Custos (Menor preço por SKU na BD_PRECOS)
  const dataPrecos = sheetPrecos.getDataRange().getValues();
  const headersPrecos = dataPrecos[0];
  const idxSkuPre = headersPrecos.indexOf("SKU_MASTER");
  const idxCustoPre = headersPrecos.indexOf("PRECO_CUSTO");

  const mapaCustos = {};
  for (let i = 1; i < dataPrecos.length; i++) {
    const sku = String(dataPrecos[i][idxSkuPre]).trim().toUpperCase();
    const custo = parseFloat(dataPrecos[i][idxCustoPre]);
    if (sku && !isNaN(custo)) {
      if (!mapaCustos[sku] || custo < mapaCustos[sku]) {
        mapaCustos[sku] = custo;
      }
    }
  }

  const resultados = [];
  const markupPadrao = 80;

  for (let i = 1; i < dataCot.length; i++) {
    const cliRow = String(dataCot[i][idxCli]).trim().toUpperCase();
    if (cliRow !== filtroCliente) continue;

    const sku = String(dataCot[i][idxSkuCot]).trim().toUpperCase();
    const desc = dataCot[i][idxDescCot] || "";
    const precoOsinfo = parseFloat(dataCot[i][idxValorCot]) || 0;
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
