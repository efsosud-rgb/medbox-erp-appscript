/**
 * MEDBOX ERP | MÓDULO LANÇAR COTAÇÃO
 *
 * Este arquivo contém a lógica para salvar e atualizar cotações de clientes.
 * Arquitetura: Centralizada na aba BD_COTACOES.
 */

/**
 * Salva um lote de itens de cotação. Realiza upsert por (CLIENTE, SKU_MASTER).
 *
 * @param {Object} payload { cliente: string, itens: [{skuMaster, descricao, quantidade, valorUnit}, ...] }
 * @return {Object} Resumo da operação { sucesso: boolean, clienteNovo: boolean, atualizados: number, novos: number, total: number }
 */
function salvarCotacaoLote(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaNome = "BD_COTACOES";
  const ws = ss.getSheetByName(abaNome);

  if (!ws) throw new Error("Aba " + abaNome + " não encontrada.");

  const nomeCliente = payload.cliente.trim().toUpperCase();
  const dataAba = ws.getDataRange().getValues();
  const headers = dataAba[0];

  const idxCliente = headers.indexOf("CLIENTE");
  const idxSku = headers.indexOf("SKU_MASTER");
  const idxQtd = headers.indexOf("QUANTIDADE");
  const idxValor = headers.indexOf("PRECO_REFERENCIA_CLIENTE");
  const idxDesc = headers.indexOf("DESCRICAO_CLIENTE");
  const idxData = headers.indexOf("DATA_ATUALIZACAO");
  const idxId = headers.indexOf("ID_COTACAO");
  const idxStatus = headers.indexOf("STATUS_COTACAO");

  if (idxCliente === -1 || idxSku === -1) {
    throw new Error("Colunas obrigatórias (CLIENTE ou SKU_MASTER) não encontradas em BD_COTACOES.");
  }

  let atualizados = 0;
  let novos = 0;
  const agora = new Date();

  // Verifica se o cliente já existia na base para o resumo
  let clienteExistente = false;

  // Mapeia chaves (CLIENTE|SKU) existentes para suas linhas (1-based)
  const entryMap = {};
  for (let i = 1; i < dataAba.length; i++) {
    const cli = String(dataAba[i][idxCliente]).trim().toUpperCase();
    if (cli === nomeCliente) clienteExistente = true;

    const sku = String(dataAba[i][idxSku]).trim().toUpperCase();
    if (cli && sku) {
      const chave = cli + "|" + sku;
      entryMap[chave] = i + 1;
    }
  }

  payload.itens.forEach(item => {
    const sku = String(item.skuMaster).trim().toUpperCase();
    const chave = nomeCliente + "|" + sku;
    const linhaExistente = entryMap[chave];

    if (linhaExistente) {
      // Atualiza
      if (idxQtd !== -1) ws.getRange(linhaExistente, idxQtd + 1).setValue(item.quantidade);
      if (idxValor !== -1) ws.getRange(linhaExistente, idxValor + 1).setValue(item.valorUnit);
      if (idxData !== -1) ws.getRange(linhaExistente, idxData + 1).setValue(agora);
      atualizados++;
    } else {
      // Adiciona nova linha respeitando os cabeçalhos
      const novaLinha = new Array(headers.length).fill("");
      novaLinha[idxCliente] = nomeCliente;
      novaLinha[idxSku] = sku;
      if (idxDesc !== -1) novaLinha[idxDesc] = item.descricao;
      if (idxQtd !== -1) novaLinha[idxQtd] = item.quantidade;
      if (idxValor !== -1) novaLinha[idxValor] = item.valorUnit;
      if (idxData !== -1) novaLinha[idxData] = agora;

      if (idxId !== -1) {
        novaLinha[idxId] = "COT-" + Utilities.formatDate(agora, "GMT-3", "yyyyMMdd") + "-" + Math.floor(Math.random()*1000);
      }
      if (idxStatus !== -1) novaLinha[idxStatus] = "NOVA";

      ws.appendRow(novaLinha);
      novos++;
    }
  });

  return {
    sucesso: true,
    clienteNovo: !clienteExistente,
    atualizados: atualizados,
    novos: novos,
    total: payload.itens.length
  };
}
