/**
 * MEDBOX ERP | MÓDULO LANÇAR COTAÇÃO
 *
 * Este arquivo contém a lógica para salvar e atualizar cotações de clientes
 * em suas respectivas abas BD_<Cliente>.
 */

/**
 * Salva um lote de itens de cotação para um cliente específico.
 * Realiza upsert por SKU_MASTER.
 *
 * @param {Object} payload { cliente: string, itens: [{skuMaster, descricao, quantidade, valorUnit}, ...] }
 * @return {Object} Resumo da operação { sucesso: boolean, clienteNovo: boolean, atualizados: number, novos: number, total: number }
 */
function salvarCotacaoLote(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeCliente = payload.cliente.trim();
  const nomeAba = "BD_" + nomeCliente;
  let clienteNovo = false;

  let sheet = ss.getSheetByName(nomeAba);

  // Se a aba não existir, cria baseada no layout de BD_COTACOES
  if (!sheet) {
    clienteNovo = true;
    sheet = ss.insertSheet(nomeAba);
    const templateSheet = ss.getSheetByName("BD_COTACOES");
    if (templateSheet) {
      const headers = templateSheet.getRange(1, 1, 1, templateSheet.getLastColumn()).getValues();
      sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
      // Aplica formatação básica de cabeçalho
      sheet.getRange(1, 1, 1, headers[0].length)
        .setBackground('#002E56')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    } else {
      // Fallback caso BD_COTACOES não exista (improvável no projeto)
      const defaultHeaders = [["SKU_MASTER", "DESCRICAO_CLIENTE", "QUANTIDADE", "PRECO_REFERENCIA_CLIENTE", "DATA_ATUALIZACAO"]];
      sheet.getRange(1, 1, 1, defaultHeaders[0].length).setValues(defaultHeaders);
    }
  }

  const dataAba = sheet.getDataRange().getValues();
  const headers = dataAba[0];
  const idxSku = headers.indexOf("SKU_MASTER");
  const idxQtd = headers.indexOf("QUANTIDADE");
  const idxValor = headers.indexOf("PRECO_REFERENCIA_CLIENTE");
  const idxDesc = headers.indexOf("DESCRICAO_CLIENTE");
  const idxData = headers.indexOf("DATA_ATUALIZACAO");

  if (idxSku === -1) throw new Error("Coluna SKU_MASTER não encontrada na aba " + nomeAba);

  let atualizados = 0;
  let novos = 0;
  const agora = new Date();

  // Mapeia SKUs existentes para suas linhas (1-based)
  const skuMap = {};
  for (let i = 1; i < dataAba.length; i++) {
    const sku = String(dataAba[i][idxSku]).trim();
    if (sku) skuMap[sku] = i + 1;
  }

  payload.itens.forEach(item => {
    const sku = String(item.skuMaster).trim();
    const linhaExistente = skuMap[sku];

    if (linhaExistente) {
      // Atualiza
      if (idxQtd !== -1) sheet.getRange(linhaExistente, idxQtd + 1).setValue(item.quantidade);
      if (idxValor !== -1) sheet.getRange(linhaExistente, idxValor + 1).setValue(item.valorUnit);
      if (idxData !== -1) sheet.getRange(linhaExistente, idxData + 1).setValue(agora);
      atualizados++;
    } else {
      // Adiciona nova linha respeitando os cabeçalhos
      const novaLinha = new Array(headers.length).fill("");
      novaLinha[idxSku] = sku;
      if (idxDesc !== -1) novaLinha[idxDesc] = item.descricao;
      if (idxQtd !== -1) novaLinha[idxQtd] = item.quantidade;
      if (idxValor !== -1) novaLinha[idxValor] = item.valorUnit;
      if (idxData !== -1) novaLinha[idxData] = agora;

      // Campos técnicos de BD_COTACOES se existirem
      const idxId = headers.indexOf("ID_COTACAO");
      if (idxId !== -1) novaLinha[idxId] = "COT-" + Utilities.formatDate(agora, "GMT-3", "yyyyMMdd") + "-" + Math.floor(Math.random()*1000);

      const idxStatus = headers.indexOf("STATUS_COTACAO");
      if (idxStatus !== -1) novaLinha[idxStatus] = "NOVA";

      sheet.appendRow(novaLinha);
      novos++;
    }
  });

  return {
    sucesso: true,
    clienteNovo: clienteNovo,
    atualizados: atualizados,
    novos: novos,
    total: payload.itens.length
  };
}
