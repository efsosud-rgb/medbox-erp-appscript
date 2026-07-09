/************************************************************
 * MEDBOX ERP | EXTRAIR COTAÇÃO POR CLIENTE
 *
 * Arquivo: ExtrairCotacao.js
 *
 * REGRAS (conforme definido por Eric em 09/07/2026):
 * - Markup padrão: 80% (editável por linha na tela)
 * - Preço MedBox = Preço Custo x (1 + Markup%)
 * - Preço Custo Máx. p/ Lance = (Preço OSINFO - 0,05) / (1 + Markup%)
 *   calculado para 80% e para 50% de markup, independente do
 *   markup escolhido na linha (são referências fixas lado a lado)
 * - Preço Custo vem de BD_PRECOS, cruzando pelo SKU_MASTER
 * - Preço OSINFO = valor da coluna VALOR_UNIT da aba BD_<cliente>
 *   (nome do campo pode precisar de ajuste conforme layout real)
 ************************************************************/
 
const REC_PREFIXO_BD_CLIENTE = 'BD_';
const REC_ABAS_BD_EXCLUIDAS = [
  'BD_PRECOS', 'BD_COTACOES', 'BD_SKU_MASTER_SEFAZ',
  'BD_IMPORTACOES_ARQUIVOS', 'BD_CONSOLIDADO', 'BD_INTS'
];
const REC_MARKUP_PADRAO = 0.80;
 
/**
 * Lista os nomes de clientes que têm aba BD_<cliente> própria
 * (exclui as abas técnicas do sistema).
 * @returns {string[]} nomes de cliente (sem o prefixo BD_)
 */
function listarClientesComCotacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todasAbas = ss.getSheets().map(s => s.getName());
 
  return todasAbas
    .filter(nome => nome.indexOf(REC_PREFIXO_BD_CLIENTE) === 0)
    .filter(nome => REC_ABAS_BD_EXCLUIDAS.indexOf(nome) === -1)
    .map(nome => nome.substring(REC_PREFIXO_BD_CLIENTE.length));
}
 
/**
 * Extrai os itens de cotação de um cliente, cruzando com
 * BD_PRECOS pelo SKU_MASTER para trazer o Preço Custo e
 * calcular Preço MedBox e os preços-limite de lance.
 * @param {string} nomeCliente
 * @returns {Object[]} lista de itens com todos os campos calculados
 */
function extrairCotacaoCliente(nomeCliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = REC_PREFIXO_BD_CLIENTE + rlcSanitizarNomeAba_(nomeCliente);
  const aba = ss.getSheetByName(nomeAba);
 
  if (!aba) {
    throw new Error('Aba ' + nomeAba + ' não encontrada.');
  }
 
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(String);
  const idx = recCriarIndice_(headers);
 
  if (idx['SKU_MASTER'] === undefined) {
    throw new Error('Coluna SKU_MASTER não encontrada em ' + nomeAba + '.');
  }
 
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return [];
 
  const dadosCliente = aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn()).getValues();
 
  // Monta mapa SKU_MASTER -> Preço Custo a partir de BD_PRECOS
  const mapaPrecoCusto = recMontarMapaPrecoCusto_(ss);
 
  const resultado = [];
 
  dadosCliente.forEach(linha => {
    const skuMaster = recGetCell_(linha, idx, 'SKU_MASTER');
    if (!skuMaster) return;
 
    const descricao = recGetCell_(linha, idx, 'DESCRICAO');
    const precoOsinfo = Number(recGetCell_(linha, idx, 'VALOR_UNIT')) || 0;
 
    const precoCusto = mapaPrecoCusto[recNorm_(skuMaster)] || null;
 
    const markup = REC_MARKUP_PADRAO;
    const precoMedbox = precoCusto !== null ? recArredondar_(precoCusto * (1 + markup)) : null;
 
    const custoMax80 = recArredondar_((precoOsinfo - 0.05) / (1 + 0.80));
    const custoMax50 = recArredondar_((precoOsinfo - 0.05) / (1 + 0.50));
 
    resultado.push({
      skuMaster: skuMaster,
      descricao: descricao,
      precoOsinfo: precoOsinfo,
      markupPercent: markup * 100,
      precoCusto: precoCusto,
      precoMedbox: precoMedbox,
      custoMaxLance80: custoMax80,
      custoMaxLance50: custoMax50,
      precoCustoDisponivel: precoCusto !== null
    });
  });
 
  return resultado;
}
 
/**
 * Recalcula Preço MedBox para um item específico quando o
 * usuário edita o % de Markup na tela (chamada individual,
 * sem precisar reprocessar a lista inteira).
 */
function recalcularPrecoMedbox(precoCusto, markupPercent) {
  if (precoCusto === null || precoCusto === undefined || precoCusto === '') return null;
  return recArredondar_(Number(precoCusto) * (1 + Number(markupPercent) / 100));
}
 
/************************************************************
 * HELPERS
 ************************************************************/
 
function recMontarMapaPrecoCusto_(ss) {
  const aba = ss.getSheetByName('BD_PRECOS');
  const mapa = {};
  if (!aba) return mapa;
 
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(String);
  const idx = recCriarIndice_(headers);
 
  if (idx['SKU_MASTER'] === undefined || idx['PRECO_CUSTO'] === undefined) return mapa;
 
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return mapa;
 
  const dados = aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn()).getValues();
 
  dados.forEach(linha => {
    const sku = recGetCell_(linha, idx, 'SKU_MASTER');
    const custo = recGetCell_(linha, idx, 'PRECO_CUSTO');
    if (sku && custo) {
      const chave = recNorm_(sku);
      // Se houver múltiplos fornecedores para o mesmo SKU, mantém o menor custo
      if (mapa[chave] === undefined || Number(custo) < mapa[chave]) {
        mapa[chave] = Number(custo);
      }
    }
  });
 
  return mapa;
}
 
function recCriarIndice_(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[String(h).trim()] = i; });
  return idx;
}
 
function recGetCell_(row, idx, nomeColuna) {
  const pos = idx[nomeColuna];
  if (pos === undefined || pos === null || pos < 0) return '';
  return row[pos];
}
 
function recNorm_(valor) {
  return String(valor || '').trim().toUpperCase();
}
 
function recArredondar_(valor) {
  return Math.round(valor * 100) / 100;
}
 