/************************************************************
 * MEDBOX ERP | DIAGNÓSTICO ROBÔ FISCAL
 *
 * Arquivo: Diagnostico.js
 *
 * Realiza varredura de diagnóstico na aba BD_PRECOS e
 * verifica o status de outras abas auxiliares.
 * Não altera dados existentes.
 ************************************************************/

/**
 * Executa o diagnóstico da aba BD_PRECOS e gera um relatório.
 */
function varredura_BD_PRECOS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBdPrecos = ss.getSheetByName('BD_PRECOS');

  if (!abaBdPrecos) {
    throw new Error("Aba 'BD_PRECOS' não encontrada.");
  }

  const lastCol = abaBdPrecos.getLastColumn();
  const headers = abaBdPrecos.getRange(1, 1, 1, lastCol).getValues()[0];

  const colunasInteresse = [
    'DESCRICAO_FORNECEDOR',
    'SKU_MASTER',
    'NCM',
    'EAN',
    'CST',
    'SUGESTAO_NCM_COSMOS',
    'SUGESTAO_EAN_COSMOS',
    'STATUS_CONSULTA_COSMOS',
    'FONTE_COSMOS_URL',
    'DATA_CONSULTA_COSMOS'
  ];

  const posicoes = {};
  headers.forEach((h, i) => {
    const nome = String(h).trim();
    if (colunasInteresse.includes(nome)) {
      posicoes[nome] = {
        indice: i,
        letra: rfcColunaParaLetra_(i + 1)
      };
    }
  });

  const lastRow = abaBdPrecos.getLastRow();
  const totalLinhasProdutos = Math.max(0, lastRow - 1);

  let contadorVaziosGeral = 0; // SKU_MASTER, NCM, EAN e CST vazios na mesma linha
  let contadorEncontradoCosmos = 0;

  if (totalLinhasProdutos > 0) {
    const data = abaBdPrecos.getRange(2, 1, totalLinhasProdutos, lastCol).getValues();

    data.forEach(row => {
      const sku = row[posicoes['SKU_MASTER']?.indice];
      const ncm = row[posicoes['NCM']?.indice];
      const ean = row[posicoes['EAN']?.indice];
      const cst = row[posicoes['CST']?.indice];
      const status = row[posicoes['STATUS_CONSULTA_COSMOS']?.indice];

      // Verifica se SKU_MASTER, NCM, EAN e CST estão vazios (todos ao mesmo tempo)
      if (!sku && !ncm && !ean && !cst) {
        contadorVaziosGeral++;
      }

      if (status && String(status).toUpperCase() === 'ENCONTRADO') {
        contadorEncontradoCosmos++;
      }
    });
  }

  const abaCadastro = ss.getSheetByName('CADASTRO_SKU_MASTER');
  const cadastroTemDados = abaCadastro && abaCadastro.getLastRow() > 1;

  const abaBdSefaz = ss.getSheetByName('BD_SKU_MASTER_SEFAZ');
  const sefazTemDados = abaBdSefaz && abaBdSefaz.getLastRow() > 1;

  // Gerar Relatório na aba DIAGNOSTICO_ROBO_FISCAL
  let abaDiag = ss.getSheetByName('DIAGNOSTICO_ROBO_FISCAL');
  if (abaDiag) {
    abaDiag.clear();
  } else {
    abaDiag = ss.insertSheet('DIAGNOSTICO_ROBO_FISCAL');
  }

  const metrics = [
    ['MÉTRICA', 'VALOR'],
    ['Total de linhas de produtos em BD_PRECOS', totalLinhasProdutos],
    ['Linhas com SKU_MASTER, NCM, EAN e CST vazios', contadorVaziosGeral],
    ['Linhas com STATUS_CONSULTA_COSMOS = ENCONTRADO', contadorEncontradoCosmos],
    ['Aba CADASTRO_SKU_MASTER possui dados?', cadastroTemDados ? 'SIM' : 'VAZIA'],
    ['Aba BD_SKU_MASTER_SEFAZ possui dados?', sefazTemDados ? 'SIM' : 'VAZIA'],
    ['', ''],
    ['MAPEAMENTO DE COLUNAS (BD_PRECOS)', '']
  ];

  colunasInteresse.forEach(col => {
    metrics.push([col, posicoes[col] ? 'Coluna ' + posicoes[col].letra : 'NÃO ENCONTRADA']);
  });

  abaDiag.getRange(1, 1, metrics.length, 2).setValues(metrics);

  // Formatação básica
  abaDiag.getRange(1, 1, 1, 2).setBackground('#002E56').setFontColor('#FFFFFF').setFontWeight('bold');
  abaDiag.getRange(8, 1, 1, 2).setBackground('#EEF1F4').setFontWeight('bold');
  abaDiag.setColumnWidth(1, 350);
  abaDiag.setColumnWidth(2, 200);

  SpreadsheetApp.getUi().alert('Diagnóstico concluído na aba DIAGNOSTICO_ROBO_FISCAL.');
}

/**
 * Converte índice de coluna para letra (ex: 1 -> A, 27 -> AA)
 */
function rfcColunaParaLetra_(coluna) {
  let temp, letra = '';
  while (coluna > 0) {
    temp = (coluna - 1) % 26;
    letra = String.fromCharCode(temp + 65) + letra;
    coluna = (coluna - temp - 1) / 26;
  }
  return letra;
}
