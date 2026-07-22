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

/**
 * Executa o diagnóstico completo do Mapa de Produtos e da Planilha Oficial.
 * NÃO realiza nenhuma alteração de gravação nas planilhas, operando de modo seguro (somente leitura).
 */
function diagnosticarMapaDeProdutosEClientes() {
  const log = [];
  log.push('============================================================');
  log.push('MEDBOX ERP | DIAGNÓSTICO DO MAPA DE PRODUTOS & CLIENTES');
  log.push('Gerado em: ' + Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm:ss'));
  log.push('============================================================\n');

  try {
    // 1. Verificar acesso à Planilha Oficial
    log.push('1. Verificando acesso à Planilha Oficial...');
    log.push('Spreadsheet ID: ' + MEDBOX_SPREADSHEET_ID);
    const ss = obterPlanilhaMedbox_();
    log.push('Planilha Oficial carregada com sucesso: "' + ss.getName() + '"\n');

    // 2. Verificar Aba MAPA_ERP
    log.push('2. Verificando aba MAPA_ERP...');
    const mapa = ss.getSheetByName('MAPA_ERP');
    if (!mapa) {
      log.push('❌ ERRO: Aba MAPA_ERP não encontrada na Planilha Oficial.');
    } else {
      const lastRow = mapa.getLastRow();
      const lastCol = mapa.getLastColumn();
      log.push('Aba MAPA_ERP localizada.');
      log.push('Linhas Totais: ' + lastRow);
      log.push('Colunas Totais: ' + lastCol);
      if (lastRow > 0) {
        const headers = mapa.getRange(1, 1, 1, lastCol).getValues()[0];
        log.push('Cabeçalhos da MAPA_ERP: ' + headers.join(' | '));
      }
    }
    log.push('');

    // 3. Descobrir Clientes de forma dinâmica
    log.push('3. Executando Descoberta Dinâmica de Clientes...');
    const abasClientes = f2dObterAbasClientes_(ss);
    log.push('Total de Abas de Clientes Detectadas: ' + abasClientes.length);

    const clientesNomes = [];
    abasClientes.forEach(sheet => {
      const name = sheet.getName();
      const clienteName = name.substring(3).trim();
      clientesNomes.push(clienteName);

      // Realizar uma verificação rápida da estrutura de cada aba de cliente
      const rows = sheet.getLastRow();
      const cols = sheet.getLastColumn();
      log.push('  - Aba: "' + name + '" -> Cliente: "' + clienteName.toUpperCase() + '" (' + rows + ' linhas, ' + cols + ' colunas)');

      const detection = f2dDetectarCabecalhoEColunas_(sheet);
      if (detection) {
        log.push('    Mapeamento de colunas (Cabeçalho na linha ' + (detection.headerRowIndex + 1) + '):');
        log.push('      SKU_MASTER: Coluna ' + (detection.colSku + 1));
        log.push('      QUANTIDADE: Coluna ' + (detection.colQtd + 1));
        log.push('      PREÇO CLIENTE: Coluna ' + (detection.colVal + 1));
      } else {
        log.push('    ❌ ERRO: Cabeçalho compatível com SKU, Quantidade e Preço não encontrado nas primeiras 5 linhas.');
      }
    });
    log.push('');

    // 4. Testar Simulação do Carregamento do Mapa de Produtos
    log.push('4. Testando Simulação do Carregamento Completo...');
    const result = obterMapaDeProdutos({ limite: 10 });
    log.push('Simulação de obterMapaDeProdutos({ limite: 10 }) executada com sucesso.');
    log.push('  - Total Encontrado: ' + result.totalEncontrado);
    log.push('  - Total Exibido: ' + result.exibidos);
    log.push('  - Quantidade Total Demandada: ' + result.quantidadeTotalDemandada);
    log.push('  - Valor Total Demandado: ' + result.valorTotalDemandado);
    log.push('  - Clientes Retornados: ' + result.clientes.join(', '));
    log.push('  - Amostra de Produtos Retornados: ' + result.produtos.length);

    result.produtos.forEach((prod, index) => {
      log.push('    ' + (index + 1) + '. SKU: ' + prod.skuMaster + ' | Produto: ' + prod.produto + ' | Qtd Total: ' + prod.quantTotal + ' | Demanda R$: ' + prod.demandaR);
    });
    log.push('');

    log.push('✅ DIAGNÓSTICO CONCLUÍDO COM SUCESSO. Nenhuma falha estrutural detectada.');
  } catch (e) {
    log.push('❌ ERRO DURANTE O DIAGNÓSTICO: ' + e.message);
    log.push(e.stack);
  }

  log.push('\n============================================================');
  log.push('FIM DO DIAGNÓSTICO');
  log.push('============================================================');

  const output = log.join('\n');
  Logger.log(output);
  return output;
}
