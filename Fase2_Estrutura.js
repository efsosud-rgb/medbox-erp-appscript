/************************************************************
 * MEDBOX ERP | FASE 2A
 * Criação da estrutura de novas abas do sistema
 *
 * Não apaga dados existentes.
 * Não altera MAPA_ERP.
 * Não altera AUX_BUSCA.
 * Não altera abas BD_ antigas.
 ************************************************************/

function criarEstruturaFase2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbaBdCotacoes_(ss);
  criarAbaBdPrecos_(ss);
  criarAbaBdSkuMasterSefaz_(ss);

  criarAbaLancamentoCotacoes_(ss);
  criarAbaLancamentoPrecos_(ss);
  criarAbaCadastroSkuMaster_(ss);
  criarAbaDashboard_(ss);
  criarAbaSimuladorPrecos_(ss);

  registrarLogFase2_(ss, 'FASE 2A - Estrutura criada/atualizada com sucesso.');

  SpreadsheetApp.getUi().alert(
    'FASE 2A concluída.\n\nForam criadas/atualizadas as abas:\n\n' +
    '- BD_COTACOES\n' +
    '- BD_PRECOS\n' +
    '- BD_SKU_MASTER_SEFAZ\n' +
    '- LANCAMENTO_COTACOES\n' +
    '- LANCAMENTO_PRECOS\n' +
    '- CADASTRO_SKU_MASTER\n' +
    '- DASHBOARD\n' +
    '- SIMULADOR_PRECOS\n\n' +
    'Nenhuma base antiga foi apagada.'
  );
}

/************************************************************
 * 1. BD_COTACOES
 ************************************************************/

function criarAbaBdCotacoes_(ss) {
  const nome = 'BD_COTACOES';

  const headers = [
    'ID_COTACAO',
    'DATA_LANCAMENTO',
    'DATA_COTACAO',
    'CLIENTE',
    'RESPONSAVEL',
    'CATEGORIA',
    'DESCRICAO_CLIENTE',
    'DESCRICAO_NORMALIZADA',
    'SKU_MASTER',
    'SKU_OPERACIONAL',
    'MARCA_FABRICANTE',
    'QUANTIDADE',
    'UNIDADE',
    'PRECO_REFERENCIA_CLIENTE',
    'VALOR_TOTAL_REFERENCIA',
    'PRECO_CUSTO_MEDBOX',
    'PRECO_MEDBOX',
    'MARKUP_PERCENTUAL',
    'STATUS_COMERCIAL',
    'STATUS_COTACAO',
    'ORIGEM',
    'OBSERVACAO',
    'USUARIO',
    'DATA_ATUALIZACAO'
  ];

  const ws = obterOuCriarAba_(ss, nome);
  prepararAbaTabela_(ws, headers, '#002E56', '#FFFFFF');

  ws.setFrozenRows(1);
  ws.setFrozenColumns(4);

  aplicarLarguras_(ws, {
    1: 130,
    2: 130,
    3: 130,
    4: 150,
    5: 140,
    6: 140,
    7: 360,
    8: 360,
    9: 260,
    10: 260,
    11: 180,
    12: 110,
    13: 90,
    14: 150,
    15: 150,
    16: 150,
    17: 150,
    18: 130,
    19: 150,
    20: 140,
    21: 160,
    22: 280,
    23: 120,
    24: 150
  });
}

/************************************************************
 * 2. BD_PRECOS
 ************************************************************/

function criarAbaBdPrecos_(ss) {
  const nome = 'BD_PRECOS';

  const headers = [
    'ID_PRECO',
    'DATA_LANCAMENTO',
    'DATA_TABELA',
    'VALIDADE_TABELA',
    'FORNECEDOR',
    'FABRICANTE_MARCA',
    'DESCRICAO_FORNECEDOR',
    'DESCRICAO_NORMALIZADA',
    'SKU_MASTER',
    'SKU_OPERACIONAL',
    'NCM',
    'EAN',
    'CST',
    'UNIDADE',
    'PRECO_CUSTO',
    'IPI_PERCENTUAL',
    'ICMS_PERCENTUAL',
    'FRETE_PERCENTUAL',
    'OUTRAS_DESPESAS_PERCENTUAL',
    'CUSTO_TOTAL_ESTIMADO',
    'ORIGEM',
    'ARQUIVO_ORIGEM',
    'OBSERVACAO',
    'USUARIO',
    'DATA_ATUALIZACAO'
  ];

  const ws = obterOuCriarAba_(ss, nome);
  prepararAbaTabela_(ws, headers, '#002E56', '#FFFFFF');

  ws.setFrozenRows(1);
  ws.setFrozenColumns(5);

  aplicarLarguras_(ws, {
    1: 130,
    2: 130,
    3: 130,
    4: 130,
    5: 180,
    6: 180,
    7: 380,
    8: 380,
    9: 260,
    10: 260,
    11: 100,
    12: 140,
    13: 90,
    14: 90,
    15: 120,
    16: 120,
    17: 120,
    18: 120,
    19: 170,
    20: 170,
    21: 170,
    22: 220,
    23: 280,
    24: 120,
    25: 150
  });
}

/************************************************************
 * 3. BD_SKU_MASTER_SEFAZ
 ************************************************************/

function criarAbaBdSkuMasterSefaz_(ss) {
  const nome = 'BD_SKU_MASTER_SEFAZ';

  const headers = [
    'ID_SKU',
    'SKU_MASTER',
    'DESCRICAO_PADRAO_SEFAZ',
    'DESCRICAO_COMERCIAL',
    'NCM',
    'EAN',
    'CST',
    'CEST',
    'CFOP_REFERENCIA',
    'UNIDADE',
    'CATEGORIA',
    'SUBCATEGORIA',
    'PRINCIPIO_ATIVO_COMPONENTE',
    'CONCENTRACAO_MEDIDA',
    'FORMA_APRESENTACAO',
    'VOLUME_TAMANHO',
    'MARCA_FABRICANTE',
    'REGISTRO_ANVISA',
    'STATUS_CADASTRO',
    'FONTE',
    'OBSERVACAO',
    'USUARIO',
    'DATA_CRIACAO',
    'DATA_ATUALIZACAO'
  ];

  const ws = obterOuCriarAba_(ss, nome);
  prepararAbaTabela_(ws, headers, '#002E56', '#FFFFFF');

  ws.setFrozenRows(1);
  ws.setFrozenColumns(3);

  aplicarLarguras_(ws, {
    1: 120,
    2: 280,
    3: 420,
    4: 420,
    5: 100,
    6: 140,
    7: 90,
    8: 90,
    9: 130,
    10: 90,
    11: 150,
    12: 160,
    13: 240,
    14: 180,
    15: 200,
    16: 160,
    17: 190,
    18: 150,
    19: 150,
    20: 150,
    21: 280,
    22: 120,
    23: 140,
    24: 150
  });
}

/************************************************************
 * 4. LANCAMENTO_COTACOES
 ************************************************************/

function criarAbaLancamentoCotacoes_(ss) {
  const nome = 'LANCAMENTO_COTACOES';

  const headers = [
    'CAMPO',
    'VALOR'
  ];

  const ws = obterOuCriarAba_(ss, nome);
  ws.clear();

  ws.getRange('A1:B1').setValues([headers]);
  formatarCabecalho_(ws.getRange('A1:B1'), '#D80024', '#FFFFFF');

  const campos = [
    ['DATA DA COTACAO', ''],
    ['CLIENTE', ''],
    ['RESPONSAVEL', ''],
    ['CATEGORIA', ''],
    ['DESCRICAO DO ITEM', ''],
    ['QUANTIDADE', ''],
    ['UNIDADE', ''],
    ['PRECO REFERENCIA CLIENTE', ''],
    ['SKU MASTER SUGERIDO', ''],
    ['MARCA / FABRICANTE', ''],
    ['STATUS', 'NOVA'],
    ['OBSERVACAO', '']
  ];

  ws.getRange(2, 1, campos.length, 2).setValues(campos);

  ws.setFrozenRows(1);
  ws.setColumnWidth(1, 260);
  ws.setColumnWidth(2, 620);

  ws.getRange('A:A').setFontWeight('bold');
  ws.getRange('B2:B13').setBackground('#FFF5C7');

  ws.getRange('D1').setValue('INSTRUCOES');
  ws.getRange('D1').setBackground('#002E56').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.getRange('D2').setValue('Esta aba é uma tela simples de lançamento. A próxima fase criará formulário HTML para salvar em BD_COTACOES.');
  ws.getRange('D2').setWrap(true);
  ws.setColumnWidth(4, 520);
}

/************************************************************
 * 5. LANCAMENTO_PRECOS
 ************************************************************/

function criarAbaLancamentoPrecos_(ss) {
  const nome = 'LANCAMENTO_PRECOS';

  const headers = [
    'CAMPO',
    'VALOR'
  ];

  const ws = obterOuCriarAba_(ss, nome);
  ws.clear();

  ws.getRange('A1:B1').setValues([headers]);
  formatarCabecalho_(ws.getRange('A1:B1'), '#D80024', '#FFFFFF');

  const campos = [
    ['DATA DO LANCAMENTO', ''],
    ['FORNECEDOR', ''],
    ['FABRICANTE / MARCA', ''],
    ['DESCRICAO PRODUTO FORNECEDOR', ''],
    ['SKU MASTER', ''],
    ['NCM', ''],
    ['EAN', ''],
    ['CST', ''],
    ['UNIDADE', ''],
    ['PRECO CUSTO', ''],
    ['VALIDADE DA TABELA', ''],
    ['ORIGEM', 'LANÇAMENTO AVULSO'],
    ['OBSERVACAO', '']
  ];

  ws.getRange(2, 1, campos.length, 2).setValues(campos);

  ws.setFrozenRows(1);
  ws.setColumnWidth(1, 280);
  ws.setColumnWidth(2, 620);

  ws.getRange('A:A').setFontWeight('bold');
  ws.getRange('B2:B14').setBackground('#FFF5C7');

  ws.getRange('D1').setValue('INSTRUCOES');
  ws.getRange('D1').setBackground('#002E56').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.getRange('D2').setValue('Esta aba é uma tela simples de lançamento. A próxima fase criará formulário HTML para salvar em BD_PRECOS.');
  ws.getRange('D2').setWrap(true);
  ws.setColumnWidth(4, 520);
}

/************************************************************
 * 6. CADASTRO_SKU_MASTER
 ************************************************************/

function criarAbaCadastroSkuMaster_(ss) {
  const nome = 'CADASTRO_SKU_MASTER';

  const headers = [
    'CAMPO',
    'VALOR'
  ];

  const ws = obterOuCriarAba_(ss, nome);
  ws.clear();

  ws.getRange('A1:B1').setValues([headers]);
  formatarCabecalho_(ws.getRange('A1:B1'), '#D80024', '#FFFFFF');

  const campos = [
    ['SKU MASTER', ''],
    ['DESCRICAO PADRAO SEFAZ', ''],
    ['DESCRICAO COMERCIAL', ''],
    ['NCM', ''],
    ['EAN', ''],
    ['CST', ''],
    ['CEST', ''],
    ['UNIDADE', ''],
    ['CATEGORIA', ''],
    ['SUBCATEGORIA', ''],
    ['PRINCIPIO ATIVO / COMPONENTE', ''],
    ['CONCENTRACAO / MEDIDA', ''],
    ['FORMA / APRESENTACAO', ''],
    ['VOLUME / TAMANHO', ''],
    ['MARCA / FABRICANTE', ''],
    ['REGISTRO ANVISA', ''],
    ['STATUS CADASTRO', 'NOVO'],
    ['OBSERVACAO', '']
  ];

  ws.getRange(2, 1, campos.length, 2).setValues(campos);

  ws.setFrozenRows(1);
  ws.setColumnWidth(1, 320);
  ws.setColumnWidth(2, 700);

  ws.getRange('A:A').setFontWeight('bold');
  ws.getRange('B2:B19').setBackground('#FFF5C7');

  ws.getRange('D1').setValue('REGRA SKU MASTER');
  ws.getRange('D1').setBackground('#002E56').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.getRange('D2').setValue('SKU MASTER deve ser comercial: produto principal + concentração/medida + forma/apresentação + volume/tamanho. SEFAZ, NCM, EAN e CST servem como base fiscal, não substituem a lógica comercial.');
  ws.getRange('D2').setWrap(true);
  ws.setColumnWidth(4, 650);
}

/************************************************************
 * 7. DASHBOARD
 ************************************************************/

function criarAbaDashboard_(ss) {
  const nome = 'DASHBOARD';

  const ws = obterOuCriarAba_(ss, nome);
  ws.clear();

  ws.getRange('A1:H1').merge();
  ws.getRange('A1').setValue('MEDBOX ERP | DASHBOARD COMERCIAL');
  ws.getRange('A1')
    .setBackground('#002E56')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  const cards = [
    ['INDICADOR', 'VALOR'],
    ['TOTAL DE SKUS', ''],
    ['TOTAL DE ITENS DEMANDADOS', ''],
    ['VALOR TOTAL DEMANDADO', ''],
    ['SKUS SEM CUSTO', ''],
    ['SKUS SEM CADASTRO FISCAL', ''],
    ['CLIENTES ATIVOS', ''],
    ['CATEGORIAS ATIVAS', '']
  ];

  ws.getRange(3, 1, cards.length, 2).setValues(cards);
  formatarCabecalho_(ws.getRange('A3:B3'), '#D80024', '#FFFFFF');
  ws.getRange('A4:A10').setFontWeight('bold');
  ws.getRange('B4:B10').setBackground('#FFF5C7');

  ws.getRange('D3:E3').setValues([['TOP 30 SKU MASTER POR QUANTIDADE', 'QUANTIDADE']]);
  formatarCabecalho_(ws.getRange('D3:E3'), '#002E56', '#FFFFFF');

  ws.getRange('G3:H3').setValues([['TOP 30 SKU MASTER POR VALOR', 'VALOR']]);
  formatarCabecalho_(ws.getRange('G3:H3'), '#002E56', '#FFFFFF');

  ws.setFrozenRows(1);

  ws.setColumnWidth(1, 260);
  ws.setColumnWidth(2, 180);
  ws.setColumnWidth(4, 420);
  ws.setColumnWidth(5, 160);
  ws.setColumnWidth(7, 420);
  ws.setColumnWidth(8, 160);

  ws.getRange('A12:H12').merge();
  ws.getRange('A12').setValue('OBSERVAÇÃO: A próxima fase preencherá o dashboard automaticamente com base no MAPA_ERP, BD_COTACOES e BD_PRECOS.');
  ws.getRange('A12')
    .setBackground('#EEF1F4')
    .setFontWeight('bold')
    .setWrap(true);
}

/************************************************************
 * 8. SIMULADOR_PRECOS
 ************************************************************/

function criarAbaSimuladorPrecos_(ss) {
  const nome = 'SIMULADOR_PRECOS';

  const headers = [
    'CAMPO',
    'VALOR'
  ];

  const ws = obterOuCriarAba_(ss, nome);
  ws.clear();

  ws.getRange('A1:B1').setValues([headers]);
  formatarCabecalho_(ws.getRange('A1:B1'), '#D80024', '#FFFFFF');

  const campos = [
    ['SKU MASTER', ''],
    ['PRODUTO', ''],
    ['PRECO CUSTO', ''],
    ['MARKUP %', '50'],
    ['FRETE %', ''],
    ['FRETE R$', ''],
    ['IMPOSTO %', ''],
    ['IMPOSTO R$', ''],
    ['OUTRAS DESPESAS %', ''],
    ['OUTRAS DESPESAS R$', ''],
    ['CUSTO TOTAL', ''],
    ['PRECO MEDBOX', ''],
    ['PRECO CLIENTE', ''],
    ['DIFERENCA R$', ''],
    ['STATUS', '']
  ];

  ws.getRange(2, 1, campos.length, 2).setValues(campos);

  ws.setFrozenRows(1);
  ws.setColumnWidth(1, 280);
  ws.setColumnWidth(2, 280);

  ws.getRange('A:A').setFontWeight('bold');
  ws.getRange('B2:B16').setBackground('#FFF5C7');

  ws.getRange('D1').setValue('FORMULA');
  ws.getRange('D1').setBackground('#002E56').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.getRange('D2').setValue('CUSTO TOTAL = PREÇO CUSTO + FRETE + IMPOSTOS + OUTRAS DESPESAS');
  ws.getRange('D3').setValue('PREÇO MEDBOX = CUSTO TOTAL × (1 + MARKUP)');
  ws.getRange('D4').setValue('A próxima fase criará a simulação automática no HTML.');
  ws.getRange('D2:D4').setWrap(true);
  ws.setColumnWidth(4, 620);
}

/************************************************************
 * LOG
 ************************************************************/

function registrarLogFase2_(ss, mensagem) {
  const nome = 'LOG_PROCESSAMENTO';
  const ws = obterOuCriarAba_(ss, nome);

  if (ws.getLastRow() === 0) {
    ws.appendRow(['DATA_HORA', 'EVENTO']);
  }

  ws.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    mensagem
  ]);
}

/************************************************************
 * HELPERS
 ************************************************************/

function obterOuCriarAba_(ss, nome) {
  let ws = ss.getSheetByName(nome);

  if (!ws) {
    ws = ss.insertSheet(nome);
  }

  return ws;
}

function prepararAbaTabela_(ws, headers, corFundo, corFonte) {
  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const headerAtual = ws.getRange(1, 1, 1, Math.max(lastCol, headers.length)).getValues()[0];

    let precisaAtualizar = false;

    for (let i = 0; i < headers.length; i++) {
      if (String(headerAtual[i] || '').trim() !== headers[i]) {
        precisaAtualizar = true;
        break;
      }
    }

    if (precisaAtualizar && lastRow <= 1) {
      ws.clear();
      ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    if (lastRow < 1) {
      ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  formatarCabecalho_(ws.getRange(1, 1, 1, headers.length), corFundo, corFonte);

  if (ws.getFilter()) {
    ws.getFilter().remove();
  }

  ws.getRange(1, 1, Math.max(ws.getLastRow(), 1), headers.length).createFilter();

  ws.getRange(1, 1, Math.max(ws.getMaxRows(), 1), headers.length)
    .setFontFamily('Arial')
    .setFontSize(10);

  ws.getRange(1, 1, 1, headers.length)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  ws.setRowHeight(1, 34);
}

function formatarCabecalho_(range, corFundo, corFonte) {
  range
    .setBackground(corFundo)
    .setFontColor(corFonte)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function aplicarLarguras_(ws, mapaLarguras) {
  Object.keys(mapaLarguras).forEach(function(col) {
    ws.setColumnWidth(Number(col), mapaLarguras[col]);
  });
}