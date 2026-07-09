/************************************************************
 * MEDBOX ERP | FASE 2 - IMPORTAR TAB PREÇOS
 *
 * Origem:
 * - Aba "Tab Preços"
 *
 * Destino:
 * - BD_PRECOS
 *
 * Regra:
 * - Valor Unitário = PRECO_CUSTO
 * - Fornecedor = Nome Fantasia Fornecedor, se existir;
 *                senão Razão Social do Fornecedor
 * - Descrição do Produto = DESCRICAO_FORNECEDOR
 * - Fabricante = FABRICANTE_MARCA
 * - Origem = TAB PREÇOS
 *
 * Não apaga a aba original.
 * Evita duplicidade básica por:
 * DATA_TABELA + FORNECEDOR + DESCRICAO + VALOR UNITÁRIO
 ************************************************************/

const F2P_ABA_ORIGEM = 'Tab Preços';
const F2P_ABA_DESTINO = 'BD_PRECOS';

function importarTabPrecosParaBdPrecos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const origem = ss.getSheetByName(F2P_ABA_ORIGEM);
  const destino = ss.getSheetByName(F2P_ABA_DESTINO);

  if (!origem) {
    throw new Error('Aba de origem não encontrada: ' + F2P_ABA_ORIGEM);
  }

  if (!destino) {
    throw new Error('Aba de destino não encontrada: ' + F2P_ABA_DESTINO + '. Execute antes a Fase 2A.');
  }

  const origemValues = origem.getDataRange().getValues();

  if (origemValues.length < 2) {
    SpreadsheetApp.getUi().alert('A aba Tab Preços não possui dados para importar.');
    return;
  }

  const origemHeaders = origemValues[0].map(String);
  const origemIdx = f2pCriarIndiceFlexivel_(origemHeaders);

  const destinoHeaders = destino.getRange(1, 1, 1, destino.getLastColumn()).getValues()[0].map(String);
  const destinoIdx = f2pCriarIndiceExato_(destinoHeaders);

  const chavesExistentes = f2pCarregarChavesExistentes_(destino, destinoIdx);

  const novasLinhas = [];

  let lidas = 0;
  let importadas = 0;
  let duplicadas = 0;
  let ignoradas = 0;

  for (let i = 1; i < origemValues.length; i++) {
    const row = origemValues[i];
    lidas++;

    const notaFiscal = f2pLimpar_(f2pGetFlex_(row, origemIdx, [
      'NOTA FISCAL',
      'NF',
      'N F',
      'NOTA'
    ]));

    const dataEmissao = f2pGetFlex_(row, origemIdx, [
      'DATA DE EMISSAO',
      'DATA DE EMISSÃO',
      'DATA EMISSAO',
      'DATA EMISSÃO'
    ]);

    const razaoFornecedor = f2pUpper_(f2pGetFlex_(row, origemIdx, [
      'RAZAO SOCIAL DO FORNECEDOR',
      'RAZÃO SOCIAL DO FORNECEDOR',
      'RAZAO SOCIAL FORNECEDOR',
      'RAZÃO SOCIAL FORNECEDOR'
    ]));

    const fantasiaFornecedor = f2pUpper_(f2pGetFlex_(row, origemIdx, [
      'NOME FANTASIA FORNECEDOR',
      'FANTASIA FORNECEDOR',
      'NOME FANTASIA'
    ]));

    const fornecedor = fantasiaFornecedor || razaoFornecedor;

    const descricaoProduto = f2pUpper_(f2pGetFlex_(row, origemIdx, [
      'DESCRICAO DO PRODUTO',
      'DESCRIÇÃO DO PRODUTO',
      'DESCRICAO PRODUTO',
      'DESCRIÇÃO PRODUTO',
      'PRODUTO'
    ]));

    const fabricante = f2pUpper_(f2pGetFlex_(row, origemIdx, [
      'FABRICANTE',
      'MARCA',
      'FABRICANTE MARCA'
    ]));

    const categoria = f2pUpper_(f2pGetFlex_(row, origemIdx, [
      'CATEGORIA'
    ]));

    const unidade = f2pUpper_(f2pGetFlex_(row, origemIdx, [
      'UNIDADE DE COMPRA',
      'UNIDADE COMPRA',
      'UNIDADE',
      'UND'
    ]));

    const quantidade = f2pToNumber_(f2pGetFlex_(row, origemIdx, [
      'QUANTIDADE',
      'QTD',
      'QTDE'
    ]));

    const valorUnitario = f2pToNumber_(f2pGetFlex_(row, origemIdx, [
      'VALOR UNITARIO',
      'VALOR UNITÁRIO',
      'VALOR UN',
      'PRECO UNITARIO',
      'PREÇO UNITÁRIO',
      'PRECO',
      'PREÇO'
    ]));

    const valorTotal = f2pToNumber_(f2pGetFlex_(row, origemIdx, [
      'VALOR TOTAL',
      'TOTAL'
    ]));

    if (!fornecedor || !descricaoProduto || valorUnitario === null) {
      ignoradas++;
      continue;
    }

    const dataTabela = dataEmissao || new Date();

    const chave = f2pMontarChave_(
      dataTabela,
      fornecedor,
      descricaoProduto,
      valorUnitario
    );

    if (chavesExistentes.has(chave)) {
      duplicadas++;
      continue;
    }

    const novaLinha = new Array(destinoHeaders.length).fill('');

    const idPreco = f2pGerarId_('PRE');

    const descricaoNormalizada = f2pNormalizarTexto_(descricaoProduto);
    const skuOperacional = descricaoNormalizada;

    f2pSetByHeader_(novaLinha, destinoIdx, 'ID_PRECO', idPreco);
    f2pSetByHeader_(novaLinha, destinoIdx, 'DATA_LANCAMENTO', new Date());
    f2pSetByHeader_(novaLinha, destinoIdx, 'DATA_TABELA', dataTabela);
    f2pSetByHeader_(novaLinha, destinoIdx, 'VALIDADE_TABELA', '');
    f2pSetByHeader_(novaLinha, destinoIdx, 'FORNECEDOR', fornecedor);
    f2pSetByHeader_(novaLinha, destinoIdx, 'FABRICANTE_MARCA', fabricante);
    f2pSetByHeader_(novaLinha, destinoIdx, 'DESCRICAO_FORNECEDOR', descricaoProduto);
    f2pSetByHeader_(novaLinha, destinoIdx, 'DESCRICAO_NORMALIZADA', descricaoNormalizada);
    f2pSetByHeader_(novaLinha, destinoIdx, 'SKU_MASTER', '');
    f2pSetByHeader_(novaLinha, destinoIdx, 'SKU_OPERACIONAL', skuOperacional);
    f2pSetByHeader_(novaLinha, destinoIdx, 'NCM', '');
    f2pSetByHeader_(novaLinha, destinoIdx, 'EAN', '');
    f2pSetByHeader_(novaLinha, destinoIdx, 'CST', '');
    f2pSetByHeader_(novaLinha, destinoIdx, 'UNIDADE', unidade);
    f2pSetByHeader_(novaLinha, destinoIdx, 'PRECO_CUSTO', valorUnitario);
    f2pSetByHeader_(novaLinha, destinoIdx, 'IPI_PERCENTUAL', 0);
    f2pSetByHeader_(novaLinha, destinoIdx, 'ICMS_PERCENTUAL', 0);
    f2pSetByHeader_(novaLinha, destinoIdx, 'FRETE_PERCENTUAL', 0);
    f2pSetByHeader_(novaLinha, destinoIdx, 'OUTRAS_DESPESAS_PERCENTUAL', 0);
    f2pSetByHeader_(novaLinha, destinoIdx, 'CUSTO_TOTAL_ESTIMADO', valorUnitario);
    f2pSetByHeader_(novaLinha, destinoIdx, 'ORIGEM', 'TAB PREÇOS');
    f2pSetByHeader_(novaLinha, destinoIdx, 'ARQUIVO_ORIGEM', F2P_ABA_ORIGEM);
    f2pSetByHeader_(
      novaLinha,
      destinoIdx,
      'OBSERVACAO',
      'NF: ' + notaFiscal +
      ' | QTD: ' + (quantidade === null ? '' : quantidade) +
      ' | TOTAL: ' + (valorTotal === null ? '' : valorTotal) +
      ' | CATEGORIA: ' + categoria
    );
    f2pSetByHeader_(novaLinha, destinoIdx, 'USUARIO', Session.getActiveUser().getEmail() || '');
    f2pSetByHeader_(novaLinha, destinoIdx, 'DATA_ATUALIZACAO', new Date());

    novasLinhas.push(novaLinha);
    chavesExistentes.add(chave);
    importadas++;
  }

  if (novasLinhas.length > 0) {
    const startRow = destino.getLastRow() + 1;
    destino.getRange(startRow, 1, novasLinhas.length, destinoHeaders.length).setValues(novasLinhas);
  }

  f2pRegistrarLog_(
    'IMPORTAÇÃO TAB PREÇOS | LIDAS: ' + lidas +
    ' | IMPORTADAS: ' + importadas +
    ' | DUPLICADAS: ' + duplicadas +
    ' | IGNORADAS: ' + ignoradas
  );

  SpreadsheetApp.getUi().alert(
    'Importação concluída.\n\n' +
    'Linhas lidas: ' + lidas + '\n' +
    'Importadas: ' + importadas + '\n' +
    'Duplicadas ignoradas: ' + duplicadas + '\n' +
    'Ignoradas por falta de dados: ' + ignoradas + '\n\n' +
    'Destino: BD_PRECOS'
  );
}

/************************************************************
 * CHAVES / DUPLICIDADE
 ************************************************************/

function f2pCarregarChavesExistentes_(destino, idx) {
  const set = new Set();

  if (destino.getLastRow() < 2) return set;

  const values = destino.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const dataTabela = f2pGetExact_(row, idx, 'DATA_TABELA');
    const fornecedor = f2pUpper_(f2pGetExact_(row, idx, 'FORNECEDOR'));
    const descricao = f2pUpper_(f2pGetExact_(row, idx, 'DESCRICAO_FORNECEDOR'));
    const preco = f2pToNumber_(f2pGetExact_(row, idx, 'PRECO_CUSTO'));

    if (!fornecedor || !descricao || preco === null) continue;

    set.add(f2pMontarChave_(dataTabela, fornecedor, descricao, preco));
  }

  return set;
}

function f2pMontarChave_(data, fornecedor, descricao, preco) {
  const dataKey = data instanceof Date
    ? Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : f2pLimpar_(data);

  return [
    dataKey,
    f2pNormalizarTexto_(fornecedor),
    f2pNormalizarTexto_(descricao),
    String(f2pToNumber_(preco))
  ].join('|');
}

/************************************************************
 * HELPERS
 ************************************************************/

function f2pCriarIndiceExato_(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    idx[String(h).trim()] = i;
  });

  return idx;
}

function f2pCriarIndiceFlexivel_(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    const normal = f2pNormalizarHeader_(h);
    idx[normal] = i;
  });

  return idx;
}

function f2pNormalizarHeader_(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function f2pGetFlex_(row, idx, aliases) {
  for (const alias of aliases) {
    const key = f2pNormalizarHeader_(alias);

    if (idx[key] !== undefined) {
      return row[idx[key]];
    }
  }

  return '';
}

function f2pGetExact_(row, idx, nomeColuna) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) return '';

  return row[pos];
}

function f2pSetByHeader_(rowArray, idx, nomeColuna, valor) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) return;

  rowArray[pos] = valor;
}

function f2pGerarId_(prefixo) {
  const data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const random = Math.floor(Math.random() * 9000) + 1000;

  return prefixo + '-' + data + '-' + random;
}

function f2pLimpar_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .trim()
    .replace(/\s+/g, ' ');
}

function f2pUpper_(valor) {
  return f2pLimpar_(valor).toUpperCase();
}

function f2pNormalizarTexto_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(',', '.')
    .replace(/([0-9])\s+(MG|ML|G|L|MCG|UI|UN)\b/g, '$1$2')
    .replace(/([A-Z])([0-9])/g, '$1 $2')
    .replace(/([0-9])([A-Z])/g, '$1$2')
    .replace(/[^A-Z0-9.%/+ X\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function f2pToNumber_(valor) {
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

function f2pRegistrarLog_(mensagem) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('LOG_PROCESSAMENTO');

  if (!ws) {
    ws = ss.insertSheet('LOG_PROCESSAMENTO');
    ws.appendRow(['DATA_HORA', 'EVENTO']);
  }

  if (ws.getLastRow() === 0) {
    ws.appendRow(['DATA_HORA', 'EVENTO']);
  }

  ws.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    mensagem
  ]);
}