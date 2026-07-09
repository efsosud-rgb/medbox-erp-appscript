/************************************************************
 * MEDBOX ERP | FASE 2H.4 - IMPORTADOR UNIVERSAL DE TABELAS
 *
 * Arquivo: Fase2_ImportadorUniversal.gs
 *
 * Objetivo:
 * - Receber arquivos de tabelas de fornecedores pelo Web App.
 * - Salvar no Google Drive.
 * - Ler XLSX, XLS, CSV, TXT, PDF, DOC, DOCX, JPEG, PNG.
 * - Bloquear XLSB com mensagem clara.
 * - Extrair produtos para REVISAO_IMPORTACAO_PRECOS.
 * - Enviar linhas aprovadas para BD_PRECOS.
 *
 * Correções consolidadas:
 * - Produto sem preço entra com PRECO_CUSTO = 0,00.
 * - Produto sem preço recebe STATUS = SEM PREÇO.
 * - Preço 0,00 pode ir para BD_PRECOS, mas não deve calcular Preço MedBox.
 * - Bloqueia linhas falsas, ex.: "61 18416 62 17362".
 * - Não confunde STATUS DE PREÇO com PREÇO CUSTO.
 * - Reconhece PREÇO CUSTO, PREÇO UNITÁRIO, VALOR UNITÁRIO, PÇ VENDA.
 * - Reconhece formatos reais:
 *   Nova Linea, Joaomed, S3MED, CIMED, Descarpack, FarmaVision, JP Farma.
 ************************************************************/

const F2H_ABA_BD_IMPORTACOES = 'BD_IMPORTACOES_ARQUIVOS';
const F2H_ABA_REVISAO = 'REVISAO_IMPORTACAO_PRECOS';
const F2H_ABA_BD_PRECOS = 'BD_PRECOS';
const F2H_ABA_LOG_IMPORTACAO = 'LOG_IMPORTACAO_PRECOS';
const F2H_PASTA_DRIVE = 'MEDBOX_IMPORTACOES_TABELAS_PRECOS';

/************************************************************
 * 1. CRIAR / ATUALIZAR ESTRUTURA
 ************************************************************/

function criarEstruturaImportadorUniversalFase2H() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  f2hCriarAbaBdImportacoes_(ss);
  f2hCriarAbaRevisao_(ss);
  f2hCriarAbaLogImportacao_(ss);

  f2hRegistrarLog_('FASE 2H.4 - Estrutura do importador universal criada/atualizada.');

  SpreadsheetApp.getUi().alert(
    'FASE 2H.4 concluída.\n\n' +
    'Abas criadas/atualizadas:\n\n' +
    '- BD_IMPORTACOES_ARQUIVOS\n' +
    '- REVISAO_IMPORTACAO_PRECOS\n' +
    '- LOG_IMPORTACAO_PRECOS\n\n' +
    'Nenhuma base antiga foi apagada.'
  );
}

function f2hCriarAbaBdImportacoes_(ss) {
  const headers = [
    'ID_IMPORTACAO',
    'DATA_UPLOAD',
    'FORNECEDOR',
    'FABRICANTE_PADRAO',
    'ORIGEM',
    'DATA_TABELA',
    'VALIDADE_TABELA',
    'NOME_ARQUIVO',
    'MIME_TYPE',
    'FILE_ID',
    'FILE_URL',
    'STATUS',
    'LINHAS_LIDAS',
    'LINHAS_EXTRAIDAS',
    'LINHAS_VALIDAS',
    'LINHAS_IGNORADAS',
    'OBSERVACAO',
    'USUARIO',
    'DATA_ATUALIZACAO'
  ];

  const ws = f2hObterOuCriarAba_(ss, F2H_ABA_BD_IMPORTACOES);
  f2hPrepararTabela_(ws, headers, '#002E56', '#FFFFFF');

  ws.setFrozenRows(1);
  ws.setFrozenColumns(3);

  f2hAplicarLarguras_(ws, {
    1: 180,
    2: 150,
    3: 220,
    4: 220,
    5: 180,
    6: 120,
    7: 120,
    8: 360,
    9: 190,
    10: 220,
    11: 360,
    12: 180,
    13: 110,
    14: 120,
    15: 110,
    16: 120,
    17: 420,
    18: 190,
    19: 150
  });
}

function f2hCriarAbaRevisao_(ss) {
  const headers = [
    'ID_REVISAO',
    'ID_IMPORTACAO',
    'STATUS',
    'CONFIANCA',
    'FORNECEDOR',
    'FABRICANTE',
    'CODIGO_FORNECEDOR',
    'DESCRICAO_ORIGINAL',
    'DESCRICAO_NORMALIZADA',
    'SKU_MASTER_SUGERIDO',
    'SKU_OPERACIONAL',
    'EAN',
    'NCM',
    'ANVISA',
    'UNIDADE',
    'QUANTIDADE',
    'PRECO_CUSTO',
    'PRECO_TOTAL',
    'IPI_PERCENTUAL',
    'ICMS_PERCENTUAL',
    'COLUNA_PRECO_USADA',
    'ORIGEM_ARQUIVO',
    'NOME_ARQUIVO',
    'TIPO_ARQUIVO',
    'DATA_TABELA',
    'VALIDADE_TABELA',
    'OBSERVACAO',
    'USUARIO',
    'DATA_EXTRACAO',
    'DATA_APROVACAO',
    'ID_PRECO_GERADO'
  ];

  const ws = f2hObterOuCriarAba_(ss, F2H_ABA_REVISAO);
  f2hPrepararTabela_(ws, headers, '#002E56', '#FFFFFF');

  ws.setFrozenRows(1);
  ws.setFrozenColumns(5);

  f2hAplicarLarguras_(ws, {
    1: 170,
    2: 170,
    3: 130,
    4: 120,
    5: 220,
    6: 200,
    7: 150,
    8: 460,
    9: 460,
    10: 300,
    11: 300,
    12: 150,
    13: 120,
    14: 150,
    15: 90,
    16: 120,
    17: 130,
    18: 130,
    19: 120,
    20: 120,
    21: 180,
    22: 180,
    23: 330,
    24: 130,
    25: 120,
    26: 120,
    27: 420,
    28: 200,
    29: 150,
    30: 150,
    31: 180
  });

  const lastRow = Math.max(ws.getMaxRows(), 2);

  ws.getRange(2, 3, lastRow - 1, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList([
          'REVISAR',
          'APROVAR',
          'APROVADO',
          'REJEITAR',
          'SEM PREÇO',
          'ENVIADO_BD_PRECOS'
        ], true)
        .setAllowInvalid(true)
        .build()
    );

  ws.getRange(2, 4, lastRow - 1, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['ALTA', 'MEDIA', 'BAIXA'], true)
        .setAllowInvalid(true)
        .build()
    );
}

function f2hCriarAbaLogImportacao_(ss) {
  const headers = [
    'DATA_HORA',
    'ID_IMPORTACAO',
    'NOME_ARQUIVO',
    'ABA_ORIGEM',
    'LINHA_ORIGEM',
    'STATUS',
    'MOTIVO',
    'DESCRICAO_DETECTADA',
    'PRECO_DETECTADO',
    'COLUNA_PRECO_USADA',
    'LINHA_ORIGINAL'
  ];

  const ws = f2hObterOuCriarAba_(ss, F2H_ABA_LOG_IMPORTACAO);
  f2hPrepararTabela_(ws, headers, '#D80024', '#FFFFFF');

  ws.setFrozenRows(1);

  f2hAplicarLarguras_(ws, {
    1: 150,
    2: 170,
    3: 330,
    4: 160,
    5: 110,
    6: 130,
    7: 300,
    8: 420,
    9: 130,
    10: 180,
    11: 600
  });
}

/************************************************************
 * 2. RECEBER ARQUIVO DO WEB APP
 ************************************************************/

function importarArquivoTabelaPrecoFase2H(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bdImportacoes = ss.getSheetByName(F2H_ABA_BD_IMPORTACOES);
  const revisao = ss.getSheetByName(F2H_ABA_REVISAO);

  if (!bdImportacoes || !revisao) {
    throw new Error('Execute primeiro criarEstruturaImportadorUniversalFase2H.');
  }

  if (!payload || !payload.base64 || !payload.nomeArquivo) {
    throw new Error('Arquivo não recebido. Verifique o upload no Web App.');
  }

  const nomeArquivoValidacao = f2hUpper_(payload.nomeArquivo);

  if (nomeArquivoValidacao.endsWith('.XLSB')) {
    throw new Error(
      'Arquivo XLSB não suportado pelo importador online. ' +
      'Abra no Excel, salve como XLSX e importe novamente.'
    );
  }

  const idImportacao = f2hGerarId_('IMP');
  const agora = new Date();

  const fornecedor = f2hUpper_(payload.fornecedor);
  const fabricantePadrao = f2hUpper_(payload.fabricantePadrao);
  const origem = f2hUpper_(payload.origem || 'IMPORTADOR UNIVERSAL');
  const dataTabela = payload.dataTabela || agora;
  const validadeTabela = payload.validadeTabela || '';
  const tipoPrecoPrincipal = f2hUpper_(payload.tipoPrecoPrincipal || 'PRECO CUSTO');

  const nomeArquivo = f2hLimpar_(payload.nomeArquivo);
  const mimeType = f2hLimpar_(payload.mimeType || 'application/octet-stream');

  const base64Limpo = String(payload.base64).includes(',')
    ? String(payload.base64).split(',').pop()
    : String(payload.base64);

  const bytes = Utilities.base64Decode(base64Limpo);
  const blob = Utilities.newBlob(bytes, mimeType, nomeArquivo);

  const pasta = f2hObterOuCriarPastaDrive_();
  const arquivo = pasta.createFile(blob);

  const fileId = arquivo.getId();
  const fileUrl = arquivo.getUrl();

  let resultado;

  try {
    resultado = f2hProcessarArquivo_({
      idImportacao,
      fileId,
      nomeArquivo,
      mimeType,
      fornecedor,
      fabricantePadrao,
      origem,
      dataTabela,
      validadeTabela,
      tipoPrecoPrincipal
    });
  } catch (erro) {
    f2hRegistrarImportacao_({
      idImportacao,
      dataUpload: agora,
      fornecedor,
      fabricantePadrao,
      origem,
      dataTabela,
      validadeTabela,
      nomeArquivo,
      mimeType,
      fileId,
      fileUrl,
      status: 'ERRO',
      linhasLidas: 0,
      linhasExtraidas: 0,
      linhasValidas: 0,
      linhasIgnoradas: 0,
      observacao: erro.message
    });

    throw erro;
  }

  f2hRegistrarImportacao_({
    idImportacao,
    dataUpload: agora,
    fornecedor,
    fabricantePadrao,
    origem,
    dataTabela,
    validadeTabela,
    nomeArquivo,
    mimeType,
    fileId,
    fileUrl,
    status: 'IMPORTADO_PARA_REVISAO',
    linhasLidas: resultado.linhasLidas,
    linhasExtraidas: resultado.linhasExtraidas,
    linhasValidas: resultado.linhasValidas,
    linhasIgnoradas: resultado.linhasIgnoradas,
    observacao: resultado.observacao
  });

  f2hRegistrarLog_(
    'IMPORTADOR UNIVERSAL FASE 2H.4 | ' +
    idImportacao +
    ' | Arquivo: ' + nomeArquivo +
    ' | Lidas: ' + resultado.linhasLidas +
    ' | Extraídas: ' + resultado.linhasExtraidas +
    ' | Válidas: ' + resultado.linhasValidas +
    ' | Ignoradas: ' + resultado.linhasIgnoradas
  );

  return {
    ok: true,
    idImportacao,
    nomeArquivo,
    fileUrl,
    linhasLidas: resultado.linhasLidas,
    linhasExtraidas: resultado.linhasExtraidas,
    linhasValidas: resultado.linhasValidas,
    linhasIgnoradas: resultado.linhasIgnoradas,
    mensagem:
      'Arquivo importado para revisão. ' +
      'Linhas válidas: ' + resultado.linhasValidas +
      ' | Ignoradas: ' + resultado.linhasIgnoradas
  };
}

/************************************************************
 * 3. PROCESSADOR PRINCIPAL
 ************************************************************/

function f2hProcessarArquivo_(ctx) {
  const nome = f2hUpper_(ctx.nomeArquivo);
  const mime = f2hUpper_(ctx.mimeType);

  let linhasBrutas = [];

  if (f2hEhPlanilha_(nome, mime)) {
    linhasBrutas = f2hExtrairDePlanilha_(ctx);
  } else if (f2hEhTexto_(nome, mime)) {
    const texto = f2hExtrairTextoDireto_(ctx.fileId);
    linhasBrutas = f2hExtrairDeTexto_(texto, ctx);
  } else {
    const textoConvertido = f2hExtrairTextoPorConversaoOuOcr_(ctx);
    linhasBrutas = f2hExtrairDeTexto_(textoConvertido, ctx);
  }

  const linhasRevisao = [];
  let linhasIgnoradas = 0;

  for (let i = 0; i < linhasBrutas.length; i++) {
    const item = linhasBrutas[i];

    const descricao = f2hUpper_(item.descricaoOriginal);

    if (!descricao || !f2hDescricaoProdutoValida_(descricao)) {
      linhasIgnoradas++;

      f2hRegistrarLogImportacaoLinha_(ctx, item, {
        status: 'IGNORADA',
        motivo: 'DESCRICAO_INVALIDA',
        descricaoDetectada: descricao,
        precoDetectado: item.precoCusto,
        colunaPrecoUsada: item.colunaPrecoUsada,
        linhaOriginal: item.linhaOriginal
      });

      continue;
    }

    let preco = f2hToNumber_(item.precoCusto);
    const semPreco = preco === null || preco <= 0;

    if (semPreco) {
      preco = 0;
    }

    const descricaoNormalizada = f2hNormalizarTexto_(descricao);
    const fornecedor = f2hUpper_(item.fornecedor || ctx.fornecedor);
    const fabricante = f2hUpper_(item.fabricante || ctx.fabricantePadrao || f2hDetectarFabricanteNaDescricao_(descricao, ''));
    const skuOperacional = f2hUpper_(item.skuOperacional || descricaoNormalizada);

    const confianca = semPreco
      ? 'BAIXA'
      : f2hCalcularConfiancaLinha_({
          descricao,
          preco,
          unidade: item.unidade,
          fabricante,
          origem: ctx.mimeType,
          veioDeOcr: item.veioDeOcr
        });

    const status = semPreco ? 'SEM PREÇO' : 'REVISAR';

    const observacaoBase = f2hUpper_(item.observacao);
    const observacaoFinal = semPreco
      ? 'PRODUTO IMPORTADO SEM PREÇO CUSTO' + (observacaoBase ? ' | ' + observacaoBase : '')
      : observacaoBase;

    linhasRevisao.push({
      idRevisao: f2hGerarId_('REV'),
      idImportacao: ctx.idImportacao,
      status,
      confianca,
      fornecedor,
      fabricante,
      codigoFornecedor: f2hSomenteTexto_(item.codigoFornecedor),
      descricaoOriginal: descricao,
      descricaoNormalizada,
      skuMasterSugerido: '',
      skuOperacional,
      ean: f2hSomenteTexto_(item.ean),
      ncm: f2hSomenteTexto_(item.ncm),
      anvisa: f2hSomenteTexto_(item.anvisa),
      unidade: f2hUpper_(item.unidade),
      quantidade: f2hToNumber_(item.quantidade),
      precoCusto: preco,
      precoTotal: f2hToNumber_(item.precoTotal),
      ipiPercentual: f2hNormalizarPercentual_(item.ipiPercentual),
      icmsPercentual: f2hNormalizarPercentual_(item.icmsPercentual),
      colunaPrecoUsada: f2hUpper_(item.colunaPrecoUsada),
      origemArquivo: ctx.origem,
      nomeArquivo: ctx.nomeArquivo,
      tipoArquivo: f2hDetectarTipoArquivo_(ctx.nomeArquivo, ctx.mimeType),
      dataTabela: ctx.dataTabela,
      validadeTabela: ctx.validadeTabela,
      observacao: observacaoFinal,
      usuario: Session.getActiveUser().getEmail() || '',
      dataExtracao: new Date()
    });

    f2hRegistrarLogImportacaoLinha_(ctx, item, {
      status: 'VALIDA',
      motivo: semPreco ? 'SEM_PRECO_CUSTO_ZERO' : 'OK',
      descricaoDetectada: descricao,
      precoDetectado: preco,
      colunaPrecoUsada: item.colunaPrecoUsada,
      linhaOriginal: item.linhaOriginal
    });
  }

  f2hInserirLinhasRevisao_(linhasRevisao);

  return {
    linhasLidas: linhasBrutas.length,
    linhasExtraidas: linhasBrutas.length,
    linhasValidas: linhasRevisao.length,
    linhasIgnoradas,
    observacao:
      'Processamento concluído. ' +
      'Produtos sem preço foram importados com PRECO_CUSTO = 0,00. ' +
      'Linhas inválidas foram registradas em LOG_IMPORTACAO_PRECOS.'
  };
}

/************************************************************
 * 4. EXTRAÇÃO DE PLANILHA XLSX / XLS / CSV
 ************************************************************/

function f2hExtrairDePlanilha_(ctx) {
  const nome = f2hUpper_(ctx.nomeArquivo);

  if (nome.endsWith('.XLSB')) {
    throw new Error(
      'Arquivo XLSB não suportado pelo importador online. ' +
      'Abra o arquivo no Excel, salve como XLSX e importe novamente.'
    );
  }

  if (nome.endsWith('.CSV')) {
    const texto = f2hExtrairTextoDireto_(ctx.fileId);
    return f2hExtrairDeCsv_(texto, ctx);
  }

  const tempId = f2hConverterArquivoParaGoogleSheets_(ctx.fileId, ctx.nomeArquivo);
  const linhas = [];

  try {
    const tempSs = SpreadsheetApp.openById(tempId);
    const sheets = tempSs.getSheets();

    sheets.forEach(sheet => {
      const values = sheet.getDataRange().getValues();

      if (!values || values.length < 2) return;

      const extraidas = f2hExtrairDeMatriz_(values, ctx, sheet.getName());
      extraidas.forEach(x => linhas.push(x));
    });
  } finally {
    try {
      DriveApp.getFileById(tempId).setTrashed(true);
    } catch (e) {}
  }

  return linhas;
}

function f2hExtrairDeCsv_(texto, ctx) {
  const delimitador = f2hDetectarDelimitadorCsv_(texto);
  const linhasCsv = Utilities.parseCsv(texto, delimitador);
  return f2hExtrairDeMatriz_(linhasCsv, ctx, 'CSV');
}

function f2hExtrairDeMatriz_(values, ctx, nomeAba) {
  const headerInfo = f2hDetectarCabecalho_(values);

  if (headerInfo.encontrou) {
    return f2hExtrairMatrizComCabecalho_(values, headerInfo, ctx, nomeAba);
  }

  const porConteudo = f2hExtrairMatrizPorConteudo_(values, ctx, nomeAba);

  if (porConteudo.length > 0) {
    return porConteudo;
  }

  return f2hExtrairMatrizSemCabecalho_(values, ctx, nomeAba);
}

function f2hDetectarCabecalho_(values) {
  const maxLinhas = Math.min(values.length, 40);

  for (let r = 0; r < maxLinhas; r++) {
    const row = values[r].map(v => f2hNormalizarHeader_(v));

    const temProduto = row.some(h => f2hHeaderEhDescricao_(h));
    const temPreco = row.some(h => f2hHeaderEhPreco_(h));
    const temEan = row.some(h => f2hHeaderEhEan_(h));
    const temNcm = row.some(h => f2hHeaderEhNcm_(h));

    if (temProduto && (temPreco || temEan || temNcm)) {
      return {
        encontrou: true,
        linhaHeader: r,
        headers: row
      };
    }
  }

  return {
    encontrou: false,
    linhaHeader: -1,
    headers: []
  };
}

function f2hExtrairMatrizComCabecalho_(values, headerInfo, ctx, nomeAba) {
  const headers = headerInfo.headers;
  const idx = f2hMapearColunas_(headers);
  const linhas = [];

  for (let r = headerInfo.linhaHeader + 1; r < values.length; r++) {
    const row = values[r];

    if (f2hLinhaVazia_(row)) continue;

    const descricao = f2hGetByIndexes_(row, idx.descricao);
    const precoInfo = f2hGetPrecoPorIndexes_(row, idx.preco, headers);
    const precoTotal = f2hGetByIndexes_(row, idx.precoTotal);
    const unidade = f2hGetByIndexes_(row, idx.unidade);
    const quantidade = f2hGetByIndexes_(row, idx.quantidade);
    const fabricante = f2hGetByIndexes_(row, idx.fabricante);
    const codigo = f2hGetByIndexes_(row, idx.codigo);
    const ean = f2hGetByIndexes_(row, idx.ean);
    const ncm = f2hGetByIndexes_(row, idx.ncm);
    const anvisa = f2hGetByIndexes_(row, idx.anvisa);
    const ipi = f2hGetByIndexes_(row, idx.ipi);
    const icms = f2hGetByIndexes_(row, idx.icms);

    if (!descricao) continue;

    linhas.push({
      codigoFornecedor: codigo,
      descricaoOriginal: descricao,
      precoCusto: precoInfo.valor,
      precoTotal: precoTotal,
      unidade: unidade,
      quantidade: quantidade,
      fabricante: fabricante || ctx.fabricantePadrao,
      fornecedor: ctx.fornecedor,
      ean: ean,
      ncm: ncm,
      anvisa: anvisa,
      ipiPercentual: ipi,
      icmsPercentual: icms,
      colunaPrecoUsada: precoInfo.coluna,
      observacao: 'EXTRAÍDO DE PLANILHA | ABA: ' + nomeAba + ' | LINHA: ' + (r + 1),
      linhaOriginal: f2hLinhaOriginalMatriz_(row),
      abaOrigem: nomeAba,
      linhaOrigem: r + 1
    });
  }

  return linhas;
}

/**
 * Parser de planilha sem cabeçalho confiável.
 * Detecta colunas pelo conteúdo:
 * - descrição = coluna com mais textos longos com letras
 * - preço = coluna numérica decimal/preço
 * - EAN = 13 dígitos
 * - NCM = 8 dígitos ou padrão 3004.90.99
 */
function f2hExtrairMatrizPorConteudo_(values, ctx, nomeAba) {
  const limiteAnalise = Math.min(values.length, 250);
  const maxCols = Math.max.apply(null, values.slice(0, limiteAnalise).map(r => r.length));

  const scores = [];

  for (let c = 0; c < maxCols; c++) {
    let textoLongo = 0;
    let precos = 0;
    let eans = 0;
    let ncms = 0;
    let unidades = 0;
    let codigos = 0;

    for (let r = 0; r < limiteAnalise; r++) {
      const valor = values[r][c];
      const txt = f2hLimpar_(valor);
      const norm = f2hNormalizarTexto_(txt);

      if (!txt) continue;

      if (f2hDescricaoProdutoValida_(txt) && norm.length >= 12) textoLongo++;
      if (f2hParecePrecoValor_(txt)) precos++;
      if (f2hExtrairEan_(txt)) eans++;
      if (f2hExtrairNcm_(txt)) ncms++;
      if (f2hTokenEhUnidade_(txt)) unidades++;
      if (/^\d{3,10}$/.test(String(txt).replace(/\D/g, ''))) codigos++;
    }

    scores.push({
      col: c,
      textoLongo,
      precos,
      eans,
      ncms,
      unidades,
      codigos
    });
  }

  const colDescricao = scores.slice().sort((a, b) => b.textoLongo - a.textoLongo)[0];
  const colPreco = scores
    .filter(s => s.col !== (colDescricao ? colDescricao.col : -1))
    .sort((a, b) => b.precos - a.precos)[0];

  const colEan = scores.slice().sort((a, b) => b.eans - a.eans)[0];
  const colNcm = scores.slice().sort((a, b) => b.ncms - a.ncms)[0];
  const colUn = scores.slice().sort((a, b) => b.unidades - a.unidades)[0];
  const colCodigo = scores.slice().sort((a, b) => b.codigos - a.codigos)[0];

  if (!colDescricao || colDescricao.textoLongo < 5) return [];

  const linhas = [];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    if (f2hLinhaVazia_(row)) continue;

    const descricao = row[colDescricao.col];
    if (!f2hDescricaoProdutoValida_(descricao)) continue;

    const preco = colPreco && colPreco.precos > 0 ? row[colPreco.col] : '';
    const ean = colEan && colEan.eans > 0 ? row[colEan.col] : '';
    const ncm = colNcm && colNcm.ncms > 0 ? row[colNcm.col] : '';
    const unidade = colUn && colUn.unidades > 0 ? row[colUn.col] : '';
    const codigo = colCodigo && colCodigo.codigos > 0 ? row[colCodigo.col] : '';

    linhas.push({
      codigoFornecedor: codigo,
      descricaoOriginal: descricao,
      precoCusto: preco,
      precoTotal: '',
      unidade: unidade,
      quantidade: '',
      fabricante: ctx.fabricantePadrao,
      fornecedor: ctx.fornecedor,
      ean: f2hExtrairEan_(ean || f2hLinhaOriginalMatriz_(row)),
      ncm: f2hExtrairNcm_(ncm || f2hLinhaOriginalMatriz_(row)),
      anvisa: '',
      ipiPercentual: '',
      icmsPercentual: '',
      colunaPrecoUsada: colPreco ? 'COLUNA ' + (colPreco.col + 1) + ' POR CONTEUDO' : '',
      observacao: 'EXTRAÍDO POR CONTEÚDO | ABA: ' + nomeAba + ' | LINHA: ' + (r + 1),
      linhaOriginal: f2hLinhaOriginalMatriz_(row),
      abaOrigem: nomeAba,
      linhaOrigem: r + 1
    });
  }

  return linhas;
}

function f2hExtrairMatrizSemCabecalho_(values, ctx, nomeAba) {
  const linhas = [];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    if (f2hLinhaVazia_(row)) continue;

    const textoLinha = row
      .map(v => f2hLimpar_(v))
      .filter(Boolean)
      .join(' ');

    const item = f2hInterpretarLinhaTexto_(textoLinha, ctx);

    if (item) {
      item.observacao = 'EXTRAÍDO DE PLANILHA SEM CABEÇALHO | ABA: ' + nomeAba + ' | LINHA: ' + (r + 1);
      item.linhaOriginal = textoLinha;
      item.abaOrigem = nomeAba;
      item.linhaOrigem = r + 1;
      linhas.push(item);
    }
  }

  return linhas;
}

/************************************************************
 * 5. EXTRAÇÃO DE TXT / PDF / DOCX / IMAGEM
 ************************************************************/

function f2hExtrairDeTexto_(texto, ctx) {
  const linhasTexto = String(texto || '')
    .split(/\r?\n/)
    .map(l => f2hLimpar_(l))
    .filter(Boolean);

  const formato = f2hDetectarFormatoTexto_(linhasTexto, ctx);

  if (formato === 'DESCARPACK') return f2hExtrairTextoDescarpack_(linhasTexto, ctx);
  if (formato === 'JP_FARMA') return f2hExtrairTextoJpFarma_(linhasTexto, ctx);
  if (formato === 'FARMAVISION') return f2hExtrairTextoFarmavision_(linhasTexto, ctx);
  if (formato === 'NOVA_LINEA') return f2hExtrairTextoNovaLinea_(linhasTexto, ctx);
  if (formato === 'TABELA_LARGA') return f2hExtrairTextoTabelaLarga_(linhasTexto, ctx);
  if (formato === 'IMAGEM_SIMPLES') return f2hExtrairTextoImagemSimples_(linhasTexto, ctx);

  return f2hExtrairTextoGenerico_(linhasTexto, ctx);
}

function f2hDetectarFormatoTexto_(linhas, ctx) {
  const texto = f2hNormalizarTexto_(linhas.slice(0, 80).join(' '));
  const nome = f2hNormalizarTexto_(ctx.nomeArquivo || '');

  if (texto.includes('DESCARPACK') && texto.includes('PRECO POR EMBALAGEM')) return 'DESCARPACK';
  if (texto.includes('JP INDUSTRIA FARMACEUTICA') || nome.includes('JP FARMA')) return 'JP_FARMA';
  if (texto.includes('FARMA VISION') || texto.includes('FARMAVISION')) return 'FARMAVISION';
  if (texto.includes('NOVA LINEA') && texto.includes('VALOR UNITARIO')) return 'NOVA_LINEA';
  if (texto.includes('INTEREST') && texto.includes('SUFRAMA') && texto.includes('ANVISA')) return 'TABELA_LARGA';
  if (texto.includes('EAN') && texto.includes('PRODUTO') && (texto.includes('PC VENDA') || texto.includes('P VENDA'))) return 'IMAGEM_SIMPLES';
  if (nome.includes('S3MED') || nome.includes('S3 MED') || nome.includes('CIMED')) return 'IMAGEM_SIMPLES';

  return 'GENERICO';
}

function f2hExtrairTextoGenerico_(linhasTexto, ctx) {
  const linhas = [];
  let bufferDescricao = '';

  for (let i = 0; i < linhasTexto.length; i++) {
    const linha = linhasTexto[i];

    if (f2hLinhaEhCabecalhoOuRodape_(linha)) continue;

    const item = f2hInterpretarLinhaTexto_(linha, ctx);

    if (item) {
      if (bufferDescricao && item.descricaoOriginal && item.descricaoOriginal.length < 45) {
        item.descricaoOriginal = bufferDescricao + ' ' + item.descricaoOriginal;
        bufferDescricao = '';
      }

      item.observacao = item.observacao || 'EXTRAÍDO DE TEXTO/OCR GENÉRICO';
      item.linhaOriginal = linha;
      item.abaOrigem = 'TEXTO/OCR';
      item.linhaOrigem = i + 1;
      linhas.push(item);
      continue;
    }

    if (f2hPodeSerContinuidadeDescricao_(linha)) {
      bufferDescricao = bufferDescricao
        ? bufferDescricao + ' ' + linha
        : linha;

      if (bufferDescricao.length > 280) bufferDescricao = '';
    }
  }

  return linhas;
}

/**
 * Nova Linea:
 * Nr Item Código Descrição Qtd Unid Valor Unitario Valor Total
 * Descrição pode quebrar linhas.
 */
function f2hExtrairTextoNovaLinea_(linhasTexto, ctx) {
  const linhas = [];
  let buffer = '';

  for (let i = 0; i < linhasTexto.length; i++) {
    const linha = linhasTexto[i];

    if (f2hLinhaEhCabecalhoOuRodape_(linha)) continue;

    const tentativas = [
      linha,
      buffer ? buffer + ' ' + linha : linha
    ];

    let encontrado = null;

    for (let t = 0; t < tentativas.length; t++) {
      encontrado = f2hParseLinhaNovaLinea_(tentativas[t], ctx);
      if (encontrado) break;
    }

    if (encontrado) {
      encontrado.observacao = 'PARSER NOVA LINEA | LINHA: ' + (i + 1);
      encontrado.linhaOriginal = buffer ? buffer + ' ' + linha : linha;
      encontrado.abaOrigem = 'PDF/OCR';
      encontrado.linhaOrigem = i + 1;
      linhas.push(encontrado);
      buffer = '';
      continue;
    }

    if (f2hPodeSerContinuidadeDescricao_(linha)) {
      buffer = buffer ? buffer + ' ' + linha : linha;
      if (buffer.length > 420) buffer = '';
    }
  }

  return linhas.length ? linhas : f2hExtrairTextoGenerico_(linhasTexto, ctx);
}

function f2hParseLinhaNovaLinea_(linha, ctx) {
  const original = f2hLimpar_(linha);
  const valores = f2hExtrairNumerosComIndices_(original);
  const precos = valores.filter(v => v.numero !== null && v.numero >= 0 && /[,\.]\d{2,4}/.test(v.texto));

  if (precos.length < 2) return null;

  const precoTotalObj = precos[precos.length - 1];
  const precoUnitObj = precos[precos.length - 2];

  const antesPreco = original.substring(0, precoUnitObj.index).trim();
  const partes = antesPreco.split(/\s+/).filter(Boolean);

  if (partes.length < 5) return null;

  let unidade = '';
  let quantidade = '';

  for (let p = partes.length - 1; p >= 1; p--) {
    if (f2hTokenEhUnidade_(partes[p]) && f2hToNumber_(partes[p - 1]) !== null) {
      unidade = partes[p];
      quantidade = partes[p - 1];
      partes.splice(p - 1, 2);
      break;
    }
  }

  let codigo = '';

  if (/^\d{1,4}$/.test(partes[0]) && /^\d{3,10}$/.test(partes[1])) {
    partes.shift();
    codigo = partes.shift();
  } else if (/^\d{3,10}$/.test(partes[0])) {
    codigo = partes.shift();
  }

  const descricao = partes.join(' ');

  if (!f2hDescricaoProdutoValida_(descricao)) return null;

  return {
    codigoFornecedor: codigo,
    descricaoOriginal: descricao,
    precoCusto: precoUnitObj.numero,
    precoTotal: precoTotalObj.numero,
    unidade,
    quantidade,
    fabricante: ctx.fabricantePadrao || f2hDetectarFabricanteNaDescricao_(descricao, ''),
    fornecedor: ctx.fornecedor,
    ean: f2hExtrairEan_(original),
    ncm: f2hExtrairNcm_(original),
    anvisa: '',
    ipiPercentual: '',
    icmsPercentual: '',
    colunaPrecoUsada: 'VALOR UNITARIO'
  };
}

/**
 * Tabela larga tipo Joaomed:
 * Cod Descricao Marca Un Origem ICMS DISTRIB INTEREST DISTRIB PARANA CLIENTE FINAL...
 */
function f2hExtrairTextoTabelaLarga_(linhasTexto, ctx) {
  const linhas = [];

  for (let i = 0; i < linhasTexto.length; i++) {
    const linha = linhasTexto[i];

    if (f2hLinhaEhCabecalhoOuRodape_(linha)) continue;

    const item = f2hParseLinhaTabelaLarga_(linha, ctx);

    if (item) {
      item.observacao = 'PARSER TABELA LARGA | LINHA: ' + (i + 1);
      item.linhaOriginal = linha;
      item.abaOrigem = 'PDF/OCR';
      item.linhaOrigem = i + 1;
      linhas.push(item);
    }
  }

  return linhas.length ? linhas : f2hExtrairTextoGenerico_(linhasTexto, ctx);
}

function f2hParseLinhaTabelaLarga_(linha, ctx) {
  const original = f2hLimpar_(linha);
  const partes = original.split(/\s+/).filter(Boolean);

  if (partes.length < 12) return null;

  const codigo = partes[0];
  if (!/[A-Z0-9]{3,}/i.test(codigo)) return null;

  const valores = f2hExtrairNumerosComIndices_(original);
  const numerosPreco = valores.filter(v => v.numero !== null && v.numero >= 0 && /[,\.]\d{2,4}/.test(v.texto));

  if (numerosPreco.length < 3) return null;

  const precoObj = f2hEscolherPrecoTabelaLarga_(numerosPreco, ctx.tipoPrecoPrincipal);
  const preco = precoObj ? precoObj.numero : 0;

  const idxPreco = precoObj ? precoObj.index : original.length;
  let textoAntesPreco = original.substring(0, idxPreco).trim();

  textoAntesPreco = textoAntesPreco.replace(/^\S+\s+/, '');

  const tokens = textoAntesPreco.split(/\s+/).filter(Boolean);

  // Remove tokens finais que parecem marca/unidade/origem/icms.
  const fimDesc = Math.max(3, tokens.length - 4);
  const descricao = tokens.slice(0, fimDesc).join(' ');
  const fabricante = tokens.length > fimDesc ? tokens.slice(fimDesc).join(' ') : ctx.fabricantePadrao;

  if (!f2hDescricaoProdutoValida_(descricao)) return null;

  return {
    codigoFornecedor: codigo,
    descricaoOriginal: descricao,
    precoCusto: preco,
    precoTotal: '',
    unidade: '',
    quantidade: '',
    fabricante: fabricante || ctx.fabricantePadrao,
    fornecedor: ctx.fornecedor,
    ean: f2hExtrairEan_(original),
    ncm: f2hExtrairNcm_(original),
    anvisa: f2hExtrairAnvisa_(original),
    ipiPercentual: '',
    icmsPercentual: '',
    colunaPrecoUsada: ctx.tipoPrecoPrincipal || 'PRECO TABELA LARGA'
  };
}

function f2hEscolherPrecoTabelaLarga_(numerosPreco, tipoPrecoPrincipal) {
  if (!numerosPreco || !numerosPreco.length) return null;

  const tipo = f2hNormalizarTexto_(tipoPrecoPrincipal || '');

  if (tipo.includes('INTEREST') && numerosPreco.length >= 2) return numerosPreco[1];
  if (tipo.includes('CLIENTE') && numerosPreco.length >= 5) return numerosPreco[4];
  if (tipo.includes('FINAL') && numerosPreco.length >= 6) return numerosPreco[5];
  if (tipo.includes('PARANA') && numerosPreco.length >= 4) return numerosPreco[3];
  if (tipo.includes('SUFRAMA') && numerosPreco.length >= 8) return numerosPreco[7];

  // Padrão: usa primeira coluna de preço comercial real depois das colunas tributárias.
  if (numerosPreco.length >= 2) return numerosPreco[1];

  return numerosPreco[0];
}

/**
 * S3MED / CIMED / imagens simples:
 * EAN Produto Laboratório Filial Pç Venda
 * ou Codigo Produto Preço.
 */
function f2hExtrairTextoImagemSimples_(linhasTexto, ctx) {
  const linhas = [];

  for (let i = 0; i < linhasTexto.length; i++) {
    const linha = linhasTexto[i];

    if (f2hLinhaEhCabecalhoOuRodape_(linha)) continue;

    const item = f2hParseLinhaImagemSimples_(linha, ctx);

    if (item) {
      item.observacao = 'PARSER IMAGEM/OCR SIMPLES | LINHA: ' + (i + 1);
      item.linhaOriginal = linha;
      item.abaOrigem = 'PDF/OCR';
      item.linhaOrigem = i + 1;
      linhas.push(item);
    }
  }

  return linhas.length ? linhas : f2hExtrairTextoGenerico_(linhasTexto, ctx);
}

function f2hParseLinhaImagemSimples_(linha, ctx) {
  const original = f2hLimpar_(linha);
  const valores = f2hExtrairNumerosComIndices_(original);
  const precos = valores.filter(v => v.numero !== null && v.numero >= 0 && /[,\.]\d{2,4}/.test(v.texto));

  if (!precos.length) return null;

  const precoObj = precos[precos.length - 1];
  const preco = precoObj.numero;

  let antesPreco = original.substring(0, precoObj.index).trim();
  antesPreco = antesPreco.replace(/\bR\$\b/gi, '').trim();

  let ean = f2hExtrairEan_(antesPreco);
  let codigo = '';

  let descricaoParte = antesPreco;

  if (ean) {
    descricaoParte = descricaoParte.replace(ean, '').trim();
    codigo = ean;
  } else {
    const mCod = descricaoParte.match(/^\d{3,10}\s+/);
    if (mCod) {
      codigo = f2hLimpar_(mCod[0]);
      descricaoParte = descricaoParte.replace(mCod[0], '').trim();
    }
  }

  const descricao = descricaoParte;

  if (!f2hDescricaoProdutoValida_(descricao)) return null;

  return {
    codigoFornecedor: codigo,
    descricaoOriginal: descricao,
    precoCusto: preco,
    precoTotal: '',
    unidade: '',
    quantidade: '',
    fabricante: ctx.fabricantePadrao || f2hDetectarFabricanteNaDescricao_(descricao, ''),
    fornecedor: ctx.fornecedor,
    ean: ean,
    ncm: f2hExtrairNcm_(original),
    anvisa: '',
    ipiPercentual: '',
    icmsPercentual: '',
    colunaPrecoUsada: 'PÇ VENDA / PREÇO OCR'
  };
}

/**
 * Descarpack:
 * CODIGO PRODUTO PREÇO POR EMBALAGEM DE EMBARQUE TABELA IPI
 */
function f2hExtrairTextoDescarpack_(linhasTexto, ctx) {
  const linhas = [];

  for (let i = 0; i < linhasTexto.length; i++) {
    const linha = linhasTexto[i];

    if (f2hLinhaEhCabecalhoOuRodape_(linha)) continue;

    const item = f2hParseLinhaDescarpack_(linha, ctx);

    if (item) {
      item.observacao = 'PARSER DESCARPACK | LINHA: ' + (i + 1);
      item.linhaOriginal = linha;
      item.abaOrigem = 'PDF/OCR';
      item.linhaOrigem = i + 1;
      linhas.push(item);
    }
  }

  return linhas.length ? linhas : f2hExtrairTextoGenerico_(linhasTexto, ctx);
}

function f2hParseLinhaDescarpack_(linha, ctx) {
  const original = f2hLimpar_(linha);
  const m = original.match(/^(\d{4,})\s+(.+)$/);

  if (!m) return null;

  const codigo = m[1];
  let resto = m[2];

  const precoMatch = resto.match(/R\$\s*[-]|\b\d+,\d{2}\b|\b\d+\.\d{2}\b/);
  let preco = 0;
  let descricao = resto;

  if (precoMatch) {
    preco = f2hToNumber_(precoMatch[0]);
    if (preco === null) preco = 0;
    descricao = resto.substring(0, precoMatch.index).trim();
  }

  descricao = descricao
    .replace(/\bunidade\b/gi, ' ')
    .replace(/\bcaixa com\b/gi, ' CAIXA COM ')
    .replace(/\bcartucho\b/gi, ' CARTUCHO ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!f2hDescricaoProdutoValida_(descricao)) return null;

  const ipiMatch = original.match(/(\d{1,2},\d{2})%|(\d{1,2}\.\d{2})%/);
  const ipi = ipiMatch ? ipiMatch[0] : '';

  return {
    codigoFornecedor: codigo,
    descricaoOriginal: descricao,
    precoCusto: preco,
    precoTotal: '',
    unidade: '',
    quantidade: '',
    fabricante: ctx.fabricantePadrao || 'DESCARPACK',
    fornecedor: ctx.fornecedor,
    ean: f2hExtrairEan_(original),
    ncm: f2hExtrairNcm_(original),
    anvisa: '',
    ipiPercentual: ipi,
    icmsPercentual: '',
    colunaPrecoUsada: 'PREÇO DESCARPACK'
  };
}

/**
 * FarmaVision:
 * MARCA DESCRIÇÃO QUANTIDADES VALOR UNITARIO TOTAL
 */
function f2hExtrairTextoFarmavision_(linhasTexto, ctx) {
  const linhas = [];
  let buffer = '';

  for (let i = 0; i < linhasTexto.length; i++) {
    const linha = linhasTexto[i];

    if (f2hLinhaEhCabecalhoOuRodape_(linha)) continue;

    const combinado = buffer ? buffer + ' ' + linha : linha;
    const item = f2hParseLinhaFarmavision_(combinado, ctx);

    if (item) {
      item.observacao = 'PARSER FARMAVISION | LINHA: ' + (i + 1);
      item.linhaOriginal = combinado;
      item.abaOrigem = 'PDF/OCR';
      item.linhaOrigem = i + 1;
      linhas.push(item);
      buffer = '';
    } else if (f2hPodeSerContinuidadeDescricao_(linha)) {
      buffer = combinado;
      if (buffer.length > 250) buffer = '';
    }
  }

  return linhas.length ? linhas : f2hExtrairTextoGenerico_(linhasTexto, ctx);
}

function f2hParseLinhaFarmavision_(linha, ctx) {
  const original = f2hLimpar_(linha);

  if (!f2hNormalizarTexto_(original).includes('FARMA VISION') && !f2hNormalizarTexto_(original).includes('CLORETO')) {
    return null;
  }

  const valores = f2hExtrairNumerosComIndices_(original);
  const precos = valores.filter(v => v.numero !== null && /[$]|[,\.]\d{2}/.test(v.texto));

  if (!precos.length) return null;

  const precoObj = precos[0];

  let descricao = original.substring(0, precoObj.index).trim();
  descricao = descricao.replace(/FARMA\s*VISION/gi, '').trim();

  if (!f2hDescricaoProdutoValida_(descricao)) return null;

  return {
    codigoFornecedor: '',
    descricaoOriginal: descricao,
    precoCusto: precoObj.numero,
    precoTotal: precos.length > 1 ? precos[precos.length - 1].numero : '',
    unidade: '',
    quantidade: '',
    fabricante: ctx.fabricantePadrao || 'FARMA VISION',
    fornecedor: ctx.fornecedor,
    ean: '',
    ncm: '',
    anvisa: '',
    ipiPercentual: '',
    icmsPercentual: '',
    colunaPrecoUsada: 'VALOR UNITARIO'
  };
}

/**
 * JP FARMA:
 * Produto Descrição Quantidade Preço Unitário Preço Total
 */
function f2hExtrairTextoJpFarma_(linhasTexto, ctx) {
  const linhas = [];

  for (let i = 0; i < linhasTexto.length; i++) {
    const linha = linhasTexto[i];

    if (f2hLinhaEhCabecalhoOuRodape_(linha)) continue;

    const item = f2hParseLinhaJpFarma_(linha, ctx);

    if (item) {
      item.observacao = 'PARSER JP FARMA | LINHA: ' + (i + 1);
      item.linhaOriginal = linha;
      item.abaOrigem = 'PDF/OCR';
      item.linhaOrigem = i + 1;
      linhas.push(item);
    }
  }

  return linhas.length ? linhas : f2hExtrairTextoGenerico_(linhasTexto, ctx);
}

function f2hParseLinhaJpFarma_(linha, ctx) {
  const original = f2hLimpar_(linha);
  const valores = f2hExtrairNumerosComIndices_(original);
  const precos = valores.filter(v => v.numero !== null && /R\$|[,\.]\d{2}/.test(v.texto));

  if (precos.length < 2) return null;

  const precoUnitObj = precos[precos.length - 2];
  const precoTotalObj = precos[precos.length - 1];

  const antesPreco = original.substring(0, precoUnitObj.index).trim();
  const tokens = antesPreco.split(/\s+/).filter(Boolean);

  if (tokens.length < 4) return null;

  let codigo = '';

  if (/^\d{3,10}$/.test(tokens[0])) {
    codigo = tokens.shift();
  }

  const quantidadeToken = tokens[tokens.length - 1];
  let quantidade = '';

  if (f2hToNumber_(quantidadeToken) !== null) {
    quantidade = quantidadeToken;
    tokens.pop();
  }

  const descricao = tokens.join(' ');

  if (!f2hDescricaoProdutoValida_(descricao)) return null;

  return {
    codigoFornecedor: codigo,
    descricaoOriginal: descricao,
    precoCusto: precoUnitObj.numero,
    precoTotal: precoTotalObj.numero,
    unidade: '',
    quantidade,
    fabricante: ctx.fabricantePadrao || 'JP FARMA',
    fornecedor: ctx.fornecedor,
    ean: f2hExtrairEan_(original),
    ncm: f2hExtrairNcm_(original),
    anvisa: '',
    ipiPercentual: '',
    icmsPercentual: '',
    colunaPrecoUsada: 'PREÇO UNITÁRIO'
  };
}

/************************************************************
 * 6. INTERPRETAÇÃO GENÉRICA DE LINHA TEXTO
 ************************************************************/

function f2hInterpretarLinhaTexto_(linha, ctx) {
  const original = f2hLimpar_(linha);

  if (!original) return null;

  const normal = f2hNormalizarTexto_(original);

  if (!normal) return null;
  if (f2hLinhaEhCabecalhoOuRodape_(normal)) return null;

  const valores = f2hExtrairNumerosComIndices_(original);

  let precoUnitarioObj = null;
  let precoTotalObj = null;

  if (valores.length > 0) {
    const candidatosPreco = valores
      .filter(v => v.numero !== null && v.numero >= 0)
      .filter(v => f2hNumeroParecePreco_(v.texto, v.numero));

    if (candidatosPreco.length >= 2) {
      precoTotalObj = candidatosPreco[candidatosPreco.length - 1];
      precoUnitarioObj = candidatosPreco[candidatosPreco.length - 2];
    } else if (candidatosPreco.length === 1) {
      precoUnitarioObj = candidatosPreco[0];
    }
  }

  const precoUnitario = precoUnitarioObj ? precoUnitarioObj.numero : 0;
  const precoTotal = precoTotalObj ? precoTotalObj.numero : '';

  let textoAntesPreco = precoUnitarioObj
    ? original.substring(0, precoUnitarioObj.index).trim()
    : original;

  textoAntesPreco = textoAntesPreco
    .replace(/^\d+\s+/, '')
    .replace(/^\d{2,}\s+/, '')
    .trim();

  const partes = textoAntesPreco.split(/\s+/).filter(Boolean);

  let unidade = '';
  let quantidade = '';

  if (partes.length >= 2) {
    const possivelUnidade = partes[partes.length - 1];
    const possivelQtd = partes[partes.length - 2];

    if (f2hTokenEhUnidade_(possivelUnidade) && f2hToNumber_(possivelQtd) !== null) {
      unidade = possivelUnidade;
      quantidade = possivelQtd;
      partes.pop();
      partes.pop();
    }
  }

  let codigo = '';

  if (partes.length && /^\d{3,10}$/.test(partes[0])) {
    codigo = partes.shift();
  }

  let descricao = partes.join(' ').trim();

  if (descricao.length < 5) {
    descricao = textoAntesPreco;
  }

  if (!descricao || descricao.length < 5) return null;
  if (!f2hDescricaoProdutoValida_(descricao)) return null;

  const fabricanteDetectado = f2hDetectarFabricanteNaDescricao_(descricao, ctx.fabricantePadrao);

  return {
    codigoFornecedor: codigo,
    descricaoOriginal: descricao,
    precoCusto: precoUnitario,
    precoTotal: precoTotal,
    unidade: unidade,
    quantidade: quantidade,
    fabricante: fabricanteDetectado,
    fornecedor: ctx.fornecedor,
    ean: f2hExtrairEan_(original),
    ncm: f2hExtrairNcm_(original),
    anvisa: f2hExtrairAnvisa_(original),
    ipiPercentual: f2hExtrairPercentualPorPalavra_(original, 'IPI'),
    icmsPercentual: f2hExtrairPercentualPorPalavra_(original, 'ICMS'),
    colunaPrecoUsada: 'HEURÍSTICA TEXTO',
    observacao: 'LINHA INTERPRETADA POR HEURÍSTICA',
    linhaOriginal: original
  };
}

/************************************************************
 * 7. APROVAR REVISÃO PARA BD_PRECOS
 ************************************************************/

function aprovarRevisaoImportacaoPrecosFase2H() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const revisao = ss.getSheetByName(F2H_ABA_REVISAO);
  const bdPrecos = ss.getSheetByName(F2H_ABA_BD_PRECOS);

  if (!revisao) {
    throw new Error('Aba REVISAO_IMPORTACAO_PRECOS não encontrada.');
  }

  if (!bdPrecos) {
    throw new Error('Aba BD_PRECOS não encontrada.');
  }

  if (revisao.getLastRow() < 2) {
    return {
      ok: false,
      aprovadas: 0,
      duplicadas: 0,
      ignoradas: 0,
      mensagem: 'Não há linhas na revisão para aprovar.'
    };
  }

  const revValues = revisao.getDataRange().getValues();
  const revHeaders = revValues[0].map(String);
  const revIdx = f2hCriarIndice_(revHeaders);

  const precoHeaders = bdPrecos
    .getRange(1, 1, 1, bdPrecos.getLastColumn())
    .getValues()[0]
    .map(String);

  const precoIdx = f2hCriarIndice_(precoHeaders);
  const chavesExistentes = f2hCarregarChavesPrecosExistentes_(bdPrecos, precoIdx);

  const novasLinhas = [];
  const updates = [];

  let aprovadas = 0;
  let duplicadas = 0;
  let ignoradas = 0;
  let invalidas = 0;

  for (let i = 1; i < revValues.length; i++) {
    const row = revValues[i];

    const status = f2hUpper_(f2hGetCell_(row, revIdx, 'STATUS'));

    if (
      status !== 'APROVAR' &&
      status !== 'APROVADO' &&
      status !== 'SEM PREÇO'
    ) {
      ignoradas++;
      continue;
    }

    const fornecedor = f2hUpper_(f2hGetCell_(row, revIdx, 'FORNECEDOR'));
    const fabricante = f2hUpper_(f2hGetCell_(row, revIdx, 'FABRICANTE'));
    const descricao = f2hUpper_(f2hGetCell_(row, revIdx, 'DESCRICAO_ORIGINAL'));

    if (!fornecedor || !descricao || !f2hDescricaoProdutoValida_(descricao)) {
      invalidas++;
      continue;
    }

    const descricaoNormalizada = f2hNormalizarTexto_(
      f2hGetCell_(row, revIdx, 'DESCRICAO_NORMALIZADA') || descricao
    );

    const skuMaster = f2hUpper_(f2hGetCell_(row, revIdx, 'SKU_MASTER_SUGERIDO'));

    const skuOperacional = f2hUpper_(
      f2hGetCell_(row, revIdx, 'SKU_OPERACIONAL') || descricaoNormalizada
    );

    let precoCusto = f2hToNumber_(f2hGetCell_(row, revIdx, 'PRECO_CUSTO'));

    if (precoCusto === null || precoCusto < 0) {
      precoCusto = 0;
    }

    const dataTabela = f2hGetCell_(row, revIdx, 'DATA_TABELA') || new Date();
    const validadeTabela = f2hGetCell_(row, revIdx, 'VALIDADE_TABELA') || '';

    const chave = f2hMontarChavePreco_(dataTabela, fornecedor, descricao, precoCusto);

    if (chavesExistentes.has(chave)) {
      duplicadas++;

      updates.push({
        rowNumber: i + 1,
        status: 'ENVIADO_BD_PRECOS',
        idPreco: 'DUPLICADO'
      });

      continue;
    }

    const idPreco = f2hGerarId_('PRE');

    const ipi = f2hNormalizarPercentual_(f2hGetCell_(row, revIdx, 'IPI_PERCENTUAL'));
    const icms = f2hNormalizarPercentual_(f2hGetCell_(row, revIdx, 'ICMS_PERCENTUAL'));

    const custoTotalEstimado = f2hArredondar_(precoCusto * (1 + ipi + icms), 4);

    const observacaoOriginal = f2hUpper_(f2hGetCell_(row, revIdx, 'OBSERVACAO'));
    const observacaoFinal = precoCusto === 0
      ? 'PREÇO CUSTO 0,00 - AGUARDANDO TABELA COM PREÇO' + (observacaoOriginal ? ' | ' + observacaoOriginal : '')
      : observacaoOriginal;

    const novaLinha = new Array(precoHeaders.length).fill('');

    f2hSetByHeader_(novaLinha, precoIdx, 'ID_PRECO', idPreco);
    f2hSetByHeader_(novaLinha, precoIdx, 'DATA_LANCAMENTO', new Date());
    f2hSetByHeader_(novaLinha, precoIdx, 'DATA_TABELA', dataTabela);
    f2hSetByHeader_(novaLinha, precoIdx, 'VALIDADE_TABELA', validadeTabela);
    f2hSetByHeader_(novaLinha, precoIdx, 'FORNECEDOR', fornecedor);
    f2hSetByHeader_(novaLinha, precoIdx, 'FABRICANTE_MARCA', fabricante);
    f2hSetByHeader_(novaLinha, precoIdx, 'DESCRICAO_FORNECEDOR', descricao);
    f2hSetByHeader_(novaLinha, precoIdx, 'DESCRICAO_NORMALIZADA', descricaoNormalizada);
    f2hSetByHeader_(novaLinha, precoIdx, 'SKU_MASTER', skuMaster);
    f2hSetByHeader_(novaLinha, precoIdx, 'SKU_OPERACIONAL', skuOperacional);
    f2hSetByHeader_(novaLinha, precoIdx, 'NCM', f2hGetCell_(row, revIdx, 'NCM'));
    f2hSetByHeader_(novaLinha, precoIdx, 'EAN', f2hGetCell_(row, revIdx, 'EAN'));
    f2hSetByHeader_(novaLinha, precoIdx, 'CST', '');
    f2hSetByHeader_(novaLinha, precoIdx, 'UNIDADE', f2hGetCell_(row, revIdx, 'UNIDADE'));
    f2hSetByHeader_(novaLinha, precoIdx, 'PRECO_CUSTO', precoCusto);
    f2hSetByHeader_(novaLinha, precoIdx, 'IPI_PERCENTUAL', ipi);
    f2hSetByHeader_(novaLinha, precoIdx, 'ICMS_PERCENTUAL', icms);
    f2hSetByHeader_(novaLinha, precoIdx, 'FRETE_PERCENTUAL', 0);
    f2hSetByHeader_(novaLinha, precoIdx, 'OUTRAS_DESPESAS_PERCENTUAL', 0);
    f2hSetByHeader_(novaLinha, precoIdx, 'CUSTO_TOTAL_ESTIMADO', custoTotalEstimado);
    f2hSetByHeader_(novaLinha, precoIdx, 'ORIGEM', f2hGetCell_(row, revIdx, 'ORIGEM_ARQUIVO'));
    f2hSetByHeader_(novaLinha, precoIdx, 'ARQUIVO_ORIGEM', f2hGetCell_(row, revIdx, 'NOME_ARQUIVO'));
    f2hSetByHeader_(novaLinha, precoIdx, 'OBSERVACAO', observacaoFinal);
    f2hSetByHeader_(novaLinha, precoIdx, 'USUARIO', Session.getActiveUser().getEmail() || '');
    f2hSetByHeader_(novaLinha, precoIdx, 'DATA_ATUALIZACAO', new Date());

    novasLinhas.push(novaLinha);
    chavesExistentes.add(chave);
    aprovadas++;

    updates.push({
      rowNumber: i + 1,
      status: 'ENVIADO_BD_PRECOS',
      idPreco: idPreco
    });
  }

  if (novasLinhas.length > 0) {
    bdPrecos
      .getRange(bdPrecos.getLastRow() + 1, 1, novasLinhas.length, precoHeaders.length)
      .setValues(novasLinhas);
  }

  const colStatus = revIdx['STATUS'] + 1;
  const colDataAprovacao = revIdx['DATA_APROVACAO'] + 1;
  const colIdPreco = revIdx['ID_PRECO_GERADO'] + 1;

  updates.forEach(function(u) {
    revisao.getRange(u.rowNumber, colStatus).setValue(u.status);
    revisao.getRange(u.rowNumber, colDataAprovacao).setValue(new Date());
    revisao.getRange(u.rowNumber, colIdPreco).setValue(u.idPreco);
  });

  f2hRegistrarLog_(
    'APROVAÇÃO IMPORTAÇÃO PREÇOS | APROVADAS: ' + aprovadas +
    ' | DUPLICADAS: ' + duplicadas +
    ' | IGNORADAS: ' + ignoradas +
    ' | INVALIDAS: ' + invalidas
  );

  return {
    ok: true,
    aprovadas: aprovadas,
    duplicadas: duplicadas,
    ignoradas: ignoradas,
    invalidas: invalidas,
    mensagem:
      'Aprovação concluída. ' +
      'Enviadas para BD_PRECOS: ' + aprovadas +
      ' | Duplicadas: ' + duplicadas +
      ' | Ignoradas: ' + ignoradas +
      ' | Inválidas: ' + invalidas
  };
}

/************************************************************
 * 8. REGISTROS EM ABAS
 ************************************************************/

function f2hRegistrarImportacao_(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(F2H_ABA_BD_IMPORTACOES);

  if (!ws) throw new Error('Aba BD_IMPORTACOES_ARQUIVOS não encontrada.');

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0].map(String);
  const idx = f2hCriarIndice_(headers);

  const row = new Array(headers.length).fill('');

  f2hSetByHeader_(row, idx, 'ID_IMPORTACAO', dados.idImportacao);
  f2hSetByHeader_(row, idx, 'DATA_UPLOAD', dados.dataUpload);
  f2hSetByHeader_(row, idx, 'FORNECEDOR', dados.fornecedor);
  f2hSetByHeader_(row, idx, 'FABRICANTE_PADRAO', dados.fabricantePadrao);
  f2hSetByHeader_(row, idx, 'ORIGEM', dados.origem);
  f2hSetByHeader_(row, idx, 'DATA_TABELA', dados.dataTabela);
  f2hSetByHeader_(row, idx, 'VALIDADE_TABELA', dados.validadeTabela);
  f2hSetByHeader_(row, idx, 'NOME_ARQUIVO', dados.nomeArquivo);
  f2hSetByHeader_(row, idx, 'MIME_TYPE', dados.mimeType);
  f2hSetByHeader_(row, idx, 'FILE_ID', dados.fileId);
  f2hSetByHeader_(row, idx, 'FILE_URL', dados.fileUrl);
  f2hSetByHeader_(row, idx, 'STATUS', dados.status);
  f2hSetByHeader_(row, idx, 'LINHAS_LIDAS', dados.linhasLidas);
  f2hSetByHeader_(row, idx, 'LINHAS_EXTRAIDAS', dados.linhasExtraidas);
  f2hSetByHeader_(row, idx, 'LINHAS_VALIDAS', dados.linhasValidas);
  f2hSetByHeader_(row, idx, 'LINHAS_IGNORADAS', dados.linhasIgnoradas);
  f2hSetByHeader_(row, idx, 'OBSERVACAO', dados.observacao);
  f2hSetByHeader_(row, idx, 'USUARIO', Session.getActiveUser().getEmail() || '');
  f2hSetByHeader_(row, idx, 'DATA_ATUALIZACAO', new Date());

  ws.appendRow(row);
}

function f2hInserirLinhasRevisao_(linhas) {
  if (!linhas || !linhas.length) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(F2H_ABA_REVISAO);

  if (!ws) throw new Error('Aba REVISAO_IMPORTACAO_PRECOS não encontrada.');

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0].map(String);
  const idx = f2hCriarIndice_(headers);

  const rows = linhas.map(item => {
    const row = new Array(headers.length).fill('');

    f2hSetByHeader_(row, idx, 'ID_REVISAO', item.idRevisao);
    f2hSetByHeader_(row, idx, 'ID_IMPORTACAO', item.idImportacao);
    f2hSetByHeader_(row, idx, 'STATUS', item.status);
    f2hSetByHeader_(row, idx, 'CONFIANCA', item.confianca);
    f2hSetByHeader_(row, idx, 'FORNECEDOR', item.fornecedor);
    f2hSetByHeader_(row, idx, 'FABRICANTE', item.fabricante);
    f2hSetByHeader_(row, idx, 'CODIGO_FORNECEDOR', item.codigoFornecedor);
    f2hSetByHeader_(row, idx, 'DESCRICAO_ORIGINAL', item.descricaoOriginal);
    f2hSetByHeader_(row, idx, 'DESCRICAO_NORMALIZADA', item.descricaoNormalizada);
    f2hSetByHeader_(row, idx, 'SKU_MASTER_SUGERIDO', item.skuMasterSugerido);
    f2hSetByHeader_(row, idx, 'SKU_OPERACIONAL', item.skuOperacional);
    f2hSetByHeader_(row, idx, 'EAN', item.ean);
    f2hSetByHeader_(row, idx, 'NCM', item.ncm);
    f2hSetByHeader_(row, idx, 'ANVISA', item.anvisa);
    f2hSetByHeader_(row, idx, 'UNIDADE', item.unidade);
    f2hSetByHeader_(row, idx, 'QUANTIDADE', item.quantidade);
    f2hSetByHeader_(row, idx, 'PRECO_CUSTO', item.precoCusto);
    f2hSetByHeader_(row, idx, 'PRECO_TOTAL', item.precoTotal);
    f2hSetByHeader_(row, idx, 'IPI_PERCENTUAL', item.ipiPercentual);
    f2hSetByHeader_(row, idx, 'ICMS_PERCENTUAL', item.icmsPercentual);
    f2hSetByHeader_(row, idx, 'COLUNA_PRECO_USADA', item.colunaPrecoUsada);
    f2hSetByHeader_(row, idx, 'ORIGEM_ARQUIVO', item.origemArquivo);
    f2hSetByHeader_(row, idx, 'NOME_ARQUIVO', item.nomeArquivo);
    f2hSetByHeader_(row, idx, 'TIPO_ARQUIVO', item.tipoArquivo);
    f2hSetByHeader_(row, idx, 'DATA_TABELA', item.dataTabela);
    f2hSetByHeader_(row, idx, 'VALIDADE_TABELA', item.validadeTabela);
    f2hSetByHeader_(row, idx, 'OBSERVACAO', item.observacao);
    f2hSetByHeader_(row, idx, 'USUARIO', item.usuario);
    f2hSetByHeader_(row, idx, 'DATA_EXTRACAO', item.dataExtracao);

    return row;
  });

  ws.getRange(ws.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function f2hRegistrarLogImportacaoLinha_(ctx, item, dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(F2H_ABA_LOG_IMPORTACAO);

  if (!ws) {
    f2hCriarAbaLogImportacao_(ss);
    ws = ss.getSheetByName(F2H_ABA_LOG_IMPORTACAO);
  }

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0].map(String);
  const idx = f2hCriarIndice_(headers);
  const row = new Array(headers.length).fill('');

  f2hSetByHeader_(row, idx, 'DATA_HORA', new Date());
  f2hSetByHeader_(row, idx, 'ID_IMPORTACAO', ctx.idImportacao);
  f2hSetByHeader_(row, idx, 'NOME_ARQUIVO', ctx.nomeArquivo);
  f2hSetByHeader_(row, idx, 'ABA_ORIGEM', item.abaOrigem || '');
  f2hSetByHeader_(row, idx, 'LINHA_ORIGEM', item.linhaOrigem || '');
  f2hSetByHeader_(row, idx, 'STATUS', dados.status);
  f2hSetByHeader_(row, idx, 'MOTIVO', dados.motivo);
  f2hSetByHeader_(row, idx, 'DESCRICAO_DETECTADA', dados.descricaoDetectada);
  f2hSetByHeader_(row, idx, 'PRECO_DETECTADO', dados.precoDetectado);
  f2hSetByHeader_(row, idx, 'COLUNA_PRECO_USADA', dados.colunaPrecoUsada);
  f2hSetByHeader_(row, idx, 'LINHA_ORIGINAL', dados.linhaOriginal);

  ws.appendRow(row);
}

/************************************************************
 * 9. CONVERSÕES GOOGLE DRIVE
 ************************************************************/

function f2hConverterArquivoParaGoogleSheets_(fileId, nomeArquivo) {
  f2hValidarDriveApi_();

  const resource = {
    title: 'TMP_SHEETS_' + nomeArquivo,
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const convertido = Drive.Files.copy(resource, fileId, {
    convert: true
  });

  return convertido.id;
}

function f2hExtrairTextoPorConversaoOuOcr_(ctx) {
  f2hValidarDriveApi_();

  const tipo = f2hDetectarTipoArquivo_(ctx.nomeArquivo, ctx.mimeType);

  const resource = {
    title: 'TMP_DOC_' + ctx.nomeArquivo,
    mimeType: MimeType.GOOGLE_DOCS
  };

  const options = {
    convert: true
  };

  if (tipo === 'PDF' || tipo === 'JPEG' || tipo === 'PNG' || tipo === 'IMAGEM') {
    options.ocr = true;
    options.ocrLanguage = 'pt';
  }

  const docConvertido = Drive.Files.copy(resource, ctx.fileId, options);
  const docId = docConvertido.id;

  try {
    const doc = DocumentApp.openById(docId);
    return doc.getBody().getText();
  } finally {
    try {
      DriveApp.getFileById(docId).setTrashed(true);
    } catch (e) {}
  }
}

function f2hExtrairTextoDireto_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();

  try {
    return blob.getDataAsString('UTF-8');
  } catch (e) {
    return blob.getDataAsString();
  }
}

function f2hValidarDriveApi_() {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error(
      'Drive API avançado não está ativado.\n\n' +
      'Ative em:\n' +
      'Apps Script > Serviços > + > Drive API > Adicionar.\n\n' +
      'Sem isso, o sistema não consegue converter XLSX, PDF, DOCX ou imagem para leitura/OCR.'
    );
  }
}

/************************************************************
 * 10. DETECÇÃO DE COLUNAS
 ************************************************************/

function f2hMapearColunas_(headers) {
  const idx = {
    codigo: [],
    descricao: [],
    preco: [],
    precoTotal: [],
    unidade: [],
    quantidade: [],
    fabricante: [],
    ean: [],
    ncm: [],
    anvisa: [],
    ipi: [],
    icms: []
  };

  headers.forEach((h, i) => {
    if (f2hHeaderEhCodigo_(h)) idx.codigo.push(i);
    if (f2hHeaderEhDescricao_(h)) idx.descricao.push(i);
    if (f2hHeaderEhPrecoTotal_(h)) idx.precoTotal.push(i);
    if (f2hHeaderEhUnidade_(h)) idx.unidade.push(i);
    if (f2hHeaderEhQuantidade_(h)) idx.quantidade.push(i);
    if (f2hHeaderEhFabricante_(h)) idx.fabricante.push(i);
    if (f2hHeaderEhEan_(h)) idx.ean.push(i);
    if (f2hHeaderEhNcm_(h)) idx.ncm.push(i);
    if (f2hHeaderEhAnvisa_(h)) idx.anvisa.push(i);
    if (f2hHeaderEhIpi_(h)) idx.ipi.push(i);
    if (f2hHeaderEhIcms_(h)) idx.icms.push(i);
  });

  const colunasPrecoOrdenadas = [];

  headers.forEach((h, i) => {
    const score = f2hScoreColunaPreco_(h);

    if (score > 0) {
      colunasPrecoOrdenadas.push({
        index: i,
        score: score
      });
    }
  });

  colunasPrecoOrdenadas
    .sort((a, b) => b.score - a.score)
    .forEach(item => idx.preco.push(item.index));

  return idx;
}

function f2hHeaderEhCodigo_(h) {
  const header = f2hNormalizarHeader_(h);

  return [
    'COD',
    'CODIGO',
    'CÓDIGO',
    'CODIGO FORNECEDOR',
    'COD FORNECEDOR',
    'PRODUTO'
  ].some(x => header === f2hNormalizarHeader_(x));
}

function f2hHeaderEhDescricao_(h) {
  const header = f2hNormalizarHeader_(h);

  return [
    'DESCRICAO',
    'DESCRICAO DO PRODUTO',
    'DESCRICAO PRODUTO',
    'PRODUTO',
    'ITEM',
    'MATERIAL',
    'NOME',
    'DESCR'
  ].some(x => header === f2hNormalizarHeader_(x) || header.includes(f2hNormalizarHeader_(x)));
}

function f2hHeaderEhPreco_(h) {
  return f2hScoreColunaPreco_(h) > 0;
}

function f2hScoreColunaPreco_(h) {
  const header = f2hNormalizarHeader_(h);

  if (!header) return 0;

  const bloqueios = [
    'STATUS DE PRECO',
    'STATUS PRECO',
    'STATUS',
    'REVISTA',
    'CMED',
    'OBS',
    'OBSERVACAO',
    'SITUACAO',
    'SITUAÇÃO'
  ];

  for (let i = 0; i < bloqueios.length; i++) {
    if (header === f2hNormalizarHeader_(bloqueios[i]) || header.includes(f2hNormalizarHeader_(bloqueios[i]))) {
      return 0;
    }
  }

  if (f2hHeaderEhPrecoTotal_(header)) {
    return 0;
  }

  const prioridadeAlta = [
    'PRECO CUSTO',
    'PREÇO CUSTO',
    'CUSTO',
    'VALOR UNITARIO',
    'VALOR UNITÁRIO',
    'PRECO UNITARIO',
    'PREÇO UNITÁRIO',
    'PRECO UNIT',
    'PREÇO UNIT',
    'VALOR UNIT',
    'VL UNIT',
    'VLR UNIT',
    'P UNIT',
    'P. UNIT',
    'PÇ VENDA',
    'PC VENDA',
    'P VENDA',
    'PRECO VENDA',
    'PREÇO VENDA'
  ];

  for (let i = 0; i < prioridadeAlta.length; i++) {
    const termo = f2hNormalizarHeader_(prioridadeAlta[i]);

    if (header === termo || header.includes(termo)) {
      return 100;
    }
  }

  const prioridadeMedia = [
    'PRECO',
    'PREÇO',
    'INTEREST',
    'CLIENTE',
    'FINAL',
    'DISTRIB',
    'PARANA',
    'SUFRAMA'
  ];

  for (let i = 0; i < prioridadeMedia.length; i++) {
    const termo = f2hNormalizarHeader_(prioridadeMedia[i]);

    if (header === termo || header.includes(termo)) {
      return 50;
    }
  }

  return 0;
}

function f2hHeaderEhPrecoTotal_(h) {
  const header = f2hNormalizarHeader_(h);

  return [
    'VALOR TOTAL',
    'PRECO TOTAL',
    'PREÇO TOTAL',
    'TOTAL'
  ].some(x => header === f2hNormalizarHeader_(x) || header.includes(f2hNormalizarHeader_(x)));
}

function f2hHeaderEhUnidade_(h) {
  const header = f2hNormalizarHeader_(h);

  return [
    'UN',
    'UND',
    'UNID',
    'UNIDADE',
    'UNIDADE DE COMPRA'
  ].some(x => header === f2hNormalizarHeader_(x) || header.includes(f2hNormalizarHeader_(x)));
}

function f2hHeaderEhQuantidade_(h) {
  const header = f2hNormalizarHeader_(h);

  return [
    'QUANTIDADE',
    'QTD',
    'QTDE',
    'CX COM',
    'CX CONTEM',
    'SALDO',
    'ESTOQUE',
    'DISPONIVEL'
  ].some(x => header === f2hNormalizarHeader_(x) || header.includes(f2hNormalizarHeader_(x)));
}

function f2hHeaderEhFabricante_(h) {
  const header = f2hNormalizarHeader_(h);

  return [
    'FABRICANTE',
    'MARCA',
    'LABORATORIO',
    'LABORATÓRIO'
  ].some(x => header === f2hNormalizarHeader_(x) || header.includes(f2hNormalizarHeader_(x)));
}

function f2hHeaderEhEan_(h) {
  const header = f2hNormalizarHeader_(h);

  return [
    'EAN',
    'COD EAN',
    'CÓD EAN',
    'COD BARRAS',
    'CODIGO BARRAS',
    'CÓDIGO BARRAS'
  ].some(x => header === f2hNormalizarHeader_(x) || header.includes(f2hNormalizarHeader_(x)));
}

function f2hHeaderEhNcm_(h) {
  const header = f2hNormalizarHeader_(h);

  return header === 'NCM' || header.includes('NCM');
}

function f2hHeaderEhAnvisa_(h) {
  const header = f2hNormalizarHeader_(h);

  return header === 'ANVISA' || header.includes('ANVISA') || header.includes('REGISTRO');
}

function f2hHeaderEhIpi_(h) {
  const header = f2hNormalizarHeader_(h);

  return header === 'IPI' || header.includes('IPI');
}

function f2hHeaderEhIcms_(h) {
  const header = f2hNormalizarHeader_(h);

  return header === 'ICMS' || header.includes('ICMS') || header.includes('ST');
}

/************************************************************
 * 11. HELPERS DE INTERPRETAÇÃO
 ************************************************************/

function f2hGetByIndexes_(row, indexes) {
  if (!indexes || !indexes.length) return '';

  for (let i = 0; i < indexes.length; i++) {
    const valor = row[indexes[i]];
    if (f2hLimpar_(valor)) return valor;
  }

  return '';
}

function f2hGetPrecoPorIndexes_(row, indexes, headers) {
  if (!indexes || !indexes.length) {
    return {
      valor: '',
      coluna: ''
    };
  }

  for (let i = 0; i < indexes.length; i++) {
    const col = indexes[i];
    const valor = row[col];

    if (f2hLimpar_(valor) || f2hToNumber_(valor) === 0) {
      return {
        valor: valor,
        coluna: headers && headers[col] ? headers[col] : 'COLUNA ' + (col + 1)
      };
    }
  }

  return {
    valor: '',
    coluna: ''
  };
}

function f2hLinhaVazia_(row) {
  return !row || row.every(v => !f2hLimpar_(v));
}

function f2hLinhaOriginalMatriz_(row) {
  return row.map(v => f2hLimpar_(v)).filter(Boolean).join(' | ');
}

function f2hLinhaEhCabecalhoOuRodape_(linha) {
  const t = f2hNormalizarTexto_(linha);

  if (!t) return true;

  const bloqueios = [
    'PAGINA',
    'EMITIDO EM',
    'CONTINUA',
    'TOTAL DOS PRODUTOS',
    'TOTAL DA PROPOSTA',
    'CONDICAO DE PAGAMENTO',
    'CONDIÇÕES',
    'CONDIÇÃO',
    'CNPJ',
    'RAZAO SOCIAL',
    'ENDERECO',
    'CLIENTE MEDBOX',
    'MEDBOX DISTRIBUIDORA',
    'NR ITEM CODIGO DESCRICAO',
    'COD DESCRICAO',
    'PRODUTO DESCRICAO QUANTIDADE',
    'TABELA DE PRECOS',
    'TABELA DE PREÇOS'
  ];

  return bloqueios.some(b => t.includes(f2hNormalizarTexto_(b)));
}

function f2hPodeSerContinuidadeDescricao_(linha) {
  const t = f2hLimpar_(linha);

  if (!t) return false;
  if (t.length < 8) return false;

  const temPreco = f2hExtrairNumerosComIndices_(t)
    .some(v => f2hNumeroParecePreco_(v.texto, v.numero));

  if (temPreco) return false;

  return true;
}

function f2hExtrairNumerosComIndices_(texto) {
  const regex = /R?\$?\s*-|\$?\s*-?\d{1,3}(?:\.\d{3})*(?:,\d{1,4})|-?\d+(?:,\d{1,4})|-?\d+\.\d{1,4}/g;
  const out = [];

  let m;

  while ((m = regex.exec(texto)) !== null) {
    const raw = m[0];

    out.push({
      texto: raw,
      numero: f2hToNumber_(raw),
      index: m.index
    });
  }

  return out;
}

function f2hParecePrecoValor_(valor) {
  const txt = f2hLimpar_(valor);
  const n = f2hToNumber_(txt);

  if (n === null) return false;
  if (txt.includes('R$')) return true;
  if (/[,\.]\d{2,4}$/.test(txt)) return true;
  if (n >= 0 && n < 100000) return true;

  return false;
}

function f2hNumeroParecePreco_(texto, numero) {
  if (numero === null || numero < 0) return false;

  const raw = String(texto || '');

  if (raw.includes('R$')) return true;
  if (raw.includes(',') || raw.includes('.')) return true;
  if (numero > 0 && numero < 10000) return true;

  return false;
}

function f2hTokenEhUnidade_(token) {
  const t = f2hNormalizarTexto_(token);

  return [
    'UN',
    'UND',
    'UNID',
    'CX',
    'FR',
    'FA',
    'AP',
    'AMP',
    'RL',
    'ENV',
    'PCT',
    'PC',
    'GL',
    'LT',
    'ML',
    'BO',
    'BOL',
    'KT',
    'TB',
    'CP',
    'CAP'
  ].includes(t);
}

function f2hDescricaoProdutoValida_(descricao) {
  const txt = f2hUpper_(descricao);

  if (!txt) return false;
  if (txt.length < 5) return false;

  if (!/[A-Z]/.test(txt)) return false;

  const letras = txt.match(/[A-Z]/g) || [];
  if (letras.length < 3) return false;

  const somenteNumerosPontuacao = txt.replace(/[0-9\s.,;:/\\\-+()]/g, '').trim();
  if (!somenteNumerosPontuacao) return false;

  const normal = f2hNormalizarTexto_(txt);

  const bloqueiosExatos = [
    'COD',
    'CODIGO',
    'CÓDIGO',
    'EAN',
    'NCM',
    'ANVISA',
    'PRECO',
    'PREÇO',
    'CUSTO',
    'TOTAL',
    'STATUS',
    'REVISTA MENOR',
    'REVISTA IGUAL',
    'SEM CMED',
    'R$',
    '-'
  ];

  for (let i = 0; i < bloqueiosExatos.length; i++) {
    if (normal === f2hNormalizarTexto_(bloqueiosExatos[i])) {
      return false;
    }
  }

  if (f2hLinhaEhCabecalhoOuRodape_(txt)) return false;

  return true;
}

function f2hDetectarFabricanteNaDescricao_(descricao, fabricantePadrao) {
  if (fabricantePadrao) return fabricantePadrao;

  const texto = f2hUpper_(descricao);

  const marcas = [
    'NOVA LINEA',
    'JP FARMA',
    'JP INDUSTRIA',
    'FRESENIUS',
    'FRESENIUS KABI',
    'DESCARPACK',
    'CIMED',
    'JOAOMED',
    'S3 MED',
    'S3MED',
    'FARMAVISION',
    'FARMA VISION',
    'HALEXISTAR',
    'HIPOLABOR',
    'CRISTALIA',
    'BLAU',
    'EMS',
    'TEUTO',
    'GEOLAB',
    'SANOFI',
    'MEDLEY',
    'NEO QUIMICA',
    'ABBOTT',
    'ABL',
    'BIOLINE',
    'VITALGOLD',
    'VITAL',
    'WILTEX',
    'SOLIDOR',
    'VIC PHARMA',
    'AMED',
    'EUROFARMA',
    'UNIÃO QUÍMICA',
    'UNIAO QUIMICA',
    'FARMACE',
    'SAMTEC',
    'POLAX',
    'BIOMEDIC',
    'MISSNER',
    'MARKMED'
  ];

  for (let i = 0; i < marcas.length; i++) {
    if (texto.includes(marcas[i])) return marcas[i];
  }

  return '';
}

function f2hExtrairEan_(texto) {
  const m = String(texto || '').match(/\b\d{13}\b/);
  return m ? m[0] : '';
}

function f2hExtrairNcm_(texto) {
  const str = String(texto || '');
  const mPonto = str.match(/\b\d{4}\.\d{2}\.\d{2}\b/);
  if (mPonto) return mPonto[0];

  const m = str.match(/\b\d{8}\b/);
  return m ? m[0] : '';
}

function f2hExtrairAnvisa_(texto) {
  const m = String(texto || '').match(/\b8\d{10,12}\b|\b1\d{10,12}\b/);
  return m ? m[0] : '';
}

function f2hExtrairPercentualPorPalavra_(texto, palavra) {
  const normal = f2hNormalizarTexto_(texto);
  const p = f2hNormalizarTexto_(palavra);

  if (!normal.includes(p)) return '';

  const regex = /(\d{1,2}(?:,\d{1,2})?)\s*%/;
  const m = String(texto || '').match(regex);

  return m ? m[1] : '';
}

function f2hCalcularConfiancaLinha_(item) {
  let score = 0;

  if (item.descricao) score += 35;
  if (item.preco !== null && item.preco > 0) score += 35;
  if (item.unidade) score += 10;
  if (item.fabricante) score += 10;
  if (!item.veioDeOcr) score += 10;

  if (score >= 80) return 'ALTA';
  if (score >= 55) return 'MEDIA';

  return 'BAIXA';
}

/************************************************************
 * 12. DUPLICIDADE BD_PRECOS
 ************************************************************/

function f2hCarregarChavesPrecosExistentes_(bdPrecos, idx) {
  const set = new Set();

  if (bdPrecos.getLastRow() < 2) return set;

  const values = bdPrecos.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const dataTabela = f2hGetCell_(row, idx, 'DATA_TABELA');
    const fornecedor = f2hUpper_(f2hGetCell_(row, idx, 'FORNECEDOR'));
    const descricao = f2hUpper_(f2hGetCell_(row, idx, 'DESCRICAO_FORNECEDOR'));
    let preco = f2hToNumber_(f2hGetCell_(row, idx, 'PRECO_CUSTO'));

    if (preco === null || preco < 0) preco = 0;

    if (!fornecedor || !descricao) continue;

    set.add(f2hMontarChavePreco_(dataTabela, fornecedor, descricao, preco));
  }

  return set;
}

function f2hMontarChavePreco_(data, fornecedor, descricao, preco) {
  const dataKey = data instanceof Date
    ? Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : f2hLimpar_(data);

  return [
    dataKey,
    f2hNormalizarTexto_(fornecedor),
    f2hNormalizarTexto_(descricao),
    String(f2hToNumber_(preco) || 0)
  ].join('|');
}

/************************************************************
 * 13. LIMPEZAS MANUAIS DE SEGURANÇA
 ************************************************************/

function limparRegistrosInvalidosBdPrecosFase2H() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(F2H_ABA_BD_PRECOS);

  if (!ws) throw new Error('Aba BD_PRECOS não encontrada.');
  if (ws.getLastRow() < 2) return 'BD_PRECOS sem dados para limpar.';

  const values = ws.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = f2hCriarIndice_(headers);

  const linhasParaExcluir = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const descricao = f2hGetCell_(row, idx, 'DESCRICAO_FORNECEDOR');
    const skuOperacional = f2hGetCell_(row, idx, 'SKU_OPERACIONAL');

    const textoBase = descricao || skuOperacional;

    if (!f2hDescricaoProdutoValida_(textoBase)) {
      linhasParaExcluir.push(i + 1);
    }
  }

  for (let i = linhasParaExcluir.length - 1; i >= 0; i--) {
    ws.deleteRow(linhasParaExcluir[i]);
  }

  f2hRegistrarLog_(
    'HOTFIX FASE 2H.4 | BD_PRECOS | Registros inválidos removidos: ' +
    linhasParaExcluir.length
  );

  return 'Registros inválidos removidos de BD_PRECOS: ' + linhasParaExcluir.length;
}

function limparRegistrosInvalidosRevisaoPrecosFase2H() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(F2H_ABA_REVISAO);

  if (!ws) throw new Error('Aba REVISAO_IMPORTACAO_PRECOS não encontrada.');
  if (ws.getLastRow() < 2) return 'REVISAO_IMPORTACAO_PRECOS sem dados para limpar.';

  const values = ws.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = f2hCriarIndice_(headers);

  const linhasParaExcluir = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const descricao = f2hGetCell_(row, idx, 'DESCRICAO_ORIGINAL');
    const skuOperacional = f2hGetCell_(row, idx, 'SKU_OPERACIONAL');

    const textoBase = descricao || skuOperacional;

    if (!f2hDescricaoProdutoValida_(textoBase)) {
      linhasParaExcluir.push(i + 1);
    }
  }

  for (let i = linhasParaExcluir.length - 1; i >= 0; i--) {
    ws.deleteRow(linhasParaExcluir[i]);
  }

  f2hRegistrarLog_(
    'HOTFIX FASE 2H.4 | REVISAO_IMPORTACAO_PRECOS | Registros inválidos removidos: ' +
    linhasParaExcluir.length
  );

  return 'Registros inválidos removidos da revisão: ' + linhasParaExcluir.length;
}

/************************************************************
 * 14. DETECÇÃO DE TIPO
 ************************************************************/

function f2hEhPlanilha_(nome, mime) {
  return (
    nome.endsWith('.XLSX') ||
    nome.endsWith('.XLS') ||
    nome.endsWith('.CSV') ||
    mime.includes('SPREADSHEET') ||
    mime.includes('EXCEL') ||
    mime.includes('CSV')
  );
}

function f2hEhTexto_(nome, mime) {
  return (
    nome.endsWith('.TXT') ||
    nome.endsWith('.CSV') ||
    mime.includes('TEXT/PLAIN') ||
    mime.includes('TEXT/CSV')
  );
}

function f2hDetectarTipoArquivo_(nomeArquivo, mimeType) {
  const nome = f2hUpper_(nomeArquivo);
  const mime = f2hUpper_(mimeType);

  if (nome.endsWith('.XLSB')) return 'XLSB';
  if (nome.endsWith('.XLSX')) return 'XLSX';
  if (nome.endsWith('.XLS')) return 'XLS';
  if (nome.endsWith('.CSV')) return 'CSV';
  if (nome.endsWith('.TXT')) return 'TXT';
  if (nome.endsWith('.PDF')) return 'PDF';
  if (nome.endsWith('.DOCX')) return 'DOCX';
  if (nome.endsWith('.DOC')) return 'DOC';
  if (nome.endsWith('.JPEG') || nome.endsWith('.JPG')) return 'JPEG';
  if (nome.endsWith('.PNG')) return 'PNG';

  if (mime.includes('PDF')) return 'PDF';
  if (mime.includes('IMAGE')) return 'IMAGEM';
  if (mime.includes('WORD')) return 'DOCX';
  if (mime.includes('SPREADSHEET') || mime.includes('EXCEL')) return 'XLSX';

  return 'DESCONHECIDO';
}

function f2hDetectarDelimitadorCsv_(texto) {
  const primeiraLinha = String(texto || '').split(/\r?\n/)[0] || '';

  const pontoVirgula = (primeiraLinha.match(/;/g) || []).length;
  const virgula = (primeiraLinha.match(/,/g) || []).length;
  const tab = (primeiraLinha.match(/\t/g) || []).length;

  if (tab > pontoVirgula && tab > virgula) return '\t';
  if (pontoVirgula >= virgula) return ';';

  return ',';
}

/************************************************************
 * 15. HELPERS GERAIS
 ************************************************************/

function f2hObterOuCriarPastaDrive_() {
  const pastas = DriveApp.getFoldersByName(F2H_PASTA_DRIVE);

  if (pastas.hasNext()) return pastas.next();

  return DriveApp.createFolder(F2H_PASTA_DRIVE);
}

function f2hObterOuCriarAba_(ss, nome) {
  let ws = ss.getSheetByName(nome);

  if (!ws) {
    ws = ss.insertSheet(nome);
  }

  return ws;
}

function f2hPrepararTabela_(ws, headers, corFundo, corFonte) {
  if (ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const atual = ws.getRange(1, 1, 1, Math.max(ws.getLastColumn(), headers.length)).getValues()[0];

    let diferente = false;

    for (let i = 0; i < headers.length; i++) {
      if (String(atual[i] || '').trim() !== headers[i]) {
        diferente = true;
        break;
      }
    }

    if (diferente && ws.getLastRow() <= 1) {
      ws.clear();
      ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  const range = ws.getRange(1, 1, 1, headers.length);

  range
    .setBackground(corFundo)
    .setFontColor(corFonte)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  ws.setRowHeight(1, 36);

  if (ws.getFilter()) {
    ws.getFilter().remove();
  }

  ws.getRange(1, 1, Math.max(ws.getLastRow(), 1), headers.length).createFilter();

  ws.getRange(1, 1, Math.max(ws.getMaxRows(), 1), headers.length)
    .setFontFamily('Arial')
    .setFontSize(10);
}

function f2hAplicarLarguras_(ws, mapa) {
  Object.keys(mapa).forEach(col => {
    ws.setColumnWidth(Number(col), mapa[col]);
  });
}

function f2hCriarIndice_(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    idx[String(h).trim()] = i;
  });

  return idx;
}

function f2hGetCell_(row, idx, nomeColuna) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) return '';

  return row[pos];
}

function f2hSetByHeader_(rowArray, idx, nomeColuna, valor) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) return;

  rowArray[pos] = valor;
}

function f2hGerarId_(prefixo) {
  const data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const random = Math.floor(Math.random() * 9000) + 1000;

  return prefixo + '-' + data + '-' + random;
}

function f2hLimpar_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .trim()
    .replace(/\s+/g, ' ');
}

function f2hUpper_(valor) {
  return f2hLimpar_(valor).toUpperCase();
}

function f2hSomenteTexto_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor).trim();
}

function f2hNormalizarHeader_(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function f2hNormalizarTexto_(valor) {
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

function f2hToNumber_(valor) {
  if (valor === null || valor === undefined || valor === '') return null;

  if (typeof valor === 'number') {
    if (isNaN(valor)) return null;
    return valor;
  }

  let txt = String(valor).trim();

  if (!txt) return null;

  if (txt === '-' || txt === 'R$ -' || txt === 'R$-' || txt === '--') return null;

  txt = txt
    .replace(/R\$/gi, '')
    .replace(/\$/g, '')
    .replace(/%/g, '')
    .replace(/\s+/g, '');

  if (txt === '-' || txt === '') return null;

  if (txt.includes(',')) {
    txt = txt.replace(/\./g, '').replace(',', '.');
  }

  txt = txt.replace(/[^0-9.\-]/g, '');

  if (!txt || txt === '-') return null;

  const n = Number(txt);

  return isNaN(n) ? null : n;
}

function f2hNormalizarPercentual_(valor) {
  const n = f2hToNumber_(valor);

  if (n === null || n < 0) return 0;

  if (n <= 1) return n;

  return n / 100;
}

function f2hArredondar_(n, casas) {
  if (n === null || n === undefined || isNaN(n)) return '';

  const f = Math.pow(10, casas || 2);

  return Math.round(n * f) / f;
}

function f2hRegistrarLog_(mensagem) {
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
