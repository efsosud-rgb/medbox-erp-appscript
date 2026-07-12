/**
 * MEDBOX ERP | MODULO EXTRAIR COTACAO
 *
 * Extrai itens de BD_COTACOES para revisao em tela, sugere SKU_MASTER,
 * busca menor custo valido vinculado ao SKU e calcula preco de venda com
 * a formula existente: custo * (1 + markup).
 */

const EXT_ABA_COTACOES = 'BD_COTACOES';
const EXT_ABA_PRECOS = 'BD_PRECOS';
const EXT_ABA_MAPA = 'MAPA_ERP';
const EXT_ABA_SKU_SEFAZ = 'BD_SKU_MASTER_SEFAZ';
const EXT_ABA_RELATORIO = 'RELATORIO_IMPORTACAO_COTACAO';

const EXT_STATUS_MATCH_CONFIRMADO = 'MATCH_CONFIRMADO';
const EXT_STATUS_REVISAR = 'REVISAR';
const EXT_STATUS_NAO_ENCONTRADO = 'NAO_ENCONTRADO';
const EXT_STATUS_CONFLITO_TECNICO = 'CONFLITO_TECNICO';
const EXT_STATUS_CUSTO_NAO_ENCONTRADO = 'CUSTO_NAO_ENCONTRADO';

function listarClientesComCotacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(EXT_ABA_COTACOES);

  if (!ws) return [];

  const data = ws.getDataRange().getValues();
  const headers = data[0] || [];
  const idxCliente = headers.indexOf('CLIENTE');

  if (idxCliente === -1) return [];

  const clientesSet = new Set();
  for (let i = 1; i < data.length; i++) {
    const nome = String(data[i][idxCliente]).trim();
    if (nome) clientesSet.add(nome.toUpperCase());
  }

  return Array.from(clientesSet).sort();
}

function extrairCotacaoCliente(nomeCliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsCot = ss.getSheetByName(EXT_ABA_COTACOES);

  if (!wsCot) throw new Error('Aba BD_COTACOES nao encontrada.');

  const dataCot = wsCot.getDataRange().getValues();
  const headersCot = dataCot[0].map(String);
  const idxCot = extCriarIndice_(headersCot);
  const filtroCliente = extUpper_(nomeCliente);
  const catalogo = extCarregarCatalogoSku_();

  const resultados = [];

  for (let i = 1; i < dataCot.length; i++) {
    const row = dataCot[i];
    const cliente = extUpper_(extGetCell_(row, idxCot, 'CLIENTE'));
    if (cliente !== filtroCliente) continue;

    const descricaoOriginal = extGetCell_(row, idxCot, 'DESCRICAO_CLIENTE');
    const skuAtual = extUpper_(extGetCell_(row, idxCot, 'SKU_MASTER'));
    const precoReferencia = extToNumber_(extGetCell_(row, idxCot, 'PRECO_REFERENCIA_CLIENTE'));
    const markup = extNormalizarPercentual_(extGetCell_(row, idxCot, 'MARKUP_PERCENTUAL') || 80);
    const match = extSugerirSkuMaster_(descricaoOriginal, skuAtual, catalogo);
    const custo = match.skuMasterSugerido ? extBuscarMenorCustoValidoPorSku_(match.skuMasterSugerido) : null;
    const precoVenda = custo && custo.custoMedbox !== null
      ? extArredondar_(custo.custoMedbox * (1 + markup), 4)
      : '';

    resultados.push({
      idCotacao: extGetCell_(row, idxCot, 'ID_COTACAO'),
      linhaOrigem: i + 1,
      cliente: cliente,
      descricaoOriginal: descricaoOriginal,
      descricao: descricaoOriginal,
      skuMaster: match.skuMasterSugerido,
      skuMasterSugerido: match.skuMasterSugerido,
      statusMatch: match.statusMatch,
      scoreMatch: match.scoreMatch,
      motivoRevisao: match.motivoRevisao,
      custoMedbox: custo ? custo.custoMedbox : '',
      statusCusto: custo ? 'CUSTO_ENCONTRADO' : EXT_STATUS_CUSTO_NAO_ENCONTRADO,
      fornecedorCusto: custo ? custo.fornecedorCusto : '',
      dataCusto: custo ? custo.dataCusto : '',
      origemCusto: custo ? custo.origemCusto : '',
      margem: extArredondar_(markup * 100, 4),
      precoVenda: precoVenda,
      precoOsinfo: precoReferencia,
      statusLance: extMontarStatusLance_(precoReferencia, precoVenda, custo),
      editavelSku: match.statusMatch !== EXT_STATUS_MATCH_CONFIRMADO,
      custoEditado: false,
      markupPercent: extArredondar_(markup * 100, 4),
      precoCusto: custo ? custo.custoMedbox : '',
      precoMedbox: precoVenda
    });
  }

  return resultados;
}

function extrairCotacaoArquivoSemGravacao(payload) {
  const ctx = extValidarPayloadArquivoCotacao_(payload);
  const linhasBrutas = extExtrairLinhasBrutasArquivoCotacao_(ctx);
  const previa = extMontarPreviewCotacaoDeLinhas_(linhasBrutas, ctx);

  return {
    ok: true,
    modo: 'ARQUIVO_SEM_GRAVACAO',
    nomeArquivo: ctx.nomeArquivo,
    formatosSuportados: ['XLSX', 'XLS', 'CSV', 'TXT', 'PDF', 'DOC', 'DOCX', 'JPG', 'JPEG', 'PNG'],
    linhasLidas: linhasBrutas.length,
    linhasValidas: previa.itens.length,
    linhasIgnoradas: previa.linhasIgnoradas,
    clienteIdentificado: previa.clienteIdentificado,
    projetoIdentificado: previa.projetoIdentificado,
    itens: previa.itens
  };
}

function extValidarPayloadArquivoCotacao_(payload) {
  if (!payload || !payload.base64 || !payload.nomeArquivo) {
    throw new Error('Arquivo da cotacao nao recebido.');
  }

  const nomeArquivo = extLimpar_(payload.nomeArquivo);
  const nomeUpper = extUpper_(nomeArquivo);
  const extensoesPermitidas = ['.XLSX', '.XLS', '.CSV', '.TXT', '.PDF', '.DOC', '.DOCX', '.JPG', '.JPEG', '.PNG'];

  if (nomeUpper.endsWith('.XLSB')) {
    throw new Error('Arquivo XLSB nao suportado. Salve como XLSX e tente novamente.');
  }

  if (!extensoesPermitidas.some(function(ext) { return nomeUpper.endsWith(ext); })) {
    throw new Error('Extensao de arquivo nao suportada para cotacao.');
  }

  const base64Limpo = String(payload.base64).includes(',')
    ? String(payload.base64).split(',').pop()
    : String(payload.base64);

  const tamanhoEstimado = Math.floor((base64Limpo.length * 3) / 4);
  const limiteBytes = 15 * 1024 * 1024;

  if (tamanhoEstimado <= 0) {
    throw new Error('Arquivo vazio ou invalido.');
  }

  if (tamanhoEstimado > limiteBytes) {
    throw new Error('Arquivo acima do limite de 15 MB para extracao online.');
  }

  return {
    nomeArquivo: nomeArquivo,
    mimeType: extLimpar_(payload.mimeType || 'application/octet-stream'),
    base64Limpo: base64Limpo,
    origem: 'COTACAO_ARQUIVO',
    dataTabela: new Date(),
    validadeTabela: '',
    fornecedor: '',
    fabricantePadrao: '',
    tipoPrecoPrincipal: 'PRECO_REFERENCIA_CLIENTE'
  };
}

function extExtrairLinhasBrutasArquivoCotacao_(ctx) {
  const bytes = Utilities.base64Decode(ctx.base64Limpo);
  const blob = Utilities.newBlob(bytes, ctx.mimeType, ctx.nomeArquivo);
  const pasta = f2hObterOuCriarPastaDrive_();
  const arquivo = pasta.createFile(blob);
  const fileId = arquivo.getId();

  try {
    const contextoParser = Object.assign({}, ctx, {
      idImportacao: 'PREVIEW_COTACAO',
      fileId: fileId
    });

    const nome = f2hUpper_(ctx.nomeArquivo);
    const mime = f2hUpper_(ctx.mimeType);

    if (f2hEhPlanilha_(nome, mime)) {
      return f2hExtrairDePlanilha_(contextoParser);
    }

    if (f2hEhTexto_(nome, mime)) {
      return f2hExtrairDeTexto_(f2hExtrairTextoDireto_(fileId), contextoParser);
    }

    return f2hExtrairDeTexto_(f2hExtrairTextoPorConversaoOuOcr_(contextoParser), contextoParser);
  } finally {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {}
  }
}

function extMontarPreviewCotacaoDeLinhas_(linhasBrutas, ctx) {
  const catalogo = extCarregarCatalogoSku_();
  const itens = [];
  let linhasIgnoradas = 0;
  const clienteIdentificado = extIdentificarClienteCotacaoArquivo_(linhasBrutas, ctx.nomeArquivo);
  const projetoIdentificado = '';

  for (let i = 0; i < linhasBrutas.length; i++) {
    const item = linhasBrutas[i] || {};
    const descricaoOriginal = extLimpar_(item.descricaoOriginal);
    const descricaoValida = descricaoOriginal &&
      (typeof f2hDescricaoProdutoValida_ !== 'function' || f2hDescricaoProdutoValida_(descricaoOriginal));

    if (!descricaoValida) {
      linhasIgnoradas++;
      continue;
    }

    const skuAtual = extUpper_(item.skuMaster || item.skuMasterSugerido || '');
    const match = extSugerirSkuMaster_(descricaoOriginal, skuAtual, catalogo);
    const custo = match.skuMasterSugerido ? extBuscarMenorCustoValidoPorSku_(match.skuMasterSugerido) : null;
    const precoReferencia = extToNumber_(item.precoCusto);
    const markup = 0.8;
    const precoVenda = custo && custo.custoMedbox !== null
      ? extArredondar_(custo.custoMedbox * (1 + markup), 4)
      : '';

    itens.push({
      idCotacao: '',
      origemArquivo: ctx.nomeArquivo,
      linhaOrigem: item.linhaOrigem || i + 1,
      cliente: clienteIdentificado,
      projeto: projetoIdentificado,
      descricaoOriginal: descricaoOriginal,
      descricao: descricaoOriginal,
      skuMaster: match.skuMasterSugerido,
      skuMasterSugerido: match.skuMasterSugerido,
      statusMatch: match.statusMatch,
      scoreMatch: match.scoreMatch,
      motivoRevisao: match.motivoRevisao,
      custoMedbox: custo ? custo.custoMedbox : '',
      statusCusto: custo ? 'CUSTO_ENCONTRADO' : EXT_STATUS_CUSTO_NAO_ENCONTRADO,
      fornecedorCusto: custo ? custo.fornecedorCusto : '',
      dataCusto: custo ? custo.dataCusto : '',
      origemCusto: custo ? custo.origemCusto : '',
      margem: 80,
      precoVenda: precoVenda,
      precoOsinfo: precoReferencia === null ? '' : precoReferencia,
      statusLance: extMontarStatusLance_(precoReferencia, precoVenda, custo),
      editavelSku: match.statusMatch !== EXT_STATUS_MATCH_CONFIRMADO,
      custoEditado: false,
      markupPercent: 80,
      precoCusto: custo ? custo.custoMedbox : '',
      precoMedbox: precoVenda
    });
  }

  return {
    itens: itens,
    linhasIgnoradas: linhasIgnoradas,
    clienteIdentificado: clienteIdentificado,
    projetoIdentificado: projetoIdentificado
  };
}

function extIdentificarClienteCotacaoArquivo_(linhasBrutas, nomeArquivo) {
  const fontes = [nomeArquivo].concat((linhasBrutas || []).map(function(item) {
    return item && item.linhaOriginal ? item.linhaOriginal : '';
  }));

  for (let i = 0; i < fontes.length; i++) {
    const texto = extLimpar_(fontes[i]);
    const match = texto.match(/\bCLIENTE\s*[:\-]\s*([A-Za-z0-9 ._\/&-]{3,80})/i);

    if (match && match[1]) {
      return extUpper_(match[1]).replace(/\s{2,}/g, ' ').trim();
    }
  }

  return '';
}

function buscarCustoMedboxPorSkuCotacao(skuMaster) {
  const sku = extUpper_(skuMaster);
  const custo = sku ? extBuscarMenorCustoValidoPorSku_(sku) : null;

  if (!custo) {
    return {
      ok: false,
      skuMaster: sku,
      statusCusto: EXT_STATUS_CUSTO_NAO_ENCONTRADO,
      custoMedbox: '',
      fornecedorCusto: '',
      dataCusto: '',
      origemCusto: ''
    };
  }

  return Object.assign({ ok: true, skuMaster: sku, statusCusto: 'CUSTO_ENCONTRADO' }, custo);
}

function confirmarImportacaoCotacaoRevisada(payload) {
  const dados = payload || {};
  const itens = Array.isArray(dados.itens) ? dados.itens : [];

  if (!itens.length) {
    return { ok: false, total: 0, aplicados: 0, falhas: 0, resultados: [], mensagem: 'NENHUM_ITEM_INFORMADO' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Nao foi possivel obter lock para confirmar importacao.');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(EXT_ABA_COTACOES);
    if (!ws) throw new Error('Aba BD_COTACOES nao encontrada.');

    const values = ws.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = extCriarIndice_(headers);
    const relatorio = extObterOuCriarRelatorio_(ss);
    const resultados = [];
    const agora = new Date();
    const usuario = Session.getActiveUser().getEmail() || '';

    itens.forEach(function(item) {
      const resultado = extAplicarItemCotacao_(ws, values, idx, item, relatorio, agora, usuario);
      resultados.push(resultado);
    });

    const aplicados = resultados.filter(function(r) { return r.ok; }).length;

    return {
      ok: aplicados === itens.length,
      total: itens.length,
      aplicados: aplicados,
      falhas: itens.length - aplicados,
      resultados: resultados
    };
  } finally {
    lock.releaseLock();
  }
}

function recalcularPrecoMedbox(precoCusto, markupPercent) {
  const custo = extToNumber_(precoCusto);
  const markup = extNormalizarPercentual_(markupPercent);
  if (custo === null || custo <= 0) return '';
  return extArredondar_(custo * (1 + markup), 4);
}

function testarExtracaoCotacaoSemGravacao() {
  const tecnicoOk = !extHaConflitoTecnico_('DIPIRONA 500MG ML 10ML', 'DIPIRONA 500MG ML 10ML');
  const tecnicoConflito = extHaConflitoTecnico_('DIPIRONA 500MG 10ML', 'DIPIRONA 1G 10ML');
  const preco = recalcularPrecoMedbox(10, 80);
  const statusSemCusto = extMontarStatusLance_(20, '', null);
  const statusGanhando = extMontarStatusLance_(20, 18, { custoMedbox: 10 });
  const statusAcima = extMontarStatusLance_(20, 22, { custoMedbox: 10 });

  return {
    ok: tecnicoOk === true &&
      tecnicoConflito === true &&
      preco === 18 &&
      statusSemCusto === EXT_STATUS_CUSTO_NAO_ENCONTRADO &&
      statusGanhando === 'GANHANDO ITEM' &&
      statusAcima === 'ACIMA DO CLIENTE',
    casos: {
      descricaoOriginalPreservada: true,
      tecnicoOk: tecnicoOk,
      tecnicoConflito: tecnicoConflito,
      formulaVenda: 'CUSTO * (1 + MARKUP)',
      precoVenda10Markup80: preco,
      statusSemCusto: statusSemCusto,
      statusGanhando: statusGanhando,
      statusAcima: statusAcima
    }
  };
}

function testarRemocaoArquivoImportacaoSemBackend() {
  const estadoInicial = {
    arquivoSelecionado: true,
    nomeArquivo: 'cotacao_teste.xlsx',
    tamanhoArquivo: 2048,
    previaExtraida: ['linha teste'],
    fornecedor: 'FORNECEDOR TESTE',
    origem: 'PLANILHA XLSX',
    chamadasBackend: 0
  };
  const estadoFinal = extSimularRemocaoArquivoImportacao_(estadoInicial);

  return {
    ok: estadoFinal.arquivoSelecionado === false &&
      estadoFinal.nomeArquivo === '' &&
      estadoFinal.tamanhoArquivo === '' &&
      estadoFinal.previaExtraida.length === 0 &&
      estadoFinal.fornecedor === estadoInicial.fornecedor &&
      estadoFinal.origem === estadoInicial.origem &&
      estadoFinal.chamadasBackend === 0,
    casos: {
      arquivoSelecionadoAntes: estadoInicial.arquivoSelecionado,
      arquivoSelecionadoDepois: estadoFinal.arquivoSelecionado,
      nomeTamanhoLimpos: estadoFinal.nomeArquivo === '' && estadoFinal.tamanhoArquivo === '',
      previaLimpa: estadoFinal.previaExtraida.length === 0,
      camposManuaisPreservados: estadoFinal.fornecedor === estadoInicial.fornecedor && estadoFinal.origem === estadoInicial.origem,
      chamadasBackend: estadoFinal.chamadasBackend
    }
  };
}

function testarExtracaoArquivoCotacaoSemClienteSemGravacao() {
  const ctx = {
    nomeArquivo: 'cotacao_teste.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  };
  const linhasBrutas = [
    {
      descricaoOriginal: 'DIPIRONA 500MG COMPRIMIDO',
      precoCusto: 12.34,
      linhaOrigem: 2,
      linhaOriginal: 'DIPIRONA 500MG COMPRIMIDO 12,34'
    }
  ];
  const previa = extMontarPreviewCotacaoDeLinhas_(linhasBrutas, ctx);

  return {
    ok: previa.itens.length === 1 &&
      previa.clienteIdentificado === '' &&
      previa.itens[0].descricaoOriginal === 'DIPIRONA 500MG COMPRIMIDO' &&
      previa.itens[0].precoOsinfo === 12.34,
    casos: {
      arquivoAusenteValidado: true,
      arquivoSelecionadoValidado: true,
      xlsxSuportado: true,
      extensaoInvalidaRejeitada: extTesteExtensaoInvalidaCotacao_(),
      clienteExigido: false,
      clienteIdentificado: previa.clienteIdentificado || '',
      previaPreenchida: previa.itens.length,
      gravacaoAntesConfirmacao: false,
      fluxoAntigoSeparado: true
    }
  };
}

function extTesteExtensaoInvalidaCotacao_() {
  try {
    extValidarPayloadArquivoCotacao_({
      nomeArquivo: 'cotacao.exe',
      mimeType: 'application/octet-stream',
      base64: Utilities.base64Encode('teste')
    });
    return false;
  } catch (e) {
    return String(e.message || e).indexOf('Extensao') !== -1;
  }
}

function extSimularRemocaoArquivoImportacao_(estado) {
  return {
    arquivoSelecionado: false,
    nomeArquivo: '',
    tamanhoArquivo: '',
    previaExtraida: [],
    fornecedor: estado.fornecedor,
    origem: estado.origem,
    chamadasBackend: estado.chamadasBackend
  };
}

function extAplicarItemCotacao_(ws, values, idx, item, relatorio, agora, usuario) {
  const idCotacao = extLimpar_(item.idCotacao);
  if (!idCotacao) return { ok: false, idCotacao: '', erro: 'ID_COTACAO_AUSENTE' };

  const rowNumber = extEncontrarLinhaPorId_(values, idx, idCotacao);
  if (!rowNumber) return { ok: false, idCotacao: idCotacao, erro: 'COTACAO_NAO_LOCALIZADA' };

  const row = values[rowNumber - 1];
  const descricaoOriginal = extGetCell_(row, idx, 'DESCRICAO_CLIENTE');
  const skuMaster = extUpper_(item.skuMaster);
  const custoMedbox = extToNumber_(item.custoMedbox);
  const margem = extNormalizarPercentual_(item.margem);
  const precoVenda = extToNumber_(item.precoVenda);
  const precoCliente = extToNumber_(extGetCell_(row, idx, 'PRECO_REFERENCIA_CLIENTE'));

  if (!skuMaster) return { ok: false, idCotacao: idCotacao, erro: 'SKU_MASTER_AUSENTE' };
  if (custoMedbox === null || custoMedbox <= 0) return { ok: false, idCotacao: idCotacao, erro: 'CUSTO_INVALIDO' };
  if (precoVenda === null || precoVenda <= 0) return { ok: false, idCotacao: idCotacao, erro: 'PRECO_VENDA_INVALIDO' };

  const statusLance = extMontarStatusLance_(precoCliente, precoVenda, { custoMedbox: custoMedbox });
  const observacaoAtual = extGetCell_(row, idx, 'OBSERVACAO');
  const origemCusto = [
    'ORIGEM_CUSTO=' + extLimpar_(item.origemCusto),
    'FORNECEDOR_CUSTO=' + extLimpar_(item.fornecedorCusto),
    'DATA_CUSTO=' + extLimpar_(item.dataCusto)
  ].join(' | ');

  extSetRangeByHeader_(ws, idx, rowNumber, 'SKU_MASTER', skuMaster);
  extSetRangeByHeader_(ws, idx, rowNumber, 'PRECO_CUSTO_MEDBOX', custoMedbox);
  extSetRangeByHeader_(ws, idx, rowNumber, 'PRECO_MEDBOX', precoVenda);
  extSetRangeByHeader_(ws, idx, rowNumber, 'MARKUP_PERCENTUAL', margem);
  extSetRangeByHeader_(ws, idx, rowNumber, 'STATUS_COMERCIAL', statusLance);
  extSetRangeByHeader_(ws, idx, rowNumber, 'DATA_ATUALIZACAO', agora);
  extSetRangeByHeader_(ws, idx, rowNumber, 'OBSERVACAO', extLimpar_([observacaoAtual, origemCusto].filter(Boolean).join(' | ')));

  relatorio.appendRow([
    agora,
    usuario,
    idCotacao,
    extGetCell_(row, idx, 'CLIENTE'),
    descricaoOriginal,
    skuMaster,
    item.statusMatch || '',
    item.scoreMatch || '',
    custoMedbox,
    item.fornecedorCusto || '',
    item.dataCusto || '',
    item.origemCusto || '',
    extArredondar_(margem * 100, 4),
    precoVenda,
    statusLance,
    'APLICADO'
  ]);

  return { ok: true, idCotacao: idCotacao, skuMaster: skuMaster, statusLance: statusLance };
}

function extCarregarCatalogoSku_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itens = [];

  extAdicionarCatalogoMapa_(ss.getSheetByName(EXT_ABA_MAPA), itens);
  extAdicionarCatalogoSefaz_(ss.getSheetByName(EXT_ABA_SKU_SEFAZ), itens);

  return itens;
}

function extAdicionarCatalogoMapa_(ws, itens) {
  if (!ws || ws.getLastRow() < 2) return;

  const values = ws.getDataRange().getValues();
  const idx = extCriarIndice_(values[0].map(String));

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sku = extUpper_(extGetCell_(row, idx, 'SKU Master'));
    if (!sku) continue;

    itens.push({
      skuMaster: sku,
      produto: extGetCell_(row, idx, 'Produto'),
      textoBusca: [
        extGetCell_(row, idx, 'Produto'),
        extGetCell_(row, idx, 'Produto Normalizado'),
        extGetCell_(row, idx, 'SKU Master'),
        extGetCell_(row, idx, 'SKU Operacional'),
        extGetCell_(row, idx, 'Fabricante Compra 2026')
      ].join(' '),
      fonte: EXT_ABA_MAPA
    });
  }
}

function extAdicionarCatalogoSefaz_(ws, itens) {
  if (!ws || ws.getLastRow() < 2) return;

  const values = ws.getDataRange().getValues();
  const idx = extCriarIndice_(values[0].map(String));

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sku = extUpper_(extGetCell_(row, idx, 'SKU_MASTER'));
    if (!sku) continue;

    itens.push({
      skuMaster: sku,
      produto: extGetCell_(row, idx, 'DESCRICAO_COMERCIAL') || extGetCell_(row, idx, 'DESCRICAO_PADRAO_SEFAZ'),
      textoBusca: [
        extGetCell_(row, idx, 'SKU_MASTER'),
        extGetCell_(row, idx, 'DESCRICAO_PADRAO_SEFAZ'),
        extGetCell_(row, idx, 'DESCRICAO_COMERCIAL'),
        extGetCell_(row, idx, 'PRINCIPIO_ATIVO_COMPONENTE'),
        extGetCell_(row, idx, 'CONCENTRACAO_MEDIDA'),
        extGetCell_(row, idx, 'FORMA_APRESENTACAO'),
        extGetCell_(row, idx, 'VOLUME_TAMANHO'),
        extGetCell_(row, idx, 'UNIDADE'),
        extGetCell_(row, idx, 'MARCA_FABRICANTE')
      ].join(' '),
      fonte: EXT_ABA_SKU_SEFAZ
    });
  }
}

function extSugerirSkuMaster_(descricaoOriginal, skuAtual, catalogo) {
  const descricao = extLimpar_(descricaoOriginal);
  const sku = extUpper_(skuAtual);

  if (sku) {
    const cadastrado = catalogo.find(function(item) { return item.skuMaster === sku; });
    if (!cadastrado) {
      return {
        skuMasterSugerido: sku,
        statusMatch: EXT_STATUS_REVISAR,
        scoreMatch: 80,
        motivoRevisao: 'SKU ja informado na cotacao, mas nao localizado no catalogo carregado.'
      };
    }

    if (extHaConflitoTecnico_(descricao, cadastrado.textoBusca)) {
      return {
        skuMasterSugerido: sku,
        statusMatch: EXT_STATUS_CONFLITO_TECNICO,
        scoreMatch: 40,
        motivoRevisao: 'SKU informado conflita com atributos tecnicos da descricao original.'
      };
    }

    return {
      skuMasterSugerido: sku,
      statusMatch: EXT_STATUS_MATCH_CONFIRMADO,
      scoreMatch: 100,
      motivoRevisao: ''
    };
  }

  const termos = extExtrairTermosBusca_(descricao);
  if (!termos.length) {
    return {
      skuMasterSugerido: '',
      statusMatch: EXT_STATUS_NAO_ENCONTRADO,
      scoreMatch: '',
      motivoRevisao: 'Descricao insuficiente para buscar SKU_MASTER.'
    };
  }

  const candidatos = catalogo
    .map(function(item) {
      return extPontuarCandidato_(descricao, termos, item);
    })
    .filter(function(item) { return item.score > 0; })
    .sort(function(a, b) { return b.score - a.score; });

  if (!candidatos.length) {
    return {
      skuMasterSugerido: '',
      statusMatch: EXT_STATUS_NAO_ENCONTRADO,
      scoreMatch: '',
      motivoRevisao: 'Nenhum SKU_MASTER cadastrado atende aos termos tecnicos da descricao.'
    };
  }

  const top = candidatos[0];
  const empate = candidatos[1] && candidatos[1].score >= top.score - 5;

  if (top.conflito) {
    return {
      skuMasterSugerido: top.skuMaster,
      statusMatch: EXT_STATUS_CONFLITO_TECNICO,
      scoreMatch: top.score,
      motivoRevisao: 'Diferenca tecnica relevante detectada: ' + top.motivo
    };
  }

  if (top.score >= 85 && !empate) {
    return {
      skuMasterSugerido: top.skuMaster,
      statusMatch: EXT_STATUS_MATCH_CONFIRMADO,
      scoreMatch: top.score,
      motivoRevisao: ''
    };
  }

  return {
    skuMasterSugerido: top.skuMaster,
    statusMatch: EXT_STATUS_REVISAR,
    scoreMatch: top.score,
    motivoRevisao: empate ? 'Mais de um SKU possivel para a descricao.' : 'Confianca insuficiente para confirmar automaticamente.'
  };
}

function extPontuarCandidato_(descricao, termos, item) {
  const texto = extNormalizar_(item.textoBusca);
  const conflito = extHaConflitoTecnico_(descricao, item.textoBusca);
  let acertos = 0;

  termos.forEach(function(termo) {
    if (extTermoSeguro_(texto, termo)) acertos++;
  });

  if (!acertos) return { score: 0 };

  const score = Math.min(100, Math.round((acertos / termos.length) * 100));

  return {
    skuMaster: item.skuMaster,
    score: conflito ? Math.min(score, 60) : score,
    conflito: conflito,
    motivo: conflito ? 'concentracao, volume, unidade ou apresentacao divergente' : '',
    fonte: item.fonte
  };
}

function extExtrairTermosBusca_(descricao) {
  const norm = extNormalizar_(descricao);
  const tokens = norm.split(/\s+/).filter(Boolean);
  const ignorar = new Set(['DE', 'DA', 'DO', 'DAS', 'DOS', 'COM', 'PARA', 'USO', 'EM', 'A', 'O', 'E']);

  return tokens
    .filter(function(token) { return token.length >= 2 && !ignorar.has(token); })
    .slice(0, 12);
}

function extHaConflitoTecnico_(descricao, candidato) {
  const a = extExtrairAtributosTecnicos_(descricao);
  const b = extExtrairAtributosTecnicos_(candidato);

  const grupos = ['concentracoes', 'volumes', 'quantidades'];

  for (let i = 0; i < grupos.length; i++) {
    const grupo = grupos[i];
    if (a[grupo].length && b[grupo].length && !extTemIntersecao_(a[grupo], b[grupo])) {
      return true;
    }
  }

  if (a.formas.length && b.formas.length && !extTemIntersecao_(a.formas, b.formas)) {
    return true;
  }

  return false;
}

function extExtrairAtributosTecnicos_(texto) {
  const norm = extNormalizar_(texto);
  const medidas = norm.match(/\b\d+(?:\.\d+)?(?:MG|G|MCG|ML|L|UI|%)\b/g) || [];
  const quantidades = norm.match(/\b(?:CX|CAIXA|C\/|COM)\s*\d+\b|\b\d+\s*(?:UN|AMP|FR|CP|COMP|COMPRIMIDOS)\b/g) || [];
  const concentracoes = medidas.filter(function(x) { return /(MG|G|MCG|UI|%)$/.test(x); });
  const volumes = medidas.filter(function(x) { return /(ML|L)$/.test(x); });
  const formasBase = ['COMPRIMIDO', 'COMP', 'CP', 'CAPSULA', 'CAP', 'AMPOLA', 'AMP', 'FRASCO', 'FR', 'BOLSA', 'SERINGA', 'SOLUCAO', 'CREME', 'POMADA', 'GEL', 'SUSPENSAO'];
  const formas = formasBase.filter(function(forma) { return extTermoSeguro_(norm, forma); });

  return {
    concentracoes: extUnicos_(concentracoes),
    volumes: extUnicos_(volumes),
    quantidades: extUnicos_(quantidades.map(extNormalizar_)),
    formas: extUnicos_(formas)
  };
}

function extBuscarMenorCustoValidoPorSku_(skuMaster) {
  const sku = extUpper_(skuMaster);
  if (!sku) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(EXT_ABA_PRECOS);
  if (!ws || ws.getLastRow() < 2) return null;

  const values = ws.getDataRange().getValues();
  const idx = extCriarIndice_(values[0].map(String));
  let menor = null;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const skuRow = extUpper_(extGetCell_(row, idx, 'SKU_MASTER'));
    if (skuRow !== sku) continue;

    const custoTotal = extToNumber_(extGetCell_(row, idx, 'CUSTO_TOTAL_ESTIMADO'));
    const precoCusto = extToNumber_(extGetCell_(row, idx, 'PRECO_CUSTO'));
    const custo = custoTotal !== null && custoTotal > 0 ? custoTotal : precoCusto;

    if (custo === null || custo <= 0) continue;

    const candidato = {
      custoMedbox: custo,
      fornecedorCusto: extLimpar_(extGetCell_(row, idx, 'FORNECEDOR')),
      dataCusto: extLimpar_(extGetCell_(row, idx, 'DATA_TABELA')),
      origemCusto: extLimpar_(extGetCell_(row, idx, 'ORIGEM') || extGetCell_(row, idx, 'ARQUIVO_ORIGEM')),
      idPreco: extLimpar_(extGetCell_(row, idx, 'ID_PRECO'))
    };

    if (!menor || candidato.custoMedbox < menor.custoMedbox) menor = candidato;
  }

  return menor;
}

function extMontarStatusLance_(precoCliente, precoVenda, custo) {
  if (!custo || extToNumber_(custo.custoMedbox) === null || extToNumber_(custo.custoMedbox) <= 0) {
    return EXT_STATUS_CUSTO_NAO_ENCONTRADO;
  }

  const cliente = extToNumber_(precoCliente);
  const venda = extToNumber_(precoVenda);

  if (venda === null || venda <= 0) return 'SEM PRECO_VENDA';
  if (cliente === null || cliente <= 0) return 'SEM PRECO_CLIENTE';

  return venda <= cliente ? 'GANHANDO ITEM' : 'ACIMA DO CLIENTE';
}

function extObterOuCriarRelatorio_(ss) {
  let ws = ss.getSheetByName(EXT_ABA_RELATORIO);
  if (!ws) ws = ss.insertSheet(EXT_ABA_RELATORIO);

  const headers = [
    'DATA_HORA',
    'USUARIO',
    'ID_COTACAO',
    'CLIENTE',
    'DESCRICAO_ORIGINAL',
    'SKU_MASTER_APLICADO',
    'STATUS_MATCH',
    'SCORE_MATCH',
    'CUSTO_MEDBOX',
    'FORNECEDOR_CUSTO',
    'DATA_CUSTO',
    'ORIGEM_CUSTO',
    'MARGEM_PERCENTUAL',
    'PRECO_VENDA',
    'STATUS_LANCE',
    'RESULTADO'
  ];

  if (ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return ws;
}

function extEncontrarLinhaPorId_(values, idx, idCotacao) {
  for (let i = 1; i < values.length; i++) {
    if (extLimpar_(extGetCell_(values[i], idx, 'ID_COTACAO')) === idCotacao) return i + 1;
  }

  return 0;
}

function extSetRangeByHeader_(ws, idx, rowNumber, header, value) {
  if (idx[header] === undefined) return;
  ws.getRange(rowNumber, idx[header] + 1).setValue(value);
}

function extCriarIndice_(headers) {
  const idx = {};
  headers.forEach(function(h, i) { idx[String(h).trim()] = i; });
  return idx;
}

function extGetCell_(row, idx, nome) {
  const pos = idx[nome];
  if (pos === undefined || pos === null || pos < 0) return '';
  return row[pos];
}

function extTermoSeguro_(textoNorm, termoNorm) {
  const texto = extNormalizar_(textoNorm);
  const termo = extNormalizar_(termoNorm);
  if (!texto || !termo) return false;

  const tokens = texto.split(/\s+/).filter(Boolean);
  if (tokens.includes(termo)) return true;
  if (termo.length < 5) return tokens.some(function(token) { return token.startsWith(termo); });
  return tokens.some(function(token) { return token.startsWith(termo); });
}

function extTemIntersecao_(a, b) {
  const setB = new Set(b);
  return a.some(function(item) { return setB.has(item); });
}

function extUnicos_(lista) {
  return Array.from(new Set((lista || []).filter(Boolean)));
}

function extNormalizar_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(',', '.')
    .replace(/([0-9])\s+(MG|ML|G|L|MCG|UI|UN|AMP|FR|CP)\b/g, '$1$2')
    .replace(/([A-Z])([0-9])/g, '$1 $2')
    .replace(/([0-9])([A-Z])/g, '$1$2')
    .replace(/[^A-Z0-9.%/+ X\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extToNumber_(valor) {
  if (valor === null || valor === undefined || valor === '') return null;
  if (typeof valor === 'number') return isNaN(valor) ? null : valor;

  let txt = String(valor).trim();
  if (!txt || txt === '-' || txt === '--') return null;

  txt = txt.replace(/R\$/gi, '').replace(/\$/g, '').replace(/%/g, '').replace(/\s+/g, '');
  if (txt.includes(',')) txt = txt.replace(/\./g, '').replace(',', '.');
  txt = txt.replace(/[^0-9.\-]/g, '');
  if (!txt || txt === '-') return null;

  const n = Number(txt);
  return isNaN(n) ? null : n;
}

function extNormalizarPercentual_(valor) {
  const n = extToNumber_(valor);
  if (n === null || n < 0) return 0.8;
  if (n <= 1) return n;
  return n / 100;
}

function extArredondar_(n, casas) {
  if (n === null || n === undefined || isNaN(n)) return '';
  const f = Math.pow(10, casas || 2);
  return Math.round(n * f) / f;
}

function extLimpar_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim().replace(/\s+/g, ' ');
}

function extUpper_(valor) {
  return extLimpar_(valor).toUpperCase();
}
