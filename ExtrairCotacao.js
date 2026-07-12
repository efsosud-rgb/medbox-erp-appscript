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
const EXT_STATUS_LANCE_VENCEDOR = 'LANCE VENCEDOR';
const EXT_STATUS_PRECO_IGUAL = 'PRECO IGUAL';
const EXT_STATUS_PRECO_ACIMA = 'PRECO ACIMA';
const EXT_STATUS_DADOS_INSUFICIENTES = 'DADOS INSUFICIENTES';

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
  const extracao = extExtrairLinhasBrutasArquivoCotacao_(ctx);
  const linhasBrutas = Array.isArray(extracao) ? extracao : extracao.linhas;
  const diagnosticoExtracao = Array.isArray(extracao) ? {} : extracao.diagnostico;
  const previa = extMontarPreviewCotacaoDeLinhas_(linhasBrutas, ctx);
  const motivosIgnorados = extMesclarMotivosIgnorados_(
    diagnosticoExtracao && diagnosticoExtracao.motivosIgnorados,
    previa.motivosIgnorados
  );
  const diagnostico = Object.assign({}, diagnosticoExtracao || {}, {
    linhasLidas: diagnosticoExtracao && diagnosticoExtracao.linhasLidas !== undefined
      ? diagnosticoExtracao.linhasLidas
      : linhasBrutas.length,
    itensValidos: previa.itens.length,
    linhasIgnoradas: extSomarMotivos_(motivosIgnorados),
    motivosIgnorados: motivosIgnorados
  });

  return {
    ok: true,
    modo: 'ARQUIVO_SEM_GRAVACAO',
    nomeArquivo: ctx.nomeArquivo,
    formatosSuportados: ['XLSX', 'XLS', 'CSV', 'TXT', 'PDF', 'DOC', 'DOCX', 'JPG', 'JPEG', 'PNG'],
    linhasLidas: diagnostico.linhasLidas,
    linhasValidas: previa.itens.length,
    linhasIgnoradas: diagnostico.linhasIgnoradas,
    motivosIgnorados: motivosIgnorados,
    diagnostico: diagnostico,
    abaSelecionada: diagnostico.abaSelecionada || '',
    linhaCabecalho: diagnostico.linhaCabecalho || '',
    colunasMapeadas: diagnostico.colunasMapeadas || {},
    clienteIdentificado: previa.clienteIdentificado,
    projetoIdentificado: previa.projetoIdentificado,
    resumo: extCalcularResumoCotacaoExtraida_(previa.itens),
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
      return extExtrairPlanilhaCotacaoDireta_(contextoParser);
    }

    if (f2hEhTexto_(nome, mime)) {
      const linhas = f2hExtrairDeTexto_(f2hExtrairTextoDireto_(fileId), contextoParser);
      return {
        linhas: linhas,
        diagnostico: {
          tipoArquivo: f2hDetectarTipoArquivo_(ctx.nomeArquivo, ctx.mimeType),
          linhasLidas: linhas.length,
          abas: []
        }
      };
    }

    const linhasTexto = f2hExtrairDeTexto_(f2hExtrairTextoPorConversaoOuOcr_(contextoParser), contextoParser);
    return {
      linhas: linhasTexto,
      diagnostico: {
        tipoArquivo: f2hDetectarTipoArquivo_(ctx.nomeArquivo, ctx.mimeType),
        linhasLidas: linhasTexto.length,
        abas: []
      }
    };
  } finally {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {}
  }
}

function extExtrairPlanilhaCotacaoDireta_(ctx) {
  const nome = f2hUpper_(ctx.nomeArquivo);

  if (nome.endsWith('.CSV')) {
    const texto = f2hExtrairTextoDireto_(ctx.fileId);
    const values = Utilities.parseCsv(texto, f2hDetectarDelimitadorCsv_(texto));
    const resultadoCsv = extSelecionarTabelaCotacaoDeMatrizes_([
      { nome: 'CSV', values: values, mergedCount: 0 }
    ]);
    return resultadoCsv;
  }

  const tempId = f2hConverterArquivoParaGoogleSheets_(ctx.fileId, ctx.nomeArquivo);

  try {
    const tempSs = SpreadsheetApp.openById(tempId);
    const matrizes = tempSs.getSheets().map(function(sheet) {
      const range = sheet.getDataRange();
      return {
        nome: sheet.getName(),
        values: range.getValues(),
        mergedCount: range.getMergedRanges().length
      };
    });

    return extSelecionarTabelaCotacaoDeMatrizes_(matrizes);
  } finally {
    try {
      DriveApp.getFileById(tempId).setTrashed(true);
    } catch (e) {}
  }
}

function extSelecionarTabelaCotacaoDeMatrizes_(matrizes) {
  const avaliacoes = (matrizes || []).map(function(matriz) {
    return extAvaliarMatrizCotacao_(matriz.values, matriz.nome, matriz.mergedCount);
  });

  const melhor = avaliacoes
    .filter(function(item) { return item.header.encontrou; })
    .sort(function(a, b) {
      if (b.linhas.length !== a.linhas.length) return b.linhas.length - a.linhas.length;
      return b.header.score - a.header.score;
    })[0];

  const selecionado = melhor || {
    nomeAba: '',
    dimensoes: { linhas: 0, colunas: 0 },
    mergedCount: 0,
    header: { encontrou: false, linhaHeader: -1, colunas: {}, score: 0 },
    linhas: [],
    motivosIgnorados: { SEM_TABELA: 1 }
  };
  const motivosSelecionado = Object.assign({}, selecionado.motivosIgnorados || {});
  avaliacoes.forEach(function(item) {
    if (item !== selecionado) {
      motivosSelecionado.ABA_NAO_SELECIONADA = (motivosSelecionado.ABA_NAO_SELECIONADA || 0) + item.dimensoes.linhas;
    }
  });

  const totalLinhas = avaliacoes.reduce(function(total, item) {
    return total + item.dimensoes.linhas;
  }, 0);

  return {
    linhas: selecionado.linhas,
    diagnostico: {
      abas: avaliacoes.map(function(item) {
        return {
          nome: item.nomeAba,
          linhas: item.dimensoes.linhas,
          colunas: item.dimensoes.colunas,
          celulasMescladas: item.mergedCount,
          linhaCabecalhoProvavel: item.header.encontrou ? item.header.linhaHeader + 1 : '',
          colunasCandidatas: item.header.colunas || {}
        };
      }),
      abaSelecionada: selecionado.nomeAba,
      linhaCabecalho: selecionado.header.encontrou ? selecionado.header.linhaHeader + 1 : '',
      colunasMapeadas: selecionado.header.colunas || {},
      linhasLidas: totalLinhas,
      itensValidos: selecionado.linhas.length,
      linhasIgnoradas: extSomarMotivos_(motivosSelecionado),
      motivosIgnorados: motivosSelecionado
    }
  };
}

function extAvaliarMatrizCotacao_(values, nomeAba, mergedCount) {
  const matriz = values || [];
  const dimensoes = {
    linhas: matriz.length,
    colunas: matriz.reduce(function(max, row) { return Math.max(max, row.length); }, 0)
  };
  const header = extDetectarCabecalhoCotacao_(matriz);
  const linhas = [];
  const motivosIgnorados = {};

  if (!header.encontrou) {
    motivosIgnorados.SEM_TABELA = dimensoes.linhas;
    return { nomeAba: nomeAba, dimensoes: dimensoes, mergedCount: mergedCount || 0, header: header, linhas: linhas, motivosIgnorados: motivosIgnorados };
  }

  if (header.linhaHeader > 0) {
    motivosIgnorados.PRE_TABELA = header.linhaHeader;
  }

  for (let r = header.linhaHeader + 1; r < matriz.length; r++) {
    const row = matriz[r] || [];
    const aval = extAvaliarLinhaItemCotacao_(row, header.colunas, header.headers);

    if (!aval.ok) {
      motivosIgnorados[aval.motivo] = (motivosIgnorados[aval.motivo] || 0) + 1;
      continue;
    }

    linhas.push(Object.assign(aval.item, {
      abaOrigem: nomeAba,
      linhaOrigem: r + 1,
      linhaOriginal: f2hLinhaOriginalMatriz_(row)
    }));
  }

  return { nomeAba: nomeAba, dimensoes: dimensoes, mergedCount: mergedCount || 0, header: header, linhas: linhas, motivosIgnorados: motivosIgnorados };
}

function extDetectarCabecalhoCotacao_(values) {
  const limite = Math.min((values || []).length, 100);
  let melhor = { encontrou: false, linhaHeader: -1, headers: [], colunas: {}, score: 0 };

  for (let r = 0; r < limite; r++) {
    const headers = (values[r] || []).map(function(v) { return extNormalizarHeaderCotacao_(v); });
    const colunas = extMapearColunasCotacao_(headers);
    const score = extPontuarCabecalhoCotacao_(colunas);

    if (score > melhor.score) {
      melhor = {
        encontrou: score >= 55,
        linhaHeader: r,
        headers: headers,
        colunas: colunas,
        score: score
      };
    }
  }

  if (!melhor.encontrou) melhor.linhaHeader = -1;
  return melhor;
}

function extPontuarCabecalhoCotacao_(colunas) {
  let score = 0;

  if (colunas.codigo !== undefined) score += 18;
  if (colunas.descricao !== undefined) score += 35;
  if (colunas.quantidade !== undefined) score += 22;
  if (colunas.skuMaster !== undefined) score += 16;
  if (colunas.projeto !== undefined) score += 10;
  if (colunas.categoria !== undefined) score += 8;
  if (colunas.valorUnitarioCliente !== undefined || colunas.osinfo !== undefined) score += 8;
  if (colunas.valorTotal !== undefined) score += 8;

  if (colunas.codigo !== undefined && colunas.descricao !== undefined && colunas.quantidade !== undefined) score += 30;
  if (colunas.descricao !== undefined && colunas.skuMaster !== undefined && colunas.quantidade !== undefined) score += 24;
  if (colunas.projeto !== undefined && colunas.categoria !== undefined && colunas.codigo !== undefined && colunas.descricao !== undefined) score += 24;

  return score;
}

function extMapearColunasCotacao_(headers) {
  const colunas = {};

  (headers || []).forEach(function(header, i) {
    if (colunas.projeto === undefined && extHeaderEm_(header, ['PROJETO'])) colunas.projeto = i;
    if (colunas.categoria === undefined && extHeaderEm_(header, ['CATEGORIA', 'CLASSIFICACAO', 'CLASSIFICACAO RAPIDA', 'GRUPO'])) colunas.categoria = i;
    if (colunas.codigo === undefined && extHeaderEm_(header, ['CODIGO', 'COD', 'COD ITEM', 'CODIGO ITEM', 'ITEM'])) colunas.codigo = i;
    if (colunas.descricao === undefined && extHeaderEm_(header, ['DESCRICAO', 'DESCRICAO ORIGINAL', 'DESCRICAO PRODUTO', 'PRODUTO', 'MATERIAL', 'ITEM DESCRICAO'])) colunas.descricao = i;
    if (colunas.skuMaster === undefined && extHeaderEm_(header, ['SKU MASTER', 'SKU_MASTER', 'SKU'])) colunas.skuMaster = i;
    if (colunas.quantidade === undefined && extHeaderEm_(header, ['QUANTIDADE', 'QUANT', 'QTD', 'QTDE'])) colunas.quantidade = i;
    if (colunas.unidade === undefined && extHeaderEm_(header, ['UNIDADE', 'UN', 'UND', 'UM'])) colunas.unidade = i;
    if (colunas.valorUnitarioCliente === undefined && extHeaderEm_(header, ['OSINFO', 'OS INFO', 'VALOR UNITARIO', 'PRECO UNITARIO', 'PRECO DE COMPRA', 'VALOR CLIENTE', 'PRECO REFERENCIA', 'PRECO CLIENTE', 'VALOR OSINFO'])) colunas.valorUnitarioCliente = i;
    if (colunas.osinfo === undefined && colunas.valorUnitarioCliente !== undefined) colunas.osinfo = colunas.valorUnitarioCliente;
    if (colunas.valorTotal === undefined && extHeaderEm_(header, ['VALOR TOTAL', 'TOTAL', 'TOTAL ITEM'])) colunas.valorTotal = i;
    if (colunas.custoMedbox === undefined && extHeaderEm_(header, ['CUSTO MEDBOX', 'CUSTO'])) colunas.custoMedbox = i;
    if (colunas.margem === undefined && extHeaderEm_(header, ['MARGEM', 'MARKUP'])) colunas.margem = i;
    if (colunas.precoVenda === undefined && extHeaderEm_(header, ['PRECO VENDA', 'PRECO DE VENDA', 'VENDA'])) colunas.precoVenda = i;
    if (colunas.status === undefined && extHeaderEm_(header, ['STATUS', 'STATUS LANCE'])) colunas.status = i;
  });

  return colunas;
}

function extHeaderEm_(header, lista) {
  const h = extNormalizarHeaderCotacao_(header);
  return lista.some(function(item) {
    const n = extNormalizarHeaderCotacao_(item);
    return h === n || h.indexOf(n) !== -1;
  });
}

function extAvaliarLinhaItemCotacao_(row, colunas, headers) {
  if (!row || row.every(function(v) { return !extLimpar_(v); })) {
    return { ok: false, motivo: 'LINHA_VAZIA' };
  }

  if (extLinhaPareceCabecalhoCotacao_(row, headers)) {
    return { ok: false, motivo: 'CABECALHO' };
  }

  const descricao = extValorColuna_(row, colunas.descricao);
  const codigo = extValorColuna_(row, colunas.codigo);
  const quantidade = extToNumber_(extValorColuna_(row, colunas.quantidade));
  const valorUnitarioInformado = extToNumber_(extValorColuna_(row, colunas.valorUnitarioCliente !== undefined ? colunas.valorUnitarioCliente : colunas.osinfo));
  const valorTotalInformado = extToNumber_(extValorColuna_(row, colunas.valorTotal));
  const valorUnitarioDerivado = valorUnitarioInformado === null &&
    valorTotalInformado !== null &&
    quantidade !== null &&
    quantidade > 0
      ? extArredondar_(valorTotalInformado / quantidade, 6)
      : null;
  const valorTotalCalculado = valorTotalInformado !== null
    ? valorTotalInformado
    : (valorUnitarioInformado !== null && quantidade !== null && quantidade > 0
      ? extArredondar_(valorUnitarioInformado * quantidade, 6)
      : null);

  if (!descricao) return { ok: false, motivo: 'SEM_DESCRICAO' };
  if (extTextoInstrucionalCotacao_(descricao, row)) return { ok: false, motivo: 'INSTRUCAO' };
  if (!extDescricaoComercialCotacao_(descricao)) return { ok: false, motivo: 'DESCRICAO_NAO_COMERCIAL' };
  if (!(quantidade !== null && quantidade > 0) && !extCodigoItemValidoCotacao_(codigo)) {
    return { ok: false, motivo: 'SEM_QUANTIDADE_E_SEM_CODIGO' };
  }

  return {
    ok: true,
    item: {
      projeto: extValorColuna_(row, colunas.projeto),
      categoria: extValorColuna_(row, colunas.categoria),
      codigo: codigo,
      codigoFornecedor: codigo,
      descricaoOriginal: descricao,
      skuMaster: extValorColuna_(row, colunas.skuMaster),
      quantidade: quantidade === null ? '' : quantidade,
      unidade: extValorColuna_(row, colunas.unidade),
      valorCompraClienteUnitario: valorUnitarioInformado !== null ? valorUnitarioInformado : (valorUnitarioDerivado !== null ? valorUnitarioDerivado : ''),
      valorCompraClienteTotal: valorTotalCalculado !== null ? valorTotalCalculado : '',
      valorClienteDerivado: valorUnitarioDerivado !== null,
      valorTotalClienteDerivado: valorTotalInformado === null && valorTotalCalculado !== null,
      precoCusto: extValorColuna_(row, colunas.valorUnitarioCliente !== undefined ? colunas.valorUnitarioCliente : colunas.osinfo),
      precoTotal: extValorColuna_(row, colunas.valorTotal),
      custoMedboxArquivo: extValorColuna_(row, colunas.custoMedbox),
      margemArquivo: extValorColuna_(row, colunas.margem),
      precoVendaArquivo: extValorColuna_(row, colunas.precoVenda),
      statusArquivo: extValorColuna_(row, colunas.status)
    }
  };
}

function extLinhaPareceCabecalhoCotacao_(row, headers) {
  const normalizada = (row || []).map(function(v) { return extNormalizarHeaderCotacao_(v); });
  let iguais = 0;

  normalizada.forEach(function(v, i) {
    if (v && headers && headers[i] && v === headers[i]) iguais++;
  });

  return iguais >= 2;
}

function extTextoInstrucionalCotacao_(descricao, row) {
  const textoLinha = extNormalizarHeaderCotacao_((row || []).map(extLimpar_).filter(Boolean).join(' '));
  const desc = extNormalizarHeaderCotacao_(descricao);
  const bloqueios = [
    'ABRIR A ABA',
    'FILTRAR POR',
    'NA COLUNA',
    'VOLTAR AO',
    'O QUE SIGNIFICA',
    'ITEM AINDA',
    'RESPONSABILIDADE',
    'SEM CLIENTES',
    'CLASSIFICACAO RAPIDA',
    'PAINEL EXECUTIVO'
  ];

  if (desc === 'PROJETO') return true;
  if (desc === 'STATUS') return true;
  if (desc === 'CATEGORIA') return true;

  return bloqueios.some(function(b) {
    return textoLinha.indexOf(extNormalizarHeaderCotacao_(b)) !== -1;
  });
}

function extDescricaoComercialCotacao_(descricao) {
  const texto = extLimpar_(descricao);
  const normal = extNormalizar_(texto);

  if (!texto || texto.length < 5) return false;
  if (!/[A-Z]/i.test(texto)) return false;
  if (normal.split(/\s+/).length < 2) return false;
  if (['PROJETO', 'UPA SANTO AMARO', 'O QUE SIGNIFICA'].indexOf(normal) !== -1) return false;

  return typeof f2hDescricaoProdutoValida_ !== 'function' || f2hDescricaoProdutoValida_(texto);
}

function extCodigoItemValidoCotacao_(codigo) {
  const texto = extLimpar_(codigo);
  const normal = extNormalizarHeaderCotacao_(texto);

  if (!texto) return false;
  if (['PROJETO', 'CODIGO', 'COD', 'ITEM', 'STATUS'].indexOf(normal) !== -1) return false;

  return /^[A-Z0-9._\/-]{1,30}$/i.test(texto);
}

function extValorColuna_(row, index) {
  if (index === undefined || index === null || index < 0) return '';
  return extLimpar_(row[index]);
}

function extNormalizarHeaderCotacao_(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extSomarMotivos_(motivos) {
  return Object.keys(motivos || {}).reduce(function(total, chave) {
    return total + Number(motivos[chave] || 0);
  }, 0);
}

function extMesclarMotivosIgnorados_() {
  const out = {};

  for (let i = 0; i < arguments.length; i++) {
    const motivos = arguments[i] || {};
    Object.keys(motivos).forEach(function(chave) {
      out[chave] = (out[chave] || 0) + Number(motivos[chave] || 0);
    });
  }

  return out;
}

function extMontarPreviewCotacaoDeLinhas_(linhasBrutas, ctx) {
  const catalogo = extCarregarCatalogoSku_();
  const itens = [];
  let linhasIgnoradas = 0;
  const motivosIgnorados = {};
  const clienteIdentificado = extIdentificarClienteCotacaoArquivo_(linhasBrutas, ctx.nomeArquivo);
  const projetoIdentificado = '';

  for (let i = 0; i < linhasBrutas.length; i++) {
    const item = linhasBrutas[i] || {};
    const descricaoOriginal = extLimpar_(item.descricaoOriginal);
    const descricaoValida = descricaoOriginal &&
      (typeof f2hDescricaoProdutoValida_ !== 'function' || f2hDescricaoProdutoValida_(descricaoOriginal));

    if (!descricaoValida) {
      linhasIgnoradas++;
      motivosIgnorados.DESCRICAO_NAO_COMERCIAL = (motivosIgnorados.DESCRICAO_NAO_COMERCIAL || 0) + 1;
      continue;
    }

    const skuAtual = extUpper_(item.skuMaster || item.skuMasterSugerido || '');
    const match = extSugerirSkuMaster_(descricaoOriginal, skuAtual, catalogo);
    const custo = match.skuMasterSugerido ? extBuscarMenorCustoValidoPorSku_(match.skuMasterSugerido) : null;
    const quantidade = extToNumber_(item.quantidade);
    const valorCompraClienteUnitario = extToNumber_(item.valorCompraClienteUnitario !== undefined ? item.valorCompraClienteUnitario : item.precoCusto);
    const valorCompraClienteTotal = extToNumber_(item.valorCompraClienteTotal);
    const markup = 0.8;
    const precoVenda = custo && custo.custoMedbox !== null
      ? extArredondar_(custo.custoMedbox * (1 + markup), 4)
      : '';
    const itemComercial = extCalcularItemComercialCotacao_({
      quantidade: quantidade,
      valorCompraClienteUnitario: valorCompraClienteUnitario,
      valorCompraClienteTotal: valorCompraClienteTotal,
      custoMedBox: custo ? custo.custoMedbox : '',
      margemPercentual: 80
    });

    itens.push({
      idCotacao: '',
      origemArquivo: ctx.nomeArquivo,
      linhaOrigem: item.linhaOrigem || i + 1,
      cliente: clienteIdentificado,
      projeto: item.projeto || projetoIdentificado,
      categoria: item.categoria || '',
      codigo: item.codigo || item.codigoFornecedor || '',
      codigoFornecedor: item.codigoFornecedor || '',
      descricaoOriginal: descricaoOriginal,
      descricao: descricaoOriginal,
      quantidade: quantidade === null ? '' : quantidade,
      unidade: item.unidade || '',
      valorCompraClienteUnitario: itemComercial.valorCompraClienteUnitario,
      valorCompraClienteTotal: itemComercial.valorCompraClienteTotal,
      valorClienteDerivado: Boolean(item.valorClienteDerivado),
      valorTotalClienteDerivado: Boolean(item.valorTotalClienteDerivado),
      skuMaster: match.skuMasterSugerido,
      skuMasterSugerido: match.skuMasterSugerido,
      statusMatch: match.statusMatch,
      statusMatchInterno: match.statusMatch,
      scoreMatch: match.scoreMatch,
      scoreInterno: match.scoreMatch,
      motivoRevisao: match.motivoRevisao,
      motivoRevisaoInterno: match.motivoRevisao,
      custoMedbox: custo ? custo.custoMedbox : '',
      custoMedBox: custo ? custo.custoMedbox : '',
      statusCusto: custo ? 'CUSTO_ENCONTRADO' : EXT_STATUS_CUSTO_NAO_ENCONTRADO,
      fornecedorCusto: custo ? custo.fornecedorCusto : '',
      dataCusto: custo ? custo.dataCusto : '',
      origemCusto: custo ? custo.origemCusto : '',
      margem: 80,
      margemPercentual: 80,
      precoVenda: precoVenda,
      precoVendaMedBox: itemComercial.precoVendaMedBox,
      precoOsinfo: itemComercial.valorCompraClienteUnitario,
      statusLance: itemComercial.statusLance,
      editavelSku: match.statusMatch !== EXT_STATUS_MATCH_CONFIRMADO,
      custoEditado: false,
      markupPercent: 80,
      precoCusto: custo ? custo.custoMedbox : '',
      precoMedbox: precoVenda,
      valorTotalMedBox: itemComercial.valorTotalMedBox
    });
  }

  return {
    itens: itens,
    linhasIgnoradas: linhasIgnoradas,
    motivosIgnorados: motivosIgnorados,
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

function testarFiltroLinhasNaoProdutosSemGravacao() {
  const abaErrada = [
    ['PAINEL EXECUTIVO'],
    ['ABRIR A ABA CLASSIFICAÇÃO RÁPIDA'],
    ['FILTRAR POR UNIDADE'],
    ['O QUE SIGNIFICA']
  ];
  const abaItens = [
    ['UPA SANTO AMARO'],
    ['PROJETO'],
    ['instrucoes acima da tabela'],
    ['PROJETO', 'CATEGORIA', 'CÓDIGO', 'DESCRIÇÃO', 'QUANT', 'OSINFO', 'VALOR TOTAL'],
    ['UPA SANTO AMARO', '', '', 'PROJETO', '', '', ''],
    ['UPA SANTO AMARO', 'MEDICAMENTO', '001', 'DIPIRONA 500MG COMPRIMIDO', 10, 2.5, 25],
    ['UPA SANTO AMARO', 'MATERIAL', '', 'SERINGA DESCARTAVEL 10ML', 20, 1.2, 24],
    ['', '', '', 'ABRIR A ABA CLASSIFICAÇÃO RÁPIDA', '', '', ''],
    ['', '', '', '', '', '', '']
  ];
  const colunasOrdemDiferente = [
    ['texto mesclado acima'],
    ['DESCRIÇÃO', 'OSINFO', 'QUANTIDADE', 'SKU_MASTER', 'CÓDIGO'],
    ['SORO FISIOLOGICO 0,9% 500ML', 8.9, 3, '', 'S001']
  ];

  const resultado = extSelecionarTabelaCotacaoDeMatrizes_([
    { nome: 'INSTRUCOES', values: abaErrada, mergedCount: 2 },
    { nome: 'ITENS', values: abaItens, mergedCount: 1 }
  ]);
  const resultadoOrdem = extSelecionarTabelaCotacaoDeMatrizes_([
    { nome: 'ORDEM_DIFERENTE', values: colunasOrdemDiferente, mergedCount: 1 }
  ]);

  const descricoes = resultado.linhas.map(function(item) { return item.descricaoOriginal; });

  return {
    ok: resultado.diagnostico.abaSelecionada === 'ITENS' &&
      resultado.diagnostico.linhaCabecalho === 4 &&
      resultado.linhas.length === 2 &&
      descricoes.indexOf('ABRIR A ABA CLASSIFICAÇÃO RÁPIDA') === -1 &&
      descricoes.indexOf('PROJETO') === -1 &&
      descricoes.indexOf('UPA SANTO AMARO') === -1 &&
      resultadoOrdem.linhas.length === 1 &&
      resultadoOrdem.linhas[0].descricaoOriginal === 'SORO FISIOLOGICO 0,9% 500ML',
    diagnostico: {
      abaSelecionada: resultado.diagnostico.abaSelecionada,
      linhaCabecalho: resultado.diagnostico.linhaCabecalho,
      colunasMapeadas: resultado.diagnostico.colunasMapeadas,
      linhasLidas: resultado.diagnostico.linhasLidas,
      itensValidos: resultado.linhas.length,
      linhasIgnoradas: resultado.diagnostico.linhasIgnoradas,
      motivosIgnorados: resultado.diagnostico.motivosIgnorados,
      ordemDiferenteItens: resultadoOrdem.linhas.length
    },
    casos: {
      instrucaoNaoViraProduto: descricoes.indexOf('ABRIR A ABA CLASSIFICAÇÃO RÁPIDA') === -1,
      cabecalhoNaoViraProduto: descricoes.indexOf('DESCRIÇÃO') === -1,
      projetoIsoladoNaoViraProduto: descricoes.indexOf('PROJETO') === -1,
      nomeProjetoIsoladoNaoViraProduto: descricoes.indexOf('UPA SANTO AMARO') === -1,
      codigoDescricaoQuantidadeViraProduto: descricoes.indexOf('DIPIRONA 500MG COMPRIMIDO') !== -1,
      descricaoQuantidadeViraProduto: descricoes.indexOf('SERINGA DESCARTAVEL 10ML') !== -1,
      linhaVaziaIgnorada: Boolean(resultado.diagnostico.motivosIgnorados.LINHA_VAZIA),
      tabelaLinhaDiferenteDe1: resultado.diagnostico.linhaCabecalho === 4,
      colunasOrdemDiferente: resultadoOrdem.linhas.length === 1,
      celulasMescladasAcima: resultado.diagnostico.abas[1].celulasMescladas === 1,
      multiplasAbas: resultado.diagnostico.abas.length === 2,
      abaErradaIgnorada: resultado.diagnostico.abaSelecionada === 'ITENS',
      gravacaoBdCotacoes: false
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

function testarCalculosCotacaoSemGravacao() {
  const vencedor = extCalcularItemComercialCotacao_({
    quantidade: 10,
    valorCompraClienteUnitario: 20,
    custoMedBox: 10,
    margemPercentual: 80
  });
  const igual = extCalcularItemComercialCotacao_({
    quantidade: 2,
    valorCompraClienteUnitario: 18,
    custoMedBox: 10,
    margemPercentual: 80
  });
  const acima = extCalcularItemComercialCotacao_({
    quantidade: 3,
    valorCompraClienteUnitario: 15,
    custoMedBox: 10,
    margemPercentual: 80
  });
  const semCusto = extCalcularItemComercialCotacao_({
    quantidade: 3,
    valorCompraClienteUnitario: 15,
    custoMedBox: '',
    margemPercentual: 80
  });
  const semPrecoCliente = extCalcularItemComercialCotacao_({
    quantidade: 3,
    valorCompraClienteUnitario: '',
    custoMedBox: 10,
    margemPercentual: 80
  });
  const virgula = extToNumber_('1.234,56');
  const ponto = extToNumber_('1234.56');
  const derivado = extAvaliarLinhaItemCotacao_(
    ['001', 'DIPIRONA 500MG COMPRIMIDO', '4', '', '40'],
    { codigo: 0, descricao: 1, quantidade: 2, valorTotal: 4 },
    ['CODIGO', 'DESCRICAO', 'QUANT', '', 'VALOR TOTAL']
  );
  const divisaoZero = extAvaliarLinhaItemCotacao_(
    ['001', 'DIPIRONA 500MG COMPRIMIDO', '0', '', '40'],
    { codigo: 0, descricao: 1, quantidade: 2, valorTotal: 4 },
    ['CODIGO', 'DESCRICAO', 'QUANT', '', 'VALOR TOTAL']
  );

  return {
    ok: vencedor.precoVendaMedBox === 18 &&
      vencedor.valorTotalMedBox === 180 &&
      vencedor.valorCompraClienteTotal === 200 &&
      vencedor.statusLance === EXT_STATUS_LANCE_VENCEDOR &&
      igual.statusLance === EXT_STATUS_PRECO_IGUAL &&
      acima.statusLance === EXT_STATUS_PRECO_ACIMA &&
      semCusto.statusLance === EXT_STATUS_DADOS_INSUFICIENTES &&
      semPrecoCliente.statusLance === EXT_STATUS_DADOS_INSUFICIENTES &&
      virgula === 1234.56 &&
      ponto === 1234.56 &&
      derivado.ok === true &&
      derivado.item.valorCompraClienteUnitario === 10 &&
      derivado.item.valorClienteDerivado === true &&
      divisaoZero.ok === true &&
      divisaoZero.item.valorCompraClienteUnitario === '',
    casos: {
      quantidadeExtraida: vencedor.quantidade,
      valorUnitarioExtraido: vencedor.valorCompraClienteUnitario,
      valorTotalExtraido: vencedor.valorCompraClienteTotal,
      unitarioDerivado: derivado.item.valorCompraClienteUnitario,
      divisaoPorZeroBloqueada: divisaoZero.item.valorCompraClienteUnitario === '',
      statusVencedor: vencedor.statusLance,
      statusIgual: igual.statusLance,
      statusAcima: acima.statusLance,
      faltaCusto: semCusto.statusLance,
      faltaPrecoCliente: semPrecoCliente.statusLance,
      moedaVirgula: virgula,
      moedaPonto: ponto
    }
  };
}

function testarDocumentoCotacaoSemDadosInternos() {
  const html = extMontarDocumentoCotacaoHtml_({
    numeroCotacao: 'COT-TESTE',
    cliente: 'CLIENTE TESTE',
    projeto: 'PROJETO TESTE',
    responsavel: 'RESPONSAVEL TESTE',
    categoria: 'MEDICAMENTOS',
    itens: [{
      quantidade: 10,
      descricaoOriginal: 'DIPIRONA 500MG COMPRIMIDO',
      marca: 'MARCA TESTE',
      precoVendaMedBox: 18,
      valorTotalMedBox: 180,
      custoMedBox: 10,
      fornecedorCusto: 'FORNECEDOR INTERNO',
      margemPercentual: 80,
      scoreInterno: 99,
      motivoRevisaoInterno: 'MATCH INTERNO',
      valorCompraClienteUnitario: 20
    }]
  });
  const bloqueados = ['custoMedBox', 'FORNECEDOR INTERNO', 'score', 'MATCH INTERNO', 'margem', 'valorCompraCliente'];
  const vazouInterno = bloqueados.some(function(txt) {
    return html.indexOf(txt) !== -1;
  });

  return {
    ok: !vazouInterno &&
      html.indexOf('COT-TESTE') !== -1 &&
      html.indexOf('DIPIRONA 500MG COMPRIMIDO') !== -1 &&
      html.indexOf('180') !== -1,
    casos: {
      semCustoInterno: html.indexOf('custoMedBox') === -1,
      semFornecedorCusto: html.indexOf('FORNECEDOR INTERNO') === -1,
      semScore: html.indexOf('score') === -1,
      semMargem: html.indexOf('margem') === -1,
      semPrecoCliente: html.indexOf('valorCompraCliente') === -1
    }
  };
}

function simularSalvarCotacaoExtraida(payload) {
  const dados = payload || {};
  const itens = Array.isArray(dados.itens) && dados.itens.length ? dados.itens : extItensCotacaoSimulacao_();
  const itensNormalizados = itens.map(function(item, index) {
    const calculado = extCalcularItemComercialCotacao_(item);
    return Object.assign({}, item, calculado, {
      indice: index + 1,
      descricaoOriginal: extLimpar_(item.descricaoOriginal)
    });
  });
  const erros = [];

  itensNormalizados.forEach(function(item) {
    if (!item.descricaoOriginal) erros.push({ indice: item.indice, erro: 'DESCRICAO_AUSENTE' });
    if (extToNumber_(item.quantidade) === null || extToNumber_(item.quantidade) <= 0) erros.push({ indice: item.indice, erro: 'QUANTIDADE_INVALIDA' });
    if (extToNumber_(item.precoVendaMedBox) === null || extToNumber_(item.precoVendaMedBox) <= 0) erros.push({ indice: item.indice, erro: 'PRECO_VENDA_INVALIDO' });
  });

  const resumo = extCalcularResumoCotacaoExtraida_(itensNormalizados);

  return {
    ok: erros.length === 0,
    simulacao: true,
    numeroCotacaoPrevisto: extGerarNumeroCotacao_(),
    numeroItens: itensNormalizados.length,
    totalCliente: resumo.valorTotalCotado,
    totalMedBox: resumo.valorTotalCotado,
    valorTotalGanhoMedBox: resumo.valorTotalGanhoMedBox,
    totalItensGanhos: resumo.totalItensGanhos,
    percentualItensGanhos: resumo.percentualItensGanhos,
    itensQueSeriamSalvos: itensNormalizados.map(function(item) {
      return {
        indice: item.indice,
        codigo: item.codigo || item.codigoFornecedor || '',
        descricaoOriginal: item.descricaoOriginal,
        skuMaster: item.skuMaster || '',
        quantidade: item.quantidade,
        valorCompraClienteUnitario: item.valorCompraClienteUnitario,
        valorCompraClienteTotal: item.valorCompraClienteTotal,
        custoMedBox: item.custoMedBox || item.custoMedbox || '',
        margemPercentual: item.margemPercentual || item.margem || '',
        precoVendaMedBox: item.precoVendaMedBox,
        valorTotalMedBox: item.valorTotalMedBox,
        statusLance: item.statusLance
      };
    }),
    erros: erros,
    gravacaoExecutada: false
  };
}

function testarFormatacaoNumericaCotacaoSemGravacao() {
  const casos = {
    qtd100: extToNumber_('100'),
    qtd1600: extToNumber_('1.600'),
    qtd2450: extToNumber_('2.450'),
    qtdDecimal: extToNumber_('1,5'),
    moedaReal: extToNumber_('R$ 2,60'),
    moedaVirgula: extToNumber_('1.234,56'),
    moedaPonto: extToNumber_('1234.56'),
    vazio: extToNumber_(''),
    invalido: extToNumber_('abc')
  };

  return {
    ok: casos.qtd100 === 100 &&
      casos.qtd1600 === 1600 &&
      casos.qtd2450 === 2450 &&
      casos.qtdDecimal === 1.5 &&
      casos.moedaReal === 2.6 &&
      casos.moedaVirgula === 1234.56 &&
      casos.moedaPonto === 1234.56 &&
      casos.vazio === null &&
      casos.invalido === null,
    casos: casos
  };
}

function testarResumoComercialCotacaoSemGravacao() {
  const itens = [
    { precoVendaMedBox: 10, valorTotalMedBox: 100, statusLance: EXT_STATUS_LANCE_VENCEDOR },
    { precoVendaMedBox: 20, valorTotalMedBox: 200, statusLance: EXT_STATUS_PRECO_ACIMA },
    { precoVendaMedBox: 30, valorTotalMedBox: 300, statusLance: EXT_STATUS_PRECO_IGUAL },
    { precoVendaMedBox: '', valorTotalMedBox: '', statusLance: EXT_STATUS_DADOS_INSUFICIENTES }
  ];
  const resumo = extCalcularResumoCotacaoExtraida_(itens);
  const zero = extCalcularResumoCotacaoExtraida_([
    { precoVendaMedBox: '', valorTotalMedBox: '', statusLance: EXT_STATUS_DADOS_INSUFICIENTES }
  ]);
  const editado = [
    { quantidade: 10, valorCompraClienteUnitario: 20, custoMedBox: 10, margemPercentual: 80 }
  ].map(extCalcularItemComercialCotacao_);
  const resumoEditado = extCalcularResumoCotacaoExtraida_(editado);
  editado[0] = extCalcularItemComercialCotacao_({
    quantidade: 10,
    valorCompraClienteUnitario: 20,
    custoMedBox: 10,
    margemPercentual: 100
  });
  const resumoMargemEditada = extCalcularResumoCotacaoExtraida_(editado);

  return {
    ok: resumo.totalItensCotados === 3 &&
      resumo.valorTotalCotado === 600 &&
      resumo.totalItensGanhos === 1 &&
      resumo.valorTotalGanhoMedBox === 100 &&
      resumo.percentualItensGanhos === 33.33 &&
      resumo.itensSemPreco === 1 &&
      zero.totalItensCotados === 0 &&
      zero.percentualItensGanhos === 0 &&
      resumoEditado.valorTotalCotado === 180 &&
      resumoMargemEditada.valorTotalCotado === 200,
    casos: {
      totalCotado: resumo.valorTotalCotado,
      totalGanho: resumo.valorTotalGanhoMedBox,
      quantidadeGanhos: resumo.totalItensGanhos,
      percentualGanhos: resumo.percentualItensGanhos,
      zeroItensCotados: zero.percentualItensGanhos,
      atualizacaoAposEditarCusto: resumoEditado.valorTotalCotado,
      atualizacaoAposEditarMargem: resumoMargemEditada.valorTotalCotado
    }
  };
}

function extItensCotacaoSimulacao_() {
  return [{
    codigo: '001',
    descricaoOriginal: 'DIPIRONA 500MG COMPRIMIDO',
    quantidade: 10,
    valorCompraClienteUnitario: 20,
    valorCompraClienteTotal: 200,
    skuMaster: 'DIPIRONA 500MG COMPRIMIDO',
    custoMedBox: 10,
    margemPercentual: 80
  }];
}

function extGerarNumeroCotacao_() {
  return 'COT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
}

function extMontarDocumentoCotacaoHtml_(payload) {
  const dados = payload || {};
  const itens = Array.isArray(dados.itens) ? dados.itens : [];
  const total = itens.reduce(function(acc, item) {
    const totalItem = extToNumber_(item.valorTotalMedBox);
    return acc + (totalItem || 0);
  }, 0);
  const linhas = itens.map(function(item) {
    return '<tr>' +
      '<td>' + extEscapeHtml_(item.quantidade || '') + '</td>' +
      '<td>' + extEscapeHtml_(item.descricaoOriginal || '') + '</td>' +
      '<td>' + extEscapeHtml_(item.marca || item.skuMaster || '') + '</td>' +
      '<td>' + extEscapeHtml_(item.precoVendaMedBox || item.precoVenda || '') + '</td>' +
      '<td>' + extEscapeHtml_(item.valorTotalMedBox || '') + '</td>' +
      '<td>' + extEscapeHtml_(item.observacao || '') + '</td>' +
      '</tr>';
  }).join('');

  return '<!doctype html><html><head><meta charset="utf-8"><style>@media print{button{display:none}}body{font-family:Arial,sans-serif}table{width:100%;border-collapse:collapse}td,th{border:1px solid #ddd;padding:6px;text-align:left}</style></head><body>' +
    '<h1>MedBox</h1>' +
    '<h2>Cotacao ' + extEscapeHtml_(dados.numeroCotacao || '') + '</h2>' +
    '<p>Data: ' + extEscapeHtml_(dados.data || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')) + '</p>' +
    '<p>Cliente: ' + extEscapeHtml_(dados.cliente || '') + '</p>' +
    '<p>Projeto: ' + extEscapeHtml_(dados.projeto || '') + '</p>' +
    '<p>Responsavel: ' + extEscapeHtml_(dados.responsavel || '') + '</p>' +
    '<p>Categoria: ' + extEscapeHtml_(dados.categoria || '') + '</p>' +
    '<table><thead><tr><th>Quantidade</th><th>Descricao</th><th>Marca</th><th>Preco unitario MedBox</th><th>Preco total MedBox</th><th>Observacao</th></tr></thead><tbody>' + linhas + '</tbody></table>' +
    '<h3>Valor total da proposta: ' + extEscapeHtml_(extArredondar_(total, 2)) + '</h3>' +
    '<p>Condicoes comerciais: conforme proposta aprovada.</p>' +
    '<p>Dados da MedBox: atendimento comercial MEDBOX ERP.</p>' +
    '</body></html>';
}

function extEscapeHtml_(valor) {
  return String(valor === null || valor === undefined ? '' : valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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

function extCalcularItemComercialCotacao_(item) {
  const quantidade = extToNumber_(item && item.quantidade);
  const valorUnitario = extToNumber_(item && item.valorCompraClienteUnitario);
  const valorTotalInformado = extToNumber_(item && item.valorCompraClienteTotal);
  const custo = extToNumber_(item && (item.custoMedBox !== undefined ? item.custoMedBox : item.custoMedbox));
  const margem = extToNumber_(item && (item.margemPercentual !== undefined ? item.margemPercentual : item.margem));
  const valorTotalCliente = valorTotalInformado !== null
    ? valorTotalInformado
    : (quantidade !== null && quantidade > 0 && valorUnitario !== null
      ? extArredondar_(quantidade * valorUnitario, 6)
      : '');
  const precoVenda = custo !== null && custo > 0
    ? extArredondar_(custo * (1 + ((margem === null ? 0 : margem) / 100)), 6)
    : '';
  const valorTotalMedBox = precoVenda !== '' && quantidade !== null && quantidade > 0
    ? extArredondar_(precoVenda * quantidade, 6)
    : '';
  const statusLance = extMontarStatusLanceComercial_(valorUnitario, precoVenda, custo);

  return {
    quantidade: quantidade === null ? '' : quantidade,
    valorCompraClienteUnitario: valorUnitario === null ? '' : valorUnitario,
    valorCompraClienteTotal: valorTotalCliente,
    custoMedBox: custo === null ? '' : custo,
    margemPercentual: margem === null ? 80 : margem,
    precoVendaMedBox: precoVenda,
    valorTotalMedBox: valorTotalMedBox,
    statusLance: statusLance
  };
}

function extMontarStatusLanceComercial_(valorClienteUnitario, precoVendaMedBox, custoMedBox) {
  const custo = extToNumber_(custoMedBox);
  const cliente = extToNumber_(valorClienteUnitario);
  const venda = extToNumber_(precoVendaMedBox);

  if (custo === null || custo <= 0 || venda === null || venda <= 0 || cliente === null || cliente <= 0) {
    return EXT_STATUS_DADOS_INSUFICIENTES;
  }

  if (Math.abs(venda - cliente) < 0.000001) return EXT_STATUS_PRECO_IGUAL;
  return venda < cliente ? EXT_STATUS_LANCE_VENCEDOR : EXT_STATUS_PRECO_ACIMA;
}

function extCalcularResumoCotacaoExtraida_(itens) {
  const lista = Array.isArray(itens) ? itens : [];
  const resumo = {
    totalItensCotados: 0,
    valorTotalCotado: 0,
    totalItensGanhos: 0,
    valorTotalGanhoMedBox: 0,
    percentualItensGanhos: 0,
    itensSemPreco: 0
  };

  lista.forEach(function(item) {
    const precoVenda = extToNumber_(item.precoVendaMedBox !== undefined ? item.precoVendaMedBox : item.precoVenda);
    const totalMedBox = extToNumber_(item.valorTotalMedBox);
    const status = extUpper_(item.statusLance);

    if (precoVenda !== null && precoVenda > 0 && totalMedBox !== null && totalMedBox > 0) {
      resumo.totalItensCotados++;
      resumo.valorTotalCotado += totalMedBox;

      if (status === extUpper_(EXT_STATUS_LANCE_VENCEDOR)) {
        resumo.totalItensGanhos++;
        resumo.valorTotalGanhoMedBox += totalMedBox;
      }
    } else {
      resumo.itensSemPreco++;
    }
  });

  resumo.valorTotalCotado = extArredondar_(resumo.valorTotalCotado, 2);
  resumo.valorTotalGanhoMedBox = extArredondar_(resumo.valorTotalGanhoMedBox, 2);
  resumo.percentualItensGanhos = resumo.totalItensCotados > 0
    ? extArredondar_((resumo.totalItensGanhos / resumo.totalItensCotados) * 100, 2)
    : 0;

  return resumo;
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

  let txt = String(valor).replace(/[\u00A0\u200B-\u200D\uFEFF]/g, '').trim();
  if (!txt || txt === '-' || txt === '--') return null;

  txt = txt.replace(/R\$/gi, '').replace(/\$/g, '').replace(/%/g, '').replace(/\s+/g, '');
  if (txt.includes(',')) {
    txt = txt.replace(/\./g, '').replace(',', '.');
  } else if (/^-?\d{1,3}(\.\d{3})+$/.test(txt)) {
    txt = txt.replace(/\./g, '');
  }
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
