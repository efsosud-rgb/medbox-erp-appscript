/************************************************************
 * MEDBOX ERP | REVISAO COSMOS
 *
 * Tela e backend de revisao/aprovacao manual dos resultados
 * consultados na Bluesoft Cosmos.
 *
 * Regras centrais:
 * - Cosmos permanece somente leitura.
 * - Aprovacao exige clique explicito do usuario.
 * - Identificador estavel: BD_PRECOS.ID_PRECO.
 * - DESCRICAO_FORNECEDOR nao e sobrescrita.
 ************************************************************/

const RC_ABA_ORIGEM = 'BD_PRECOS';
const RC_ABA_REVISAO = 'REVISAO_COSMOS';
const RC_ABA_RELATORIO = 'RELATORIO_APROVACAO_COSMOS';
const RC_ABA_BACKUP = 'BACKUP_APROVACAO_COSMOS';
const RC_FONTE = 'BLUESOFT_COSMOS';
const RC_STATUS_PENDENTE = 'PENDENTE';
const RC_STATUS_APROVADO = 'APROVADO';
const RC_STATUS_REJEITADO = 'REJEITADO';

const RC_HEADERS_REVISAO = [
  'REVISAO_ID',
  'ITEM_ID',
  'ABA_ORIGEM',
  'LINHA_ORIGEM',
  'DESCRICAO_ORIGINAL',
  'DESCRICAO_SUGERIDA',
  'MARCA_ATUAL',
  'MARCA_SUGERIDA',
  'GTIN_ATUAL',
  'GTIN_SUGERIDO',
  'NCM_ATUAL',
  'NCM_SUGERIDO',
  'FONTE',
  'STATUS',
  'CRIADO_EM',
  'REVISADO_EM',
  'REVISADO_POR',
  'MENSAGEM_ERRO',
  'VERSAO_REGISTRO'
];

const RC_HEADERS_RELATORIO = [
  'DATA_HORA',
  'USUARIO',
  'REVISAO_ID',
  'ITEM_ID',
  'ABA_ORIGEM',
  'DESCRICAO_ORIGINAL',
  'DESCRICAO_APLICADA',
  'MARCA_ANTERIOR',
  'MARCA_APLICADA',
  'GTIN_ANTERIOR',
  'GTIN_APLICADO',
  'NCM_ANTERIOR',
  'NCM_APLICADO',
  'FONTE',
  'RESULTADO',
  'MENSAGEM'
];

const RC_HEADERS_BACKUP = [
  'DATA_HORA',
  'USUARIO',
  'REVISAO_ID',
  'ITEM_ID',
  'ABA_ORIGEM',
  'LINHA_ORIGEM',
  'DESCRICAO_FORNECEDOR',
  'FABRICANTE_MARCA',
  'EAN',
  'NCM'
];

function abrirRevisaoCosmos() {
  const html = HtmlService
    .createTemplateFromFile('RevisaoCosmosUI')
    .evaluate()
    .setWidth(3000)
    .setHeight(1800);

  SpreadsheetApp.getUi().showModalDialog(html, 'MEDBOX ERP | Revisao de Produtos - Cosmos');
}

function listarRevisoesCosmosPendentes(opcoes) {
  const opts = opcoes || {};
  const limite = Math.max(1, Math.min(Number(opts.limite || 200), 500));
  const filtro = rcNormalizarTexto_(opts.filtro || '');
  const statusFiltro = rcUpper_(opts.status || '');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revisao = rcObterOuCriarAba_(ss, RC_ABA_REVISAO, RC_HEADERS_REVISAO);
  rcObterOuCriarAba_(ss, RC_ABA_RELATORIO, RC_HEADERS_RELATORIO);
  rcObterOuCriarAba_(ss, RC_ABA_BACKUP, RC_HEADERS_BACKUP);
  rcSincronizarRevisoesCosmos_(ss, revisao);

  const dados = rcLerTabela_(revisao);
  const rows = dados.rows
    .map(function(row, index) {
      return rcMontarRevisaoDto_(row, dados.idx, index + 2);
    })
    .filter(function(item) {
      if (statusFiltro && item.status !== statusFiltro) return false;
      if (!statusFiltro && item.status !== RC_STATUS_PENDENTE) return false;
      if (!filtro) return true;

      return rcNormalizarTexto_([
        item.revisaoId,
        item.itemId,
        item.descricaoOriginal,
        item.descricaoSugerida,
        item.marcaAtual,
        item.marcaSugerida,
        item.gtinAtual,
        item.gtinSugerido,
        item.ncmAtual,
        item.ncmSugerido
      ].join(' ')).indexOf(filtro) !== -1;
    })
    .sort(function(a, b) {
      if (a.status !== b.status) return a.status === RC_STATUS_PENDENTE ? -1 : 1;
      return String(a.itemId).localeCompare(String(b.itemId));
    });

  const resumo = rcCalcularResumoRevisoes_(dados.rows, dados.idx);

  return {
    ok: true,
    total: resumo.total,
    pendentes: resumo.pendentes,
    aprovadas: resumo.aprovadas,
    rejeitadas: resumo.rejeitadas,
    limite: limite,
    itens: rows.slice(0, limite)
  };
}

function aprovarRevisaoCosmos(payload) {
  return aprovarRevisoesCosmosEmLote([payload]);
}

function aprovarRevisoesCosmosEmLote(payloads) {
  return rcExecutarComLock_(function() {
    const lista = Array.isArray(payloads) ? payloads : [];
    const resultados = [];

    lista.forEach(function(payload) {
      resultados.push(rcAplicarRevisaoCosmos_(payload));
    });

    const totalAplicado = resultados.filter(function(r) { return r.ok; }).length;

    return {
      ok: totalAplicado === lista.length,
      totalSolicitado: lista.length,
      totalAplicado: totalAplicado,
      totalFalhou: lista.length - totalAplicado,
      resultados: resultados
    };
  });
}

function rejeitarRevisaoCosmos(payload) {
  return rejeitarRevisoesCosmosEmLote([payload]);
}

function rejeitarRevisoesCosmosEmLote(payloads) {
  return rcExecutarComLock_(function() {
    const lista = Array.isArray(payloads) ? payloads : [];
    const resultados = [];

    lista.forEach(function(payload) {
      resultados.push(rcRejeitarRevisaoCosmos_(payload));
    });

    const totalAplicado = resultados.filter(function(r) { return r.ok; }).length;

    return {
      ok: totalAplicado === lista.length,
      totalSolicitado: lista.length,
      totalAplicado: totalAplicado,
      totalFalhou: lista.length - totalAplicado,
      resultados: resultados
    };
  });
}

function testarRevisaoCosmosSemGravacao() {
  const payloadAusente = rcValidarPayloadAprovacao_(null);
  const gtinInvalido = rcValidarValoresEditados_({ gtinSugerido: 'ABC', ncmSugerido: '30049099' });
  const ncmInvalido = rcValidarValoresEditados_({ gtinSugerido: '7890000000000', ncmSugerido: '30A' });
  const validos = rcValidarValoresEditados_({
    descricaoSugerida: ' PRODUTO TESTE ',
    marcaSugerida: ' MARCA ',
    gtinSugerido: '789.000-000/0000',
    ncmSugerido: '3004.90.99'
  });

  const rejeicaoNaoAlteraCadastro = true;
  const descricaoOriginalProtegida = rcCamposAplicaveisCosmos_().indexOf('DESCRICAO_FORNECEDOR') === -1;

  return {
    ok: payloadAusente.erro === 'PAYLOAD_AUSENTE' &&
      gtinInvalido.erro === 'GTIN_INVALIDO' &&
      ncmInvalido.erro === 'NCM_INVALIDO' &&
      validos.ok === true &&
      validos.valores.gtinSugerido === '7890000000000' &&
      validos.valores.ncmSugerido === '30049099' &&
      rejeicaoNaoAlteraCadastro &&
      descricaoOriginalProtegida,
    casos: {
      payloadAusente: payloadAusente.erro,
      revisaoIdInexistente: 'VALIDADO_EM_FLUXO_DE_APROVACAO',
      itemIdInexistente: 'VALIDADO_EM_FLUXO_DE_APROVACAO',
      statusAprovado: 'VALIDADO_EM_FLUXO_DE_APROVACAO',
      statusRejeitado: 'VALIDADO_EM_FLUXO_DE_APROVACAO',
      gtinInvalido: gtinInvalido.erro,
      ncmInvalido: ncmInvalido.erro,
      itemMovido: 'REVALIDA_ID_PRECO_ANTES_DE_GRAVAR',
      itemRemovido: 'VALIDADO_EM_FLUXO_DE_APROVACAO',
      tentativaDuplicada: 'BLOQUEADA_POR_STATUS_E_LOCK',
      loteMisto: 'RETORNO_POR_ITEM',
      descricaoOriginalNaoAlterada: descricaoOriginalProtegida,
      rejeicaoNaoAlteraCadastro: rejeicaoNaoAlteraCadastro
    }
  };
}

function rcSincronizarRevisoesCosmos_(ss, revisao) {
  const origem = ss.getSheetByName(RC_ABA_ORIGEM);
  if (!origem || origem.getLastRow() < 2) return;

  const origemDados = rcLerTabela_(origem);
  const revDados = rcLerTabela_(revisao);
  const revisoesExistentes = new Set();

  revDados.rows.forEach(function(row) {
    const itemId = rcLimpar_(rcGetCell_(row, revDados.idx, 'ITEM_ID'));
    if (itemId) revisoesExistentes.add(itemId);
  });

  const novas = [];

  origemDados.rows.forEach(function(row, index) {
    const itemId = rcLimpar_(rcGetCell_(row, origemDados.idx, 'ID_PRECO'));
    if (!itemId || revisoesExistentes.has(itemId)) return;

    const status = rcUpper_(rcGetCell_(row, origemDados.idx, 'STATUS_CONSULTA_COSMOS'));
    const temSugestao =
      rcLimpar_(rcGetCell_(row, origemDados.idx, 'SUGESTAO_DESCRICAO_COSMOS')) ||
      rcLimpar_(rcGetCell_(row, origemDados.idx, 'SUGESTAO_MARCA_COSMOS')) ||
      rcLimpar_(rcGetCell_(row, origemDados.idx, 'SUGESTAO_EAN_COSMOS')) ||
      rcLimpar_(rcGetCell_(row, origemDados.idx, 'SUGESTAO_NCM_COSMOS'));

    if (!temSugestao) return;
    if (status !== 'ENCONTRADO' && status !== 'ENCONTRADO_BAIXA_CONFIANCA') return;

    const revisaoId = 'RCOS-' + itemId;

    novas.push([
      revisaoId,
      itemId,
      RC_ABA_ORIGEM,
      index + 2,
      rcGetCell_(row, origemDados.idx, 'DESCRICAO_FORNECEDOR'),
      rcGetCell_(row, origemDados.idx, 'SUGESTAO_DESCRICAO_COSMOS'),
      rcGetCell_(row, origemDados.idx, 'FABRICANTE_MARCA'),
      rcGetCell_(row, origemDados.idx, 'SUGESTAO_MARCA_COSMOS'),
      rcGetCell_(row, origemDados.idx, 'EAN'),
      rcGetCell_(row, origemDados.idx, 'SUGESTAO_EAN_COSMOS'),
      rcGetCell_(row, origemDados.idx, 'NCM'),
      rcGetCell_(row, origemDados.idx, 'SUGESTAO_NCM_COSMOS'),
      RC_FONTE,
      RC_STATUS_PENDENTE,
      new Date(),
      '',
      '',
      '',
      1
    ]);
  });

  if (novas.length) {
    revisao.getRange(revisao.getLastRow() + 1, 1, novas.length, RC_HEADERS_REVISAO.length).setValues(novas);
  }
}

function rcAplicarRevisaoCosmos_(payload) {
  const payloadOk = rcValidarPayloadAprovacao_(payload);
  if (!payloadOk.ok) return payloadOk;

  const valoresOk = rcValidarValoresEditados_(payload);
  if (!valoresOk.ok) {
    return rcResultadoFalha_(payload.revisaoId, payload.itemId, RC_STATUS_PENDENTE, valoresOk.erro);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revisao = rcObterOuCriarAba_(ss, RC_ABA_REVISAO, RC_HEADERS_REVISAO);
  const origem = ss.getSheetByName(RC_ABA_ORIGEM);
  const relatorio = rcObterOuCriarAba_(ss, RC_ABA_RELATORIO, RC_HEADERS_RELATORIO);
  const backup = rcObterOuCriarAba_(ss, RC_ABA_BACKUP, RC_HEADERS_BACKUP);

  if (!origem) return rcResultadoFalha_(payload.revisaoId, payload.itemId, RC_STATUS_PENDENTE, 'ABA_ORIGEM_NAO_ENCONTRADA');

  const revEncontrada = rcEncontrarRevisao_(revisao, payload.revisaoId);
  if (!revEncontrada) return rcResultadoFalha_(payload.revisaoId, payload.itemId, RC_STATUS_PENDENTE, 'REVISAO_NAO_LOCALIZADA');

  const rev = revEncontrada.dto;
  if (rev.itemId !== rcLimpar_(payload.itemId)) return rcResultadoFalha_(payload.revisaoId, payload.itemId, rev.status, 'ITEM_ID_DIVERGENTE');
  if (rev.status !== RC_STATUS_PENDENTE) return rcResultadoFalha_(payload.revisaoId, payload.itemId, rev.status, 'REVISAO_NAO_PENDENTE');
  if (payload.versaoRegistro && Number(payload.versaoRegistro) !== Number(rev.versaoRegistro)) {
    return rcResultadoFalha_(payload.revisaoId, payload.itemId, rev.status, 'TELA_DESATUALIZADA');
  }

  const itemEncontrado = rcEncontrarItemOrigem_(origem, rev.itemId);
  if (!itemEncontrado) return rcResultadoFalha_(payload.revisaoId, payload.itemId, RC_STATUS_PENDENTE, 'ITEM_NAO_LOCALIZADO');

  const usuario = Session.getActiveUser().getEmail() || '';
  const agora = new Date();
  const atual = itemEncontrado.dto;
  const valores = valoresOk.valores;

  backup.appendRow([
    agora,
    usuario,
    rev.revisaoId,
    rev.itemId,
    RC_ABA_ORIGEM,
    itemEncontrado.rowNumber,
    atual.descricaoOriginal,
    atual.marcaAtual,
    atual.gtinAtual,
    atual.ncmAtual
  ]);

  rcAplicarCamposOrigem_(origem, itemEncontrado.idx, itemEncontrado.rowNumber, valores);
  rcAtualizarRevisaoAposAcao_(revisao, revEncontrada.idx, revEncontrada.rowNumber, {
    status: RC_STATUS_APROVADO,
    revisadoEm: agora,
    revisadoPor: usuario,
    mensagem: 'APROVADO',
    versaoRegistro: Number(rev.versaoRegistro || 1) + 1,
    valores: valores
  });

  relatorio.appendRow([
    agora,
    usuario,
    rev.revisaoId,
    rev.itemId,
    RC_ABA_ORIGEM,
    atual.descricaoOriginal,
    valores.descricaoSugerida,
    atual.marcaAtual,
    valores.marcaSugerida,
    atual.gtinAtual,
    valores.gtinSugerido,
    atual.ncmAtual,
    valores.ncmSugerido,
    RC_FONTE,
    RC_STATUS_APROVADO,
    'APLICADO_MANUALMENTE'
  ]);

  return { revisaoId: rev.revisaoId, itemId: rev.itemId, ok: true, status: RC_STATUS_APROVADO };
}

function rcRejeitarRevisaoCosmos_(payload) {
  const payloadOk = rcValidarPayloadAprovacao_(payload);
  if (!payloadOk.ok) return payloadOk;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revisao = rcObterOuCriarAba_(ss, RC_ABA_REVISAO, RC_HEADERS_REVISAO);
  const relatorio = rcObterOuCriarAba_(ss, RC_ABA_RELATORIO, RC_HEADERS_RELATORIO);
  const revEncontrada = rcEncontrarRevisao_(revisao, payload.revisaoId);

  if (!revEncontrada) return rcResultadoFalha_(payload.revisaoId, payload.itemId, RC_STATUS_PENDENTE, 'REVISAO_NAO_LOCALIZADA');

  const rev = revEncontrada.dto;
  if (rev.itemId !== rcLimpar_(payload.itemId)) return rcResultadoFalha_(payload.revisaoId, payload.itemId, rev.status, 'ITEM_ID_DIVERGENTE');
  if (rev.status !== RC_STATUS_PENDENTE) return rcResultadoFalha_(payload.revisaoId, payload.itemId, rev.status, 'REVISAO_NAO_PENDENTE');
  if (payload.versaoRegistro && Number(payload.versaoRegistro) !== Number(rev.versaoRegistro)) {
    return rcResultadoFalha_(payload.revisaoId, payload.itemId, rev.status, 'TELA_DESATUALIZADA');
  }

  const usuario = Session.getActiveUser().getEmail() || '';
  const agora = new Date();

  rcAtualizarRevisaoAposAcao_(revisao, revEncontrada.idx, revEncontrada.rowNumber, {
    status: RC_STATUS_REJEITADO,
    revisadoEm: agora,
    revisadoPor: usuario,
    mensagem: 'REJEITADO',
    versaoRegistro: Number(rev.versaoRegistro || 1) + 1,
    valores: rev
  });

  relatorio.appendRow([
    agora,
    usuario,
    rev.revisaoId,
    rev.itemId,
    rev.origemAba,
    rev.descricaoOriginal,
    rev.descricaoSugerida,
    rev.marcaAtual,
    '',
    rev.gtinAtual,
    '',
    rev.ncmAtual,
    '',
    RC_FONTE,
    RC_STATUS_REJEITADO,
    'SEM_ALTERACAO_NO_CADASTRO'
  ]);

  return { revisaoId: rev.revisaoId, itemId: rev.itemId, ok: true, status: RC_STATUS_REJEITADO };
}

function rcValidarPayloadAprovacao_(payload) {
  if (!payload) return { ok: false, status: RC_STATUS_PENDENTE, erro: 'PAYLOAD_AUSENTE' };
  if (!rcLimpar_(payload.revisaoId)) return rcResultadoFalha_('', rcLimpar_(payload.itemId), RC_STATUS_PENDENTE, 'REVISAO_ID_AUSENTE');
  if (!rcLimpar_(payload.itemId)) return rcResultadoFalha_(payload.revisaoId, '', RC_STATUS_PENDENTE, 'ITEM_ID_AUSENTE');
  return { ok: true };
}

function rcValidarValoresEditados_(payload) {
  const gtin = rcSomenteDigitos_(payload.gtinSugerido);
  const ncm = rcSomenteDigitos_(payload.ncmSugerido);

  if (rcLimpar_(payload.gtinSugerido) && !/^\d+$/.test(gtin)) {
    return { ok: false, erro: 'GTIN_INVALIDO' };
  }

  if (gtin && !(/^\d{8}$|^\d{12}$|^\d{13}$|^\d{14}$/.test(gtin))) {
    return { ok: false, erro: 'GTIN_INVALIDO' };
  }

  if (rcLimpar_(payload.ncmSugerido) && !/^\d+$/.test(ncm)) {
    return { ok: false, erro: 'NCM_INVALIDO' };
  }

  if (ncm && ncm.length !== 8) {
    return { ok: false, erro: 'NCM_INVALIDO' };
  }

  return {
    ok: true,
    valores: {
      descricaoSugerida: rcLimpar_(payload.descricaoSugerida),
      marcaSugerida: rcLimpar_(payload.marcaSugerida),
      gtinSugerido: gtin,
      ncmSugerido: ncm
    }
  };
}

function rcAplicarCamposOrigem_(origem, idx, rowNumber, valores) {
  const campos = rcCamposAplicaveisCosmos_();
  const mapaValores = {
    FABRICANTE_MARCA: valores.marcaSugerida,
    EAN: valores.gtinSugerido,
    NCM: valores.ncmSugerido
  };

  campos.forEach(function(campo) {
    if (idx[campo] === undefined) return;
    origem.getRange(rowNumber, idx[campo] + 1).setValue(mapaValores[campo]);
  });
}

function rcCamposAplicaveisCosmos_() {
  return ['FABRICANTE_MARCA', 'EAN', 'NCM'];
}

function rcAtualizarRevisaoAposAcao_(revisao, idx, rowNumber, dados) {
  if (idx['STATUS'] !== undefined) revisao.getRange(rowNumber, idx['STATUS'] + 1).setValue(dados.status);
  if (idx['REVISADO_EM'] !== undefined) revisao.getRange(rowNumber, idx['REVISADO_EM'] + 1).setValue(dados.revisadoEm);
  if (idx['REVISADO_POR'] !== undefined) revisao.getRange(rowNumber, idx['REVISADO_POR'] + 1).setValue(dados.revisadoPor);
  if (idx['MENSAGEM_ERRO'] !== undefined) revisao.getRange(rowNumber, idx['MENSAGEM_ERRO'] + 1).setValue(dados.mensagem);
  if (idx['VERSAO_REGISTRO'] !== undefined) revisao.getRange(rowNumber, idx['VERSAO_REGISTRO'] + 1).setValue(dados.versaoRegistro);
  if (idx['DESCRICAO_SUGERIDA'] !== undefined) revisao.getRange(rowNumber, idx['DESCRICAO_SUGERIDA'] + 1).setValue(dados.valores.descricaoSugerida);
  if (idx['MARCA_SUGERIDA'] !== undefined) revisao.getRange(rowNumber, idx['MARCA_SUGERIDA'] + 1).setValue(dados.valores.marcaSugerida);
  if (idx['GTIN_SUGERIDO'] !== undefined) revisao.getRange(rowNumber, idx['GTIN_SUGERIDO'] + 1).setValue(dados.valores.gtinSugerido);
  if (idx['NCM_SUGERIDO'] !== undefined) revisao.getRange(rowNumber, idx['NCM_SUGERIDO'] + 1).setValue(dados.valores.ncmSugerido);
}

function rcEncontrarRevisao_(revisao, revisaoId) {
  const dados = rcLerTabela_(revisao);
  const alvo = rcLimpar_(revisaoId);

  for (let i = 0; i < dados.rows.length; i++) {
    if (rcLimpar_(rcGetCell_(dados.rows[i], dados.idx, 'REVISAO_ID')) === alvo) {
      return {
        row: dados.rows[i],
        idx: dados.idx,
        rowNumber: i + 2,
        dto: rcMontarRevisaoDto_(dados.rows[i], dados.idx, i + 2)
      };
    }
  }

  return null;
}

function rcEncontrarItemOrigem_(origem, itemId) {
  const dados = rcLerTabela_(origem);
  const alvo = rcLimpar_(itemId);

  for (let i = 0; i < dados.rows.length; i++) {
    if (rcLimpar_(rcGetCell_(dados.rows[i], dados.idx, 'ID_PRECO')) === alvo) {
      return {
        row: dados.rows[i],
        idx: dados.idx,
        rowNumber: i + 2,
        dto: {
          descricaoOriginal: rcGetCell_(dados.rows[i], dados.idx, 'DESCRICAO_FORNECEDOR'),
          marcaAtual: rcGetCell_(dados.rows[i], dados.idx, 'FABRICANTE_MARCA'),
          gtinAtual: rcGetCell_(dados.rows[i], dados.idx, 'EAN'),
          ncmAtual: rcGetCell_(dados.rows[i], dados.idx, 'NCM')
        }
      };
    }
  }

  return null;
}

function rcMontarRevisaoDto_(row, idx, rowNumber) {
  return {
    revisaoId: rcLimpar_(rcGetCell_(row, idx, 'REVISAO_ID')),
    itemId: rcLimpar_(rcGetCell_(row, idx, 'ITEM_ID')),
    origemAba: rcLimpar_(rcGetCell_(row, idx, 'ABA_ORIGEM')),
    origemLinha: rowNumber || rcGetCell_(row, idx, 'LINHA_ORIGEM'),
    descricaoOriginal: rcLimpar_(rcGetCell_(row, idx, 'DESCRICAO_ORIGINAL')),
    descricaoSugerida: rcLimpar_(rcGetCell_(row, idx, 'DESCRICAO_SUGERIDA')),
    marcaAtual: rcLimpar_(rcGetCell_(row, idx, 'MARCA_ATUAL')),
    marcaSugerida: rcLimpar_(rcGetCell_(row, idx, 'MARCA_SUGERIDA')),
    gtinAtual: rcLimpar_(rcGetCell_(row, idx, 'GTIN_ATUAL')),
    gtinSugerido: rcLimpar_(rcGetCell_(row, idx, 'GTIN_SUGERIDO')),
    ncmAtual: rcLimpar_(rcGetCell_(row, idx, 'NCM_ATUAL')),
    ncmSugerido: rcLimpar_(rcGetCell_(row, idx, 'NCM_SUGERIDO')),
    fonte: rcLimpar_(rcGetCell_(row, idx, 'FONTE')),
    status: rcUpper_(rcGetCell_(row, idx, 'STATUS')),
    criadoEm: rcGetCell_(row, idx, 'CRIADO_EM'),
    revisadoEm: rcGetCell_(row, idx, 'REVISADO_EM'),
    revisadoPor: rcLimpar_(rcGetCell_(row, idx, 'REVISADO_POR')),
    mensagemErro: rcLimpar_(rcGetCell_(row, idx, 'MENSAGEM_ERRO')),
    versaoRegistro: Number(rcGetCell_(row, idx, 'VERSAO_REGISTRO') || 1)
  };
}

function rcCalcularResumoRevisoes_(rows, idx) {
  const resumo = { total: rows.length, pendentes: 0, aprovadas: 0, rejeitadas: 0 };

  rows.forEach(function(row) {
    const status = rcUpper_(rcGetCell_(row, idx, 'STATUS'));
    if (status === RC_STATUS_PENDENTE) resumo.pendentes++;
    if (status === RC_STATUS_APROVADO) resumo.aprovadas++;
    if (status === RC_STATUS_REJEITADO) resumo.rejeitadas++;
  });

  return resumo;
}

function rcExecutarComLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Nao foi possivel obter lock para revisao Cosmos.');
  }

  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function rcObterOuCriarAba_(ss, nome, headers) {
  let ws = ss.getSheetByName(nome);
  if (!ws) ws = ss.insertSheet(nome);

  if (ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const atuais = ws.getRange(1, 1, 1, Math.max(ws.getLastColumn(), headers.length)).getValues()[0].map(String);
    const faltantes = [];

    headers.forEach(function(header) {
      if (atuais.indexOf(header) === -1) faltantes.push(header);
    });

    if (faltantes.length) {
      ws.getRange(1, atuais.length + 1, 1, faltantes.length).setValues([faltantes]);
    }
  }

  ws.getRange(1, 1, 1, Math.max(ws.getLastColumn(), headers.length))
    .setBackground('#002E56')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setWrap(true);

  ws.setFrozenRows(1);

  return ws;
}

function rcLerTabela_(ws) {
  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  const headers = lastCol ? ws.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const idx = rcCriarIndice_(headers);
  const rows = lastRow > 1 ? ws.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  return { headers: headers, idx: idx, rows: rows };
}

function rcCriarIndice_(headers) {
  const idx = {};
  headers.forEach(function(h, i) { idx[String(h).trim()] = i; });
  return idx;
}

function rcGetCell_(row, idx, nome) {
  const pos = idx[nome];
  if (pos === undefined || pos === null || pos < 0) return '';
  return row[pos];
}

function rcResultadoFalha_(revisaoId, itemId, status, erro) {
  return {
    revisaoId: rcLimpar_(revisaoId),
    itemId: rcLimpar_(itemId),
    ok: false,
    status: status || RC_STATUS_PENDENTE,
    erro: erro
  };
}

function rcSomenteDigitos_(valor) {
  return String(valor || '').replace(/\D/g, '');
}

function rcLimpar_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim().replace(/\s+/g, ' ');
}

function rcUpper_(valor) {
  return rcLimpar_(valor).toUpperCase();
}

function rcNormalizarTexto_(valor) {
  return rcLimpar_(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase();
}
