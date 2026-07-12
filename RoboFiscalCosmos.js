/************************************************************
 * MEDBOX ERP | ROBÔ FISCAL — API COSMOS (BLUESOFT)
 * BUSCA POR DESCRIÇÃO DO PRODUTO (SEM DEPENDER DE EAN)
 *
 * Arquivo: RoboFiscalCosmos.js
 *
 * NÃO altera Código.gs, RoboFiscal.gs ou qualquer outro arquivo.
 * NÃO sobrescreve SKU_MASTER, NCM, EAN ou CST diretamente.
 * Grava apenas em colunas de SUGESTÃO. Revisão e aplicação
 * em campos oficiais são SEMPRE manuais — nunca automáticas.
 *
 * SEGURANÇA: o token da API Cosmos NÃO fica neste arquivo.
 * É lido via PropertiesService.getScriptProperties(), chave
 * 'RFC_COSMOS_TOKEN'. Configure rodando configurarTokenCosmos()
 * uma vez, manualmente, pelo editor do Apps Script.
 *
 * NÃO possui gatilho automático (trigger onChange). O robô só
 * roda quando chamado manualmente pelo menu ou pelo Codex/CLI —
 * nunca disparado sozinho por edição da planilha.
 *
 * Estratégia:
 * - Lê DESCRICAO_FORNECEDOR (coluna G) de cada linha.
 * - Consulta a API Cosmos por texto: /products?query=...
 * - Compara palavras da descrição original com a do resultado,
 *   IGNORANDO tokens numéricos/dosagem (ex.: "3ML", "0,9%"),
 *   que geram falso positivo entre produtos diferentes do
 *   mesmo fabricante/volume.
 * - EXIGE que a primeira palavra significativa da descrição
 *   original (normalmente o nome do produto/princípio ativo)
 *   apareça no candidato. Sem isso, nunca marca ENCONTRADO.
 * - Sinaliza ALERTA_NCM_FORA_CAPITULO_30 quando o NCM sugerido
 *   não começa em "30", para revisão manual extra.
 *
 * Processa em lotes (respeita limite de 6 min do Apps Script) e
 * retoma de onde parou usando PropertiesService.
 ************************************************************/

const RFC_ABA_BD_PRECOS = 'BD_PRECOS';

const RFC_COSMOS_USER_AGENT = 'Cosmos-API-Request';

const RFC_LIMITE_TEMPO_MS = 4.5 * 60 * 1000;
const RFC_PAUSA_MS = 1200;
const RFC_PROP_ULTIMA_LINHA = 'RFC_ULTIMA_LINHA_PROCESSADA';
const RFC_PROP_TOKEN = 'RFC_COSMOS_TOKEN';
const RFC_RESULTADOS_POR_BUSCA = 5;
const RFC_MIN_PALAVRAS_COINCIDENTES = 2;
const RFC_PREFIXO_NCM_FARMACEUTICO = '30';

const RFC_COLUNAS_NOVAS = [
  'SUGESTAO_DESCRICAO_COSMOS',
  'SUGESTAO_NCM_COSMOS',
  'SUGESTAO_EAN_COSMOS',
  'SUGESTAO_MARCA_COSMOS',
  'STATUS_CONSULTA_COSMOS',
  'ALERTA_NCM_FORA_CAPITULO_30',
  'FONTE_COSMOS_URL',
  'DATA_CONSULTA_COSMOS'
];

const RFC_STOPWORDS = [
  'DE', 'DA', 'DO', 'DAS', 'DOS', 'COM', 'PARA', 'EM', 'E', 'OU',
  'UN', 'UND', 'UNID', 'CX', 'CAIXA', 'FR', 'FRASCO', 'AMP', 'AMPOLA',
  'AMPOLAS', 'ML', 'MG', 'MCG', 'UI', 'CADA', 'INDIVIDUAL', 'EMBALADO',
  'PC', 'PCT'
];

/************************************************************
 * FUNÇÃO PRINCIPAL — chamada pelo menu (executarLoteVarreduraFiscalCosmos)
 ************************************************************/

function executarLoteVarreduraFiscalCosmos() {
  processarLoteCosmos(50);
}

/**
 * Processa um lote de produtos da BD_PRECOS na API Cosmos.
 * NUNCA escreve em SKU_MASTER, NCM, EAN ou CST oficiais —
 * apenas nas colunas SUGESTAO_*, para revisão manual.
 */
function processarLoteCosmos(tamanhoLote) {
  tamanhoLote = tamanhoLote || 50;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(RFC_ABA_BD_PRECOS);

  if (!aba) {
    Logger.log("Aba 'BD_PRECOS' não encontrada.");
    return;
  }

  const token = PropertiesService.getScriptProperties().getProperty(RFC_PROP_TOKEN);

  if (!token) {
    const msg = 'Token da API Cosmos não configurado. Rode configurarTokenCosmos() primeiro.';
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
    return;
  }

  const headers = rfcGarantirColunas_(aba);
  const idx = rfcCriarIndice_(headers);

  const colunasCriticas = ['DESCRICAO_FORNECEDOR', 'NCM', 'EAN', 'STATUS_CONSULTA_COSMOS'];
  for (const col of colunasCriticas) {
    if (idx[col] === undefined) {
      Logger.log('Erro: Coluna ' + col + ' não encontrada na BD_PRECOS.');
      return;
    }
  }

  const inicio = Date.now();
  const props = PropertiesService.getScriptProperties();
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;

  let linhaInicial = Number(props.getProperty(RFC_PROP_ULTIMA_LINHA)) || 1;
  linhaInicial = Math.max(linhaInicial + 1, 2);
  if (linhaInicial > ultimaLinha) linhaInicial = 2;

  let processados = 0;
  let encontradas = 0;
  let baixaConfianca = 0;
  let naoEncontradas = 0;
  let semDescricao = 0;
  let alertasNcm = 0;
  let erros = 0;

  for (let linha = linhaInicial; linha <= ultimaLinha && processados < tamanhoLote; linha++) {

    if (Date.now() - inicio > RFC_LIMITE_TEMPO_MS) {
      props.setProperty(RFC_PROP_ULTIMA_LINHA, String(linha - 1));
      break;
    }

    const linhaData = aba.getRange(linha, 1, 1, headers.length).getValues()[0];

    const status = rfcUpper_(rfcGetCell_(linhaData, idx, 'STATUS_CONSULTA_COSMOS'));
    if (status) continue; // já processada antes

    const descricao = rfcLimpar_(rfcGetCell_(linhaData, idx, 'DESCRICAO_FORNECEDOR'));

    if (!descricao || descricao.length < 5) {
      rfcEscreverResultado_(aba, linha, idx, {
        status: 'SEM_DESCRICAO', descricao: '', ncm: '', ean: '', marca: '', url: '', data: new Date()
      });
      semDescricao++;
      continue;
    }

    try {
      const resultado = rfcConsultarCosmosPorDescricao_(descricao, token);
      rfcEscreverResultado_(aba, linha, idx, resultado);

      if (resultado.status === 'ENCONTRADO') encontradas++;
      if (resultado.status === 'ENCONTRADO_BAIXA_CONFIANCA') baixaConfianca++;
      if (resultado.status === 'NAO_ENCONTRADO') naoEncontradas++;
      if (rfcNcmForaCapituloFarmaceutico_(resultado.ncm)) alertasNcm++;

      processados++;
      Utilities.sleep(RFC_PAUSA_MS);

    } catch (e) {
      Logger.log('Erro na linha ' + linha + ': ' + e.message);
      rfcSetCell_(aba, linha, idx, 'STATUS_CONSULTA_COSMOS', 'ERRO');
      erros++;
    }
  }

  if (processados < tamanhoLote) {
    props.deleteProperty(RFC_PROP_ULTIMA_LINHA); // percorreu tudo disponível no lote
  }

  const resumo =
    'Processamento Cosmos concluído.\n\n' +
    'Processados: ' + processados + '\n' +
    'Encontrados (alta confiança): ' + encontradas + '\n' +
    'Encontrados (baixa confiança): ' + baixaConfianca + '\n' +
    'Não encontrados: ' + naoEncontradas + '\n' +
    'Sem descrição: ' + semDescricao + '\n' +
    'Alertas NCM fora capítulo 30: ' + alertasNcm + '\n' +
    'Erros: ' + erros + '\n\n' +
    'Nada foi aplicado em SKU_MASTER/NCM/EAN/CST oficiais — revise as colunas SUGESTAO_* manualmente.';

  Logger.log(resumo);
  try { SpreadsheetApp.getUi().alert(resumo); } catch (e) {}
}

/************************************************************
 * CONFIGURAR TOKEN (rode uma vez, manualmente, pelo editor)
 ************************************************************/

function configurarTokenCosmos() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt(
    'Configurar Token Cosmos',
    'Cole aqui o token gerado em cosmos.bluesoft.com.br/api:',
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) return;

  const token = resposta.getResponseText().trim();

  if (!token) {
    ui.alert('Nenhum token informado. Nada foi salvo.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(RFC_PROP_TOKEN, token);
  ui.alert('Token salvo com sucesso nas Propriedades do Script.');
}

function verificarTokenCosmosConfigurado() {
  const token = PropertiesService.getScriptProperties().getProperty(RFC_PROP_TOKEN);
  const configurado = Boolean(token && String(token).trim());

  return {
    ok: configurado,
    propriedade: RFC_PROP_TOKEN,
    configurado: configurado
  };
}

function rfcTokenCosmosSeguro_() {
  const token = PropertiesService.getScriptProperties().getProperty(RFC_PROP_TOKEN);
  let tokenSeguro = token ? String(token).trim() : '';

  if ((tokenSeguro.charAt(0) === '"' && tokenSeguro.charAt(tokenSeguro.length - 1) === '"') ||
      (tokenSeguro.charAt(0) === "'" && tokenSeguro.charAt(tokenSeguro.length - 1) === "'")) {
    tokenSeguro = tokenSeguro.substring(1, tokenSeguro.length - 1).trim();
  }

  return tokenSeguro;
}

function rfcOpcoesGetCosmos_(tokenSeguro) {
  return {
    method: 'get',
    headers: {
      'X-Cosmos-Token': tokenSeguro,
      'User-Agent': 'Cosmos-API-Request',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };
}

function rfcNormalizarConsultaDescricao_(descricao) {
  const consulta = String(descricao || '').trim().replace(/\s+/g, ' ');
  return consulta.substring(0, 80).trim();
}

function rfcNormalizarConsultaGtin_(gtin) {
  return String(gtin || '').replace(/[\s.\-/]/g, '');
}

function rfcGtinValido_(gtin) {
  return /^\d{8}$|^\d{12}$|^\d{13}$|^\d{14}$/.test(gtin);
}

function rfcErroCosmosSeguro_(codigoStatus, tipoConsulta, encontradoQuando404) {
  if (codigoStatus === 401) {
    return { ok: false, autenticado: false, httpStatus: 401, erro: 'NAO_AUTORIZADO' };
  }

  if (codigoStatus === 403) {
    return { ok: false, autenticado: false, httpStatus: 403, erro: 'ACESSO_NEGADO' };
  }

  if (codigoStatus === 404) {
    const resposta404 = { ok: true, httpStatus: 404, encontrado: false };
    if (tipoConsulta === 'DESCRICAO') resposta404.resultados = [];
    if (encontradoQuando404 !== undefined) resposta404.encontrado = encontradoQuando404;
    return resposta404;
  }

  if (codigoStatus === 429) {
    return { ok: false, httpStatus: 429, erro: 'LIMITE_DE_REQUISICOES' };
  }

  if (codigoStatus >= 500) {
    return { ok: false, httpStatus: codigoStatus, erro: 'SERVICO_COSMOS_INDISPONIVEL' };
  }

  return { ok: false, httpStatus: codigoStatus, erro: 'ERRO_HTTP' };
}

function rfcSanitizarProdutoCosmos_(produto) {
  const seguro = {};

  if (produto && produto.description) seguro.descricao = String(produto.description);
  if (produto && produto.brand && produto.brand.name) seguro.marca = String(produto.brand.name);
  if (produto && produto.gtin) seguro.gtin = String(produto.gtin);
  if (produto && produto.ncm && produto.ncm.code) seguro.ncm = String(produto.ncm.code);

  return seguro;
}

function consultarProdutoCosmosPorDescricaoSeguro(descricao) {
  const consulta = rfcNormalizarConsultaDescricao_(descricao);

  if (!consulta) {
    return { ok: false, erro: 'CONSULTA_INVALIDA' };
  }

  const tokenSeguro = rfcTokenCosmosSeguro_();
  if (!tokenSeguro) {
    return { ok: false, autenticado: false, erro: 'TOKEN_NAO_CONFIGURADO' };
  }

  const url = 'https://api.cosmos.bluesoft.com.br/products?query=' +
    encodeURIComponent(consulta) +
    '&per_page=5';

  try {
    const resposta = UrlFetchApp.fetch(url, rfcOpcoesGetCosmos_(tokenSeguro));
    const codigoStatus = resposta.getResponseCode();

    if (codigoStatus !== 200) {
      const erro = rfcErroCosmosSeguro_(codigoStatus, 'DESCRICAO');
      erro.tipoConsulta = 'DESCRICAO';
      return erro;
    }

    const dados = JSON.parse(resposta.getContentText());
    const resultados = rfcExtrairListaProdutos_(dados)
      .slice(0, 5)
      .map(rfcSanitizarProdutoCosmos_);

    return {
      ok: true,
      consulta: consulta,
      tipoConsulta: 'DESCRICAO',
      httpStatus: codigoStatus,
      quantidadeRetornada: resultados.length,
      resultados: resultados
    };
  } catch (erro) {
    return { ok: false, tipoConsulta: 'DESCRICAO', erro: 'ERRO_CONEXAO' };
  }
}

function consultarProdutoCosmosPorGtinSeguro(gtin) {
  const consulta = rfcNormalizarConsultaGtin_(gtin);

  if (!rfcGtinValido_(consulta)) {
    return { ok: false, erro: 'CONSULTA_INVALIDA' };
  }

  const tokenSeguro = rfcTokenCosmosSeguro_();
  if (!tokenSeguro) {
    return { ok: false, autenticado: false, erro: 'TOKEN_NAO_CONFIGURADO' };
  }

  const url = 'https://api.cosmos.bluesoft.com.br/products?query=' +
    encodeURIComponent(consulta) +
    '&per_page=1';

  try {
    const resposta = UrlFetchApp.fetch(url, rfcOpcoesGetCosmos_(tokenSeguro));
    const codigoStatus = resposta.getResponseCode();

    if (codigoStatus !== 200) {
      const erro = rfcErroCosmosSeguro_(codigoStatus, 'GTIN', false);
      erro.tipoConsulta = 'GTIN';
      return erro;
    }

    const produtos = rfcExtrairListaProdutos_(JSON.parse(resposta.getContentText()));
    const produto = produtos.length ? rfcSanitizarProdutoCosmos_(produtos[0]) : null;

    return {
      ok: true,
      consulta: consulta,
      tipoConsulta: 'GTIN',
      httpStatus: codigoStatus,
      encontrado: Boolean(produto),
      produto: produto || {}
    };
  } catch (erro) {
    return { ok: false, tipoConsulta: 'GTIN', erro: 'ERRO_CONEXAO' };
  }
}

function testarValidacoesConsultaCosmosSemApi() {
  const descricaoVazia = consultarProdutoCosmosPorDescricaoSeguro('');
  const descricaoEspacos = consultarProdutoCosmosPorDescricaoSeguro('   ');
  const gtinPontuado = rfcNormalizarConsultaGtin_('789.000-000/0000');
  const gtinInvalido = consultarProdutoCosmosPorGtinSeguro('123');
  const descricaoAusente = consultarProdutoCosmosPorDescricaoSeguro();
  const gtinAusente = consultarProdutoCosmosPorGtinSeguro();

  return {
    ok: descricaoVazia.erro === 'CONSULTA_INVALIDA' &&
      descricaoEspacos.erro === 'CONSULTA_INVALIDA' &&
      gtinPontuado === '7890000000000' &&
      gtinInvalido.erro === 'CONSULTA_INVALIDA' &&
      descricaoAusente.erro === 'CONSULTA_INVALIDA' &&
      gtinAusente.erro === 'CONSULTA_INVALIDA',
    casos: {
      descricaoVazia: descricaoVazia.erro,
      descricaoSomenteEspacos: descricaoEspacos.erro,
      gtinComPontuacao: gtinPontuado,
      gtinInvalido: gtinInvalido.erro,
      descricaoAusente: descricaoAusente.erro,
      gtinAusente: gtinAusente.erro
    }
  };
}

function testarConexaoCosmosSomenteLeitura() {
  const endpoint = 'products_query';
  const metodo = 'GET';
  const tokenSeguro = rfcTokenCosmosSeguro_();

  if (!tokenSeguro) {
    return {
      ok: false,
      autenticado: false,
      erro: 'TOKEN_NAO_CONFIGURADO'
    };
  }

  const url = 'https://api.cosmos.bluesoft.com.br/products?query=' +
    encodeURIComponent('DIPIRONA') +
    '&per_page=1';

  const opcoes = {
    method: 'get',
    headers: rfcOpcoesGetCosmos_(tokenSeguro).headers,
    muteHttpExceptions: true
  };

  try {
    const resposta = UrlFetchApp.fetch(url, opcoes);
    const codigoStatus = resposta.getResponseCode();

    if (codigoStatus === 200) {
      return {
        ok: true,
        autenticado: true,
        httpStatus: codigoStatus,
        endpoint: endpoint,
        metodo: metodo
      };
    }

    if (codigoStatus === 401 || codigoStatus === 403) {
      return {
        ok: false,
        autenticado: false,
        httpStatus: codigoStatus,
        endpoint: endpoint,
        metodo: metodo,
        erro: 'NAO_AUTORIZADO',
        orientacao: 'VALIDAR_TOKEN_NO_PORTAL_COSMOS'
      };
    }

    return {
      ok: false,
      autenticado: true,
      httpStatus: codigoStatus,
      endpoint: endpoint,
      metodo: metodo,
      erro: 'ERRO_HTTP'
    };
  } catch (erro) {
    return {
      ok: false,
      autenticado: false,
      endpoint: endpoint,
      metodo: metodo,
      erro: 'ERRO_CONEXAO'
    };
  }
}

/************************************************************
 * CONSULTA À API COSMOS POR TEXTO/DESCRIÇÃO
 ************************************************************/

function rfcConsultarCosmosPorDescricao_(descricaoOriginal, token) {
  let tokenSeguro = token ? String(token).trim() : '';

  if ((tokenSeguro.charAt(0) === '"' && tokenSeguro.charAt(tokenSeguro.length - 1) === '"') ||
      (tokenSeguro.charAt(0) === "'" && tokenSeguro.charAt(tokenSeguro.length - 1) === "'")) {
    tokenSeguro = tokenSeguro.substring(1, tokenSeguro.length - 1).trim();
  }

  const url = 'https://api.cosmos.bluesoft.com.br/products?query=' +
    encodeURIComponent(descricaoOriginal) +
    '&per_page=' + RFC_RESULTADOS_POR_BUSCA;

  const opcoes = {
    method: 'get',
    headers: rfcOpcoesGetCosmos_(tokenSeguro).headers,
    muteHttpExceptions: true
  };

  try {
    const resposta = UrlFetchApp.fetch(url, opcoes);
    const codigoStatus = resposta.getResponseCode();

    if (codigoStatus === 401 || codigoStatus === 403) return rfcResultadoVazio_('TOKEN_INVALIDO_OU_EXPIRADO', url);
    if (codigoStatus === 429) return rfcResultadoVazio_('LIMITE_API_ATINGIDO', url);
    if (codigoStatus !== 200) return rfcResultadoVazio_('ERRO_HTTP_' + codigoStatus, url);

    const dados = JSON.parse(resposta.getContentText());
    const candidatos = rfcExtrairListaProdutos_(dados);

    if (!candidatos.length) return rfcResultadoVazio_('NAO_ENCONTRADO', url);

    const melhor = rfcEscolherMelhorCandidato_(descricaoOriginal, candidatos);
    if (!melhor) return rfcResultadoVazio_('NAO_ENCONTRADO', url);

    return {
      status: melhor.confiavel ? 'ENCONTRADO' : 'ENCONTRADO_BAIXA_CONFIANCA',
      descricao: melhor.descricao || '',
      ncm: melhor.ncm || '',
      ean: melhor.ean || '',
      marca: melhor.marca || '',
      url: url,
      data: new Date()
    };
  } catch (erro) {
    return rfcResultadoVazio_('ERRO_CONEXAO', url);
  }
}

function rfcResultadoVazio_(status, url) {
  return { status: status, descricao: '', ncm: '', ean: '', marca: '', url: url, data: new Date() };
}

function rfcExtrairListaProdutos_(dados) {
  if (Array.isArray(dados)) return dados;
  if (dados && Array.isArray(dados.products)) return dados.products;
  if (dados && Array.isArray(dados.data)) return dados.data;
  return [];
}

function rfcEscolherMelhorCandidato_(descricaoOriginal, candidatos) {
  const palavrasOriginal = rfcExtrairPalavrasSignificativas_(descricaoOriginal);
  if (!palavrasOriginal.length) return null;

  const palavraChave = palavrasOriginal[0];
  let melhor = null;
  let melhorScore = -1;

  candidatos.forEach(function(c) {
    const descricaoCandidato = (c && c.description) || '';
    const palavrasCandidato = rfcExtrairPalavrasSignificativas_(descricaoCandidato);
    const temPalavraChave = palavrasCandidato.indexOf(palavraChave) !== -1;

    if (!temPalavraChave) return;

    let score = 0;
    palavrasOriginal.forEach(function(p) {
      if (palavrasCandidato.indexOf(p) !== -1) score++;
    });

    if (score > melhorScore) {
      melhorScore = score;
      melhor = {
        descricao: descricaoCandidato,
        ncm: (c.ncm && c.ncm.code) || '',
        ean: c.gtin ? String(c.gtin) : '',
        marca: (c.brand && c.brand.name) || '',
        confiavel: score >= RFC_MIN_PALAVRAS_COINCIDENTES
      };
    }
  });

  return melhor;
}

function rfcExtrairPalavrasSignificativas_(texto) {
  const normalizado = rfcNormalizarTexto_(texto);
  return normalizado
    .split(/\s+/)
    .filter(Boolean)
    .filter(function(p) { return p.length >= 3; })
    .filter(function(p) { return !/\d/.test(p); })
    .filter(function(p) { return RFC_STOPWORDS.indexOf(p) === -1; });
}

function rfcNcmForaCapituloFarmaceutico_(ncm) {
  const limpo = String(ncm || '').replace(/\D/g, '');
  if (!limpo) return false;
  return limpo.substring(0, 2) !== RFC_PREFIXO_NCM_FARMACEUTICO;
}

/************************************************************
 * ESCRITA SEGURA — só nas colunas de sugestão
 ************************************************************/

function rfcEscreverResultado_(aba, linha, idx, resultado) {
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_DESCRICAO_COSMOS', resultado.descricao);
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_NCM_COSMOS', resultado.ncm);
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_EAN_COSMOS', resultado.ean);
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_MARCA_COSMOS', resultado.marca);
  rfcSetCell_(aba, linha, idx, 'STATUS_CONSULTA_COSMOS', resultado.status);
  rfcSetCell_(aba, linha, idx, 'ALERTA_NCM_FORA_CAPITULO_30', rfcNcmForaCapituloFarmaceutico_(resultado.ncm) ? 'SIM' : '');
  rfcSetCell_(aba, linha, idx, 'FONTE_COSMOS_URL', resultado.url);
  rfcSetCell_(aba, linha, idx, 'DATA_CONSULTA_COSMOS', resultado.data);
}

/************************************************************
 * GARANTIR COLUNAS NOVAS (sem apagar nada existente)
 ************************************************************/

function rfcGarantirColunas_(aba) {
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(String);
  let atualizou = false;

  RFC_COLUNAS_NOVAS.forEach(nome => {
    const existe = headers.some(h => rfcNorm_(h) === rfcNorm_(nome));
    if (!existe) {
      aba.getRange(1, headers.length + 1).setValue(nome);
      headers.push(nome);
      atualizou = true;
    }
  });

  if (atualizou) {
    aba.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#002E56')
      .setFontColor('#FFFFFF');
  }

  return headers;
}

/************************************************************
 * HELPERS
 ************************************************************/

function rfcCriarIndice_(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[String(h).trim()] = i; });
  return idx;
}

function rfcGetCell_(row, idx, nomeColuna) {
  const pos = idx[nomeColuna];
  if (pos === undefined || pos === null || pos < 0) return '';
  return row[pos];
}

function rfcSetCell_(aba, linha, idx, nomeColuna, valor) {
  const pos = idx[nomeColuna];
  if (pos === undefined || pos === null || pos < 0) return;
  aba.getRange(linha, pos + 1).setValue(valor);
}

function rfcUpper_(valor) {
  return String(valor || '').trim().toUpperCase();
}

function rfcLimpar_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim().replace(/\s+/g, ' ');
}

function rfcNorm_(valor) {
  return String(valor || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function rfcNormalizarTexto_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/************************************************************
 * RESET MANUAL (opcional) — zera o ponteiro de retomada
 ************************************************************/

function resetarVarreduraFiscalCosmos() {
  PropertiesService.getScriptProperties().deleteProperty(RFC_PROP_ULTIMA_LINHA);
  SpreadsheetApp.getUi().alert('Ponteiro de retomada zerado. Próxima execução começa do início.');
}
