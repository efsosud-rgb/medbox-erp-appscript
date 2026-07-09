/************************************************************
 * MEDBOX ERP | ROBÔ FISCAL — API COSMOS (BLUESOFT)
 * BUSCA POR DESCRIÇÃO DO PRODUTO (SEM DEPENDER DE EAN)
 *
 * Arquivo: RoboFiscalCosmos.gs
 *
 * NÃO altera Código.gs, RoboFiscal.gs ou qualquer outro arquivo.
 * NÃO sobrescreve SKU_MASTER, NCM, EAN ou CST diretamente.
 * Grava apenas em colunas de SUGESTÃO, criadas automaticamente
 * em BD_PRECOS caso não existam. Revisão e aplicação são manuais.
 *
 * Estratégia:
 * - Lê DESCRICAO_FORNECEDOR (coluna G) de cada linha.
 * - Consulta a API Cosmos por texto: /products?query=...
 * - Compara palavras da descrição original com a do resultado,
 *   IGNORANDO tokens numéricos/dosagem (ex.: "3ML", "0,9%"),
 *   que geram falso positivo entre produtos diferentes do mesmo
 *   fabricante/volume.
 * - EXIGE que a primeira palavra significativa da descrição
 *   original (normalmente o nome do produto/princípio ativo)
 *   apareça no candidato. Sem isso, nunca marca ENCONTRADO,
 *   mesmo que outras palavras (marca, embalagem) batam.
 *
 * Processa em lotes (respeita limite de 6 min do Apps Script) e
 * retoma de onde parou usando PropertiesService.
 ************************************************************/

const RFC_ABA_BD_PRECOS = 'BD_PRECOS';

const RFC_COSMOS_TOKEN = '-2xvD0ybVF4HdsR8MDVDsQ';
const RFC_COSMOS_USER_AGENT = 'Cosmos-API-Request';

const RFC_LIMITE_TEMPO_MS = 4.5 * 60 * 1000; // margem de segurança (limite real é 6 min)
const RFC_PAUSA_MS = 1200;
const RFC_PROP_ULTIMA_LINHA = 'RFC_ULTIMA_LINHA_PROCESSADA';
const RFC_RESULTADOS_POR_BUSCA = 5;
const RFC_MIN_PALAVRAS_COINCIDENTES = 2; // mínimo de palavras (sem contar a 1ª obrigatória)

const RFC_COLUNAS_NOVAS = [
  'SUGESTAO_DESCRICAO_COSMOS',
  'SUGESTAO_NCM_COSMOS',
  'SUGESTAO_EAN_COSMOS',
  'SUGESTAO_MARCA_COSMOS',
  'STATUS_CONSULTA_COSMOS',
  'FONTE_COSMOS_URL',
  'DATA_CONSULTA_COSMOS'
];

// Palavras muito comuns que não ajudam a validar o match.
const RFC_STOPWORDS = [
  'DE', 'DA', 'DO', 'DAS', 'DOS', 'COM', 'PARA', 'EM', 'E', 'OU',
  'UN', 'UND', 'UNID', 'CX', 'CAIXA', 'FR', 'FRASCO', 'AMP', 'AMPOLA',
  'AMPOLAS', 'ML', 'MG', 'MCG', 'UI', 'CADA', 'INDIVIDUAL', 'EMBALADO',
  'PC', 'PCT'
];

/************************************************************
 * FUNÇÃO PRINCIPAL — chamada pelo menu
 ************************************************************/

function executarLoteVarreduraFiscalCosmos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(RFC_ABA_BD_PRECOS);

  if (!aba) {
    SpreadsheetApp.getUi().alert("Aba 'BD_PRECOS' não encontrada.");
    return;
  }

  if (!RFC_COSMOS_TOKEN) {
    SpreadsheetApp.getUi().alert(
      'Configure RFC_COSMOS_TOKEN e RFC_COSMOS_USER_AGENT no início do arquivo ' +
      'RoboFiscalCosmos.gs antes de rodar.'
    );
    return;
  }

  const headers = rfcGarantirColunas_(aba);
  const idx = rfcCriarIndice_(headers);

  const inicio = Date.now();
  const props = PropertiesService.getScriptProperties();
  let linhaAtual = Number(props.getProperty(RFC_PROP_ULTIMA_LINHA)) || 1; // 1 = cabeçalho
  linhaAtual = Math.max(linhaAtual + 1, 2);

  const ultimaLinha = aba.getLastRow();

  if (linhaAtual > ultimaLinha) {
    linhaAtual = 2; // já passou por tudo pelo menos uma vez; reinicia
  }

  let processadas = 0;
  let encontradas = 0;
  let baixaConfianca = 0;
  let semResultado = 0;
  let semDescricao = 0;
  let jaProcessadasPuladas = 0;

  for (let linha = linhaAtual; linha <= ultimaLinha; linha++) {

    if (Date.now() - inicio > RFC_LIMITE_TEMPO_MS) {
      props.setProperty(RFC_PROP_ULTIMA_LINHA, String(linha - 1));

      SpreadsheetApp.getUi().alert(
        'Lote parcial concluído (limite de tempo do Apps Script).\n\n' +
        'Processadas nesta execução: ' + processadas + '\n' +
        'Encontradas (alta confiança): ' + encontradas + '\n' +
        'Encontradas (baixa confiança): ' + baixaConfianca + '\n' +
        'Sem resultado: ' + semResultado + '\n' +
        'Sem descrição: ' + semDescricao + '\n' +
        'Já processadas antes (puladas): ' + jaProcessadasPuladas + '\n\n' +
        'Rode "Buscar Dados Fiscais (Cosmos)" novamente para continuar de onde parou.'
      );
      return;
    }

    const linhaValores = aba.getRange(linha, 1, 1, headers.length).getValues()[0];

    const statusAtual = rfcUpper_(rfcGetCell_(linhaValores, idx, 'STATUS_CONSULTA_COSMOS'));

    if (statusAtual) {
      jaProcessadasPuladas++;
      continue;
    }

    const descricao = rfcLimpar_(rfcGetCell_(linhaValores, idx, 'DESCRICAO_FORNECEDOR'));

    if (!descricao || descricao.length < 5) {
      rfcEscreverResultado_(aba, linha, idx, {
        status: 'SEM_DESCRICAO',
        descricao: '', ncm: '', ean: '', marca: '', url: '', data: new Date()
      });
      semDescricao++;
      continue;
    }

    const resultado = rfcConsultarCosmosPorDescricao_(descricao);

    rfcEscreverResultado_(aba, linha, idx, resultado);

    processadas++;
    if (resultado.status === 'ENCONTRADO') encontradas++;
    if (resultado.status === 'ENCONTRADO_BAIXA_CONFIANCA') baixaConfianca++;
    if (resultado.status === 'NAO_ENCONTRADO') semResultado++;

    Utilities.sleep(RFC_PAUSA_MS);
  }

  props.deleteProperty(RFC_PROP_ULTIMA_LINHA);

  SpreadsheetApp.getUi().alert(
    'Varredura fiscal concluída (BD_PRECOS inteira percorrida).\n\n' +
    'Processadas: ' + processadas + '\n' +
    'Encontradas (alta confiança): ' + encontradas + '\n' +
    'Encontradas (baixa confiança): ' + baixaConfianca + '\n' +
    'Sem resultado: ' + semResultado + '\n' +
    'Sem descrição: ' + semDescricao + '\n' +
    'Já processadas antes (puladas): ' + jaProcessadasPuladas + '\n\n' +
    'Revise as colunas SUGESTAO_* (especialmente as marcadas BAIXA_CONFIANCA) ' +
    'antes de aplicar em SKU_MASTER/NCM/EAN/CST.'
  );
}

/************************************************************
 * DIAGNÓSTICO — conta quantas linhas de BD_PRECOS têm
 * DESCRICAO_FORNECEDOR preenchida.
 ************************************************************/

function diagnosticarPreenchimentoDescricaoCosmos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(RFC_ABA_BD_PRECOS);

  if (!aba) {
    Logger.log("Aba 'BD_PRECOS' não encontrada.");
    return;
  }

  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(String);
  const idx = rfcCriarIndice_(headers);

  if (idx['DESCRICAO_FORNECEDOR'] === undefined) {
    Logger.log('Coluna DESCRICAO_FORNECEDOR não encontrada nos cabeçalhos atuais: ' + headers.join(' | '));
    return;
  }

  const ultimaLinha = aba.getLastRow();
  const valores = aba.getRange(2, 1, Math.max(ultimaLinha - 1, 0), headers.length).getValues();

  let comDescricao = 0;
  let semDescricao = 0;
  const exemplos = [];

  valores.forEach((row, i) => {
    const descricao = rfcLimpar_(rfcGetCell_(row, idx, 'DESCRICAO_FORNECEDOR'));

    if (descricao && descricao.length >= 5) {
      comDescricao++;
      if (exemplos.length < 5) {
        exemplos.push('Linha ' + (i + 2) + ': "' + descricao + '"');
      }
    } else {
      semDescricao++;
    }
  });

  Logger.log('Total de linhas de dados: ' + valores.length);
  Logger.log('Com descrição válida: ' + comDescricao);
  Logger.log('Sem descrição: ' + semDescricao);
  Logger.log('Exemplos:\n' + (exemplos.join('\n') || 'nenhum'));
}

/************************************************************
 * TESTE COM ATÉ 5 LINHAS QUE TENHAM DESCRIÇÃO PREENCHIDA
 ************************************************************/

function testarVarreduraFiscalCosmos5Linhas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(RFC_ABA_BD_PRECOS);

  if (!aba) {
    Logger.log("Aba 'BD_PRECOS' não encontrada.");
    return;
  }

  const headers = rfcGarantirColunas_(aba);
  const idx = rfcCriarIndice_(headers);

  const ultimaLinha = aba.getLastRow();
  let testadas = 0;
  const resumo = [];

  for (let linha = 2; linha <= ultimaLinha && testadas < 5; linha++) {
    const linhaValores = aba.getRange(linha, 1, 1, headers.length).getValues()[0];
    const descricao = rfcLimpar_(rfcGetCell_(linhaValores, idx, 'DESCRICAO_FORNECEDOR'));

    if (!descricao || descricao.length < 5) continue;

    const resultado = rfcConsultarCosmosPorDescricao_(descricao);
    rfcEscreverResultado_(aba, linha, idx, resultado);
    testadas++;

    resumo.push(
      'Linha ' + linha + ' | Original: "' + descricao + '"' +
      ' | Status: ' + resultado.status +
      ' | Cosmos: "' + resultado.descricao + '"' +
      ' | NCM: ' + resultado.ncm +
      ' | EAN: ' + resultado.ean
    );

    Utilities.sleep(RFC_PAUSA_MS);
  }

  if (testadas === 0) {
    Logger.log(
      'Nenhuma linha com DESCRICAO_FORNECEDOR válida foi encontrada. ' +
      'Rode diagnosticarPreenchimentoDescricaoCosmos() para confirmar.'
    );
    return;
  }

  Logger.log('Teste concluído. Linhas testadas: ' + testadas);
  Logger.log(resumo.join('\n'));
}

/************************************************************
 * CONSULTA À API COSMOS POR TEXTO/DESCRIÇÃO
 ************************************************************/

function rfcConsultarCosmosPorDescricao_(descricaoOriginal) {
  const url = 'https://api.cosmos.bluesoft.com.br/products?query=' +
    encodeURIComponent(descricaoOriginal) +
    '&per_page=' + RFC_RESULTADOS_POR_BUSCA;

  const opcoes = {
    method: 'get',
    headers: {
      'X-Cosmos-Token': RFC_COSMOS_TOKEN,
      'User-Agent': RFC_COSMOS_USER_AGENT,
      'Content-Type': 'application/json;charset=UTF-8'
    },
    muteHttpExceptions: true
  };

  try {
    const resposta = UrlFetchApp.fetch(url, opcoes);
    const codigoStatus = resposta.getResponseCode();

    if (codigoStatus === 401 || codigoStatus === 403) {
      return rfcResultadoVazio_('TOKEN_INVALIDO_OU_EXPIRADO', url);
    }

    if (codigoStatus === 429) {
      return rfcResultadoVazio_('LIMITE_API_ATINGIDO', url);
    }

    if (codigoStatus !== 200) {
      return rfcResultadoVazio_('ERRO_HTTP_' + codigoStatus, url);
    }

    const dados = JSON.parse(resposta.getContentText());
    const candidatos = rfcExtrairListaProdutos_(dados);

    if (!candidatos.length) {
      return rfcResultadoVazio_('NAO_ENCONTRADO', url);
    }

    const melhor = rfcEscolherMelhorCandidato_(descricaoOriginal, candidatos);

    if (!melhor) {
      return rfcResultadoVazio_('NAO_ENCONTRADO', url);
    }

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

/**
 * A resposta de /products pode vir como array direto ou dentro
 * de uma propriedade (ex.: "products"). Trata os dois formatos
 * sem assumir estrutura fixa.
 */
function rfcExtrairListaProdutos_(dados) {
  if (Array.isArray(dados)) return dados;
  if (dados && Array.isArray(dados.products)) return dados.products;
  if (dados && Array.isArray(dados.data)) return dados.data;
  return [];
}

/**
 * Compara a descrição original com cada candidato.
 *
 * REGRA DE SEGURANÇA (evita cruzar produtos diferentes, ex.:
 * Acetilcisteína marcado como Diclofenaco só por bater marca
 * e volume): a PRIMEIRA palavra significativa da descrição
 * original — normalmente o nome do produto/princípio ativo —
 * precisa aparecer no candidato. Sem isso, o candidato é
 * descartado, mesmo que outras palavras batam.
 *
 * Tokens que contêm dígitos (dosagem, volume, concentração,
 * ex.: "3ML", "0,9%", "25MG") NÃO contam para a pontuação,
 * porque múltiplos produtos diferentes compartilham a mesma
 * dosagem/volume.
 */
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

    if (!temPalavraChave) {
      return; // descarta candidato: nome do produto não bate
    }

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
    .filter(function(p) { return !/\d/.test(p); }) // remove tokens com dígito (dosagem/volume)
    .filter(function(p) { return RFC_STOPWORDS.indexOf(p) === -1; });
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
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
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