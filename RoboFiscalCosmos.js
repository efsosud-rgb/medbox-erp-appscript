/************************************************************
 * MEDBOX ERP | ROBÔ FISCAL — API COSMOS (BLUESOFT)
 * BUSCA POR DESCRIÇÃO DO PRODUTO (SEM DEPENDER DE EAN)
 *
 * Arquivo: RoboFiscalCosmos.js
 *
 * NÃO altera Código.gs, RoboFiscal.gs ou qualquer outro arquivo.
 * NÃO sobrescreve SKU_MASTER, NCM, EAN ou CST diretamente se já preenchidos.
 * Grava em colunas de SUGESTÃO e atualiza NCM/EAN oficiais quando vazios.
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
 * Processa em lotes e permite acionamento automático por trigger.
 ************************************************************/

const RFC_ABA_BD_PRECOS = 'BD_PRECOS';

const RFC_COSMOS_TOKEN = '-2xvD0ybVF4HdsR8MDVDsQ';
const RFC_COSMOS_USER_AGENT = 'Cosmos-API-Request';

const RFC_LIMITE_TEMPO_MS = 4.5 * 60 * 1000;
const RFC_PAUSA_MS = 1200;
const RFC_PROP_ULTIMA_LINHA = 'RFC_ULTIMA_LINHA_PROCESSADA';
const RFC_RESULTADOS_POR_BUSCA = 5;
const RFC_MIN_PALAVRAS_COINCIDENTES = 2;

const RFC_COLUNAS_NOVAS = [
  'SUGESTAO_DESCRICAO_COSMOS',
  'SUGESTAO_NCM_COSMOS',
  'SUGESTAO_EAN_COSMOS',
  'SUGESTAO_MARCA_COSMOS',
  'STATUS_CONSULTA_COSMOS',
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
 * PROCESSAMENTO EM LOTE
 ************************************************************/

/**
 * Processa um lote de produtos da BD_PRECOS na API Cosmos.
 * @param {number|Object} tamanhoLoteOuEvento Quantidade de linhas ou objeto de evento do trigger.
 */
function processarLoteCosmos(tamanhoLoteOuEvento) {
  // Trata o caso de ser chamado por trigger (onde o primeiro arg é o objeto de evento)
  let tamanhoLote = 50;
  if (typeof tamanhoLoteOuEvento === 'number') {
    tamanhoLote = tamanhoLoteOuEvento;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(RFC_ABA_BD_PRECOS);

  if (!aba) {
    Logger.log("Aba 'BD_PRECOS' não encontrada.");
    return;
  }

  const headers = rfcGarantirColunas_(aba);
  const idx = rfcCriarIndice_(headers);

  const inicio = Date.now();
  const ultimaLinha = aba.getLastRow();

  let processados = 0;

  for (let linha = 2; linha <= ultimaLinha && processados < tamanhoLote; linha++) {
    if (Date.now() - inicio > RFC_LIMITE_TEMPO_MS) break;

    const rowRange = aba.getRange(linha, 1, 1, headers.length);
    const rowValues = rowRange.getValues()[0];

    const status = rfcUpper_(rfcGetCell_(rowValues, idx, 'STATUS_CONSULTA_COSMOS'));

    // Pula se já processado com sucesso ou não encontrado
    if (status === 'ENCONTRADO' || status === 'NAO_ENCONTRADO') continue;

    const ncmOficial = rfcGetCell_(rowValues, idx, 'NCM');
    const eanOficial = rfcGetCell_(rowValues, idx, 'EAN');

    // Só processa se NCM ou EAN estiverem vazios
    if (ncmOficial && eanOficial) continue;

    const descricao = rfcLimpar_(rfcGetCell_(rowValues, idx, 'DESCRICAO_FORNECEDOR'));

    if (!descricao || descricao.length < 5) {
      rfcEscreverResultado_(aba, linha, idx, {
        status: 'SEM_DESCRICAO',
        descricao: '', ncm: '', ean: '', marca: '', url: '', data: new Date()
      });
      continue;
    }

    try {
      const resultado = rfcConsultarCosmosPorDescricao_(descricao);
      rfcEscreverResultado_(aba, linha, idx, resultado);

      // Se encontrado, preenche as colunas oficiais caso estejam vazias
      if (resultado.status === 'ENCONTRADO') {
        if (!ncmOficial && resultado.ncm) {
          aba.getRange(linha, idx['NCM'] + 1).setValue(resultado.ncm);
        }
        if (!eanOficial && resultado.ean) {
          aba.getRange(linha, idx['EAN'] + 1).setValue(resultado.ean);
        }
      }

      processados++;
      Utilities.sleep(RFC_PAUSA_MS);

    } catch (e) {
      Logger.log('Erro na linha ' + linha + ': ' + e.message);
      rfcSetCell_(aba, linha, idx, 'STATUS_CONSULTA_COSMOS', 'ERRO');
      continue;
    }
  }
}

/************************************************************
 * TRIGGER AUTOMÁTICO
 ************************************************************/

/**
 * Trigger disparado por alterações na planilha.
 * Detecta adições na BD_PRECOS e agenda o processamento.
 */
function rfcTriggerOnChange(e) {
  if (!e) return;

  const tipo = e.changeType;

  if (tipo === 'INSERT_ROW' || tipo === 'EDIT' || tipo === 'OTHER') {
    const triggers = ScriptApp.getProjectTriggers();
    const jaTemTrigger = triggers.some(t => t.getHandlerFunction() === 'processarLoteCosmos');

    if (!jaTemTrigger) {
      ScriptApp.newTrigger('processarLoteCosmos')
        .timeBased()
        .after(30 * 1000)
        .create();
    }
  }
}

/************************************************************
 * FUNÇÕES DE VARREDURA E TESTE
 ************************************************************/

function executarLoteVarreduraFiscalCosmos() {
  processarLoteCosmos(50);
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

    if (codigoStatus === 401 || codigoStatus === 403) return rfcResultadoVazio_('TOKEN_INVALIDO', url);
    if (codigoStatus === 429) return rfcResultadoVazio_('LIMITE_API', url);
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

function rfcEscreverResultado_(aba, linha, idx, resultado) {
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_DESCRICAO_COSMOS', resultado.descricao);
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_NCM_COSMOS', resultado.ncm);
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_EAN_COSMOS', resultado.ean);
  rfcSetCell_(aba, linha, idx, 'SUGESTAO_MARCA_COSMOS', resultado.marca);
  rfcSetCell_(aba, linha, idx, 'STATUS_CONSULTA_COSMOS', resultado.status);
  rfcSetCell_(aba, linha, idx, 'FONTE_COSMOS_URL', resultado.url);
  rfcSetCell_(aba, linha, idx, 'DATA_CONSULTA_COSMOS', resultado.data);
}

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

function resetarVarreduraFiscalCosmos() {
  PropertiesService.getScriptProperties().deleteProperty(RFC_PROP_ULTIMA_LINHA);
  SpreadsheetApp.getUi().alert('Ponteiro de retomada zerado.');
}
