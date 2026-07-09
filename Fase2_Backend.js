/************************************************************
 * MEDBOX ERP | FASE 2B - BACKEND
 *
 * Funções de lançamento e simulação:
 * - Salvar nova cotação em BD_COTACOES
 * - Salvar preço avulso/tabela em BD_PRECOS
 * - Buscar preço de custo por SKU Master, SKU Operacional ou descrição
 * - Simular preço MedBox com markup, frete e impostos
 *
 * Não apaga dados existentes.
 * Não altera BD_ antigas.
 ************************************************************/

const F2_ABA_BD_COTACOES = 'BD_COTACOES';
const F2_ABA_BD_PRECOS = 'BD_PRECOS';
const F2_ABA_BD_SKU_SEFAZ = 'BD_SKU_MASTER_SEFAZ';
const F2_ABA_MAPA = 'MAPA_ERP';

/************************************************************
 * OPÇÕES PARA TELAS DE LANÇAMENTO
 ************************************************************/

function obterOpcoesFase2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mapa = ss.getSheetByName(F2_ABA_MAPA);
  const bdPrecos = ss.getSheetByName(F2_ABA_BD_PRECOS);
  const bdSku = ss.getSheetByName(F2_ABA_BD_SKU_SEFAZ);

  const clientesSet = new Set();
  const categoriasSet = new Set();
  const skuSet = new Set();
  const fornecedoresSet = new Set();
  const fabricantesSet = new Set();

  if (mapa && mapa.getLastRow() > 1) {
    const values = mapa.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = f2CriarIndice_(headers);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const categoria = f2Upper_(f2GetCell_(row, idx, 'Categoria'));
      const skuMaster = f2Upper_(f2GetCell_(row, idx, 'SKU Master'));
      const fornecedor = f2Upper_(f2GetCell_(row, idx, 'Fornecedor Compra 2026'));
      const fabricante = f2Upper_(f2GetCell_(row, idx, 'Fabricante Compra 2026'));
      const clientes = f2Upper_(f2GetCell_(row, idx, 'Clientes com Consumo'));

      if (categoria) categoriasSet.add(categoria);
      if (skuMaster) skuSet.add(skuMaster);
      if (fornecedor) fornecedoresSet.add(fornecedor);
      if (fabricante) fabricantesSet.add(fabricante);

      if (clientes) {
        clientes.split(/[;,]/).forEach(c => {
          const cliente = f2Upper_(c);
          if (cliente) clientesSet.add(cliente);
        });
      }
    }
  }

  if (bdPrecos && bdPrecos.getLastRow() > 1) {
    const values = bdPrecos.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = f2CriarIndice_(headers);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const fornecedor = f2Upper_(f2GetCell_(row, idx, 'FORNECEDOR'));
      const fabricante = f2Upper_(f2GetCell_(row, idx, 'FABRICANTE_MARCA'));
      const skuMaster = f2Upper_(f2GetCell_(row, idx, 'SKU_MASTER'));
      const skuOperacional = f2Upper_(f2GetCell_(row, idx, 'SKU_OPERACIONAL'));

      if (fornecedor) fornecedoresSet.add(fornecedor);
      if (fabricante) fabricantesSet.add(fabricante);
      if (skuMaster) skuSet.add(skuMaster);
      if (skuOperacional) skuSet.add(skuOperacional);
    }
  }

  if (bdSku && bdSku.getLastRow() > 1) {
    const values = bdSku.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = f2CriarIndice_(headers);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const skuMaster = f2Upper_(f2GetCell_(row, idx, 'SKU_MASTER'));
      const categoria = f2Upper_(f2GetCell_(row, idx, 'CATEGORIA'));
      const fabricante = f2Upper_(f2GetCell_(row, idx, 'MARCA_FABRICANTE'));

      if (skuMaster) skuSet.add(skuMaster);
      if (categoria) categoriasSet.add(categoria);
      if (fabricante) fabricantesSet.add(fabricante);
    }
  }

  return {
    atualizadoEm: f2FormatarDataHora_(new Date()),
    clientes: Array.from(clientesSet).sort(),
    categorias: Array.from(categoriasSet).sort(),
    skuMasters: Array.from(skuSet).sort().slice(0, 1000),
    fornecedores: Array.from(fornecedoresSet).sort().slice(0, 1000),
    fabricantes: Array.from(fabricantesSet).sort().slice(0, 1000)
  };
}

/************************************************************
 * SALVAR NOVA COTAÇÃO
 ************************************************************/

function salvarCotacaoFase2(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(F2_ABA_BD_COTACOES);

  if (!ws) {
    throw new Error('Aba BD_COTACOES não encontrada. Execute primeiro a Fase 2A.');
  }

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0].map(String);
  const idx = f2CriarIndice_(headers);

  const agora = new Date();

  const idCotacao = f2GerarId_('COT');

  const cliente = f2Upper_(dados.cliente);
  const responsavel = f2Upper_(dados.responsavel);
  const categoria = f2Upper_(dados.categoria);
  const descricaoCliente = f2Upper_(dados.descricaoCliente);
  const descricaoNormalizada = f2NormalizarTexto_(descricaoCliente);
  const skuMaster = f2Upper_(dados.skuMaster);
  const skuOperacional = f2Upper_(dados.skuOperacional || skuMaster || descricaoNormalizada);
  const marcaFabricante = f2Upper_(dados.marcaFabricante);
  const quantidade = f2ToNumber_(dados.quantidade);
  const unidade = f2Upper_(dados.unidade);
  const precoReferencia = f2ToNumber_(dados.precoReferenciaCliente);
  const markup = f2NormalizarPercentual_(dados.markupPercentual);

  const termoBuscaCusto = skuMaster || skuOperacional || descricaoNormalizada;
  const precoCusto = buscarMenorPrecoCustoSku_(termoBuscaCusto);

  const precoMedbox = precoCusto !== null ? f2Arredondar_(precoCusto * (1 + markup), 4) : '';
  const valorTotalReferencia = precoReferencia !== null && quantidade !== null
    ? f2Arredondar_(precoReferencia * quantidade, 4)
    : '';

  const statusComercial = f2MontarStatusComercial_(precoReferencia, precoMedbox);
  const statusCotacao = f2Upper_(dados.statusCotacao || 'NOVA');
  const origem = f2Upper_(dados.origem || 'LANÇAMENTO MANUAL');
  const observacao = f2Upper_(dados.observacao);
  const usuario = Session.getActiveUser().getEmail() || '';

  const novaLinha = new Array(headers.length).fill('');

  f2SetByHeader_(novaLinha, idx, 'ID_COTACAO', idCotacao);
  f2SetByHeader_(novaLinha, idx, 'DATA_LANCAMENTO', agora);
  f2SetByHeader_(novaLinha, idx, 'DATA_COTACAO', dados.dataCotacao || agora);
  f2SetByHeader_(novaLinha, idx, 'CLIENTE', cliente);
  f2SetByHeader_(novaLinha, idx, 'RESPONSAVEL', responsavel);
  f2SetByHeader_(novaLinha, idx, 'CATEGORIA', categoria);
  f2SetByHeader_(novaLinha, idx, 'DESCRICAO_CLIENTE', descricaoCliente);
  f2SetByHeader_(novaLinha, idx, 'DESCRICAO_NORMALIZADA', descricaoNormalizada);
  f2SetByHeader_(novaLinha, idx, 'SKU_MASTER', skuMaster);
  f2SetByHeader_(novaLinha, idx, 'SKU_OPERACIONAL', skuOperacional);
  f2SetByHeader_(novaLinha, idx, 'MARCA_FABRICANTE', marcaFabricante);
  f2SetByHeader_(novaLinha, idx, 'QUANTIDADE', quantidade);
  f2SetByHeader_(novaLinha, idx, 'UNIDADE', unidade);
  f2SetByHeader_(novaLinha, idx, 'PRECO_REFERENCIA_CLIENTE', precoReferencia);
  f2SetByHeader_(novaLinha, idx, 'VALOR_TOTAL_REFERENCIA', valorTotalReferencia);
  f2SetByHeader_(novaLinha, idx, 'PRECO_CUSTO_MEDBOX', precoCusto);
  f2SetByHeader_(novaLinha, idx, 'PRECO_MEDBOX', precoMedbox);
  f2SetByHeader_(novaLinha, idx, 'MARKUP_PERCENTUAL', markup);
  f2SetByHeader_(novaLinha, idx, 'STATUS_COMERCIAL', statusComercial);
  f2SetByHeader_(novaLinha, idx, 'STATUS_COTACAO', statusCotacao);
  f2SetByHeader_(novaLinha, idx, 'ORIGEM', origem);
  f2SetByHeader_(novaLinha, idx, 'OBSERVACAO', observacao);
  f2SetByHeader_(novaLinha, idx, 'USUARIO', usuario);
  f2SetByHeader_(novaLinha, idx, 'DATA_ATUALIZACAO', agora);

  ws.appendRow(novaLinha);

  f2RegistrarLog_('COTAÇÃO SALVA: ' + idCotacao + ' | ' + cliente + ' | ' + descricaoCliente);

  return {
    ok: true,
    idCotacao,
    mensagem: 'Cotação salva com sucesso.',
    precoCusto,
    precoMedbox,
    statusComercial
  };
}

/************************************************************
 * SALVAR PREÇO AVULSO / TABELA
 ************************************************************/

function salvarPrecoFase2(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(F2_ABA_BD_PRECOS);

  if (!ws) {
    throw new Error('Aba BD_PRECOS não encontrada. Execute primeiro a Fase 2A.');
  }

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0].map(String);
  const idx = f2CriarIndice_(headers);

  const agora = new Date();

  const idPreco = f2GerarId_('PRE');

  const fornecedor = f2Upper_(dados.fornecedor);
  const fabricante = f2Upper_(dados.fabricanteMarca);
  const descricaoFornecedor = f2Upper_(dados.descricaoFornecedor);
  const descricaoNormalizada = f2NormalizarTexto_(descricaoFornecedor);
  const skuMaster = f2Upper_(dados.skuMaster);
  const skuOperacional = f2Upper_(dados.skuOperacional || skuMaster || descricaoNormalizada);

  const ncm = f2SomenteTexto_(dados.ncm);
  const ean = f2SomenteTexto_(dados.ean);
  const cst = f2SomenteTexto_(dados.cst);
  const unidade = f2Upper_(dados.unidade);

  const precoCusto = f2ToNumber_(dados.precoCusto);
  const ipi = f2NormalizarPercentual_(dados.ipiPercentual);
  const icms = f2NormalizarPercentual_(dados.icmsPercentual);
  const frete = f2NormalizarPercentual_(dados.fretePercentual);
  const outras = f2NormalizarPercentual_(dados.outrasDespesasPercentual);

  const custoTotalEstimado = precoCusto !== null
    ? f2Arredondar_(precoCusto * (1 + ipi + icms + frete + outras), 4)
    : '';

  const origem = f2Upper_(dados.origem || 'LANÇAMENTO AVULSO');
  const arquivoOrigem = f2Upper_(dados.arquivoOrigem);
  const observacao = f2Upper_(dados.observacao);
  const usuario = Session.getActiveUser().getEmail() || '';

  const novaLinha = new Array(headers.length).fill('');

  f2SetByHeader_(novaLinha, idx, 'ID_PRECO', idPreco);
  f2SetByHeader_(novaLinha, idx, 'DATA_LANCAMENTO', agora);
  f2SetByHeader_(novaLinha, idx, 'DATA_TABELA', dados.dataTabela || agora);
  f2SetByHeader_(novaLinha, idx, 'VALIDADE_TABELA', dados.validadeTabela || '');
  f2SetByHeader_(novaLinha, idx, 'FORNECEDOR', fornecedor);
  f2SetByHeader_(novaLinha, idx, 'FABRICANTE_MARCA', fabricante);
  f2SetByHeader_(novaLinha, idx, 'DESCRICAO_FORNECEDOR', descricaoFornecedor);
  f2SetByHeader_(novaLinha, idx, 'DESCRICAO_NORMALIZADA', descricaoNormalizada);
  f2SetByHeader_(novaLinha, idx, 'SKU_MASTER', skuMaster);
  f2SetByHeader_(novaLinha, idx, 'SKU_OPERACIONAL', skuOperacional);
  f2SetByHeader_(novaLinha, idx, 'NCM', ncm);
  f2SetByHeader_(novaLinha, idx, 'EAN', ean);
  f2SetByHeader_(novaLinha, idx, 'CST', cst);
  f2SetByHeader_(novaLinha, idx, 'UNIDADE', unidade);
  f2SetByHeader_(novaLinha, idx, 'PRECO_CUSTO', precoCusto);
  f2SetByHeader_(novaLinha, idx, 'IPI_PERCENTUAL', ipi);
  f2SetByHeader_(novaLinha, idx, 'ICMS_PERCENTUAL', icms);
  f2SetByHeader_(novaLinha, idx, 'FRETE_PERCENTUAL', frete);
  f2SetByHeader_(novaLinha, idx, 'OUTRAS_DESPESAS_PERCENTUAL', outras);
  f2SetByHeader_(novaLinha, idx, 'CUSTO_TOTAL_ESTIMADO', custoTotalEstimado);
  f2SetByHeader_(novaLinha, idx, 'ORIGEM', origem);
  f2SetByHeader_(novaLinha, idx, 'ARQUIVO_ORIGEM', arquivoOrigem);
  f2SetByHeader_(novaLinha, idx, 'OBSERVACAO', observacao);
  f2SetByHeader_(novaLinha, idx, 'USUARIO', usuario);
  f2SetByHeader_(novaLinha, idx, 'DATA_ATUALIZACAO', agora);

  ws.appendRow(novaLinha);

  f2RegistrarLog_('PREÇO SALVO: ' + idPreco + ' | ' + fornecedor + ' | ' + descricaoFornecedor);

  return {
    ok: true,
    idPreco,
    mensagem: 'Preço salvo com sucesso.',
    custoTotalEstimado
  };
}

/************************************************************
 * BUSCAR MENOR PREÇO DE CUSTO
 *
 * Busca por:
 * 1. SKU_MASTER
 * 2. SKU_OPERACIONAL
 * 3. DESCRICAO_NORMALIZADA
 * 4. DESCRICAO_FORNECEDOR
 * 5. Produto/SKU operacional do MAPA_ERP
 *
 * Usa CUSTO_TOTAL_ESTIMADO quando existir.
 ************************************************************/

function buscarMenorPrecoCustoSku_(skuMaster) {
  const termoBusca = f2NormalizarTexto_(skuMaster);

  if (!termoBusca) return null;

  let menor = null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /************************************************************
   * 1. BUSCAR NA BD_PRECOS
   ************************************************************/

  const bdPrecos = ss.getSheetByName(F2_ABA_BD_PRECOS);

  if (bdPrecos && bdPrecos.getLastRow() > 1) {
    const values = bdPrecos.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = f2CriarIndice_(headers);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const skuMasterRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'SKU_MASTER'));
      const skuOperacionalRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'SKU_OPERACIONAL'));
      const descricaoNormalizadaRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'DESCRICAO_NORMALIZADA'));
      const descricaoFornecedorRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'DESCRICAO_FORNECEDOR'));

      const precoCusto = f2ToNumber_(f2GetCell_(row, idx, 'PRECO_CUSTO'));
      const custoTotalEstimado = f2ToNumber_(f2GetCell_(row, idx, 'CUSTO_TOTAL_ESTIMADO'));

      const custoUsado = custoTotalEstimado !== null
        ? custoTotalEstimado
        : precoCusto;

      if (custoUsado === null) continue;

      const encontrou =
        skuMasterRow === termoBusca ||
        skuOperacionalRow === termoBusca ||
        descricaoNormalizadaRow === termoBusca ||
        descricaoFornecedorRow === termoBusca ||
        f2TextoContemTodosTermos_(skuMasterRow, termoBusca) ||
        f2TextoContemTodosTermos_(skuOperacionalRow, termoBusca) ||
        f2TextoContemTodosTermos_(descricaoNormalizadaRow, termoBusca) ||
        f2TextoContemTodosTermos_(descricaoFornecedorRow, termoBusca);

      if (encontrou) {
        if (menor === null || custoUsado < menor) {
          menor = custoUsado;
        }
      }
    }
  }

  /************************************************************
   * 2. BUSCAR NA MAPA_ERP
   ************************************************************/

  const mapa = ss.getSheetByName(F2_ABA_MAPA);

  if (mapa && mapa.getLastRow() > 1) {
    const values = mapa.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = f2CriarIndice_(headers);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const skuMasterRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'SKU Master'));
      const skuOperacionalRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'SKU Operacional'));
      const produtoNormalizadoRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'Produto Normalizado'));
      const produtoRow = f2NormalizarTexto_(f2GetCell_(row, idx, 'Produto'));

      const preco = f2ToNumber_(f2GetCell_(row, idx, 'Custo Compra 2026'));

      if (preco === null) continue;

      const encontrou =
        skuMasterRow === termoBusca ||
        skuOperacionalRow === termoBusca ||
        produtoNormalizadoRow === termoBusca ||
        produtoRow === termoBusca ||
        f2TextoContemTodosTermos_(skuMasterRow, termoBusca) ||
        f2TextoContemTodosTermos_(skuOperacionalRow, termoBusca) ||
        f2TextoContemTodosTermos_(produtoNormalizadoRow, termoBusca) ||
        f2TextoContemTodosTermos_(produtoRow, termoBusca);

      if (encontrou) {
        if (menor === null || preco < menor) {
          menor = preco;
        }
      }
    }
  }

  return menor;
}

function buscarPrecoCustoPorSkuFase2(skuMaster) {
  const termoBusca = f2Upper_(skuMaster);
  const preco = buscarMenorPrecoCustoSku_(termoBusca);

  return {
    skuMaster: termoBusca,
    precoCusto: preco,
    encontrado: preco !== null
  };
}

/************************************************************
 * SIMULADOR DE PREÇOS
 ************************************************************/

function simularPrecoFase2(dados) {
  const skuMaster = f2Upper_(dados.skuMaster);
  const produto = f2Upper_(dados.produto);

  let precoCusto = f2ToNumber_(dados.precoCusto);

  const termoBuscaCusto = skuMaster || produto;

  if (precoCusto === null && termoBuscaCusto) {
    precoCusto = buscarMenorPrecoCustoSku_(termoBuscaCusto);
  }

  const markup = f2NormalizarPercentual_(dados.markupPercentual);
  const fretePercentual = f2NormalizarPercentual_(dados.fretePercentual);
  const impostoPercentual = f2NormalizarPercentual_(dados.impostoPercentual);
  const outrasPercentual = f2NormalizarPercentual_(dados.outrasDespesasPercentual);

  const freteValorInformado = f2ToNumber_(dados.freteValor);
  const impostoValorInformado = f2ToNumber_(dados.impostoValor);
  const outrasValorInformado = f2ToNumber_(dados.outrasDespesasValor);

  const precoCliente = f2ToNumber_(dados.precoCliente);

  if (precoCusto === null) {
    return {
      ok: false,
      skuMaster,
      produto,
      status: 'SEM CUSTO',
      mensagem: 'Não há preço de custo para simular.'
    };
  }

  const freteValor = freteValorInformado !== null
    ? freteValorInformado
    : precoCusto * fretePercentual;

  const impostoValor = impostoValorInformado !== null
    ? impostoValorInformado
    : precoCusto * impostoPercentual;

  const outrasValor = outrasValorInformado !== null
    ? outrasValorInformado
    : precoCusto * outrasPercentual;

  const custoTotal = f2Arredondar_(precoCusto + freteValor + impostoValor + outrasValor, 4);
  const precoMedbox = f2Arredondar_(custoTotal * (1 + markup), 4);

  let diferenca = '';
  let status = 'SEM PREÇO CLIENTE';

  if (precoCliente !== null) {
    diferenca = f2Arredondar_(precoCliente - precoMedbox, 4);
    status = precoMedbox <= precoCliente ? 'GANHANDO ITEM' : 'ACIMA DO CLIENTE';
  }

  return {
    ok: true,
    skuMaster,
    produto,
    precoCusto,
    markupPercentual: markup,
    freteValor: f2Arredondar_(freteValor, 4),
    impostoValor: f2Arredondar_(impostoValor, 4),
    outrasDespesasValor: f2Arredondar_(outrasValor, 4),
    custoTotal,
    precoMedbox,
    precoCliente,
    diferenca,
    status
  };
}

/************************************************************
 * HELPERS
 ************************************************************/

function f2CriarIndice_(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    idx[String(h).trim()] = i;
  });

  return idx;
}

function f2GetCell_(row, idx, nomeColuna) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) return '';

  return row[pos];
}

function f2SetByHeader_(rowArray, idx, nomeColuna, valor) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) return;

  rowArray[pos] = valor;
}

function f2GerarId_(prefixo) {
  const data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const random = Math.floor(Math.random() * 9000) + 1000;

  return prefixo + '-' + data + '-' + random;
}

function f2Limpar_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .trim()
    .replace(/\s+/g, ' ');
}

function f2Upper_(valor) {
  return f2Limpar_(valor).toUpperCase();
}

function f2SomenteTexto_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor).trim();
}

function f2NormalizarTexto_(valor) {
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

function f2ToNumber_(valor) {
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

function f2NormalizarPercentual_(valor) {
  const n = f2ToNumber_(valor);

  if (n === null || n < 0) return 0;

  if (n <= 1) return n;

  return n / 100;
}

function f2Arredondar_(n, casas) {
  if (n === null || n === undefined || isNaN(n)) return '';

  const f = Math.pow(10, casas || 2);

  return Math.round(n * f) / f;
}

function f2MontarStatusComercial_(precoCliente, precoMedbox) {
  const cliente = f2ToNumber_(precoCliente);
  const medbox = f2ToNumber_(precoMedbox);

  if (medbox === null || medbox === '') return 'SEM CUSTO';
  if (cliente === null || cliente === '') return 'SEM PREÇO CLIENTE';

  return medbox <= cliente ? 'GANHANDO ITEM' : 'ACIMA DO CLIENTE';
}

function f2TextoContemTodosTermos_(texto, termoBusca) {
  const textoNorm = f2NormalizarTexto_(texto);
  const termoNorm = f2NormalizarTexto_(termoBusca);

  if (!textoNorm || !termoNorm) return false;

  const termos = termoNorm
    .split(/\s+/)
    .map(t => t.trim())
    .filter(Boolean);

  if (!termos.length) return false;

  for (const termo of termos) {
    if (!textoNorm.includes(termo)) {
      return false;
    }
  }

  return true;
}

function f2FormatarDataHora_(data) {
  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );
}

function f2RegistrarLog_(mensagem) {
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
    f2FormatarDataHora_(new Date()),
    mensagem
  ]);
}