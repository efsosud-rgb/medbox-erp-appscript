// Este repositório está conectado ao Jules e ao GitHub.
/************************************************************
 * MEDBOX ERP | Google Sheets + Apps Script + HTML
 *
 * Arquivo: Código.gs
 *
 * Telas HTML:
 * - Busca_Avançada
 * - Mapa_de_Produtos
 * - Lançar Cotação
 * - Lançar Preço
 * - Importar Tabela
 * - Simulador
 * - Dashboard
 *
 * Abas técnicas usadas:
 * - MAPA_ERP
 * - AUX_BUSCA
 *
 * Regras:
 * - Não altera BD_
 * - Não altera MAPA_ERP
 * - Não altera AUX_BUSCA
 * - Busca direta na MAPA_ERP para evitar falha do AUX_BUSCA
 * - Filtro Fabricante / Marca busca em produto, SKU, fabricante e fornecedor
 * - ENOX pode encontrar ENOXAPARINA
 * - ENOXAPARINA não pode encontrar TENOXICAM
 ************************************************************/

const ABA_MAPA = 'MAPA_ERP';
const ABA_AUX = 'AUX_BUSCA';

const MAX_RESULTADOS_BUSCA = 300;
const MAX_RESULTADOS_MAPA = 500;

/************************************************************
 * MENU GOOGLE SHEETS
 ************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MEDBOX ERP')
    .addItem('🖥️ Abrir Sistema MEDBOX', 'abrirSistemaMedbox')
    .addSeparator()
    .addItem('🔎 Abrir Busca_Avançada', 'abrirBuscaAvancada')
    .addItem('🧭 Abrir Mapa_de_Produtos', 'abrirMapaProdutos')
    .addItem('📥 Abrir Importar Tabela', 'abrirImportarTabela')
    .addItem('🧾 Abrir Lançar Cotação', 'abrirLancarCotacao')
    .addItem('💰 Abrir Lançar Preço', 'abrirLancarPreco')
    .addItem('🧮 Abrir Simulador', 'abrirSimulador')
    .addItem('📊 Abrir Dashboard', 'abrirDashboard')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Cosmos')
        .addItem('Consultar produtos', 'executarLoteVarreduraFiscalCosmos')
        .addItem('Revisar e aprovar resultados', 'abrirRevisaoCosmos')
    )
    .addSeparator()
    .addItem('🔍 Buscar Dados Fiscais (Cosmos)', 'executarLoteVarreduraFiscalCosmos')
    .addItem('📦 Processar Lote Cosmos (50 produtos)', 'processarLoteCosmos50Manual')
    .addSeparator()
    .addItem('📄 Ir para aba MAPA_ERP', 'abrirAbaMapaErp')
    .addItem('📄 Ir para aba BD_PRECOS', 'abrirAbaBdPrecos')
    .addItem('📄 Ir para aba REVISAO_IMPORTACAO_PRECOS', 'abrirAbaRevisaoPrecos')
    .addToUi();
}

/************************************************************
 * WEB APP
 ************************************************************/

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('MedboxERP');

  template.telaInicial = e && e.parameter && e.parameter.tela
    ? e.parameter.tela
    : 'Busca_Avançada';

  return template
    .evaluate()
    .setTitle('MEDBOX ERP | Sistema Comercial')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/************************************************************
 * ABRIR MODAIS NO GOOGLE SHEETS
 ************************************************************/

function abrirSistemaMedbox() {
  abrirModalMedbox_('Busca_Avançada', 'MEDBOX ERP | Sistema Comercial');
}

function abrirBuscaAvancada() {
  abrirModalMedbox_('Busca_Avançada', 'MEDBOX ERP | Busca_Avançada');
}

function abrirMapaProdutos() {
  abrirModalMedbox_('Mapa_de_Produtos', 'MEDBOX ERP | Mapa_de_Produtos');
}

function abrirImportarTabela() {
  abrirModalMedbox_('Importar_Tabela', 'MEDBOX ERP | Importar Tabela');
}

function abrirLancarCotacao() {
  abrirModalMedbox_('Lancar_Cotacao', 'MEDBOX ERP | Lançar Cotação');
}

function abrirLancarPreco() {
  abrirModalMedbox_('Lancar_Preco', 'MEDBOX ERP | Lançar Preço');
}

function abrirSimulador() {
  abrirModalMedbox_('Simulador', 'MEDBOX ERP | Simulador');
}

function abrirDashboard() {
  abrirModalMedbox_('Dashboard', 'MEDBOX ERP | Dashboard');
}

function abrirModalMedbox_(telaInicial, titulo) {
  const template = HtmlService.createTemplateFromFile('MedboxERP');
  template.telaInicial = telaInicial;

  const html = template
    .evaluate()
    .setWidth(3000)
    .setHeight(1800);

  SpreadsheetApp.getUi().showModalDialog(html, titulo);
}

/************************************************************
 * ABRIR ABAS
 ************************************************************/

function abrirAbaMapaErp() {
  abrirAbaPorNome_('MAPA_ERP');
}

function abrirAbaBdPrecos() {
  abrirAbaPorNome_('BD_PRECOS');
}

function abrirAbaRevisaoPrecos() {
  abrirAbaPorNome_('REVISAO_IMPORTACAO_PRECOS');
}

function processarLoteCosmos50Manual() {
  processarLoteCosmos(50);
}

function abrirAbaPorNome_(nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(nomeAba);

  if (ws) {
    ss.setActiveSheet(ws);
  } else {
    SpreadsheetApp.getUi().alert('Aba não encontrada: ' + nomeAba);
  }
}

/************************************************************
 * OPÇÕES DO SISTEMA
 ************************************************************/

function obterOpcoesSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mapa = ss.getSheetByName(ABA_MAPA);
  const aux = ss.getSheetByName(ABA_AUX);

  if (!mapa) throw new Error('Aba MAPA_ERP não encontrada.');

  const mapaValues = mapa.getDataRange().getValues();
  const auxValues = aux ? aux.getDataRange().getValues() : [[]];

  if (mapaValues.length < 2) throw new Error('MAPA_ERP está vazio.');

  const headers = mapaValues[0].map(String);
  const idx = criarIndiceColunas(headers);

  const clientes = obterClientesDoMapa(headers);

  const categoriasSet = new Set();
  const fabricantesSet = new Set();
  const fornecedoresSet = new Set();

  for (let i = 1; i < mapaValues.length; i++) {
    const row = mapaValues[i];

    const categoria = upperText(getCell(row, idx, 'Categoria'));
    const fabricante = upperText(getCell(row, idx, 'Fabricante Compra 2026'));
    const fornecedor = upperText(getCell(row, idx, 'Fornecedor Compra 2026'));

    if (categoria) categoriasSet.add(categoria);
    if (fabricante) fabricantesSet.add(fabricante);
    if (fornecedor) fornecedoresSet.add(fornecedor);
  }

  return {
    atualizadoEm: formatarDataHora(new Date()),
    totalProdutosMapa: Math.max(0, mapaValues.length - 1),
    totalAuxBusca: Math.max(0, auxValues.length - 1),
    clientes: clientes,
    categorias: Array.from(categoriasSet).sort(),
    fabricantes: Array.from(fabricantesSet).sort().slice(0, 500),
    fornecedores: Array.from(fornecedoresSet).sort().slice(0, 500)
  };
}

/************************************************************
 * TELA 1 — BUSCA_AVANÇADA
 ************************************************************/

function pesquisarBuscaAvancada(filtros) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapa = ss.getSheetByName(ABA_MAPA);

  if (!mapa) throw new Error('Aba MAPA_ERP não encontrada.');

  const values = mapa.getDataRange().getValues();

  if (values.length < 2) throw new Error('MAPA_ERP está vazio.');

  const headers = values[0].map(String);
  const idx = criarIndiceColunas(headers);

  const filtrosTratados = filtros || {};

  const termos = [
    filtrosTratados.termo1,
    filtrosTratados.termo2,
    filtrosTratados.termo3,
    filtrosTratados.termo4
  ].map(limparTexto).filter(Boolean);

  const categoria = normalizarTexto(filtrosTratados.categoria || '');
  const clienteFiltro = normalizarTexto(filtrosTratados.cliente || '');
  const sku = normalizarTexto(filtrosTratados.sku || '');
  const fabricante = normalizarTexto(filtrosTratados.fabricante || '');
  const markup = normalizarMarkup(filtrosTratados.markup);

  const encontrados = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const produto = getCell(row, idx, 'Produto');
    const skuMaster = getCell(row, idx, 'SKU Master');
    const skuOperacional = getCell(row, idx, 'SKU Operacional');
    const produtoNormalizado = getCell(row, idx, 'Produto Normalizado');
    const categoriaRow = getCell(row, idx, 'Categoria');
    const fabricanteRow = getCell(row, idx, 'Fabricante Compra 2026');
    const fornecedorRow = getCell(row, idx, 'Fornecedor Compra 2026');
    const clientesRow = getCell(row, idx, 'Clientes com Consumo');

    const textoCompleto = normalizarTexto([
      produto,
      skuMaster,
      skuOperacional,
      produtoNormalizado,
      categoriaRow,
      fabricanteRow,
      fornecedorRow,
      clientesRow
    ].join(' '));

    if (termos.length > 0 && !textoAtendeTermos(textoCompleto, produto, termos)) continue;

    if (categoria) {
      if (!normalizarTexto(categoriaRow).includes(categoria)) continue;
    }

    if (clienteFiltro) {
      if (!normalizarTexto(clientesRow).includes(clienteFiltro)) continue;
    }

    if (sku) {
      const skuMasterNorm = normalizarTexto(skuMaster);
      const skuOperacionalNorm = normalizarTexto(skuOperacional);

      if (!skuMasterNorm.includes(sku) && !skuOperacionalNorm.includes(sku)) continue;
    }

    if (fabricante) {
      if (!medboxAtendeFiltroFabricanteMarca_(row, idx, fabricante)) continue;
    }

    encontrados.push({
      row,
      qtd: toNumber(getCell(row, idx, 'Quant. total')) || 0,
      preco: toNumber(getCell(row, idx, 'Preço Médio Clientes')) || 999999999
    });
  }

  encontrados.sort((a, b) => {
    if (b.qtd !== a.qtd) return b.qtd - a.qtd;
    if (a.preco !== b.preco) return a.preco - b.preco;

    const pa = normalizarTexto(getCell(a.row, idx, 'Produto'));
    const pb = normalizarTexto(getCell(b.row, idx, 'Produto'));

    return pa.localeCompare(pb);
  });

  const totalEncontrado = encontrados.length;
  const limitados = encontrados.slice(0, MAX_RESULTADOS_BUSCA);

  const resultados = limitados.map(obj => {
    const row = obj.row;

    const skuMaster = limparTexto(getCell(row, idx, 'SKU Master'));
    const precoCusto = toNumber(getCell(row, idx, 'Custo Compra 2026'));
    const precoMedBox = calcularPrecoMedBox(skuMaster, precoCusto, markup);

    return montarLinhaResultadoBusca(row, idx, precoMedBox, precoCusto);
  });

  const sugestoes = gerarSugestoesSkuBusca(limitados, idx);
  const resumo = calcularResumoBusca(resultados);

  return {
    tela: 'Busca_Avançada',
    atualizadoEm: formatarDataHora(new Date()),
    markupUsado: markup,
    totalEncontrado,
    exibidos: resultados.length,
    quantidadeTotalDemandada: resumo.quantidadeTotalDemandada,
    valorTotalDemandado: resumo.valorTotalDemandado,
    sugestoes,
    resultados
  };
}

/************************************************************
 * TELA 2 — MAPA_DE_PRODUTOS
 ************************************************************/

function obterMapaDeProdutos(filtros) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapa = ss.getSheetByName(ABA_MAPA);

  if (!mapa) throw new Error('Aba MAPA_ERP não encontrada.');

  const values = mapa.getDataRange().getValues();

  if (values.length < 2) throw new Error('MAPA_ERP está vazio.');

  const headers = values[0].map(String);
  const idx = criarIndiceColunas(headers);
  const clientes = obterClientesDoMapa(headers);

  const filtrosTratados = filtros || {};

  const termos = [
    filtrosTratados.termo1,
    filtrosTratados.termo2,
    filtrosTratados.termo3,
    filtrosTratados.termo4
  ].map(limparTexto).filter(Boolean);

  const categoria = normalizarTexto(filtrosTratados.categoria || '');
  const clienteFiltro = normalizarTexto(filtrosTratados.cliente || '');
  const sku = normalizarTexto(filtrosTratados.sku || '');
  const fabricante = normalizarTexto(filtrosTratados.fabricante || '');
  const fornecedor = normalizarTexto(filtrosTratados.fornecedor || '');
  const markup = normalizarMarkup(filtrosTratados.markup);
  const limite = normalizarLimite(filtrosTratados.limite, MAX_RESULTADOS_MAPA);

  const grupos = new Map();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const produto = getCell(row, idx, 'Produto');
    const skuMaster = getCell(row, idx, 'SKU Master');
    const skuOperacional = getCell(row, idx, 'SKU Operacional');
    const produtoNormalizado = getCell(row, idx, 'Produto Normalizado');
    const categoriaRow = getCell(row, idx, 'Categoria');
    const fabricanteRow = getCell(row, idx, 'Fabricante Compra 2026');
    const fornecedorRow = getCell(row, idx, 'Fornecedor Compra 2026');
    const clientesRow = getCell(row, idx, 'Clientes com Consumo');

    const textoCompleto = normalizarTexto([
      produto,
      skuMaster,
      skuOperacional,
      produtoNormalizado,
      categoriaRow,
      fabricanteRow,
      fornecedorRow,
      clientesRow
    ].join(' '));

    if (termos.length > 0 && !textoAtendeTermos(textoCompleto, produto, termos)) continue;

    if (categoria) {
      if (!normalizarTexto(categoriaRow).includes(categoria)) continue;
    }

    if (clienteFiltro) {
      if (!normalizarTexto(clientesRow).includes(clienteFiltro)) continue;
    }

    if (sku) {
      const skuMasterNorm = normalizarTexto(skuMaster);
      const skuOperacionalNorm = normalizarTexto(skuOperacional);

      if (!skuMasterNorm.includes(sku) && !skuOperacionalNorm.includes(sku)) continue;
    }

    if (fabricante) {
      if (!medboxAtendeFiltroFabricanteMarca_(row, idx, fabricante)) continue;
    }

    if (fornecedor) {
      if (!normalizarTexto(fornecedorRow).includes(fornecedor)) continue;
    }

    const chave = montarChaveProduto(row, idx);

    if (!grupos.has(chave)) {
      grupos.set(chave, criarGrupoProduto(row, idx, clientes));
    }

    acumularGrupoProduto(grupos.get(chave), row, idx, clientes);
  }

  const produtosAgrupados = Array.from(grupos.values()).map(g => finalizarGrupoProduto(g, markup));

  produtosAgrupados.sort((a, b) => {
    if ((b.quantTotal || 0) !== (a.quantTotal || 0)) {
      return (b.quantTotal || 0) - (a.quantTotal || 0);
    }

    const precoA = a.precoMedioClientes || 999999999;
    const precoB = b.precoMedioClientes || 999999999;

    if (precoA !== precoB) return precoA - precoB;

    return normalizarTexto(a.produto).localeCompare(normalizarTexto(b.produto));
  });

  const totalEncontrado = produtosAgrupados.length;
  const limitados = produtosAgrupados.slice(0, limite);

  const resumo = calcularResumoMapa(limitados);

  return {
    tela: 'Mapa_de_Produtos',
    atualizadoEm: formatarDataHora(new Date()),
    markupUsado: markup,
    totalEncontrado,
    exibidos: limitados.length,
    quantidadeTotalDemandada: resumo.quantidadeTotalDemandada,
    valorTotalDemandado: resumo.valorTotalDemandado,
    clientes: clientes,
    produtos: limitados
  };
}

/************************************************************
 * AGRUPAMENTO DO MAPA_DE_PRODUTOS
 ************************************************************/

function montarChaveProduto(row, idx) {
  const skuMaster = limparTexto(getCell(row, idx, 'SKU Master'));
  const skuOperacional = limparTexto(getCell(row, idx, 'SKU Operacional'));
  const produtoNormalizado = limparTexto(getCell(row, idx, 'Produto Normalizado'));
  const produto = limparTexto(getCell(row, idx, 'Produto'));

  const base = skuMaster || skuOperacional || produtoNormalizado || produto;

  return normalizarTexto(base);
}

function criarGrupoProduto(row, idx, clientes) {
  const grupo = {
    categoria: upperText(getCell(row, idx, 'Categoria')),
    skuMaster: upperText(getCell(row, idx, 'SKU Master')),
    produto: upperText(getCell(row, idx, 'Produto')),
    precoMedioSomaPonderada: 0,
    precoMedioPeso: 0,
    quantTotalFallback: 0,
    menorValorUnit: null,
    maiorValorUnit: null,
    precoCusto: null,
    fornecedorCompra2026: upperText(getCell(row, idx, 'Fornecedor Compra 2026')),
    fabricanteCompra2026: upperText(getCell(row, idx, 'Fabricante Compra 2026')),
    dataCompra2026: upperText(getCell(row, idx, 'Data Compra 2026')),
    produtoCompra2026: upperText(getCell(row, idx, 'Produto Compra 2026')),
    clientesComConsumoSet: new Set(),
    skuOperacional: upperText(getCell(row, idx, 'SKU Operacional')),
    produtoNormalizado: upperText(getCell(row, idx, 'Produto Normalizado')),
    clientesValores: {}
  };

  clientes.forEach(c => {
    grupo.clientesValores[c] = {
      quant: 0,
      valorUnit: null,
      demandaR: 0
    };
  });

  return grupo;
}

function acumularGrupoProduto(grupo, row, idx, clientes) {
  const quantTotalRow = toNumber(getCell(row, idx, 'Quant. total')) || 0;
  const precoMedioRow = toNumber(getCell(row, idx, 'Preço Médio Clientes'));
  const menorRow = toNumber(getCell(row, idx, 'Menor Valor Unit'));
  const maiorRow = toNumber(getCell(row, idx, 'Maior Valor Unit'));
  const custoRow = toNumber(getCell(row, idx, 'Custo Compra 2026'));

  grupo.quantTotalFallback += quantTotalRow;

  if (precoMedioRow !== null && quantTotalRow > 0) {
    grupo.precoMedioSomaPonderada += precoMedioRow * quantTotalRow;
    grupo.precoMedioPeso += quantTotalRow;
  }

  if (menorRow !== null) {
    if (grupo.menorValorUnit === null || menorRow < grupo.menorValorUnit) {
      grupo.menorValorUnit = menorRow;
    }
  }

  if (maiorRow !== null) {
    if (grupo.maiorValorUnit === null || maiorRow > grupo.maiorValorUnit) {
      grupo.maiorValorUnit = maiorRow;
    }
  }

  if (custoRow !== null && custoRow > 0) {
    if (grupo.precoCusto === null || custoRow < grupo.precoCusto) {
      grupo.precoCusto = custoRow;
    }
  }

  const clientesTexto = upperText(getCell(row, idx, 'Clientes com Consumo'));

  if (clientesTexto) {
    clientesTexto.split(/[;,]/).forEach(c => {
      const t = upperText(c);
      if (t) grupo.clientesComConsumoSet.add(t);
    });
  }

  clientes.forEach(cliente => {
    const q = toNumber(getCell(row, idx, `${cliente} Quant`)) || 0;
    const v = toNumber(getCell(row, idx, `${cliente} Valor Unit`));

    if (q > 0) {
      grupo.clientesValores[cliente].quant += q;

      if (v !== null) {
        if (
          grupo.clientesValores[cliente].valorUnit === null ||
          v < grupo.clientesValores[cliente].valorUnit
        ) {
          grupo.clientesValores[cliente].valorUnit = v;
        }

        grupo.clientesValores[cliente].demandaR += q * v;
      }

      grupo.clientesComConsumoSet.add(upperText(cliente));
    }
  });

  if (!grupo.fornecedorCompra2026) {
    grupo.fornecedorCompra2026 = upperText(getCell(row, idx, 'Fornecedor Compra 2026'));
  }

  if (!grupo.fabricanteCompra2026) {
    grupo.fabricanteCompra2026 = upperText(getCell(row, idx, 'Fabricante Compra 2026'));
  }

  if (!grupo.produtoCompra2026) {
    grupo.produtoCompra2026 = upperText(getCell(row, idx, 'Produto Compra 2026'));
  }
}

function finalizarGrupoProduto(grupo, markup) {
  let quantTotalClientes = 0;

  Object.keys(grupo.clientesValores).forEach(cliente => {
    quantTotalClientes += grupo.clientesValores[cliente].quant || 0;
  });

  const quantTotal = quantTotalClientes > 0 ? quantTotalClientes : grupo.quantTotalFallback;

  let precoMedioClientes = null;

  if (grupo.precoMedioPeso > 0) {
    precoMedioClientes = grupo.precoMedioSomaPonderada / grupo.precoMedioPeso;
  }

  const precoMedBox = calcularPrecoMedBox(grupo.skuMaster, grupo.precoCusto, markup);
  const demandaR = grupo.menorValorUnit === null ? '' : arredondar(grupo.menorValorUnit * quantTotal, 4);
  const statusComercial = montarStatusComercialMapa(grupo.menorValorUnit, precoMedBox);

  return {
    categoria: grupo.categoria,
    skuMaster: grupo.skuMaster,
    produto: grupo.produto,
    quantTotal: arredondar(quantTotal, 4),
    precoMedioClientes: precoMedioClientes === null ? '' : arredondar(precoMedioClientes, 4),
    menorValorUnit: grupo.menorValorUnit,
    maiorValorUnit: grupo.maiorValorUnit,
    precoMedBox: precoMedBox,
    precoCusto: grupo.precoCusto,
    demandaR: demandaR,
    statusComercial: statusComercial,
    fornecedorCompra2026: grupo.fornecedorCompra2026,
    fabricanteCompra2026: grupo.fabricanteCompra2026,
    dataCompra2026: grupo.dataCompra2026,
    produtoCompra2026: grupo.produtoCompra2026,
    clientesComConsumo: Array.from(grupo.clientesComConsumoSet).sort().join('; '),
    skuOperacional: grupo.skuOperacional,
    produtoNormalizado: grupo.produtoNormalizado,
    clientesValores: grupo.clientesValores
  };
}

/************************************************************
 * PREÇO MEDBOX / LINHAS / SUGESTÕES
 ************************************************************/

function calcularPrecoMedBox(skuMaster, precoCusto, markup) {
  const sku = limparTexto(skuMaster);
  const custo = toNumber(precoCusto);

  if (!sku) return '';
  if (custo === null || custo <= 0) return '';

  return arredondar(custo * (1 + markup), 4);
}

function montarLinhaResultadoBusca(row, idx, precoMedBox, precoCusto) {
  const precoMedio = toNumber(getCell(row, idx, 'Preço Médio Clientes'));
  const menorValor = toNumber(getCell(row, idx, 'Menor Valor Unit'));
  const quantTotal = toNumber(getCell(row, idx, 'Quant. total')) || 0;
  const demandaR = menorValor === null ? '' : arredondar(menorValor * quantTotal, 4);

  return {
    categoria: upperText(getCell(row, idx, 'Categoria')),
    skuMaster: upperText(getCell(row, idx, 'SKU Master')),
    produto: upperText(getCell(row, idx, 'Produto')),
    quantTotal: quantTotal,
    precoMedioClientes: precoMedio,
    menorValorUnit: menorValor,
    maiorValorUnit: toNumber(getCell(row, idx, 'Maior Valor Unit')),
    precoMedBox: precoMedBox,
    precoCusto: precoCusto,
    demandaR: demandaR,
    fornecedorCompra2026: upperText(getCell(row, idx, 'Fornecedor Compra 2026')),
    fabricanteCompra2026: upperText(getCell(row, idx, 'Fabricante Compra 2026')),
    dataCompra2026: upperText(getCell(row, idx, 'Data Compra 2026')),
    produtoCompra2026: upperText(getCell(row, idx, 'Produto Compra 2026')),
    clientesComConsumo: upperText(getCell(row, idx, 'Clientes com Consumo')),
    skuOperacional: upperText(getCell(row, idx, 'SKU Operacional')),
    produtoNormalizado: upperText(getCell(row, idx, 'Produto Normalizado')),
    statusComercial: montarStatusComercialMapa(menorValor, precoMedBox)
  };
}

function gerarSugestoesSkuBusca(limitados, idx) {
  const mapa = new Map();

  for (const obj of limitados) {
    const row = obj.row;

    const skuMaster = upperText(getCell(row, idx, 'SKU Master'));
    const skuOperacional = upperText(getCell(row, idx, 'SKU Operacional'));
    const produto = upperText(getCell(row, idx, 'Produto'));
    const qtd = toNumber(getCell(row, idx, 'Quant. total')) || 0;

    const skuFinal = skuMaster || skuOperacional;

    if (!skuFinal) continue;

    if (!mapa.has(skuFinal) || mapa.get(skuFinal).qtd < qtd) {
      mapa.set(skuFinal, {
        sku: skuFinal,
        tipo: skuMaster ? 'SKU MASTER EXISTENTE' : 'SKU OPERACIONAL',
        produto,
        score: skuMaster ? 100 : 80,
        qtd
      });
    }
  }

  return Array.from(mapa.values())
    .sort((a, b) => b.qtd - a.qtd)
    .slice(0, 5);
}

function montarStatusComercialMapa(menorValorCliente, precoMedBox) {
  const menor = toNumber(menorValorCliente);
  const medbox = toNumber(precoMedBox);

  if (medbox === null || medbox === '' || medbox <= 0) return 'SEM CUSTO';
  if (menor === null || menor === '') return 'SEM REFERÊNCIA';

  if (medbox <= menor) return 'GANHANDO ITEM';

  return 'ACIMA DO CLIENTE';
}

function calcularResumoBusca(resultados) {
  let quantidadeTotalDemandada = 0;
  let valorTotalDemandado = 0;

  resultados.forEach(r => {
    const q = toNumber(r.quantTotal) || 0;
    const menor = toNumber(r.menorValorUnit);

    quantidadeTotalDemandada += q;

    if (menor !== null) {
      valorTotalDemandado += q * menor;
    }
  });

  return {
    quantidadeTotalDemandada: arredondar(quantidadeTotalDemandada, 4),
    valorTotalDemandado: arredondar(valorTotalDemandado, 4)
  };
}

function calcularResumoMapa(produtos) {
  let quantidadeTotalDemandada = 0;
  let valorTotalDemandado = 0;

  produtos.forEach(r => {
    quantidadeTotalDemandada += toNumber(r.quantTotal) || 0;
    valorTotalDemandado += toNumber(r.demandaR) || 0;
  });

  return {
    quantidadeTotalDemandada: arredondar(quantidadeTotalDemandada, 4),
    valorTotalDemandado: arredondar(valorTotalDemandado, 4)
  };
}

/************************************************************
 * FILTRO GLOBAL FABRICANTE / MARCA
 ************************************************************/

function medboxAtendeFiltroFabricanteMarca_(row, idx, filtroFabricanteMarca) {
  const filtro = normalizarTexto(filtroFabricanteMarca);

  if (!filtro) return true;

  const textoCombinado = medboxMontarTextoFabricanteMarca_(row, idx);

  if (!textoCombinado) return false;

  const termos = filtro
    .split(/\s+/)
    .map(t => t.trim())
    .filter(Boolean);

  if (!termos.length) return true;

  for (const termo of termos) {
    if (!textoCombinado.includes(termo)) {
      return false;
    }
  }

  return true;
}

function medboxMontarTextoFabricanteMarca_(row, idx) {
  const campos = [
    'Fabricante Compra 2026',
    'Fornecedor Compra 2026',
    'Fabricante',
    'Marca',
    'Marca / Fabricante',
    'Fabricante / Marca',
    'FABRICANTE',
    'MARCA',
    'FABRICANTE_MARCA',
    'FORNECEDOR',
    'Fornecedor',
    'Produto',
    'PRODUTO',
    'Produto Normalizado',
    'DESCRICAO_NORMALIZADA',
    'DESCRICAO_FORNECEDOR',
    'Descrição',
    'Descricao',
    'Descrição do Produto',
    'SKU Master',
    'SKU_MASTER',
    'SKU Operacional',
    'SKU_OPERACIONAL',
    'Produto Compra 2026'
  ];

  const partes = [];

  campos.forEach(function(nomeColuna) {
    const valor = medboxGetCellFlex_(row, idx, nomeColuna);

    if (valor !== null && valor !== undefined && String(valor).trim() !== '') {
      partes.push(String(valor));
    }
  });

  return normalizarTexto(partes.join(' '));
}

function medboxGetCellFlex_(row, idx, nomeColuna) {
  if (!idx || !row) return '';

  if (idx[nomeColuna] !== undefined && idx[nomeColuna] !== null) {
    return row[idx[nomeColuna]];
  }

  const nomeNorm = normalizarTexto(nomeColuna);
  const chaves = Object.keys(idx);

  for (let i = 0; i < chaves.length; i++) {
    const chave = chaves[i];

    if (normalizarTexto(chave) === nomeNorm) {
      return row[idx[chave]];
    }
  }

  return '';
}

/************************************************************
 * BUSCA / NORMALIZAÇÃO
 ************************************************************/

function textoAtendeTermos(textoBusca, produtoOriginal, termos) {
  const textoNorm = normalizarTexto(textoBusca);
  const produtoNorm = normalizarTexto(produtoOriginal);

  for (const termo of termos) {
    const volume = extrairVolumeTermo(termo);

    if (volume) {
      if (
        !textoTemVolumeExato(produtoNorm, volume) &&
        !textoTemVolumeExato(textoNorm, volume)
      ) {
        return false;
      }

      continue;
    }

    const alternativas = expandirTermoBusca(termo);
    let encontrou = false;

    for (const alt of alternativas) {
      const altNorm = normalizarTexto(alt);

      if (!altNorm) continue;

      if (termoBuscaSeguro(textoNorm, altNorm)) {
        encontrou = true;
        break;
      }
    }

    if (!encontrou) return false;
  }

  return true;
}

function termoBuscaSeguro(textoNorm, termoNorm) {
  if (!textoNorm || !termoNorm) return false;

  const termo = normalizarTexto(termoNorm);
  const texto = normalizarTexto(textoNorm);

  if (!termo) return false;

  const tokens = texto.split(/\s+/).filter(Boolean);

  if (termo.length >= 5) {
    if (tokens.includes(termo)) return true;

    for (const token of tokens) {
      if (token.startsWith(termo)) return true;
    }

    return false;
  }

  for (const token of tokens) {
    if (token.startsWith(termo)) return true;
  }

  return false;
}

function expandirTermoBusca(termo) {
  const t = normalizarTexto(termo);

  const grupos = {
    'SORO': ['SORO', 'SORO FISIOLOGICO', 'CLORETO DE SODIO', 'FISIOLOGICO', 'SF'],
    'FISIOLOGICO': ['FISIOLOGICO', 'SORO FISIOLOGICO', 'CLORETO DE SODIO', 'SF'],
    'SF': ['SF', 'SORO FISIOLOGICO', 'CLORETO DE SODIO'],
    'GLICOSADO': ['GLICOSADO', 'GLICOSE', 'SORO GLICOSADO', 'SOLUCAO GLICOSADA'],
    'GLICOSE': ['GLICOSE', 'GLICOSADO', 'SORO GLICOSADO', 'SOLUCAO GLICOSADA'],
    'API': ['API', 'AGUA PARA INJECAO', 'AGUA INJETAVEL'],
    'RL': ['RL', 'RINGER LACTATO', 'RINGER COM LACTATO'],
    'RINGER': ['RINGER', 'RINGER LACTATO', 'RINGER COM LACTATO', 'RINGER SIMPLES']
  };

  return grupos[t] || [t];
}

function extrairVolumeTermo(termo) {
  const t = normalizarTexto(termo);

  const ml = t.match(/(^|[^0-9])(\d{1,4})\s*ML([^0-9]|$)/);
  if (ml) return `${ml[2]}ML`;

  const l = t.match(/(^|[^0-9])(\d{1,2})\s*L([^0-9]|$)/);
  if (l) return `${l[2]}L`;

  return null;
}

function textoTemVolumeExato(texto, volume) {
  const textoNorm = normalizarTexto(texto);
  const volNorm = normalizarTexto(volume);

  if (volNorm.endsWith('ML')) {
    const numero = volNorm.replace('ML', '');
    const regex = new RegExp(`(^|[^0-9])${escapeRegExp(numero)}\\s*ML([^0-9]|$)`);
    return regex.test(textoNorm);
  }

  if (volNorm.endsWith('L')) {
    const numero = volNorm.replace('L', '');
    const regex = new RegExp(`(^|[^0-9])${escapeRegExp(numero)}\\s*L([^0-9]|$)`);
    return regex.test(textoNorm);
  }

  return termoBuscaSeguro(textoNorm, volNorm);
}

/************************************************************
 * HELPERS GERAIS
 ************************************************************/

function obterClientesDoMapa(headers) {
  return headers
    .filter(h => String(h).trim().endsWith(' Quant'))
    .map(h => upperText(String(h).replace(' Quant', '').trim()))
    .filter(Boolean)
    .sort();
}

function criarIndiceColunas(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    idx[String(h).trim()] = i;
  });

  return idx;
}

function getCell(row, idx, nomeColuna) {
  const pos = idx[nomeColuna];

  if (pos === undefined || pos === null || pos < 0) {
    return '';
  }

  return row[pos];
}

function limparTexto(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .trim()
    .replace(/\s+/g, ' ');
}

function upperText(valor) {
  return limparTexto(valor).toUpperCase();
}

function normalizarTexto(valor) {
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

function toNumber(valor) {
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

function normalizarMarkup(valor) {
  const n = toNumber(valor);

  if (n === null || n < 0) return 0.8;
  if (n <= 1) return n;

  return n / 100;
}

function normalizarLimite(valor, padrao) {
  const n = toNumber(valor);

  if (n === null || n <= 0) return padrao;

  return Math.min(Math.floor(n), 1000);
}

function arredondar(n, casas) {
  if (n === null || n === undefined || isNaN(n)) return '';

  const f = Math.pow(10, casas || 2);

  return Math.round(n * f) / f;
}

function formatarDataHora(data) {
  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );
}

function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
