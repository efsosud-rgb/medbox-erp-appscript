// Este repositório está conectado ao Jules e ao GitHub.
/************************************************************
 * MEDBOX ERP | Google Sheets + Apps Script + HTML
 *
 * Arquivo: Código.gs
 *
 * Telas HTML:
 * - Busca_Avançada
 * - Mapa_de_Produtos
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

const MEDBOX_SPREADSHEET_ID = '1QLurrnKCXorjXUoUtPl6MoSP0Gk32Yz0AWBHxebifCw';

function obterPlanilhaMedbox_() {
  return SpreadsheetApp.openById(MEDBOX_SPREADSHEET_ID);
}

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
    .addItem('📊 Abrir Dashboard', 'abrirDashboard')
    .addSeparator()
    .addItem('📄 Ir para aba MAPA_ERP', 'abrirAbaMapaErp')
    .addItem('📄 Ir para aba BD_PRECOS', 'abrirAbaBdPrecos')
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

function abrirAbaPorNome_(nomeAba) {
  const ss = obterPlanilhaMedbox_();
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
  const ss = obterPlanilhaMedbox_();

  const mapa = ss.getSheetByName(ABA_MAPA);
  const aux = ss.getSheetByName(ABA_AUX);

  if (!mapa) throw new Error('Aba MAPA_ERP não encontrada.');

  const mapaValues = mapa.getDataRange().getValues();
  const auxValues = aux ? aux.getDataRange().getValues() : [[]];

  if (mapaValues.length < 2) throw new Error('MAPA_ERP está vazio.');

  const headers = mapaValues[0].map(String);
  const mapCols = f2dMapearColunasMapaErp_(headers);

  const clientes = obterClientesDoMapa(headers);

  const categoriasSet = new Set();
  const fabricantesSet = new Set();
  const fornecedoresSet = new Set();

  for (let i = 1; i < mapaValues.length; i++) {
    const row = mapaValues[i];

    const categoria = upperText(f2dGetMapCell_(row, mapCols, 'categoria'));
    const fabricante = upperText(f2dGetMapCell_(row, mapCols, 'fabricante'));
    const fornecedor = upperText(f2dGetMapCell_(row, mapCols, 'fornecedor'));

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
  const ss = obterPlanilhaMedbox_();
  const mapa = ss.getSheetByName(ABA_MAPA);

  if (!mapa) throw new Error('Aba MAPA_ERP não encontrada.');

  const values = mapa.getDataRange().getValues();

  if (values.length < 2) throw new Error('MAPA_ERP está vazio.');

  const headers = values[0].map(String);
  const mapCols = f2dMapearColunasMapaErp_(headers);

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

    const produto = f2dGetMapCell_(row, mapCols, 'produto');
    const skuMaster = f2dGetMapCell_(row, mapCols, 'skuMaster');
    const skuOperacional = f2dGetMapCell_(row, mapCols, 'skuOperacional');
    const produtoNormalizado = f2dGetMapCell_(row, mapCols, 'produtoNormalizado');
    const categoriaRow = f2dGetMapCell_(row, mapCols, 'categoria');
    const fabricanteRow = f2dGetMapCell_(row, mapCols, 'fabricante');
    const fornecedorRow = f2dGetMapCell_(row, mapCols, 'fornecedor');
    const clientesRow = f2dGetMapCell_(row, mapCols, 'clientesComConsumo');

    if (!produto && !skuMaster && !skuOperacional) {
      continue;
    }

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
      if (!medboxAtendeFiltroFabricanteMarca_(row, mapCols, fabricante)) continue;
    }

    encontrados.push({
      row,
      qtd: toNumber(f2dGetMapCell_(row, mapCols, 'quantTotal')) || 0,
      preco: toNumber(f2dGetMapCell_(row, mapCols, 'precoMedioClientes')) || 999999999
    });
  }

  encontrados.sort((a, b) => {
    if (b.qtd !== a.qtd) return b.qtd - a.qtd;
    if (a.preco !== b.preco) return a.preco - b.preco;

    const pa = normalizarTexto(f2dGetMapCell_(a.row, mapCols, 'produto'));
    const pb = normalizarTexto(f2dGetMapCell_(b.row, mapCols, 'produto'));

    return pa.localeCompare(pb);
  });

  const totalEncontrado = encontrados.length;
  const limitados = encontrados.slice(0, MAX_RESULTADOS_BUSCA);

  const resultados = limitados.map(obj => {
    const row = obj.row;

    const skuMaster = limparTexto(f2dGetMapCell_(row, mapCols, 'skuMaster'));
    const precoCusto = toNumber(f2dGetMapCell_(row, mapCols, 'custo'));
    const precoMedBox = calcularPrecoMedBox(skuMaster, precoCusto, markup);

    return montarLinhaResultadoBusca(row, mapCols, precoMedBox, precoCusto);
  });

  const sugestoes = gerarSugestoesSkuBusca(limitados, mapCols);
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
  const ss = obterPlanilhaMedbox_();
  const mapa = ss.getSheetByName(ABA_MAPA);

  if (!mapa) throw new Error('Aba MAPA_ERP não encontrada.');

  const values = mapa.getDataRange().getValues();

  if (values.length < 2) throw new Error('MAPA_ERP está vazio.');

  const headers = values[0].map(String);
  const mapCols = f2dMapearColunasMapaErp_(headers);
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

  // --- CARREGAMENTO DOS DADOS DE CADA CLIENTE DAS RESPECTIVAS ABAS BD_<CLIENTE> ---
  const dadosClientesMap = {};
  const abasClientes = f2dObterAbasClientes_(ss);

  abasClientes.forEach(sheet => {
    const name = sheet.getName();
    const clienteName = upperText(name.substring(3).trim());

    if (clienteFiltro && clienteFiltro !== normalizarTexto(clienteName)) {
      return;
    }

    dadosClientesMap[clienteName] = f2dCarregarDadosClienteSheet_(sheet);
  });

  const grupos = new Map();

  // Resolve key mismatch: find the matching key from dynamic customers list for lookup
  let matchingClientKey = "";
  if (clienteFiltro) {
    matchingClientKey = clientes.find(c => normalizarTexto(c) === clienteFiltro) || "";
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const produto = f2dGetMapCell_(row, mapCols, 'produto');
    const skuMaster = f2dGetMapCell_(row, mapCols, 'skuMaster');
    const skuOperacional = f2dGetMapCell_(row, mapCols, 'skuOperacional');
    const produtoNormalizado = f2dGetMapCell_(row, mapCols, 'produtoNormalizado');
    const categoriaRow = f2dGetMapCell_(row, mapCols, 'categoria');
    const fabricanteRow = f2dGetMapCell_(row, mapCols, 'fabricante');
    const fornecedorRow = f2dGetMapCell_(row, mapCols, 'fornecedor');

    if (!skuMaster && !skuOperacional && !produtoNormalizado && !produto) {
      continue;
    }

    // Identificação dinâmica de consumo por cliente
    let temConsumoQualquerCliente = false;
    const consumosPorCliente = {};

    clientes.forEach(cliente => {
      const clientData = dadosClientesMap[cliente];
      let q = 0;
      let v = null;
      let d = 0;

      if (clientData) {
        // Prioridade 1: SKU_MASTER da aba cliente × SKU Master da MAPA_ERP
        const skuMasterKey = normalizarTexto(skuMaster);
        const skuOperacionalKey = normalizarTexto(skuOperacional);

        let found = false;

        if (skuMasterKey && clientData.skuMap[skuMasterKey]) {
          const dataItem = clientData.skuMap[skuMasterKey];
          q = dataItem.quant;
          v = dataItem.valorUnit;
          d = dataItem.demandaR;
          found = true;
        } else if (skuOperacionalKey && clientData.skuMap[skuOperacionalKey]) {
          const dataItem = clientData.skuMap[skuOperacionalKey];
          q = dataItem.quant;
          v = dataItem.valorUnit;
          d = dataItem.demandaR;
          found = true;
        }

        // Prioridade 2: Produtos normalizado × Produto Normalizado da MAPA_ERP
        if (!found) {
          const prodNormKey = normalizarTexto(produtoNormalizado);
          if (prodNormKey && clientData.productMap[prodNormKey]) {
            const dataItem = clientData.productMap[prodNormKey];
            q = dataItem.quant;
            v = dataItem.valorUnit;
            d = dataItem.demandaR;
            found = true;
          }
        }

        // Prioridade 3: Produtos normalizado × Produto da MAPA_ERP
        if (!found) {
          const prodKey = normalizarTexto(produto);
          if (prodKey && clientData.productMap[prodKey]) {
            const dataItem = clientData.productMap[prodKey];
            q = dataItem.quant;
            v = dataItem.valorUnit;
            d = dataItem.demandaR;
            found = true;
          }
        }
      }

      if (q > 0) {
        temConsumoQualquerCliente = true;
      }

      consumosPorCliente[cliente] = {
        quant: q,
        valorUnit: v,
        demandaR: d
      };
    });

    if (clienteFiltro) {
      const consumoCli = matchingClientKey ? consumosPorCliente[matchingClientKey] : null;
      if (!consumoCli || consumoCli.quant <= 0) {
        continue;
      }
    }

    const textoCompleto = normalizarTexto([
      produto,
      skuMaster,
      skuOperacional,
      produtoNormalizado,
      categoriaRow,
      fabricanteRow,
      fornecedorRow
    ].join(' '));

    if (termos.length > 0 && !textoAtendeTermos(textoCompleto, produto, termos)) continue;

    if (categoria) {
      if (!normalizarTexto(categoriaRow).includes(categoria)) continue;
    }

    if (sku) {
      const skuMasterNorm = normalizarTexto(skuMaster);
      const skuOperacionalNorm = normalizarTexto(skuOperacional);

      if (!skuMasterNorm.includes(sku) && !skuOperacionalNorm.includes(sku)) continue;
    }

    if (fabricante) {
      if (!medboxAtendeFiltroFabricanteMarca_(row, mapCols, fabricante)) continue;
    }

    if (fornecedor) {
      if (!normalizarTexto(fornecedorRow).includes(fornecedor)) continue;
    }

    const chave = skuMaster ? normalizarTexto(skuMaster) : montarChaveProduto(row, mapCols);
    if (!chave) continue;

    if (!grupos.has(chave)) {
      grupos.set(chave, criarGrupoProduto(row, mapCols, clientes));
    }

    acumularGrupoProduto(grupos.get(chave), row, mapCols, clientes, consumosPorCliente);
  }

  const produtosAgrupados = Array.from(grupos.values()).map(g => finalizarGrupoProduto(g, markup, matchingClientKey));

  produtosAgrupados.sort((a, b) => {
    if ((b.quantTotal || 0) !== (a.quantTotal || 0)) {
      return (b.quantTotal || 0) - (a.quantTotal || 0);
    }

    const precoA = a.precoCustoAtual || 999999999;
    const precoB = b.precoCustoAtual || 999999999;

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
  const skuMaster = f2dGetMapCell_(row, idx, 'skuMaster');
  const skuOperacional = f2dGetMapCell_(row, idx, 'skuOperacional');
  const produtoNormalizado = f2dGetMapCell_(row, idx, 'produtoNormalizado');
  const produto = f2dGetMapCell_(row, idx, 'produto');

  const base = skuMaster || skuOperacional || produtoNormalizado || produto;

  return normalizarTexto(base);
}

function criarGrupoProduto(row, mapCols, clientes) {
  const grupo = {
    categoria: upperText(f2dGetMapCell_(row, mapCols, 'categoria')),
    skuMaster: upperText(f2dGetMapCell_(row, mapCols, 'skuMaster')),
    produto: upperText(f2dGetMapCell_(row, mapCols, 'produto')),
    precoMedioSomaPonderada: 0,
    precoMedioPeso: 0,
    quantTotalFallback: 0,
    menorValorUnit: null,
    maiorValorUnit: null,
    precoCusto: null,
    fornecedorCompra2026: upperText(f2dGetMapCell_(row, mapCols, 'fornecedor')),
    fabricanteCompra2026: upperText(f2dGetMapCell_(row, mapCols, 'fabricante')),
    dataCompra2026: upperText(f2dGetMapCell_(row, mapCols, 'dataCompra')),
    produtoCompra2026: upperText(f2dGetMapCell_(row, mapCols, 'produto')),
    clientesComConsumoSet: new Set(),
    skuOperacional: upperText(f2dGetMapCell_(row, mapCols, 'skuOperacional')),
    produtoNormalizado: upperText(f2dGetMapCell_(row, mapCols, 'produtoNormalizado')),
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

function acumularGrupoProduto(grupo, row, mapCols, clientes, consumosPorCliente) {
  const quantTotalRow = toNumber(f2dGetMapCell_(row, mapCols, 'quantTotal')) || 0;
  const precoMedioRow = toNumber(f2dGetMapCell_(row, mapCols, 'precoMedioClientes'));
  const menorRow = toNumber(f2dGetMapCell_(row, mapCols, 'menorValorUnit'));
  const custoRow = toNumber(f2dGetMapCell_(row, mapCols, 'custo'));

  grupo.quantTotalFallback += quantTotalRow;

  if (precoMedioRow !== null && quantTotalRow > 0) {
    grupo.precoMedioSomaPonderada += precoMedioRow * quantTotalRow;
    grupo.precoMedioPeso += quantTotalRow;
  }

  if (menorRow !== null) {
    if (grupo.menorValorUnit === null || menorRow < grupo.menorValorUnit) {
      grupo.menorValorUnit = menorRow;
    }
    if (grupo.maiorValorUnit === null || menorRow > grupo.maiorValorUnit) {
      grupo.maiorValorUnit = menorRow;
    }
  }

  if (custoRow !== null && custoRow > 0) {
    if (grupo.precoCusto === null || custoRow < grupo.precoCusto) {
      grupo.precoCusto = custoRow;
    }
  }

  if (consumosPorCliente) {
    clientes.forEach(cliente => {
      const consumo = consumosPorCliente[cliente];
      if (consumo && consumo.quant > 0) {
        grupo.clientesValores[cliente].quant += consumo.quant;

        if (consumo.valorUnit !== null) {
          if (
            grupo.clientesValores[cliente].valorUnit === null ||
            consumo.valorUnit < grupo.clientesValores[cliente].valorUnit
          ) {
            grupo.clientesValores[cliente].valorUnit = consumo.valorUnit;
          }
        }
        grupo.clientesValores[cliente].demandaR += consumo.demandaR;
        grupo.clientesComConsumoSet.add(upperText(cliente));
      }
    });
  }

  if (!grupo.fornecedorCompra2026) {
    grupo.fornecedorCompra2026 = upperText(f2dGetMapCell_(row, mapCols, 'fornecedor'));
  }

  if (!grupo.fabricanteCompra2026) {
    grupo.fabricanteCompra2026 = upperText(f2dGetMapCell_(row, mapCols, 'fabricante'));
  }

  if (!grupo.produtoCompra2026) {
    grupo.produtoCompra2026 = upperText(f2dGetMapCell_(row, mapCols, 'produto'));
  }
}

function finalizarGrupoProduto(grupo, markup, matchingClientKey) {
  let quantTotal = 0;
  let demandaR = 0;
  let menorValorUnit = grupo.menorValorUnit;

  if (matchingClientKey) {
    const cliVal = grupo.clientesValores[matchingClientKey];
    if (cliVal) {
      quantTotal = cliVal.quant || 0;
      menorValorUnit = cliVal.valorUnit;
      demandaR = cliVal.demandaR || 0;
    }
  } else {
    let sumQuant = 0;
    let sumDemanda = 0;
    let minVal = null;

    Object.keys(grupo.clientesValores).forEach(cliente => {
      const cliVal = grupo.clientesValores[cliente];
      sumQuant += cliVal.quant || 0;
      sumDemanda += cliVal.demandaR || 0;
      if (cliVal.valorUnit !== null) {
        if (minVal === null || cliVal.valorUnit < minVal) {
          minVal = cliVal.valorUnit;
        }
      }
    });

    quantTotal = sumQuant > 0 ? sumQuant : grupo.quantTotalFallback;
    menorValorUnit = minVal !== null ? minVal : grupo.menorValorUnit;
    demandaR = sumDemanda > 0 ? sumDemanda : (menorValorUnit !== null ? menorValorUnit * quantTotal : 0);
  }

  let precoMedioClientes = null;
  if (matchingClientKey) {
    const cliVal = grupo.clientesValores[matchingClientKey];
    precoMedioClientes = cliVal ? cliVal.valorUnit : null;
  } else {
    if (grupo.precoMedioPeso > 0) {
      precoMedioClientes = grupo.precoMedioSomaPonderada / grupo.precoMedioPeso;
    } else {
      precoMedioClientes = menorValorUnit;
    }
  }

  const precoMedBox = calcularPrecoMedBox(grupo.skuMaster, grupo.precoCusto, markup);
  const statusCadastro = grupo.skuMaster ? 'ATIVO' : 'PENDENTE';
  const statusPreco = grupo.precoCusto ? 'CONCLUÍDO' : 'SEM CUSTO';

  // Format clientesValores to follow client contract exactly
  const formattedClientesValores = {};
  Object.keys(grupo.clientesValores).forEach(c => {
    const val = grupo.clientesValores[c];
    formattedClientesValores[c] = {
      quantidadeAtual: val.quant || 0,
      valorUnitarioAtual: val.valorUnit || null,
      demandaAtual: val.demandaR || 0,
      dataUltimaCotacao: null, // Placeholder
      mediaMensalQuantidade: null, // Placeholder
      mediaMensalValorUnitario: null // Placeholder
    };
  });

  return {
    categoria: grupo.categoria,
    skuMaster: grupo.skuMaster,
    produto: grupo.produto,
    fabricante: grupo.fabricanteCompra2026,
    fornecedor: grupo.fornecedorCompra2026,
    precoCustoAtual: grupo.precoCusto || null,
    precoMedbox: precoMedBox || null,
    dataUltimoPreco: grupo.dataCompra2026 ? new Date(grupo.dataCompra2026).toISOString() : null,
    clientesValores: formattedClientesValores,
    statusCadastro: statusCadastro,
    statusPreco: statusPreco,

    // keeps compatibility for sums
    quantTotal: quantTotal,
    precoCusto: grupo.precoCusto,
    precoMedbox: precoMedBox,
    demandaR: demandaR
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
  const skuMaster = upperText(f2dGetMapCell_(row, idx, 'skuMaster'));
  const statusCadastro = skuMaster ? 'ATIVO' : 'PENDENTE';
  const statusPreco = precoCusto ? 'CONCLUÍDO' : 'SEM CUSTO';
  const dataCompra = f2dGetMapCell_(row, idx, 'dataCompra');

  return {
    categoria: upperText(f2dGetMapCell_(row, idx, 'categoria')),
    produto: upperText(f2dGetMapCell_(row, idx, 'produto')),
    skuMaster: skuMaster,
    fabricante: upperText(f2dGetMapCell_(row, idx, 'fabricante')),
    fornecedor: upperText(f2dGetMapCell_(row, idx, 'fornecedor')),
    precoCustoAtual: precoCusto || null,
    precoMedbox: precoMedBox || null,
    dataUltimoPreco: dataCompra ? new Date(dataCompra).toISOString() : null,
    clientesValores: {},
    statusCadastro: statusCadastro,
    statusPreco: statusPreco,
    ultimaQuantidadeCotada: null, // Placeholder
    ultimoValorUnitarioCotado: null, // Placeholder
    clientesComDemanda: upperText(f2dGetMapCell_(row, idx, 'clientesComConsumo')),

    // keeps compatibility
    quantTotal: toNumber(f2dGetMapCell_(row, idx, 'quantTotal')) || 0,
    menorValorUnit: toNumber(f2dGetMapCell_(row, idx, 'menorValorUnit'))
  };
}

function gerarSugestoesSkuBusca(limitados, idx) {
  const mapa = new Map();

  for (const obj of limitados) {
    const row = obj.row;

    const skuMaster = upperText(f2dGetMapCell_(row, idx, 'skuMaster'));
    const skuOperacional = upperText(f2dGetMapCell_(row, idx, 'skuOperacional'));
    const produto = upperText(f2dGetMapCell_(row, idx, 'produto'));
    const qtd = toNumber(f2dGetMapCell_(row, idx, 'quantTotal')) || 0;

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
    'fabricante',
    'fornecedor',
    'produto',
    'produtoNormalizado',
    'skuMaster',
    'skuOperacional'
  ];

  const partes = [];

  campos.forEach(function(nomeColuna) {
    const valor = f2dGetMapCell_(row, idx, nomeColuna);

    if (valor !== null && valor !== undefined && String(valor).trim() !== '') {
      partes.push(String(valor));
    }
  });

  return normalizarTexto(partes.join(' '));
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
  try {
    const ss = obterPlanilhaMedbox_();
    const abasClientes = f2dObterAbasClientes_(ss);
    const clientes = abasClientes.map(sheet => {
      const name = sheet.getName();
      return upperText(name.substring(3).trim());
    });
    const uniqueClientes = Array.from(new Set(clientes)).filter(Boolean).sort();
    if (uniqueClientes.length > 0) {
      return uniqueClientes;
    }
  } catch (e) {
    Logger.log("Erro ao carregar clientes dinamicamente: " + e.message);
  }

  if (headers && headers.length > 0) {
    return headers
      .filter(h => String(h).trim().endsWith(' Quant'))
      .map(h => upperText(String(h).replace(' Quant', '').trim()))
      .filter(Boolean)
      .sort();
  }
  return [];
}

function f2dObterAbasClientes_(ss) {
  const sheets = ss.getSheets();
  const abasClientes = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const upperName = name.trim().toUpperCase();

    // Deve começar com BD_
    if (!upperName.startsWith("BD_")) return;

    // Excluir abas de sistema e referência técnica por nome normalizado
    // Deve ocorrer antes da validação de schema
    // Termos longos e específicos: substring no nome normalizado
    const longTerms = [
      "COTACOES", "COTAÇÕES",
      "SKU_MASTER", "SKU MASTER",
      "PRECOS", "PREÇOS",
      "SEFAZ",
      "ADMIN", "BACKUP",
      "REVISAO", "REVISÃO",
      "RELATORIO", "RELATÓRIO"
    ];
    for (const term of longTerms) {
      if (upperName.indexOf(term) !== -1) {
        return;
      }
    }
    // Termos curtos e genéricos: correspondência por token completo
    const shortTerms = ["LOG", "AUX", "MAPA"];
    const tokenName = upperName
      .replace(/[_\-.\/\\]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const tokens = tokenName.split(" ");
    for (const term of shortTerms) {
      if (tokens.indexOf(term) !== -1) {
        return;
      }
    }

    // Se passar nas exclusões, verificar se tem estrutura válida
    const detection = f2dDetectarCabecalhoEColunas_(sheet);
    if (detection !== null) {
      abasClientes.push(sheet);
    }
  });

  return abasClientes;
}

function f2dDetectarCabecalhoEColunas_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return null;

  const range = sheet.getRange(1, 1, Math.min(lastRow, 5), lastCol);
  const values = range.getValues();

  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(String);

    // 1. Look for SKU columns in priority order: SKU_MASTER, SKU_OPERACIONAL, SKU
    const colSkuMaster = f2dEncontrarColunaFlex_(row, ['SKU_MASTER', 'SKU MASTER']);
    const colSkuOperacional = f2dEncontrarColunaFlex_(row, ['SKU_OPERACIONAL', 'SKU OPERACIONAL']);
    const colSkuGeneric = f2dEncontrarColunaFlex_(row, ['SKU']);

    // 2. Look for Product/Description column
    const colProduto = f2dEncontrarColunaFlex_(row, ['PRODUTOS', 'PRODUTO', 'DESCRICAO DO PRODUTO', 'DESCRIÇÃO DO PRODUTO', 'DESCRICAO', 'DESCRIÇÃO', 'PRODUTO NORMALIZADO', 'DESCRICAO_NORMALIZADA']);

    // 3. Look for Quantity column (REQUIRED)
    const colQtd = f2dEncontrarColunaFlex_(row, ['QUANTIDADE', 'QUANT', 'QTD', 'QTDE', 'QUANT_TOTAL', 'QUANT. TOTAL']);

    // 4. Look for Unit Price column (OPTIONAL)
    const colVal = f2dEncontrarColunaFlex_(row, ['PRECO_REFERENCIA_CLIENTE', 'PRECO_REFERENCIA', 'VALOR_UNITARIO', 'VALOR UNITARIO', 'VALOR_UNIT', 'VALOR UNIT', 'PRECO', 'PRECO_MEDBOX', 'MENOR_VALOR', 'MENOR PRECO', 'PRECO_UNITARIO', 'PRECO UNITARIO', 'V UNIT', 'V. UNIT']);

    // Determine target column of identification according to required priority
    let colId = -1;
    let tipoId = "";

    if (colSkuMaster !== -1) {
      colId = colSkuMaster;
      tipoId = "SKU_MASTER";
    } else if (colSkuOperacional !== -1) {
      colId = colSkuOperacional;
      tipoId = "SKU";
    } else if (colSkuGeneric !== -1) {
      colId = colSkuGeneric;
      tipoId = "SKU";
    } else if (colProduto !== -1) {
      colId = colProduto;
      tipoId = "PRODUTO";
    }

    if (colId !== -1 && colQtd !== -1) {
      return {
        headerRowIndex: r,
        colId: colId,
        colSkuMaster: colSkuMaster,
        colProduto: colProduto,
        tipoId: tipoId,
        colQtd: colQtd,
        colVal: colVal
      };
    }
  }
  return null;
}

function f2dCarregarDadosClienteSheet_(sheet) {
  const skuMap = {};
  const productMap = {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      tipoId: "",
      headerRowIndex: -1,
      colId: -1,
      colQtd: -1,
      colVal: -1,
      validRowsCount: 0,
      ignoredRowsCount: 0,
      skuMap: skuMap,
      productMap: productMap
    };
  }

  const detection = f2dDetectarCabecalhoEColunas_(sheet);
  if (!detection) {
    return {
      tipoId: "",
      headerRowIndex: -1,
      colId: -1,
      colQtd: -1,
      colVal: -1,
      validRowsCount: 0,
      ignoredRowsCount: 0,
      skuMap: skuMap,
      productMap: productMap
    };
  }

  const result = {
    tipoId: detection.tipoId,
    headerRowIndex: detection.headerRowIndex,
    colId: detection.colId,
    colQtd: detection.colQtd,
    colVal: detection.colVal,
    validRowsCount: 0,
    ignoredRowsCount: 0,
    skuMap: skuMap,
    productMap: productMap
  };

  const values = sheet.getDataRange().getValues();
  const headers = values[detection.headerRowIndex].map(String);

  // Mapear colunas exatas do cliente
  const colSkuMaster = f2dEncontrarColunaFlex_(headers, ['SKU_MASTER', 'SKU MASTER']);
  const colProdutos = f2dEncontrarColunaFlex_(headers, ['Produtos', 'Produto', 'DESCRICAO']);
  const colQuant = f2dEncontrarColunaFlex_(headers, ['Quant', 'QUANTIDADE', 'QTD']);
  const colVUnit = f2dEncontrarColunaFlex_(headers, ['V. Unit', 'VALOR_UNITARIO', 'V UNIT', 'PRECO']);
  const colVTotal = f2dEncontrarColunaFlex_(headers, ['V. Total', 'VALOR_TOTAL', 'V TOTAL', 'VALOR TOTAL']);

  const startRow = detection.headerRowIndex + 1;

  for (let i = startRow; i < values.length; i++) {
    const row = values[i];

    // Ler Quant
    const q = colQuant !== -1 ? toNumber(row[colQuant]) : null;
    if (q === null || q <= 0) {
      result.ignoredRowsCount++;
      continue;
    }

    // Ler V. Unit
    const vUnit = colVUnit !== -1 ? toNumber(row[colVUnit]) : null;

    // Ler V. Total
    let vTotal = colVTotal !== -1 ? toNumber(row[colVTotal]) : null;
    if (vTotal === null && q !== null && vUnit !== null) {
      vTotal = q * vUnit;
    }

    // Usar SKU_MASTER quando preenchido, Produtos quando vazio, nunca usar Código
    const skuMasterVal = colSkuMaster !== -1 ? String(row[colSkuMaster]).trim() : "";
    const produtoVal = colProdutos !== -1 ? String(row[colProdutos]).trim() : "";

    let keySku = "";
    let keyProduct = "";

    if (skuMasterVal) {
      keySku = normalizarTexto(skuMasterVal);
    } else if (produtoVal) {
      keyProduct = normalizarTexto(produtoVal);
    }

    if (!keySku && !keyProduct) {
      result.ignoredRowsCount++;
      continue;
    }

    result.validRowsCount++;

    const targetMap = keySku ? result.skuMap : result.productMap;
    const targetKey = keySku || keyProduct;

    if (!targetMap[targetKey]) {
      targetMap[targetKey] = {
        quant: 0,
        valorUnit: null,
        demandaR: 0
      };
    }

    targetMap[targetKey].quant += q;
    if (vUnit !== null && vUnit > 0) {
      if (targetMap[targetKey].valorUnit === null || vUnit < targetMap[targetKey].valorUnit) {
        targetMap[targetKey].valorUnit = vUnit;
      }
    }
    if (vTotal !== null && vTotal > 0) {
      targetMap[targetKey].demandaR += vTotal;
    }
  }

  // Se demandaR não foi preenchida ou é 0, calcula como quant * valorUnit
  [result.skuMap, result.productMap].forEach(m => {
    Object.keys(m).forEach(key => {
      const item = m[key];
      if (item.demandaR === 0 && item.quant > 0 && item.valorUnit !== null) {
        item.demandaR = item.quant * item.valorUnit;
      }
    });
  });

  return result;
}

function f2dEncontrarColunaFlex_(headers, candidates) {
  const headersNorm = headers.map(h => String(h).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
  for (const candidate of candidates) {
    const candidateNorm = candidate.toUpperCase().replace(/[^A-Z0-9]/g, '');
    const idx = headersNorm.indexOf(candidateNorm);
    if (idx !== -1) {
      const orig = headers[idx].toUpperCase();
      if (candidateNorm.indexOf('SKU') !== -1 && (orig.indexOf('CODIGO') !== -1 || orig.indexOf('CÓDIGO') !== -1)) {
        continue;
      }
      return idx;
    }
    for (let i = 0; i < headersNorm.length; i++) {
      if (headersNorm[i] !== "" && (headersNorm[i].indexOf(candidateNorm) !== -1 || candidateNorm.indexOf(headersNorm[i]) !== -1)) {
        const orig = headers[i].toUpperCase();
        if (candidateNorm.indexOf('SKU') !== -1 && (orig.indexOf('CODIGO') !== -1 || orig.indexOf('CÓDIGO') !== -1)) {
          continue;
        }
        return i;
      }
    }
  }
  return -1;
}

function f2dMapearColunasMapaErp_(headers) {
  const map = {
    categoria: f2dEncontrarColunaFlex_(headers, ['Categoria']),
    produto: f2dEncontrarColunaFlex_(headers, ['Produtos', 'Produto']),
    skuMaster: f2dEncontrarColunaFlex_(headers, ['SKU_MASTER', 'SKU MASTER']),
    skuOperacional: f2dEncontrarColunaFlex_(headers, ['SKU_OPERACIONAL', 'SKU OPERACIONAL']),
    produtoNormalizado: f2dEncontrarColunaFlex_(headers, ['Produto Normalizado', 'DESCRICAO_NORMALIZADA']),
    custo: f2dEncontrarColunaFlex_(headers, ['Custo Compra 2026']),
    fornecedor: f2dEncontrarColunaFlex_(headers, ['Fornecedor Compra 2026']),
    fabricante: f2dEncontrarColunaFlex_(headers, ['Fabricante Compra 2026']),
    dataCompra: f2dEncontrarColunaFlex_(headers, ['Data Compra 2026']),
    quantTotal: f2dEncontrarColunaFlex_(headers, ['Quant. total', 'Quant total', 'QUANTIDADE', 'QUANT']),
    precoMedioClientes: f2dEncontrarColunaFlex_(headers, ['Preço Médio Clientes']),
    menorValorUnit: f2dEncontrarColunaFlex_(headers, ['Menor Valor Unit']),
    clientesComConsumo: f2dEncontrarColunaFlex_(headers, ['Clientes com Consumo', 'Clientes com Consumo Set'])
  };
  return map;
}

function f2dGetMapCell_(row, map, colName) {
  const idx = map[colName];
  if (idx === undefined || idx === null || idx < 0) return '';
  return row[idx];
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
