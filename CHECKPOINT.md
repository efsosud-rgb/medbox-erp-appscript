# CHECKPOINT

## MVP Emergencial — Dados Reais das Bases BD_ (branch fix/mvp-dados-reais)

- **Objetivo emergencial:** substituir a camada de dados de Busca Avançada, Mapa de Produtos e Dashboard por um motor que lê exclusivamente `BD_SKU MASTER`, `BD_PRECOS` e as abas `BD_<CLIENTE>`, sem depender de `MAPA_ERP` (diagnosticada com estrutura não confiável) nem de `AUX_BUSCA`.
- **Motivo:** `MAPA_ERP`, na Planilha Oficial centralizada, estava retornando linhas inválidas (títulos/rótulos como "CATEGORIA", "TOTAL SKU MASTER" tratados como produtos) e zerando quantidade/valor/SKU em todas as telas — bloqueador para apresentação.
- **Prazo:** apresentação no dia seguinte à criação deste checkpoint.
- **Branch:** `fix/mvp-dados-reais` (a partir de `main` @ `34e458b`). Não integrada na `main` nesta etapa.

### Arquitetura do MVP
- **Arquivo novo:** `DadosMVP.js` — único motor de consolidação, somente leitura.
  - `mvpConsolidarDados_(markup)`: ponto de entrada único; abre a planilha oficial via `obterPlanilhaMedbox_()` (reaproveitado de `Codigo.js`), carrega `BD_SKU MASTER`, `BD_PRECOS` e as abas de cliente, e devolve `{ clientes, produtos, atualizadoEm }`.
  - `mvpDescobrirAbasClientes_(ss)`: prefixo `BD_` obrigatório + exclusão por termos longos (substring) e termos curtos `LOG/AUX/MAPA` (token completo, evita falso positivo em nomes como Cardiologia/Auxiliadora/Amapá) + validação de schema (`Produtos` + `Quant` + `V. Unit` ou `V. Total`) antes de aceitar a aba.
  - `mvpDetectarCabecalhoCliente_(sheet)`: varre as 5 primeiras linhas à procura do cabeçalho real (não assume linha 1), replicando a proteção que faltava em `MAPA_ERP`.
  - `mvpCarregarSkuMaster_(ss)`: le `BD_SKU MASTER` A:C, indexa por descrição normalizada e por SKU Master normalizado.
  - `mvpCarregarPrecos_(ss)`: le `BD_PRECOS` A:G, indexa por SKU (prioridade) ou descrição, mantendo o menor Valor Unitário positivo.
  - `mvpCarregarAbaCliente_`, `mvpCriarProdutoBase_`, `mvpAcumularProduto_`, `mvpAplicarPreco_`, `mvpFinalizarProduto_`: pipeline de consolidação por produto/cliente.
  - `mvpNormalizarTexto_`, `mvpConverterNumero_`, `mvpCriarChaveProduto_`, `mvpEncontrarColuna_`: utilitários somente leitura, sem fuzzy matching.
  - Funções públicas do motor: `mvpObterOpcoesSistema()`, `mvpPesquisarBuscaAvancada(filtros)`, `mvpObterMapaDeProdutos(filtros)`, `mvpCalcularDashboard()`.
- **Reaproveitado sem duplicar:** `obterPlanilhaMedbox_()`, `normalizarMarkup()`, `normalizarLimite()`, `arredondar()`, `calcularPrecoMedBox()`, `formatarDataHora()`, `upperText()` (já existentes em `Codigo.js`, genéricos, não leem `MAPA_ERP`).

### Funções públicas redirecionadas (nomes preservados, contrato do HTML preservado)
- `Codigo.js :: obterOpcoesSistema()` → `mvpObterOpcoesSistema()`.
- `Codigo.js :: pesquisarBuscaAvancada(filtros)` → `mvpPesquisarBuscaAvancada(filtros)`.
- `Codigo.js :: obterMapaDeProdutos(filtros)` → `mvpObterMapaDeProdutos(filtros)`.
- `Fase2_Dashboard.js :: calcularDashboardFase2()` → `mvpCalcularDashboard()`.
- Nenhuma dessas quatro funções lê `MAPA_ERP`/`ABA_MAPA`/`AUX_BUSCA` (confirmado por `git grep`). `MedboxERP.html` não foi alterado.

### Campos reais entregues
categoria, produto, skuMaster, descricaoOrigem, unidade, precoCustoAtual/precoCusto, precoMedbox, fornecedor, fabricante, quantidadeTotalDemandada/quantTotal, valorTotalDemandado/demandaR, clientesComDemanda, clientesValores (quantidadeAtual, valorUnitarioAtual, demandaAtual), statusCadastro, statusPreco; no Dashboard: totalSkuMaster, totalProdutos, totalClientes, quantidadeTotalDemandada, valorTotalDemandado, itensComPrecoCusto, itensSemPreco, itensSemSkuMaster, topQuantidade, topValor, topClientes, semSku, semPreco.

### Campos ainda placeholder (sem base histórica no MVP)
dataUltimaCotacao, mediaMensalQuantidade, mediaMensalValorUnitario (por cliente), dataUltimoPreco/ultimaQuantidadeCotada/ultimoValorUnitarioCotado (Busca), itensCotadosNoMes, valorCotadoNoMes, ultimosPrecos, ultimasCotacoes (Dashboard) — todos retornados como `0`/`null`/`[]`, nunca inventados.

### Testes
- 30 testes locais em Node puro (sem tocar a planilha): conversão pt-BR de moeda, normalização de acentos/espaços, chave de produto, busca de coluna por cabeçalho, produto com SKU preenchido, correlação por `BD_SKU MASTER`, produto sem correlação, menor preço positivo entre dois registros, soma de duas linhas do mesmo cliente sem sobrescrever, dois clientes no mesmo produto, regra de Preço MedBox = Custo × (1 + Margem/100). **0 falhas em 30.**
- Leitura real contra a planilha (contagens de abas/linhas/produtos): **não executada nesta etapa** — bloqueada porque `clasp push`/publicação estão proibidos nesta fase, e a sessão de diagnóstico anterior já havia confirmado `clasp run` indisponível por erro de permissão mesmo para funções já publicadas.

### Estado da publicação
Nenhum `clasp push`, criação de versão, atualização de deployment ou merge na `main` foi executado. Trabalho isolado na branch `fix/mvp-dados-reais`.

### Próxima ação
Validar manualmente no Editor do Apps Script (ou resolver o bloqueio de permissão do `clasp run`) rodando `mvpConsolidarDados_()`/`calcularDashboardFase2()` contra a planilha real antes de publicar, para confirmar contagens reais de clientes, produtos, SKUs e valores.

---

## Reestruturação e Integração do Mapa de Clientes (Fase 2026)

- **Objetivo:** Adaptar o módulo do "MAPA DE PRODUTOS" para ler de forma centralizada e dinâmica os dados de consumo de clientes de suas respectivas abas na Planilha Oficial, sem alterar o design visual aprovado do `MedboxERP.html`.

### Fontes de Dados
- **Fonte Anterior (Legada):** `SpreadsheetApp.getActiveSpreadsheet()` (planilha container-bound), com colunas de clientes hardcoded inseridas diretamente na aba `MAPA_ERP` (como `<Cliente> Quant` e `<Cliente> Valor Unit`).
- **Nova Fonte Oficial:** Planilha `MEDBOX_MAPA_CLIENTES_2026` centralizada com o ID `1QLurrnKCXorjXUoUtPl6MoSP0Gk32Yz0AWBHxebifCw` (Aba central `MAPA_ERP` e as abas de cotação individuais `BD_<CLIENTE>`).

### Arquivos e Funções Alteradas
- **`Codigo.js` (Backend Principal):**
  - Adicionada constante: `MEDBOX_SPREADSHEET_ID` para centralizar o ID da planilha.
  - Adicionadas funções: `obterPlanilhaMedbox_()` para centralizar o acesso.
  - Atualizada `obterOpcoesSistema()`: para carregar as opções de filtros lendo da planilha oficial por ID.
  - Atualizada `pesquisarBuscaAvancada(filtros)`: para pesquisar produtos na planilha centralizada.
  - Atualizada `obterMapaDeProdutos(filtros)`: para realizar a descoberta de abas baseada na `LISTA_CLIENTES_OFICIAIS`, mapeamento flexível de colunas por cliente e consolidação na memória para o Mapa de Produtos.
  - Atualizada `obterClientesDoMapa(headers)`: agora utiliza a descoberta das abas oficiais configuradas com fallback seguro para colunas da planilha.
  - Adicionada `f2dObterAbasClientes_(ss)`: para retornar apenas as abas presentes na configuração `LISTA_CLIENTES_OFICIAIS` que passarem na validação estrutural de colunas.
  - Adicionada `f2dDetectarCabecalhoEColunas_(sheet)`: para realizar a descoberta dinâmica de cabeçalhos de clientes nas primeiras 5 linhas de cada aba de dados, validando que contenham colunas compatíveis com SKU/produto e quantidade.
  - Adicionada `f2dCarregarDadosClienteSheet_(sheet)`: para carregar quantidade e menor valor unitário por SKU para cada cliente, suportando cabeçalhos na linha 1, 2 ou 3 (com valor unitário opcional).
  - Adicionada `f2dEncontrarColunaFlex_(headers, candidates)`: para mapeamento flexível (fuzzy matching) de cabeçalhos de clientes.
  - Atualizada `acumularGrupoProduto()`: suporta o fluxo de mesclagem baseado em consumos por cliente.
  - Atualizada `finalizarGrupoProduto()`: suporta o filtro dinâmico de cliente e a consolidação de "TODOS".
- **`Fase2_Dashboard.js`:**
  - Atualizadas referências à planilha ativa para utilizar `obterPlanilhaMedbox_()`.
  - Revertidas mensagens e alertas modificados acidentalmente para inglês.
- **`Diagnostico.js`:**
  - Adicionada a função `diagnosticarMapaDeProdutosEClientes()` que realiza diagnóstico estrutural completo e simula o processamento do Mapa de Produtos em modo somente leitura (sem gravação).

### Clientes Detectados & Abas Usadas
- **Clientes Detectados Dinamicamente baseados na Configuração Oficial:**
  - `VivaRio` (aba `BD_VivaRio`)
  - `AFNE` (aba `BD_AFNE`)
  - `INTS UPA Santo Amaro` (aba `BD_INTS UPA Santo Amaro`)
  - `INTS Hospital Dia Cidade Ade` (aba `BD_INTS Hospital Dia Cidade Ade`)
  - `INTS Hospital Dia Santo Amar` (aba `BD_INTS Hospital Dia Santo Amar`)
  - `INTS Hospital Guarapiranga` (aba `BD_INTS Hospital Guarapiranga`)
  - `Solar` (aba `BD_Solar`)
  - `IGEDES` (aba `BD_IGEDES`)
  - `SPDM` (aba `BD_SPDM`)
- **Abas do Sistema Usadas:**
  - `MAPA_ERP` (aba central de produtos com seus SKUs Masters)

### Testes Executados
- **Verificação de Sintaxe:** Executado `node -c Codigo.js Fase2_Dashboard.js Diagnostico.js` com sucesso (zero erros de sintaxe).
- **Testes de Unidade Locais:**
  - Validada a detecção dinâmica de cabeçalhos compatíveis na linha 1, 2 ou 3.
  - Validada a consolidação do Mapa de Produtos no modo "TODOS" (soma das quantidades, menor valor unitário dos clientes e demanda total calculada corretamente).
  - Validada a filtragem individual por cliente (exibição apenas de linhas do cliente, soma local e cards de resumo atualizados perfeitamente).

### Riscos e Pendências
- **Garantias de Performance:** Como a planilha oficial possui mais de 1.800 SKUs e dezenas de abas dinâmicas de clientes, o uso de `sheet.getDataRange().getValues()` em lote é essencial. O código implementado foi otimizado para carregar os dados de cada aba de uma só vez (batch loading), minimizando o risco de exceder o limite de execução de 6 minutos do Google Apps Script.
- **Visualização de Layout:** A interface `MedboxERP.html` e a navegação permanecem intocadas, preservando com 100% de exatidão o visual aprovado anterior.

---

## Integracao Cosmos

- Objetivo: validar a integracao Cosmos do MEDBOX ERP em modo somente leitura.
- OAuth e `clasp run` estao configurados para execucao em dev e no Executavel da API.
- Project ID: `hidden-terrain-496418-c4`.
- Deployment correto do Executavel da API: `AKfycbweUomzdMhwsVlBlB8V4Fo4xt4JsGLCUPZ_pWojHQXL_cSadEgCwVCUKHdKA45C_en6`.
- `verificarTokenCosmosConfigurado` foi validada em dev e non-dev.
- `testarConexaoCosmosSomenteLeitura` foi validada em dev e non-dev.
- `consultarProdutoCosmosPorDescricaoSeguro` foi criada e validada em dev e non-dev.
- `consultarProdutoCosmosPorGtinSeguro` foi criada com validacao local; consulta real por GTIN nao foi executada porque nao ha GTIN publico confirmado no projeto.
- Endpoint logico validado: `products_query`.
- Metodo HTTP: `GET`.
- Autenticacao Cosmos: cabecalho `X-Cosmos-Token`.
- `User-Agent` obrigatorio: `Cosmos-API-Request`.
- Campos reais retornados de forma sanitizada pela consulta Cosmos: `descricao`, `marca`, `gtin`, `ncm`.
- Validacoes implementadas: descricao vazia, descricao somente com espacos, parametros ausentes, GTIN pontuado normalizado e GTIN invalido rejeitado.
- Testes executados: validacoes sem API, consulta por descricao em dev e consulta por descricao em non-dev.
- As funcoes remotas de consulta nao usam `SpreadsheetApp.getUi()` e nao gravam dados em planilhas.
- Nunca expor, registrar, retornar, medir ou gerar hash de `RFC_COSMOS_TOKEN`.
- Operacoes fiscais de escrita, emissao, cancelamento, envio ou alteracao de documentos sao proibidas sem autorizacao explicita.
- Versao Apps Script validada: `20`.
- Deployment validado: `AKfycbweUomzdMhwsVlBlB8V4Fo4xt4JsGLCUPZ_pWojHQXL_cSadEgCwVCUKHdKA45C_en6 @20`.
- Proximo passo recomendado: criar tela de revisao e aprovacao antes de alimentar o cadastro interno.

## Revisao e aprovacao Cosmos

- Tela HTML criada: `RevisaoCosmosUI.html`.
- Backend criado: `RevisaoCosmos.js`.
- Funcao de abertura no Google Sheets: `abrirRevisaoCosmos`.
- Menu MEDBOX ERP atualizado com submenu `Cosmos` e opcao `Revisar e aprovar resultados`.
- Aba de origem identificada: `BD_PRECOS`.
- Identificador estavel usado: `ID_PRECO`.
- Fila tecnica de revisao: `REVISAO_COSMOS`.
- Relatorio append-only: `RELATORIO_APROVACAO_COSMOS`.
- Backup append-only antes de aprovacao: `BACKUP_APROVACAO_COSMOS`.
- Campos exibidos/editaveis na tela: marca aprovada, GTIN aprovado e NCM aprovado.
- Descricao sugerida pela Cosmos e exibida somente como referencia visual; nao e editavel e nao e aplicada ao cadastro.
- Nao foi encontrada coluna real e apropriada em `BD_PRECOS` para descricao padronizada/corrigida (`DESCRICAO_PADRONIZADA`, `DESCRICAO_CORRIGIDA`, `DESCRICAO_CADASTRO` ou `DESCRICAO_PRODUTO`).
- Campos aplicados no cadastro apos aprovacao explicita: `FABRICANTE_MARCA`, `EAN`, `NCM`.
- `DESCRICAO_FORNECEDOR` e `SKU_MASTER` nao sao sobrescritos pela revisao Cosmos.
- A interface identifica o campo de destino de cada valor aprovavel: `FABRICANTE_MARCA`, `EAN` e `NCM`.
- Limpeza de campo preenchido exige confirmacao especifica na aprovacao individual; aprovacao em lote nao permite apagamento silencioso.
- Funcao de simulacao criada: `simularAprovacaoRevisaoCosmos`, reutilizando as validacoes sem escrever em `BD_PRECOS`, `REVISAO_COSMOS`, backup ou relatorio.
- Itens rejeitados apenas mudam status para `REJEITADO` na fila de revisao e nao alteram `BD_PRECOS`.
- Itens pendentes nao alteram `BD_PRECOS`.
- Aprovacao usa `LockService`, `REVISAO_ID`, `ITEM_ID`, status atual e `VERSAO_REGISTRO` para evitar duplicidade e tela desatualizada.
- Revalidacao antes da escrita: a origem e localizada novamente por `BD_PRECOS.ID_PRECO`; `LINHA_ORIGEM` e apenas referencia visual.
- Validacoes implementadas: payload ausente, revisao ausente, item ausente, status nao pendente, tela desatualizada, GTIN 8/12/13/14, GTIN invalido, NCM com 8 digitos, NCM invalido, limpeza sem confirmacao, item movido/removido e tentativa duplicada.
- Teste executado sem gravacao: `testarRevisaoCosmosSemGravacao`.
- Validacao visual e aprovacao/rejeicao real permanecem pendentes para evitar teste sobre produto real sem autorizacao explicita.
- Limitacao: a fila `REVISAO_COSMOS` e criada/sincronizada ao carregar a tela a partir de sugestoes Cosmos existentes em `BD_PRECOS`.
- Proximo passo: autorizar um unico item de teste controlled ou copia controlada para validar carregamento visual, simulacao, aprovacao, rejeicao, backup e relatorio automatico ponta a ponta.

## Extracao e revisao de cotacao

- A extracao por cliente em `BD_COTACOES` preserva `DESCRICAO_CLIENTE` como descricao original, sem sobrescrever esse campo.
- A grade comercial exibe codigo, descricao original, quantidade, valor unitario de compra do cliente, valor total do cliente, `SKU_MASTER`, custo MedBox, fornecedor/data do custo, Margem %, preco de venda MedBox, valor total MedBox e status do lance.
- Diagnosticos tecnicos de match (`statusMatch`, score e motivo de revisao) continuam no backend e podem aparecer como detalhe/tooltip, mas nao ocupam colunas permanentes da grade principal.
- Regra de negocio: o rotulo exibido na grade e `Margem`, mas o calculo interno preserva markup sobre custo.
- Regra comercial aplicada na grade: `Preco de Venda MedBox = Custo MedBox * (1 + Margem / 100)`.
- Valor total do cliente e valor total MedBox sao calculados por quantidade; quando o arquivo traz apenas quantidade e valor total, o valor unitario do cliente e derivado por `valor total / quantidade` somente se a quantidade for numerica e maior que zero.
- O preco total do cliente nao deve ser usado como preco unitario.
- Status do lance comercial: `LANCE VENCEDOR`, `PRECO IGUAL`, `PRECO ACIMA` ou `DADOS INSUFICIENTES`.
- Quantidades na grade usam formato pt-BR sem moeda, inteiras sem casas decimais e fracionadas com ate 3 casas.
- Quantidade vem de uma unica coluna mapeada; quando houver mais de uma candidata, a escolha prioriza `QUANTIDADE`, `QUANT`, `QTD` ou `QTDE` e a proporcao de celulas numericas positivas.
- Valores ambiguos de quantidade com dois numeros separados por espaco, como `0 1` ou `30 100`, sao rejeitados como `QUANTIDADE_AMBIGUA` e nao sao interpretados por tentativa.
- Valores financeiros usam formato pt-BR `R$ 1.234,56`; valores vazios ou invalidos aparecem como `—`.
- Resumo comercial da cotacao exibe: total de itens cotados, valor total cotado, total de itens ganhos, valor total ganho MedBox, percentual de itens ganhos e itens sem preco.
- `TOTAL DE ITENS COTADOS` e a quantidade de linhas comerciais validas extraidas, independente de SKU, custo, preco MedBox ou status do lance.
- `ITENS SEM PRECO` contabiliza itens validos sem `precoVendaMedBox` maior que zero.
- Itens ganhos no resumo consideram somente status `LANCE VENCEDOR`; `PRECO IGUAL`, `PRECO ACIMA`, `DADOS INSUFICIENTES` e `CUSTO_NAO_ENCONTRADO` nao entram como ganhos.
- Status de match implementados: `MATCH_CONFIRMADO`, `REVISAR`, `NAO_ENCONTRADO`, `CONFLITO_TECNICO`.
- O SKU sugerido usa catalogo existente em `MAPA_ERP` e `BD_SKU_MASTER_SEFAZ`, com normalizacao e validacao de atributos tecnicos como concentracao, volume, quantidade, forma e apresentacao.
- Custos sao buscados em `BD_PRECOS` apenas por `SKU_MASTER` exato e somente quando positivos; custos vazios, zero ou negativos sao rejeitados.
- Se nao houver custo valido, a grade mostra `CUSTO_NAO_ENCONTRADO`, deixa o custo vazio e nao calcula preco de venda.
- Formula de preco de venda reaproveitada: `CUSTO * (1 + MARKUP)`.
- Alteracoes de SKU, custo ou margem recalculam preco de venda e status do lance na grade antes da gravacao.
- Nenhum SKU, custo, margem ou preco e gravado antes de salvamento explicito da cotacao.
- Nesta etapa, `SALVAR COTACAO` executa `simularSalvarCotacaoExtraida` para validar a operacao sem escrever em `BD_COTACOES`; a gravacao real permanece pendente de autorizacao controlada.
- `IMPRIMIR COTACAO` e `SALVAR EM PDF` ficam desabilitados ate haver cotacao salva no estado da tela.
- O documento comercial de impressao/PDF nao exibe custo MedBox, fornecedor de custo, score, motivo de revisao, margem ou preco de compra do cliente.
- Na confirmacao, `BD_COTACOES` recebe exatamente o SKU, custo, margem e preco revisados; `DESCRICAO_CLIENTE` nao e alterada.
- Relatorio append-only criado na confirmacao: `RELATORIO_IMPORTACAO_COTACAO`.
- Teste sem gravacao executado: `testarExtracaoCotacaoSemGravacao`.
- Testes sem gravacao adicionados: `testarCalculosCotacaoSemGravacao`, `testarDocumentoCotacaoSemDadosInternos` e `simularSalvarCotacaoExtraida`.
- Teste sem backend para remocao isolada de arquivo: `testarRemocaoArquivoImportacaoSemBackend`.
- Botao dedicado `REMOVER ARQUIVO` limpa apenas o arquivo selecionado, nome/tamanho, mensagens e previa, sem chamar backend e sem limpar campos manuais.
- A extracao direta por arquivo usa `FileReader` no Web App e chama `extrairCotacaoArquivoSemGravacao`, sem exigir cliente previamente cadastrado.
- O fluxo antigo por cliente foi separado como consulta de cotacoes ja cadastradas.
- A extracao XLSX/CSV para cotacao agora detecta a tabela real por cabecalhos comerciais, mapeia colunas por nome e ignora instrucoes, paineis, filtros, nomes de projeto isolados e linhas fora da tabela.
- Diagnostico seguro retornado na previa: aba selecionada, linha de cabecalho, colunas mapeadas, linhas lidas, itens validos, linhas ignoradas e motivos agregados.
- Teste sem gravacao executado: `testarFiltroLinhasNaoProdutosSemGravacao`.
- Proximo passo: validar visualmente a grade com o XLSX controlado no Web App e executar confirmacao somente com um item de teste controlado.

---

## Correção de Exclusão das Abas de Sistema (commit 36d6862)

- **Origem:** Commit Jules 36d6862 — refatoração do módulo de mapa de produtos com integração dinâmica de clientes 2026.
- **Correção:** Substituída a exclusão parcial de abas de sistema por uma lista unificada de termos de exclusão por substring do nome normalizado, executada antes da validação de schema.
- **Preservação:** Descoberta dinâmica por prefixo `BD_`, validação de colunas de produto/SKU, validação de coluna de quantidade, leitura dinâmica dos clientes reais. Nenhuma whitelist fixa foi reintroduzida.
- **Arquivos alterados:** `Codigo.js` (função `f2dObterAbasClientes_`), `CHECKPOINT.md`.
- **Testes executados:** `node --check Codigo.js`, `git diff --check`, verificação manual de exclusão por nome para todos os termos do enunciado.
- **Pendência:** `calcularDashboardFase2` (em `Fase2_Dashboard.js`) ainda lê `MAPA_ERP` diretamente e não consome `obterMapaDeProdutos`. Essa decisão será tratada separadamente após validação visual dos dados.
- **Refinamento posterior:** Termos curtos LOG, AUX e MAPA passaram a usar correspondência por token completo para evitar falsos positivos em nomes como Cardiologia, Auxiliadora e Amapá.
