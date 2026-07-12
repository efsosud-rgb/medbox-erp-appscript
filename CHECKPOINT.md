# CHECKPOINT

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
- Proximo passo: autorizar um unico item de teste controlado ou copia controlada para validar carregamento visual, simulacao, aprovacao, rejeicao, backup e relatorio automatico ponta a ponta.

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
- Valores financeiros usam formato pt-BR `R$ 1.234,56`; valores vazios ou invalidos aparecem como `—`.
- Resumo comercial da cotacao exibe: total de itens cotados, valor total cotado, total de itens ganhos, valor total ganho MedBox, percentual de itens ganhos e itens sem preco.
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
