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
- Campos exibidos/editaveis na tela: descricao sugerida, marca sugerida, GTIN sugerido e NCM sugerido.
- Campos aplicados no cadastro apos aprovacao explicita: `FABRICANTE_MARCA`, `EAN`, `NCM`.
- `DESCRICAO_FORNECEDOR` e `SKU_MASTER` nao sao sobrescritos pela revisao Cosmos.
- Itens rejeitados apenas mudam status para `REJEITADO` na fila de revisao e nao alteram `BD_PRECOS`.
- Itens pendentes nao alteram `BD_PRECOS`.
- Aprovacao usa `LockService`, `REVISAO_ID`, `ITEM_ID`, status atual e `VERSAO_REGISTRO` para evitar duplicidade e tela desatualizada.
- Revalidacao antes da escrita: a origem e localizada novamente por `BD_PRECOS.ID_PRECO`; `LINHA_ORIGEM` e apenas referencia visual.
- Validacoes implementadas: payload ausente, revisao ausente, item ausente, status nao pendente, GTIN invalido, NCM invalido, item movido/removido e tentativa duplicada.
- Teste executado sem gravacao: `testarRevisaoCosmosSemGravacao`.
- Validacao visual e aprovacao/rejeicao real nao foram executadas nesta etapa para evitar teste sobre produto real sem autorizacao explicita.
- Limitacao: a fila `REVISAO_COSMOS` e criada/sincronizada ao carregar a tela a partir de sugestoes Cosmos existentes em `BD_PRECOS`.
- Proximo passo: autorizar um item de teste ou copia controlada para validar carregamento visual, aprovacao, rejeicao e relatorio automatico ponta a ponta.
