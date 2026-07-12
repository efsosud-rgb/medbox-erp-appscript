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
