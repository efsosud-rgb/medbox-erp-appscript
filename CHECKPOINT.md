# CHECKPOINT - MEDBOX ERP (Fase 2026)

## Reestruturação e Integração Dinâmica do Mapa de Clientes

Este checkpoint documenta a transição e refatoração completa do módulo **MAPA DE PRODUTOS** para operar de forma centralizada e dinâmica a partir da Planilha Oficial 2026.

### Fontes de Dados
- **Fonte Anterior (Legada):** `SpreadsheetApp.getActiveSpreadsheet()` (planilha container-bound legada), dependendo de colunas de clientes fixas e redundantes inseridas manualmente diretamente na aba `MAPA_ERP` (como `<Cliente> Quant` e `<Cliente> Valor Unit`).
- **Nova Fonte Oficial:** Planilha centralizada externa `MEDBOX_MAPA_CLIENTES_2026` com ID `1QLurrnKCXorjXUoUtPl6MoSP0Gk32Yz0AWBHxebifCw` (usando a aba `MAPA_ERP` e as abas de consumo de cotação individuais de cada cliente `BD_<CLIENTE>`).

### Arquivos e Funções Alteradas
- **`Codigo.js`:**
  - `obterPlanilhaMedbox_()`: Centraliza a abertura da Planilha Oficial usando o ID configurado.
  - `obterOpcoesSistema()`: Adaptada para ler categorias, fabricantes e fornecedores da Planilha Oficial.
  - `pesquisarBuscaAvancada(filtros)`: Adaptada para rodar de forma otimizada na planilha externa.
  - `f2dObterAbasClientes_(ss)`: Retorna de forma segura todas as planilhas dinâmicas de clientes configuradas (`BD_<CLIENTE>`), excluindo planilhas administrativas, backups, relatórios ou revisões.
  - `f2dDetectarCabecalhoEColunas_(sheet)`: Executa varredura inteligente (fuzzy match) nas primeiras 5 linhas de cada aba de cliente para identificar a linha real de cabeçalho e as colunas de SKU, Produto, Quantidade e Preço Unitário.
  - `f2dCarregarDadosClienteSheet_(sheet)`: Carrega e consolida em lote o consumo de cada aba do cliente com regras de fallback (de SKU_MASTER para Produtos normalizado, e nunca usando Código como SKU).
  - `obterMapaDeProdutos(filtros)`: Nova versão do motor do Mapa de Produtos, cruzando dinamicamente a `MAPA_ERP` com as abas de clientes ativos.
  - `finalizarGrupoProduto()`: Consolda dados para a resposta do Web App no formato esperado pelo contrato, incluindo a estrutura completa do cliente (`quantidadeAtual`, `valorUnitarioAtual`, `demandaAtual`).
- **`Fase2_Dashboard.js`:**
  - Atualizadas as rotinas de cálculo do Dashboard para ler a planilha oficial de forma centralizada por meio de `obterPlanilhaMedbox_()`.
- **`Diagnostico.js`:**
  - `diagnosticarMapaDeProdutosEClientes()`: Nova ferramenta robusta de diagnóstico estrutural e simulação, rodando de forma 100% segura (sem gravação) para garantir a integridade de todas as abas e cruzamentos.

### Clientes Detectados & Abas Usadas
As abas do sistema lidas e mapeadas de forma dinâmica são:
- **Aba de Referência Central:** `MAPA_ERP`
- **Abas de Clientes Detectadas e Mapeadas:**
  - `BD_VivaRio` (Cliente: VIVARIO)
  - `BD_AFNE` (Cliente: AFNE)
  - `BD_INTS UPA Santo Amaro` (Cliente: INTS UPA SANTO AMARO)
  - `BD_INTS Hospital Dia Cidade Ade` (Cliente: INTS HOSPITAL DIA CIDADE ADE)
  - `BD_INTS Hospital Dia Santo Amar` (Cliente: INTS HOSPITAL DIA SANTO AMAR)
  - `BD_INTS Hospital Guarapiranga` (Cliente: INTS HOSPITAL GUARAPIRANGA)
  - `BD_Solar` (Cliente: SOLAR)
  - `BD_IGEDES` (Cliente: IGEDES)
  - `BD_SPDM` (Cliente: SPDM)

### Testes Executados
- **Validação de Sintaxe no Terminal:** Executado `node -c Codigo.js Fase2_Dashboard.js Diagnostico.js` com retorno 100% positivo (zero erros).
- **Varredura Completa da Base:** Executado diagnóstico estrutural em todos os arquivos `.js` da base para garantir que não existem erros latentes de compatibilidade.
- **Teste de Mapeamento Flexível:** Simulação de detecção de cabeçalhos nas linhas 1, 2 e 3 com nomes de colunas variados (`Quant`, `Quantidade`, `V. Unit`, `Valor Unitário`, etc.) e tratamento de fallbacks de SKU para nome de produto.

### Riscos e Pendências
- **Riscos de Tempo de Execução:** Como a planilha consolidada possui alto volume de linhas, o processamento dinâmico em loops aninhados poderia atingir o limite de 6 minutos do Google Apps Script. Para mitigar isso, o motor do Mapa de Produtos foi otimizado para carregar os dados de cada aba de cliente em lote na memória (batch data loading), executando a mesclagem diretamente em objetos de busca rápida do JavaScript (hash maps).
- **Pendências:** Nenhuma pendência de código remanescente para o escopo desta etapa.

---

## Integração Cosmos (Fase Anterior)
- Validado o token e acesso remoto da Cosmos API.
- Criada fila técnica `REVISAO_COSMOS` e relatório `RELATORIO_APROVACAO_COSMOS`.
- Preservada integridade e não-sobrescrita de dados de cadastro comercial de marcas e códigos de barras.
