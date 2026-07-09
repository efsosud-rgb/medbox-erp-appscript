/**
 * MEDBOX ERP | ROBÔ SKU MASTER (CONFORMIDADE CMED/ANVISA)
 *
 * Este arquivo contém as funções de inspeção e processamento da lista de conformidade
 * da CMED (Câmara de Regulação do Mercado de Medicamentos).
 */

/**
 * Baixa o arquivo Excel da CMED, converte para Google Sheets e registra cabeçalhos e amostra de dados.
 * Etapa exploratória para diagnóstico da estrutura do arquivo público.
 */
function inspecionarListaCMED() {
  const url = "https://www.gov.br/anvisa/pt-br/assuntos/medicamentos/cmed/precos/arquivos/xls_conformidade_site_20260610_121627707.xlsx/@@download/file";

  Logger.log("Iniciando download do arquivo CMED da URL: " + url);

  let spreadsheetId = null;

  try {
    const response = UrlFetchApp.fetch(url);
    const blob = response.getBlob();

    // Configura metadados para criação e conversão automática para Google Sheets
    const fileMetadata = {
      name: 'CMED_TEMP.xlsx',
      mimeType: 'application/vnd.google-apps.spreadsheet'
    };

    Logger.log("Salvando e convertendo arquivo no Google Drive via Advanced Drive Service (v3)...");

    // Cria o arquivo no Drive convertendo para formato Google Sheets
    const file = Drive.Files.create(fileMetadata, blob);
    spreadsheetId = file.id;

    Logger.log("Arquivo convertido com sucesso. ID: " + spreadsheetId);

    // Abre a planilha convertida
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheets()[0];

    // Define o limite de leitura: cabeçalho + 5 linhas de dados
    const lastColumn = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const rowsToRead = Math.min(6, lastRow); // 1 (header) + 5 (data) ou menos se o arquivo for pequeno

    if (rowsToRead === 0) {
      Logger.log("Planilha vazia.");
    } else {
      Logger.log("Lendo cabeçalhos e primeiras " + (rowsToRead - 1) + " linhas...");

      const range = sheet.getRange(1, 1, rowsToRead, lastColumn);
      const values = range.getValues();

      if (values && values.length > 0) {
        const headers = values[0];
        Logger.log("CABEÇALHOS ENCONTRADOS:");
        Logger.log(JSON.stringify(headers));

        Logger.log("AMOSTRA DE DADOS (5 PRIMEIRAS LINHAS):");
        for (let i = 1; i < values.length; i++) {
          Logger.log("Linha " + i + ": " + JSON.stringify(values[i]));
        }
      } else {
        Logger.log("Nenhum dado encontrado na primeira aba da planilha.");
      }
    }

  } catch (error) {
    Logger.log("Erro ao processar lista CMED: " + error.toString());
    throw error;
  } finally {
    // Limpeza: move o arquivo temporário para a lixeira, mesmo em caso de erro
    if (spreadsheetId) {
      Logger.log("Limpando arquivo temporário do Drive...");
      DriveApp.getFileById(spreadsheetId).setTrashed(true);
    }
  }
}
