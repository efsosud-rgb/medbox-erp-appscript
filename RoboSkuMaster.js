function inspecionarListaCMED() {
  const url = 'https://www.gov.br/anvisa/pt-br/assuntos/medicamentos/cmed/precos/arquivos/xls_conformidade_site_20260610_121627707.xlsx/@@download/file';

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const blob = response.getBlob().setName('CMED_TEMP.xlsx');

  const file = DriveApp.createFile(blob);
  const converted = Drive.Files.copy(
    { title: 'CMED_TEMP_CONVERTED', mimeType: MimeType.GOOGLE_SHEETS },
    file.getId()
  );

  const ss = SpreadsheetApp.openById(converted.id);
  const sheet = ss.getSheets()[0];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sample = sheet.getRange(2, 1, 5, sheet.getLastColumn()).getValues();

  Logger.log('CABECALHOS: ' + JSON.stringify(headers));
  Logger.log('AMOSTRA (5 linhas): ' + JSON.stringify(sample));

  DriveApp.getFileById(file.getId()).setTrashed(true);
  DriveApp.getFileById(converted.id).setTrashed(true);
}