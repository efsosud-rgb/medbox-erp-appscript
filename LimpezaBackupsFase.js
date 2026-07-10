/************************************************************
 * MEDBOX ERP | LIMPEZA CONTROLADA DE BACKUPS 20260707
 *
 * Execute manualmente pelo editor do Apps Script.
 * Nao adicionar ao menu.
 ************************************************************/

const LBF_NOME_ARQUIVO_BACKUPS = 'MEDBOX_ARQUIVO_BACKUPS_20260707';
const LBF_ID_ARQUIVO_BACKUPS_CONFIRMADO = '1YXdmn_pFkUzGX04wbmBRYDU7y3Z6bPeg7pe1yb0sdAM';

const LBF_ABAS_BACKUP_20260707 = [
  'BACKUP_PRE_RESET_PAINEL_ERP_20260707_222213',
  'BACKUP_PRE_RESET_BUSCA_AVANCADA_20260707_222154',
  'BACKUP_PRE_RESET_MAPA_ERP_20260707_222042',
  'BACKUP_LAYOUT_PAINEL_ERP_20260707_220933',
  'BACKUP_LAYOUT_BUSCA_AVANCADA_20260707_220759',
  'BACKUP_LAYOUT_BUSCA_AVANCADA_20260707_212730',
  'BACKUP_CONFIG_20260707_194456',
  'BACKUP_LAYOUT_BUSCA_BUSCA_AVANCADA_20260707_190147',
  'BACKUP_LAYOUT_BUSCA_BUSCA_AVANCADA_20260707_184027',
  'BACKUP_LAYOUT_BUSCA_BUSCA_AVANCADA_20260707_182338',
  'BACKUP_BUSCA_AVANCADA_BUSCA_AVANCADA_20260707_180605',
  'BACKUP_BUSCA_AVANCADA_20260707_174937',
  'BACKUP_BUSCA_AVANCADA_20260707_174314',
  'BACKUP_MAPA_ERP_20260707_173707',
  'BACKUP_BD_INTS_20260707_171642'
];

const LBF_ABAS_REMOVER_REDUNDANTES = [
  'BACKUP_LAYOUT_BUSCA_AVANCADA_20260707_212730',
  'BACKUP_LAYOUT_BUSCA_BUSCA_AVANCADA_20260707_182338',
  'BACKUP_LAYOUT_BUSCA_BUSCA_AVANCADA_20260707_184027',
  'BACKUP_BUSCA_AVANCADA_BUSCA_AVANCADA_20260707_180605',
  'BACKUP_BUSCA_AVANCADA_20260707_174937',
  'BACKUP_MAPA_ERP_20260707_173707'
];

const LBF_ABAS_CONVERTER_PARA_VALORES = [
  'BACKUP_BUSCA_AVANCADA_20260707_174314',
  'BACKUP_BD_INTS_20260707_171642'
];

const LBF_ABAS_MANTER_INTACTAS = [
  'BACKUP_PRE_RESET_PAINEL_ERP_20260707_222213',
  'BACKUP_PRE_RESET_BUSCA_AVANCADA_20260707_222154',
  'BACKUP_PRE_RESET_MAPA_ERP_20260707_222042',
  'BACKUP_LAYOUT_PAINEL_ERP_20260707_220933',
  'BACKUP_LAYOUT_BUSCA_AVANCADA_20260707_220759',
  'BACKUP_CONFIG_20260707_194456',
  'BACKUP_LAYOUT_BUSCA_BUSCA_AVANCADA_20260707_190147'
];


const LBF_ABAS_ARQUIVADAS_CONFIRMADAS_REMOVER = [
  'BACKUP_PRE_RESET_PAINEL_ERP_20260707_222213',
  'BACKUP_PRE_RESET_BUSCA_AVANCADA_20260707_222154',
  'BACKUP_PRE_RESET_MAPA_ERP_20260707_222042',
  'BACKUP_LAYOUT_PAINEL_ERP_20260707_220933',
  'BACKUP_LAYOUT_BUSCA_AVANCADA_20260707_220759',
  'BACKUP_CONFIG_20260707_194456',
  'BACKUP_LAYOUT_BUSCA_BUSCA_AVANCADA_20260707_190147',
  'BACKUP_BUSCA_AVANCADA_20260707_174314',
  'BACKUP_BD_INTS_20260707_171642',
  'RELATORIO_DIAGNOSTICO_ESTRUTURA'
];

const LBF_ABAS_GRUPO_INCERTO_ARQUIVAR_REMOVER = [
  'AUX_BUSCA',
  'BD_CONSOLIDADO',
  'MAPA_BASE_ORIGINAL',
  'BUSCA_AVANCADA',
  'PAINEL_ERP',
  'LANCAMENTO_COTACOES',
  'LANCAMENTO_PRECOS',
  'CADASTRO_SKU_MASTER',
  'SIMULADOR_PRECOS',
  'DIAGNOSTICO_ROBO_FISCAL'
];
const LBF_TAMANHO_BLOCO_LINHAS = 500;

function arquivarBackupsAntigos() {
  const ssOrigem = SpreadsheetApp.getActiveSpreadsheet();
  const ssArquivo = lbfObterOuCriarArquivoBackups_();
  const logs = [];
  const divergencias = [];

  logs.push('MEDBOX ERP | ARQUIVAMENTO DE BACKUPS 20260707');
  logs.push('Origem: ' + ssOrigem.getName());
  logs.push('Arquivo: ' + ssArquivo.getName());
  logs.push('URL arquivo: ' + ssArquivo.getUrl());
  logs.push('------------------------------------------------------------');

  LBF_ABAS_BACKUP_20260707.forEach(function(nomeAba) {
    const abaOrigem = ssOrigem.getSheetByName(nomeAba);

    if (!abaOrigem) {
      const msgAusente = nomeAba + ' | AUSENTE NA PLANILHA ATIVA';
      logs.push(msgAusente);
      divergencias.push(msgAusente);
      return;
    }

    const linhasOrigem = abaOrigem.getLastRow();
    const colunasOrigem = abaOrigem.getLastColumn();
    const abaArquivo = lbfObterOuCriarAbaNoArquivo_(ssArquivo, nomeAba);

    abaArquivo.clear();

    if (linhasOrigem > 0 && colunasOrigem > 0) {
      lbfGarantirTamanhoAba_(abaArquivo, linhasOrigem, colunasOrigem);
      lbfCopiarDisplayValuesEmBlocos_(abaOrigem, abaArquivo, linhasOrigem, colunasOrigem);
    }

    const linhasArquivo = abaArquivo.getLastRow();
    const colunasArquivo = abaArquivo.getLastColumn();
    const ok = linhasArquivo === linhasOrigem && colunasArquivo === colunasOrigem;
    const status = ok ? 'OK' : 'DIVERGENTE';

    logs.push(
      nomeAba +
      ' | linhas copiadas: ' + linhasOrigem +
      ' | colunas copiadas: ' + colunasOrigem +
      ' | arquivo linhas/colunas: ' + linhasArquivo + '/' + colunasArquivo +
      ' | ' + status
    );

    if (!ok) {
      divergencias.push(nomeAba);
    }
  });

  logs.push('------------------------------------------------------------');
  logs.push('URL DA PLANILHA DE ARQUIVO: ' + ssArquivo.getUrl());

  const textoLog = logs.join('\n');
  Logger.log(textoLog);

  if (divergencias.length > 0) {
    throw new Error(
      'Arquivamento com divergencias. Nao remova backups antes de corrigir: ' +
      divergencias.join(', ')
    );
  }

  return {
    ok: true,
    arquivoUrl: ssArquivo.getUrl(),
    abasArquivadas: LBF_ABAS_BACKUP_20260707.length
  };
}

function removerBackupsDuplicadosRedundantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssArquivo = lbfAbrirArquivoBackupsExistente_();
  const validacao = lbfValidarArquivoContraOrigem_(ss, ssArquivo);

  if (!validacao.ok) {
    Logger.log(validacao.logs.join('\n'));
    throw new Error(
      'Arquivo de backups nao esta validado. Rode arquivarBackupsAntigos() e confira divergencias antes de remover.'
    );
  }

  const logs = [];
  const convertidas = [];
  const removidas = [];
  const mantidasIntactas = [];

  logs.push('MEDBOX ERP | REMOCAO CONTROLADA DE BACKUPS REDUNDANTES');
  logs.push('Arquivo validado: ' + ssArquivo.getUrl());
  logs.push('------------------------------------------------------------');

  LBF_ABAS_CONVERTER_PARA_VALORES.forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);

    if (!aba) {
      logs.push(nomeAba + ' | CONVERSAO IGNORADA: aba nao encontrada');
      return;
    }

    const resultado = lbfConverterFormulasEmValores_(aba);
    convertidas.push(
      nomeAba +
      ' | formulas antes: ' + resultado.formulasAntes +
      ' | formulas depois: ' + resultado.formulasDepois
    );
  });

  LBF_ABAS_REMOVER_REDUNDANTES.forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);

    if (!aba) {
      logs.push(nomeAba + ' | REMOCAO IGNORADA: aba nao encontrada');
      return;
    }

    ss.deleteSheet(aba);
    removidas.push(nomeAba);
  });

  LBF_ABAS_MANTER_INTACTAS.forEach(function(nomeAba) {
    if (ss.getSheetByName(nomeAba)) {
      mantidasIntactas.push(nomeAba);
    } else {
      logs.push(nomeAba + ' | MANTER INTACTA: aba nao encontrada');
    }
  });

  logs.push('CONVERTIDAS PARA VALORES');
  logs.push(convertidas.length ? convertidas.join('\n') : 'Nenhuma aba convertida.');
  logs.push('------------------------------------------------------------');
  logs.push('REMOVIDAS');
  logs.push(removidas.length ? removidas.join('\n') : 'Nenhuma aba removida.');
  logs.push('------------------------------------------------------------');
  logs.push('MANTIDAS INTACTAS');
  logs.push(mantidasIntactas.length ? mantidasIntactas.join('\n') : 'Nenhuma aba mantida intacta encontrada.');
  logs.push('------------------------------------------------------------');
  logs.push('OBS: BACKUP_BUSCA_AVANCADA_20260707_174314 e BACKUP_BD_INTS_20260707_171642 foram mantidas apos conversao para valores.');

  Logger.log(logs.join('\n'));

  return {
    ok: true,
    removidas: removidas,
    convertidas: convertidas,
    mantidasIntactas: mantidasIntactas
  };
}

function removerAbasArquivadasConfirmadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssArquivo = SpreadsheetApp.openById(LBF_ID_ARQUIVO_BACKUPS_CONFIRMADO);
  const removidas = [];
  const puladas = [];
  const logs = [];

  logs.push('MEDBOX ERP | REMOCAO DE ABAS ARQUIVADAS CONFIRMADAS');
  logs.push('Arquivo validado: ' + ssArquivo.getUrl());
  logs.push('------------------------------------------------------------');

  LBF_ABAS_ARQUIVADAS_CONFIRMADAS_REMOVER.forEach(function(nomeAba) {
    const abaAtiva = ss.getSheetByName(nomeAba);

    if (!abaAtiva) {
      const motivoAusente = 'aba nao existe na planilha ativa';
      puladas.push(nomeAba + ' | ' + motivoAusente);
      logs.push(nomeAba + ' | PULADA: ' + motivoAusente);
      return;
    }

    if (nomeAba === 'RELATORIO_DIAGNOSTICO_ESTRUTURA') {
      ss.deleteSheet(abaAtiva);
      removidas.push(nomeAba);
      logs.push(nomeAba + ' | REMOVIDA: relatorio descartavel sem checagem cruzada');
      return;
    }

    const abaArquivo = ssArquivo.getSheetByName(nomeAba);

    if (!abaArquivo) {
      const motivoSemArquivo = 'aba nao encontrada no arquivo de backups';
      puladas.push(nomeAba + ' | ' + motivoSemArquivo);
      logs.push(nomeAba + ' | PULADA: ' + motivoSemArquivo);
      return;
    }

    const linhasAtiva = abaAtiva.getLastRow();
    const colunasAtiva = abaAtiva.getLastColumn();
    const linhasArquivo = abaArquivo.getLastRow();
    const colunasArquivo = abaArquivo.getLastColumn();

    if (linhasAtiva !== linhasArquivo || colunasAtiva !== colunasArquivo) {
      const motivoDivergente =
        'contagem divergente ativa ' + linhasAtiva + '/' + colunasAtiva +
        ' vs arquivo ' + linhasArquivo + '/' + colunasArquivo;

      puladas.push(nomeAba + ' | ' + motivoDivergente);
      logs.push(nomeAba + ' | PULADA: ' + motivoDivergente);
      return;
    }

    ss.deleteSheet(abaAtiva);
    removidas.push(nomeAba);
    logs.push(nomeAba + ' | REMOVIDA: validada contra arquivo');
  });

  logs.push('------------------------------------------------------------');
  logs.push('REMOVIDAS COM SUCESSO');
  logs.push(removidas.length ? removidas.join('\n') : 'Nenhuma aba removida.');
  logs.push('------------------------------------------------------------');
  logs.push('PULADAS POR FALHA NA VERIFICACAO');
  logs.push(puladas.length ? puladas.join('\n') : 'Nenhuma aba pulada.');

  Logger.log(logs.join('\n'));

  return {
    ok: true,
    removidas: removidas,
    puladas: puladas
  };
}
function arquivarERemoverGrupoIncerto() {
  const ssOrigem = SpreadsheetApp.getActiveSpreadsheet();
  const ssArquivo = SpreadsheetApp.openById(LBF_ID_ARQUIVO_BACKUPS_CONFIRMADO);
  const logs = [];
  const removidas = [];
  const puladas = [];

  logs.push('MEDBOX ERP | ARQUIVAR E REMOVER GRUPO INCERTO');
  logs.push('Origem: ' + ssOrigem.getName());
  logs.push('Arquivo: ' + ssArquivo.getUrl());
  logs.push('------------------------------------------------------------');

  LBF_ABAS_GRUPO_INCERTO_ARQUIVAR_REMOVER.forEach(function(nomeAba) {
    const abaOrigem = ssOrigem.getSheetByName(nomeAba);

    if (!abaOrigem) {
      const msgAusente = nomeAba + ' | AUSENTE NA PLANILHA ATIVA | removida NAO';
      logs.push(msgAusente);
      puladas.push(nomeAba + ' | ausente na planilha ativa');
      return;
    }

    const linhasOrigem = abaOrigem.getLastRow();
    const colunasOrigem = abaOrigem.getLastColumn();
    const abaArquivo = lbfObterOuCriarAbaNoArquivo_(ssArquivo, nomeAba);

    abaArquivo.clear();

    if (linhasOrigem > 0 && colunasOrigem > 0) {
      lbfGarantirTamanhoAba_(abaArquivo, linhasOrigem, colunasOrigem);
      lbfCopiarDisplayValuesEmBlocos_(abaOrigem, abaArquivo, linhasOrigem, colunasOrigem);
    }

    SpreadsheetApp.flush();

    const linhasArquivo = abaArquivo.getLastRow();
    const colunasArquivo = abaArquivo.getLastColumn();
    const ok = linhasOrigem === linhasArquivo && colunasOrigem === colunasArquivo;

    if (!ok) {
      const motivoDivergente =
        'DIVERGENTE origem ' + linhasOrigem + '/' + colunasOrigem +
        ' vs arquivo ' + linhasArquivo + '/' + colunasArquivo;

      logs.push(
        nomeAba +
        ' | linhas copiadas ' + linhasOrigem +
        ' | validacao DIVERGENTE' +
        ' | removida NAO | ' + motivoDivergente
      );
      puladas.push(nomeAba + ' | ' + motivoDivergente);
      return;
    }

    ssOrigem.deleteSheet(abaOrigem);
    removidas.push(nomeAba);
    logs.push(
      nomeAba +
      ' | linhas copiadas ' + linhasOrigem +
      ' | validacao OK' +
      ' | removida SIM'
    );
  });

  logs.push('------------------------------------------------------------');
  logs.push('REMOVIDAS');
  logs.push(removidas.length ? removidas.join('\n') : 'Nenhuma aba removida.');
  logs.push('------------------------------------------------------------');
  logs.push('PULADAS');
  logs.push(puladas.length ? puladas.join('\n') : 'Nenhuma aba pulada.');

  Logger.log(logs.join('\n'));

  return {
    ok: true,
    removidas: removidas,
    puladas: puladas,
    arquivoUrl: ssArquivo.getUrl()
  };
}
function lbfObterOuCriarArquivoBackups_() {
  const arquivos = DriveApp.getFilesByName(LBF_NOME_ARQUIVO_BACKUPS);

  if (arquivos.hasNext()) {
    return SpreadsheetApp.openById(arquivos.next().getId());
  }

  return SpreadsheetApp.create(LBF_NOME_ARQUIVO_BACKUPS);
}

function lbfAbrirArquivoBackupsExistente_() {
  const arquivos = DriveApp.getFilesByName(LBF_NOME_ARQUIVO_BACKUPS);

  if (!arquivos.hasNext()) {
    throw new Error('Planilha de arquivo nao encontrada: ' + LBF_NOME_ARQUIVO_BACKUPS);
  }

  return SpreadsheetApp.openById(arquivos.next().getId());
}

function lbfObterOuCriarAbaNoArquivo_(ssArquivo, nomeAba) {
  let aba = ssArquivo.getSheetByName(nomeAba);

  if (!aba) {
    aba = ssArquivo.insertSheet(nomeAba);
  }

  return aba;
}

function lbfGarantirTamanhoAba_(aba, linhas, colunas) {
  const maxRows = aba.getMaxRows();
  const maxColumns = aba.getMaxColumns();

  if (maxRows < linhas) {
    aba.insertRowsAfter(maxRows, linhas - maxRows);
  }

  if (maxColumns < colunas) {
    aba.insertColumnsAfter(maxColumns, colunas - maxColumns);
  }
}

function lbfCopiarDisplayValuesEmBlocos_(abaOrigem, abaDestino, linhas, colunas) {
  for (let linha = 1; linha <= linhas; linha += LBF_TAMANHO_BLOCO_LINHAS) {
    const altura = Math.min(LBF_TAMANHO_BLOCO_LINHAS, linhas - linha + 1);
    const valores = abaOrigem.getRange(linha, 1, altura, colunas).getDisplayValues();
    abaDestino.getRange(linha, 1, altura, colunas).setValues(valores);
  }
}

function lbfConverterFormulasEmValores_(aba) {
  const linhas = aba.getLastRow();
  const colunas = aba.getLastColumn();

  if (linhas === 0 || colunas === 0) {
    return { formulasAntes: 0, formulasDepois: 0 };
  }

  const range = aba.getRange(1, 1, linhas, colunas);
  const formulasAntes = lbfContarNaoVazios_(range.getFormulas());

  for (let linha = 1; linha <= linhas; linha += LBF_TAMANHO_BLOCO_LINHAS) {
    const altura = Math.min(LBF_TAMANHO_BLOCO_LINHAS, linhas - linha + 1);
    const bloco = aba.getRange(linha, 1, altura, colunas);
    bloco.setValues(bloco.getValues());
  }

  SpreadsheetApp.flush();

  const formulasDepois = lbfContarNaoVazios_(aba.getRange(1, 1, linhas, colunas).getFormulas());

  return {
    formulasAntes: formulasAntes,
    formulasDepois: formulasDepois
  };
}

function lbfValidarArquivoContraOrigem_(ssOrigem, ssArquivo) {
  const logs = [];
  const divergencias = [];

  LBF_ABAS_BACKUP_20260707.forEach(function(nomeAba) {
    const abaOrigem = ssOrigem.getSheetByName(nomeAba);
    const abaArquivo = ssArquivo.getSheetByName(nomeAba);

    if (!abaOrigem || !abaArquivo) {
      const msgAusente = nomeAba + ' | DIVERGENTE: origem ou arquivo ausente';
      logs.push(msgAusente);
      divergencias.push(nomeAba);
      return;
    }

    const origemLinhas = abaOrigem.getLastRow();
    const origemColunas = abaOrigem.getLastColumn();
    const arquivoLinhas = abaArquivo.getLastRow();
    const arquivoColunas = abaArquivo.getLastColumn();
    const ok = origemLinhas === arquivoLinhas && origemColunas === arquivoColunas;

    logs.push(
      nomeAba +
      ' | origem: ' + origemLinhas + '/' + origemColunas +
      ' | arquivo: ' + arquivoLinhas + '/' + arquivoColunas +
      ' | ' + (ok ? 'OK' : 'DIVERGENTE')
    );

    if (!ok) {
      divergencias.push(nomeAba);
    }
  });

  return {
    ok: divergencias.length === 0,
    divergencias: divergencias,
    logs: logs
  };
}

function lbfContarNaoVazios_(matriz) {
  let total = 0;

  matriz.forEach(function(row) {
    row.forEach(function(valor) {
      if (valor !== null && valor !== undefined && String(valor) !== '') {
        total++;
      }
    });
  });

  return total;
}
