/************************************************************
 * MEDBOX ERP | DIAGNOSTICO DE ESTRUTURA DA PLANILHA
 *
 * Funcao somente leitura.
 * Execute manualmente pelo editor do Apps Script.
 ************************************************************/

function diagnosticarEstruturaCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const diagnosticos = sheets.map(function(sheet) {
    return deDiagnosticarAba_(sheet, sheets);
  });

  const candidatas = diagnosticos.filter(function(info) {
    return info.vazia || info.tecnicaOuAuxiliar || info.possiveisDuplicatas.length > 0 || info.menosDeCincoLinhas;
  });

  const principais = diagnosticos.filter(function(info) {
    return !info.vazia &&
      !info.tecnicaOuAuxiliar &&
      info.possiveisDuplicatas.length === 0 &&
      !info.menosDeCincoLinhas &&
      info.linhasComDados >= 5 &&
      info.colunasComDados > 0;
  });

  const relatorio = [];

  relatorio.push('============================================================');
  relatorio.push('MEDBOX ERP | DIAGNOSTICO COMPLETO DA ESTRUTURA');
  relatorio.push('Planilha: ' + ss.getName());
  relatorio.push('Total de abas: ' + diagnosticos.length);
  relatorio.push('Gerado em: ' + deFormatarDataHora_(new Date()));
  relatorio.push('============================================================');
  relatorio.push('');

  relatorio.push('ABAS ENCONTRADAS');
  relatorio.push('------------------------------------------------------------');
  diagnosticos.forEach(function(info, index) {
    relatorio.push(deFormatarResumoAba_(info, index + 1));
  });

  relatorio.push('');
  relatorio.push('CANDIDATAS A ABA DESNECESSARIA/OBSOLETA');
  relatorio.push('------------------------------------------------------------');
  if (candidatas.length === 0) {
    relatorio.push('Nenhuma candidata encontrada pelos criterios definidos.');
  } else {
    candidatas.forEach(function(info) {
      relatorio.push('- ' + info.nome + ' | Motivos: ' + info.motivos.join('; '));
    });
  }

  relatorio.push('');
  relatorio.push('ABAS PRINCIPAIS EM USO');
  relatorio.push('------------------------------------------------------------');
  if (principais.length === 0) {
    relatorio.push('Nenhuma aba principal identificada automaticamente.');
  } else {
    principais.forEach(function(info) {
      relatorio.push(
        '- ' + info.nome +
        ' | Linhas: ' + info.linhasComDados +
        ' | Colunas: ' + info.colunasComDados +
        ' | Valores estaticos: ' + info.valoresEstaticos +
        ' | Formulas: ' + info.formulas
      );
    });
  }

  relatorio.push('');
  relatorio.push('============================================================');
  relatorio.push('FIM DO DIAGNOSTICO');
  relatorio.push('============================================================');

  const texto = relatorio.join('\n');
  Logger.log(texto);

  return texto;
}

function deDiagnosticarAba_(sheet, todasAsAbas) {
  const nome = sheet.getName();
  const linhasComDados = sheet.getLastRow();
  const colunasComDados = sheet.getLastColumn();
  const linhasDadosAlemCabecalho = Math.max(0, linhasComDados - 1);
  const vazia = linhasComDados <= 1;
  const menosDeCincoLinhas = linhasComDados < 5;

  const cabecalhos = colunasComDados > 0
    ? sheet.getRange(1, 1, 1, colunasComDados).getValues()[0].map(function(valor) {
        return String(valor || '').trim();
      })
    : [];

  const rangeDados = linhasComDados > 0 && colunasComDados > 0
    ? sheet.getRange(1, 1, linhasComDados, colunasComDados)
    : null;

  const formulas = rangeDados ? deContarNaoVazios_(rangeDados.getFormulas()) : 0;
  const valoresNaoVazios = rangeDados ? deContarValoresNaoVazios_(rangeDados.getValues()) : 0;
  const valoresEstaticos = Math.max(0, valoresNaoVazios - formulas);

  const tecnicaOuAuxiliar = deEhAbaTecnicaOuAuxiliar_(nome);
  const possiveisDuplicatas = deEncontrarAbasParecidas_(nome, todasAsAbas);

  const motivos = [];
  if (vazia) motivos.push('vazia ou sem dados alem do cabecalho');
  if (tecnicaOuAuxiliar) motivos.push('nome indica aba tecnica/auxiliar');
  if (possiveisDuplicatas.length > 0) motivos.push('nome parecido com: ' + possiveisDuplicatas.join(', '));
  if (menosDeCincoLinhas) motivos.push('menos de 5 linhas totais');

  return {
    nome: nome,
    linhasComDados: linhasComDados,
    colunasComDados: colunasComDados,
    cabecalhos: cabecalhos,
    vazia: vazia,
    linhasDadosAlemCabecalho: linhasDadosAlemCabecalho,
    tecnicaOuAuxiliar: tecnicaOuAuxiliar,
    possiveisDuplicatas: possiveisDuplicatas,
    formulas: formulas,
    valoresEstaticos: valoresEstaticos,
    menosDeCincoLinhas: menosDeCincoLinhas,
    motivos: motivos
  };
}

function deFormatarResumoAba_(info, numero) {
  const linhas = [];

  linhas.push(numero + '. ' + info.nome);
  linhas.push('   Linhas com dados: ' + info.linhasComDados);
  linhas.push('   Colunas com dados: ' + info.colunasComDados);
  linhas.push('   Cabecalhos: ' + (info.cabecalhos.length ? info.cabecalhos.join(' | ') : '(sem cabecalhos)'));
  linhas.push('   Vazia: ' + (info.vazia ? 'SIM' : 'NAO'));
  linhas.push('   Linhas alem do cabecalho: ' + info.linhasDadosAlemCabecalho);
  linhas.push('   Tecnica/auxiliar pelo nome: ' + (info.tecnicaOuAuxiliar ? 'SIM' : 'NAO'));
  linhas.push('   Possiveis duplicatas: ' + (info.possiveisDuplicatas.length ? info.possiveisDuplicatas.join(', ') : 'nenhuma'));
  linhas.push('   Formulas: ' + info.formulas);
  linhas.push('   Valores estaticos nao vazios: ' + info.valoresEstaticos);
  linhas.push('   Menos de 5 linhas totais: ' + (info.menosDeCincoLinhas ? 'SIM' : 'NAO'));
  linhas.push('   Classificacao automatica: ' + (info.motivos.length ? info.motivos.join('; ') : 'sem alerta estrutural'));
  linhas.push('');

  return linhas.join('\n');
}

function deEhAbaTecnicaOuAuxiliar_(nome) {
  const n = deNormalizarNomeAba_(nome);

  const marcadores = [
    'LOG_',
    'TEMP_',
    'TESTE_',
    'BACKUP_',
    'RASCUNHO_',
    'ANTIGO_',
    'OLD_',
    'COPIA_'
  ];

  const contemMarcador = marcadores.some(function(marcador) {
    return n.indexOf(marcador) !== -1;
  });

  const terminaComVersao = /(_\d+|_V\d+|_FINAL|_FINAL\d+)$/.test(n);

  return contemMarcador || terminaComVersao;
}

function deEncontrarAbasParecidas_(nome, todasAsAbas) {
  const base = deBaseNomeDuplicata_(nome);

  return todasAsAbas
    .map(function(sheet) { return sheet.getName(); })
    .filter(function(outroNome) {
      if (outroNome === nome) return false;

      const outraBase = deBaseNomeDuplicata_(outroNome);

      if (!base || !outraBase) return false;
      if (base === outraBase) return true;
      if (base.length >= 6 && outraBase.indexOf(base) === 0) return true;
      if (outraBase.length >= 6 && base.indexOf(outraBase) === 0) return true;

      return false;
    });
}

function deBaseNomeDuplicata_(nome) {
  return deNormalizarNomeAba_(nome)
    .replace(/\s*\(\d+\)\s*$/g, '')
    .replace(/_(ANTIGO|OLD|BACKUP|COPIA|TEMP|TESTE|RASCUNHO)$/g, '')
    .replace(/(_\d+|_V\d+|_FINAL|_FINAL\d+)$/g, '')
    .trim();
}

function deNormalizarNomeAba_(nome) {
  return String(nome || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/\s+/g, '_')
    .replace(/[^A-Z0-9_()]/g, '_')
    .replace(/_+/g, '_')
    .trim();
}

function deContarNaoVazios_(matriz) {
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

function deContarValoresNaoVazios_(matriz) {
  let total = 0;

  matriz.forEach(function(row) {
    row.forEach(function(valor) {
      if (valor !== null && valor !== undefined && String(valor).trim() !== '') {
        total++;
      }
    });
  });

  return total;
}

function deFormatarDataHora_(data) {
  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );
}
