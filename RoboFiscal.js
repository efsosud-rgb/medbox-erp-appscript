function executarVarreduraFiscal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("BD_PRECOS");
  
  if (!aba) {
    SpreadsheetApp.getUi().alert("Erro: A aba 'BD_PRECOS' não foi encontrada!");
    return;
  }
  
  // ==========================================
  // CONFIGURAÇÃO DAS COLUNAS (Atualizado por você)
  // ==========================================
  var COLUNA_SKU = 9;   // Coluna I
  var COLUNA_NCM = 11;  // Coluna K
  var COLUNA_EAN = 12;  // Coluna L
  var COLUNA_CST = 13;  // Coluna M
  
  // Insira o token da API de consulta que você contratou/ativou
  var API_TOKEN = "SEU_TOKEN_DA_API_AQUI"; 
  
  var ultimaLinha = aba.getLastRow();
  var contadorSucesso = 0;
  
  // Varre a planilha da linha 2 até o final
  for (var i = 2; i <= ultimaLinha; i++) {
    var valorCelulaEan = aba.getRange(i, COLUNA_EAN).getValue();
    var codigoEan = formatarEanCorreto(valorCelulaEan);
    
    // Só executa a busca se houver um EAN válido na célula
    if (codigoEan !== "" && codigoEan !== "0") {
      
      // Modelo de exemplo utilizando a API Bluesoft Cosmos
      var urlApi = "https://bluesoft.com.br" + codigoEan + ".json"; 
      
      var opcoes = {
        "method" : "get",
        "headers" : {
          "X-Cosmos-Token": API_TOKEN,
          "User-Agent": "Cosmos-API-Request"
        },
        "muteHttpExceptions": true
      };
      
      try {
        var resposta = UrlFetchApp.fetch(urlApi, opcoes);
        var codigoStatus = resposta.getResponseCode();
        
        if (codigoStatus === 200) {
          var dados = JSON.parse(resposta.getContentText());
          
          // Preenche as respectivas colunas com os dados retornados
          if (dados.description) aba.getRange(i, COLUNA_SKU).setValue(dados.description); 
          if (dados.ncm && dados.ncm.code) aba.getRange(i, COLUNA_NCM).setValue(dados.ncm.code);
          
          // Se a API retornar dados de imposto/CST padrão
          if (dados.cst_piscofins) {
             aba.getRange(i, COLUNA_CST).setValue(dados.cst_piscofins);
          }
          
          contadorSucesso++;
          
        } else if (codigoStatus === 404) {
          aba.getRange(i, COLUNA_SKU).setValue("Não encontrado no banco fiscal");
        } else if (codigoStatus === 429) {
          SpreadsheetApp.getUi().alert("Limite de consultas da API atingido!");
          return;
        } else {
          aba.getRange(i, COLUNA_SKU).setValue("Erro HTTP: " + codigoStatus);
        }
        
      } catch(erro) {
        aba.getRange(i, COLUNA_SKU).setValue("Erro de Conexão");
      }
      
      // Pausa de 1.2 segundos para respeitar o limite de requisições por minuto de APIs padrão
      Utilities.sleep(1200); 
    }
  }
  
  SpreadsheetApp.getUi().alert("Processo concluído! Itens atualizados: " + contadorSucesso);
}

// Função auxiliar para tratar erros comuns de formatação de código de barras no Sheets
function formatarEanCorreto(valor) {
  if (!valor) return "";
  
  // Se o número veio convertido em notação científica (ex: 7.89E+12)
  if (typeof valor === 'number') {
    valor = valor.toFixed(0); 
  }
  
  var eanLimpo = valor.toString().replace(/\D/g, ""); // Remove letras ou traços estranhos
  
  // Se o código tiver 12 dígitos, adiciona o 0 faltante na frente (padrão EAN-13)
  if (eanLimpo.length === 12) {
    eanLimpo = "0" + eanLimpo;
  }
  
  return eanLimpo;
}
