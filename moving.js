/*
  > função para criar nova aba com o plano mensal nas planilhas de projeto
  
  > histórico de revisões
      - 20250512 - R03
        - autor: Henrique
        - observações:
        - 
*/

function modeloR05() {
    var idPastaTodosProjetos = 'ID-DA-PASTA';
    var pastaTodosProjetos = DriveApp.getFolderById(idPastaTodosProjetos);
    var arquivosProjetos = pastaTodosProjetos.getFilesByType(MimeType.GOOGLE_SHEETS);
  
    var masterPlan = SpreadsheetApp.getActiveSpreadsheet();
    var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    var mesAtual = meses[new Date().getMonth()];
    var abaMasterPlan = masterPlan.getSheetByName(mesAtual);
    if (!abaMasterPlan) {
      Logger.log("Erro: Aba do mês '" + mesAtual + "' não encontrada no MasterPlan.");
      return;
    }
  
    var ultimaLinha = abaMasterPlan.getLastRow();
    var colA = abaMasterPlan.getRange(1, 1, ultimaLinha).getValues();
    var fontColors = abaMasterPlan.getRange(1, 1, ultimaLinha).getFontColors();
    var backgroundColors = abaMasterPlan.getRange(1, 1, ultimaLinha).getBackgrounds();
  
    var corReferencia = "#d0cece";
    var blocosPorProjeto = {};
  
    Logger.log("Iniciando varredura no MasterPlan...");
    for (var i = 0; i < colA.length; i++) {
      if (fontColors[i][0].toLowerCase() === corReferencia) {
        var nomeProjeto = colA[i][0].trim();
        var chaveProjeto = nomeProjeto.toLowerCase();
        Logger.log("Projeto encontrado: '" + nomeProjeto + "'");
  
        var bloco = [];
  
        // Adiciona a linha com o nome do projeto
        var linhaNomeProjeto = abaMasterPlan.getRange(i + 1, 1, 1, 6).getValues()[0];
        bloco.push(linhaNomeProjeto);
  
        var j = i + 1;
        while (
          j < colA.length &&
          fontColors[j][0].toLowerCase() !== corReferencia
        ) {
          var linhaParcial = abaMasterPlan.getRange(j + 1, 1, 1, 6).getValues()[0];
          var corFundoLinha = backgroundColors[j][0];
          if (corFundoLinha === "#ffffff") {
            bloco.push(linhaParcial);
          }
          j++;
        }
  
        if (bloco.length > 0) {
          blocosPorProjeto[chaveProjeto] = bloco;
          Logger.log("  Linhas coletadas: " + bloco.length);
        }
  
        i = j - 1;
      }
    }
  
    Logger.log("Buscando arquivos de projeto...");
    while (arquivosProjetos.hasNext()) {
      var arquivo = arquivosProjetos.next();
      var nomeArquivo = arquivo.getName().trim();
      var chaveArquivo = nomeArquivo.toLowerCase();
  
      Logger.log("Verificando arquivo: " + nomeArquivo);
  
      if (blocosPorProjeto[chaveArquivo]) {
        var bloco = blocosPorProjeto[chaveArquivo];
        var planilha = SpreadsheetApp.open(arquivo);
        var abaDestino = planilha.getSheetByName("MSP");
  
        if (!abaDestino) {
          abaDestino = planilha.insertSheet("MSP");
          Logger.log("Aba 'MSP' criada.");
        } else {
          abaDestino.clearContents();
          Logger.log("Aba 'MSP' limpa.");
        }
  
        abaDestino.getRange(2, 1, bloco.length, 6).setValues(bloco);
        Logger.log("Bloco colado no projeto: " + nomeArquivo);
      } else {
        Logger.log("Nenhum bloco correspondente ao arquivo: " + nomeArquivo);
      }
    }
  
    Logger.log("Processo concluído.");
  }
  