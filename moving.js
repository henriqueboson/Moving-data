  /*
    > função para criar nova aba com o masterplan do mês atual
    
    > histórico de revisões
        - 20250519 - R02
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

        var blocoValores = [];
        var blocoBgs = [];
        var blocoFontColors = [];
        var blocoFontWeights = [];
        var blocoFontStyles = [];

        var j = i;
        while (
          j < colA.length &&
          (j === i || fontColors[j][0].toLowerCase() !== corReferencia)
        ) {
          var linhaRange = abaMasterPlan.getRange(j + 1, 1, 1, 6);
          var valores = linhaRange.getValues()[0];
          var bg = linhaRange.getBackgrounds()[0];
          var fontColor = linhaRange.getFontColors()[0];
          var fontWeight = linhaRange.getFontWeights()[0];
          var fontStyle = linhaRange.getFontStyles()[0];

          var corFonteLinha = fontColor[0].toLowerCase();
          if (corFonteLinha.toLowerCase() !== "#ffff27" && corFonteLinha.toLowerCase() !== "#00b050") {
            blocoValores.push(valores);
            blocoBgs.push(bg);
            blocoFontColors.push(fontColor);
            blocoFontWeights.push(fontWeight);
            blocoFontStyles.push(fontStyle);
          }  else {
                // Se a linha for amarela ou verde, adicionar valores padrão
                blocoValores.push(["", "", "", "", "", ""]);
                blocoBgs.push(["#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff"]);
                blocoFontColors.push(["#000000", "#000000", "#000000", "#000000", "#000000", "#000000"]);
                blocoFontWeights.push(["normal", "normal", "normal", "normal", "normal", "normal"]);
                blocoFontStyles.push(["normal", "normal", "normal", "normal", "normal", "normal"]);
              }

          j++;
        }

        if (blocoValores.length > 0) {
          blocosPorProjeto[chaveProjeto] = {
            valores: blocoValores,
            bgs: blocoBgs,
            fontColors: blocoFontColors,
            fontWeights: blocoFontWeights,
            fontStyles: blocoFontStyles
          };
          Logger.log("  Linhas coletadas: " + blocoValores.length);
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
          abaDestino.getRange(2, 1, abaDestino.getMaxRows() - 1, abaDestino.getMaxColumns()).clearContent();
          Logger.log("Aba 'MSP' limpa.");
        }

        var linhas = bloco.valores.length;
        var colunas = bloco.valores[0].length;
        var destinoRange = abaDestino.getRange(2, 1, linhas, colunas);

        destinoRange.setValues(bloco.valores);
        destinoRange.setBackgrounds(bloco.bgs);
        destinoRange.setFontColors(bloco.fontColors);
        destinoRange.setFontWeights(bloco.fontWeights);
        destinoRange.setFontStyles(bloco.fontStyles);

        Logger.log("Bloco colado com formatação no projeto: " + nomeArquivo);
      } else {
        Logger.log("Nenhum bloco correspondente ao arquivo: " + nomeArquivo);
      }
    }

    Logger.log("Processo concluído.");
  }
