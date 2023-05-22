function enviarLembretePagamento() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var recipientEmail = "lucasvt01@gmail.com";
    var recipientEmail2 = "ozana93@gmail.com";
    var abaAtiva = planilha.getSheetByName("Geral");
    var intervalo = abaAtiva.getRange("D1:L");
    
    var filtros = intervalo.getFontWeights();
    var linhasVisiveis = [];
  
    filtros.forEach(function (linha, index) {
      var linhaVisivel = true;
  
      linha.forEach(function (fontWeight, coluna) {
        if (coluna >= 3 && coluna <= 11) {
          if (fontWeight !== "bold") {
            linhaVisivel = false;
          }
        }
      });
  
      if (linhaVisivel) {
        var dataColunaF = abaAtiva.getRange(index + 1, 6).getValue();
        var dataAtual = new Date();
        var dataLimite = new Date();
        dataLimite.setDate(dataLimite.getDate() + 7);
        
        if (dataColunaF >= dataAtual && dataColunaF <= dataLimite) {
          linhasVisiveis.push(index + 1);
        }
      }
    });
  
    if (linhasVisiveis.length > 0) {
      var tabelaFormatada = formatarTabela(linhasVisiveis, intervalo);
      var mensagem = "<html><body>";
      mensagem += "<h2>Lembrete de pagamento:</h2>";
      mensagem += tabelaFormatada;
      mensagem += "</body></html>";
  
      enviarEmail(recipientEmail, mensagem);
      enviarEmail(recipientEmail2, mensagem);
    }
  }
  
  function formatarTabela(linhasVisiveis, intervalo) {
    var tabela = "<table>";
    
    // Cabeçalho da tabela
    tabela += "<tr>";
    for (var coluna = 0; coluna < 9; coluna++) {
      tabela += "<th>" + intervalo.getSheet().getRange(1, coluna + 4).getValue() + "</th>";
    }
    tabela += "</tr>";
    
    // Linhas da tabela
    linhasVisiveis.forEach(function (linhaIndex) {
      var linha = intervalo.getSheet().getRange(linhaIndex, 1, 1, intervalo.getNumColumns()).getValues()[0];
      
      tabela += "<tr>";
      for (var coluna = 5; coluna < 12; coluna++) {
        // Formatar coluna de data para "dd/mm/yyyy"
        if (coluna === 6 || coluna === 5 ) {
          var data = linha[coluna - 1];
          var dataFormatada = Utilities.formatDate(data, "GMT", "dd/MM/yyyy");
          tabela += "<td>" + dataFormatada + "</td>";
        } else {
          tabela += "<td>" + linha[coluna - 1] + "</td>";
        }
      }
      tabela += "</tr>";
    });
  
    tabela += "</table>";
  
    return tabela;
  }
  
  function enviarEmail(recipientEmail, mensagem) {
    var assunto = "Lembrete de pagamento";
  
    MailApp.sendEmail({
      to: recipientEmail,
      subject: assunto,
      htmlBody: mensagem  // Usar htmlBody em vez de body para enviar o e-mail como HTML
    });
  }
  