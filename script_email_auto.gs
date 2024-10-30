// Função principal para enviar emails com anexos
function sendEmails() {
  // Obtém a planilha ativa e define a linha inicial e o número de linhas a processar
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Linha inicial (ignora o cabeçalho)
  var numRows = sheet.getLastRow() - 1; // Número de linhas com dados

  // Define o intervalo de dados (linhas e colunas a serem processadas)
  var dataRange = sheet.getRange(startRow, 1, numRows, 4); // Colunas A a D
  var data = dataRange.getValues(); // Armazena os dados em um array 2D

  // ID da pasta no Google Drive onde os arquivos PDF estão armazenados
  var folderId = '12rYbthkunxY0l9CedWKsgwqvh7n2Aisk'; // Insira aqui os ID da pasta com os arquivos
  var folder = DriveApp.getFolderById(folderId);

  // Loop pelos dados de cada linha da planilha
  data.forEach(function(row) {
    // Extrai informações de cada coluna (email, nome do arquivo, assunto e corpo da mensagem)
    var emailAddress = row[0];
    var pdfFileName = row[1];
    var subject = row[2];
    var body = row[3];

    // Cria uma mensagem personalizada de corpo do email com o conteúdo adicional
    var customBody = `
Prezado cliente,

Segue em anexo relatório trimestral de acompanhamento de sua carteira de investimentos, referente ao 3º trimestre de 2024.

Ficamos à disposição.

Atenciosamente,

Bruno Xavier, CEA - FH Advisors
` + body;

    // Verifica se o arquivo PDF existe na pasta, com base no nome especificado
    var files = folder.getFilesByName(pdfFileName);
    if (files.hasNext()) {
      var file = files.next();
      var attachment = file.getBlob(); // Obtém o arquivo em formato Blob para anexar

      // Envia o email com o anexo
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        body: customBody,
        attachments: [attachment]
      });
      
      // Registra no Logger que o email foi enviado com sucesso
      Logger.log('Email enviado para ' + emailAddress);
    } else {
      // Caso o arquivo não seja encontrado, registra no Logger
      Logger.log('Arquivo não encontrado: ' + pdfFileName);
    }
  });
}
