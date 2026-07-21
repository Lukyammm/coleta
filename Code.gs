function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Coleta de Dados')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no'); 
    // user-scalable=no is very important for iOS Safari to feel like a native app and avoid zooming on input focus
}

function processForm(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    
    // Auto-configuration: verify if headers exist. If not, create them.
    const headers = ["Data e Hora", "Nome do Paciente", "Prontuário", "Justificativa do Procedimento", "Número do Conselho Profissional"];
    
    // Check if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      // Optional: make headers bold
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
      sheet.setFrozenRows(1);
    } else {
      // Just to be sure, check if first row matches headers. If it's completely empty but has rows, we overwrite row 1.
      const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      if (firstRow[0] === "") {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
        sheet.setFrozenRows(1);
      }
    }
    
    // Insert data with server-side timestamp
    const timestamp = new Date();
    sheet.appendRow([
      timestamp,
      formData.nome,
      formData.prontuario,
      formData.justificativa,
      formData.conselho
    ]);
    
    return { success: true, message: "Coleta registrada com sucesso!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
