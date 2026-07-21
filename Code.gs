function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Clinify')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function processForm(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    const headers = ["ID", "Data e Hora", "Paciente", "Prontuário", "Conselho Profissional", "Justificativa"];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f4f6f3");
      sheet.setFrozenRows(1);
    } else {
      const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      if (firstRow[0] === "") {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f4f6f3");
        sheet.setFrozenRows(1);
      }
    }
    
    const timestamp = new Date();
    // Unique ID generation (similar to diet-request patterns)
    const id = Math.floor(timestamp.getTime()/1000).toString(16).toUpperCase();
    
    sheet.appendRow([
      '#' + id,
      timestamp,
      formData.nome,
      formData.prontuario,
      formData.conselho,
      formData.justificativa
    ]);
    
    return { ok: true, id: '#' + id };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function getRegistros(senha) {
  if (senha !== "adm") {
    return { ok: false, erro: "Senha administrativa incorreta." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    if(sheet.getLastRow() === 0) return {ok: true, data: []};
    const data = sheet.getDataRange().getDisplayValues();
    return { ok: true, data: data };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}
