function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Clinify')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function processForm(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    const headers = ["ID", "Data e Hora", "Paciente", "Prontuário", "Conselho Profissional", "Justificativa", "Status"];
    
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
      } else if (firstRow.length < 7 || firstRow[6] !== "Status") {
        // If the table exists but doesn't have the Status column (migrating from old version)
        sheet.getRange(1, 7).setValue("Status").setFontWeight("bold").setBackground("#f4f6f3");
      }
    }
    
    const timestamp = new Date();
    const id = Math.floor(timestamp.getTime()/1000).toString(16).toUpperCase();
    
    sheet.appendRow([
      '#' + id,
      timestamp,
      formData.nome,
      formData.prontuario,
      formData.conselho,
      formData.justificativa,
      'Pendente'
    ]);
    
    return { ok: true, id: '#' + id };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function getRegistros(senha) {
  if (senha !== "adm") return { ok: false, erro: "Senha administrativa incorreta." };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if(sheet.getLastRow() === 0) return {ok: true, data: []};
    const data = sheet.getDataRange().getDisplayValues();
    return { ok: true, data: data };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function updateStatus(payload) {
  if (payload.senha !== "adm") return { ok: false, erro: "Senha administrativa incorreta." };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Find row by ID
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        rowIndex = i + 1; // 1-indexed for Apps Script ranges
        break;
      }
    }
    
    if (rowIndex === -1) return { ok: false, erro: "Registro não encontrado." };
    
    // Update the Status column (column 7)
    sheet.getRange(rowIndex, 7).setValue(payload.novoStatus);
    
    return { ok: true };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}
