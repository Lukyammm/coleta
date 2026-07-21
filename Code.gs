function doGet(e) {
  try { initDB(); } catch(e){}
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Clinify')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function initDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet: Coletas
  let sColetas = ss.getSheetByName("Coletas");
  if (!sColetas) {
    let defaultSheet = ss.getSheets()[0];
    if (defaultSheet.getName() === "Página1" || defaultSheet.getName() === "Sheet1" || defaultSheet.getName() === "Página 1") {
      defaultSheet.setName("Coletas");
      sColetas = defaultSheet;
    } else {
      sColetas = ss.insertSheet("Coletas");
    }
  }
  
  // Assegura que todas as colunas necessárias existem
  const expectedHeaders = ["ID", "Data e Hora", "Paciente", "Prontuário", "Setor", "Conselho Profissional", "Justificativa", "Status"];
  if (sColetas.getLastRow() === 0) {
    sColetas.appendRow(expectedHeaders);
    sColetas.getRange(1, 1, 1, expectedHeaders.length).setFontWeight("bold").setBackground("#f4f6f3");
    sColetas.setFrozenRows(1);
  } else {
    // Adiciona colunas ausentes
    let headers = sColetas.getRange(1, 1, 1, sColetas.getLastColumn()).getValues()[0];
    for (let h of expectedHeaders) {
      if (headers.indexOf(h) === -1 && headers.indexOf(h.replace(' e Hora', '')) === -1 && headers.indexOf(h.replace(' Profissional', '')) === -1) {
        sColetas.getRange(1, headers.length + 1).setValue(h).setFontWeight("bold").setBackground("#f4f6f3");
        headers.push(h);
      }
    }
  }

  // Sheet: Setores
  let sSetores = ss.getSheetByName("Setores");
  if (!sSetores) {
    sSetores = ss.insertSheet("Setores");
    sSetores.appendRow(["ID", "Nome do Setor", "Criado Em"]);
    sSetores.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f4f6f3");
    sSetores.setFrozenRows(1);
    sSetores.appendRow(["#SETOR_GERAL", "Geral", new Date()]);
  }

  // Sheet: Logs
  let sLogs = ss.getSheetByName("Logs");
  if (!sLogs) {
    sLogs = ss.insertSheet("Logs");
    sLogs.appendRow(["Data e Hora", "Ação", "Detalhes"]);
    sLogs.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f4f6f3");
    sLogs.setFrozenRows(1);
  }
}

function logAction(action, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sLogs = ss.getSheetByName("Logs");
    if (sLogs) sLogs.appendRow([new Date(), action, details]);
  } catch(e){}
}

function getBootData() {
  initDB();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sSetores = ss.getSheetByName("Setores");
  let setores = [];
  if (sSetores && sSetores.getLastRow() > 1) {
    const data = sSetores.getRange(2, 2, sSetores.getLastRow() - 1, 1).getValues();
    setores = data.map(r => r[0]).filter(Boolean);
  }
  if (setores.length === 0) setores = ["Geral"];
  return { ok: true, setores: setores };
}

function processForm(formData) {
  initDB();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Coletas");
    if (!sheet) throw new Error("Banco de dados não inicializado.");
    
    const timestamp = new Date();
    const id = '#' + Math.floor(timestamp.getTime()/1000).toString(16).toUpperCase();
    
    // Mapeamento dinâmico de colunas para suportar planilhas antigas do usuário
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill("");
    
    const hIdx = (name, alt) => {
      let i = headers.indexOf(name);
      if (i === -1 && alt) i = headers.indexOf(alt);
      return i;
    };
    
    const map = {
      "ID": id,
      "Data e Hora": timestamp,
      "Data": timestamp,
      "Paciente": formData.nome,
      "Nome": formData.nome,
      "Prontuário": formData.prontuario,
      "Setor": formData.setor || "Geral",
      "Conselho Profissional": formData.conselho,
      "Conselho": formData.conselho,
      "Justificativa": formData.justificativa,
      "Status": "Pendente"
    };
    
    for (let i = 0; i < headers.length; i++) {
      if (map[headers[i]] !== undefined) {
        newRow[i] = map[headers[i]];
      }
    }
    
    sheet.appendRow(newRow);
    
    logAction('NOVA_COLETA', `Paciente: ${formData.nome} | ID: ${id}`);
    return { ok: true, id: id };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function getAdminData(senha) {
  initDB();
  if (senha !== "adm") return { ok: false, erro: "Senha administrativa incorreta." };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sColetas = ss.getSheetByName("Coletas");
    const coletas = sColetas && sColetas.getLastRow() > 0 ? sColetas.getDataRange().getDisplayValues() : [];
    
    const sSetores = ss.getSheetByName("Setores");
    const setores = sSetores && sSetores.getLastRow() > 0 ? sSetores.getDataRange().getDisplayValues() : [];

    const sLogs = ss.getSheetByName("Logs");
    const logs = sLogs && sLogs.getLastRow() > 0 ? sLogs.getDataRange().getDisplayValues() : [];
    
    logAction('LOGIN_ADMIN', 'Acesso autorizado ao painel.');
    return { ok: true, coletas, setores, logs };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function updateStatus(payload) {
  if (payload.senha !== "adm") return { ok: false, erro: "Sem permissão." };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coletas");
    const displayData = sheet.getDataRange().getDisplayValues();
    const headers = displayData[0];
    
    let idIdx = headers.indexOf("ID");
    if (idIdx === -1) idIdx = 0; // fallback se o usuário não tem coluna ID
    
    let statusIdx = headers.indexOf("Status");
    if (statusIdx === -1) {
      statusIdx = headers.length;
      sheet.getRange(1, statusIdx + 1).setValue("Status").setFontWeight("bold");
    }
    
    let rowIndex = -1;
    for (let i = 1; i < displayData.length; i++) {
      if (String(displayData[i][idIdx]).trim() === String(payload.id).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { ok: false, erro: "Registro não encontrado." };
    
    sheet.getRange(rowIndex, statusIdx + 1).setValue(payload.novoStatus);
    logAction('ATUALIZAR_STATUS', `Coleta ${payload.id} alterada para ${payload.novoStatus}`);
    return { ok: true };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function addSetor(payload) {
  if (payload.senha !== "adm") return { ok: false, erro: "Sem permissão." };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setores");
    const id = '#S_' + Math.floor(new Date().getTime()/1000).toString(16).toUpperCase();
    sheet.appendRow([id, payload.nome, new Date()]);
    logAction('CRIAR_SETOR', `Novo setor cadastrado: ${payload.nome}`);
    return { ok: true, novoSetor: [id, payload.nome, new Date().toLocaleString()] };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function deleteSetor(payload) {
  if (payload.senha !== "adm") return { ok: false, erro: "Sem permissão." };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setores");
    const displayData = sheet.getDataRange().getDisplayValues();
    let rowIndex = -1;
    let nome = "";
    
    for (let i = 0; i < displayData.length; i++) {
      if (String(displayData[i][0]).trim() === String(payload.id).trim()) {
        rowIndex = i + 1;
        nome = displayData[i][1];
        break;
      }
    }
    
    if (rowIndex === -1) return { ok: false, erro: "Setor não encontrado." };
    sheet.deleteRow(rowIndex);
    logAction('EXCLUIR_SETOR', `Setor excluído: ${nome}`);
    return { ok: true };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}
