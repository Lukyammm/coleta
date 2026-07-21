function doGet(e) {
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
    if (defaultSheet.getName() === "Página1" || defaultSheet.getName() === "Sheet1") {
      defaultSheet.setName("Coletas");
      sColetas = defaultSheet;
    } else {
      sColetas = ss.insertSheet("Coletas");
    }
  }
  const headersColetas = ["ID", "Data e Hora", "Paciente", "Prontuário", "Setor", "Conselho Profissional", "Justificativa", "Status"];
  if (sColetas.getLastRow() === 0) {
    sColetas.appendRow(headersColetas);
    sColetas.getRange(1, 1, 1, headersColetas.length).setFontWeight("bold").setBackground("#f4f6f3");
    sColetas.setFrozenRows(1);
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
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Coletas");
    if (!sheet) throw new Error("Banco de dados (Coletas) não inicializado.");
    
    const timestamp = new Date();
    const id = Math.floor(timestamp.getTime()/1000).toString(16).toUpperCase();
    const fullId = '#' + id;
    
    sheet.appendRow([
      fullId,
      timestamp,
      formData.nome,
      formData.prontuario,
      formData.setor,
      formData.conselho,
      formData.justificativa,
      'Pendente'
    ]);
    
    logAction('NOVA_COLETA', `Paciente: ${formData.nome} | ID: ${fullId}`);
    return { ok: true, id: fullId };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

// --- Admin Panel Endpoints ---

function getAdminData(senha) {
  if (senha !== "adm") return { ok: false, erro: "Senha administrativa incorreta." };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sColetas = ss.getSheetByName("Coletas");
    const coletas = sColetas && sColetas.getLastRow() > 0 ? sColetas.getDataRange().getDisplayValues() : [];
    
    const sSetores = ss.getSheetByName("Setores");
    const setores = sSetores && sSetores.getLastRow() > 0 ? sSetores.getDataRange().getDisplayValues() : [];

    const sLogs = ss.getSheetByName("Logs");
    const logs = sLogs && sLogs.getLastRow() > 0 ? sLogs.getDataRange().getDisplayValues() : [];
    
    logAction('LOGIN_ADMIN', 'Acesso autorizado ao painel administrativo.');
    return { ok: true, coletas, setores, logs };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}

function updateStatus(payload) {
  if (payload.senha !== "adm") return { ok: false, erro: "Sem permissão." };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coletas");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { ok: false, erro: "Registro não encontrado." };
    sheet.getRange(rowIndex, 8).setValue(payload.novoStatus);
    
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
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let nome = "";
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        rowIndex = i + 1;
        nome = data[i][1];
        break;
      }
    }
    
    if (rowIndex === -1) return { ok: false, erro: "Setor não encontrado." };
    sheet.deleteRow(rowIndex);
    
    logAction('EXCLUIR_SETOR', `Setor excluído: ${nome} (${payload.id})`);
    return { ok: true };
  } catch (error) {
    return { ok: false, erro: error.toString() };
  }
}
