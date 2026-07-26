/**
 * FICHEIRO: Utils.js
 * Funções globais e utilitárias usadas por todo o sistema.
 */

function getDatabaseConnection() {
  try {
    // MAGIA AQUI: O sistema lê a Variável de Ambiente do servidor onde ele está rodando
    const sheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    
    if (!sheetId) {
      throw new Error("ALERTA DE SISTEMA: SPREADSHEET_ID não está configurado nas Propriedades do Script no painel do Google.");
    }
    
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    throw new Error("Erro de permissão ou falha ao conectar no Banco de Dados. Verifique o ID e o compartilhamento da planilha.");
  }
}

function saveSystemLog(logData) {
  const dbConnection = getDatabaseConnection();
  let logSheet = dbConnection.getSheetByName("Logs");
  
  if (!logSheet) {
    logSheet = dbConnection.insertSheet("Logs");
    logSheet.appendRow(["Data/Hora", "Chave J", "Operação", "Perfil", "Resultado"]);
  }
  
  logSheet.appendRow([
    new Date(), 
    logData.userKey, 
    logData.operationType, 
    logData.userProfile, 
    logData.operationResult
  ]);
  // teste
}

/**
 * Função: Carrega os dados para a tela do Administrador (Atribuição de Leads)
 */
function getDadosPainelAdmin() {
  try {
    const ss = getDatabaseConnection();

    // --- 1. BUSCAR PROMOTORES ATIVOS ---
    const abaPromotores = ss.getSheetByName("Promotores");
    const dadosPromotores = abaPromotores.getDataRange().getValues();
    const headersPromotores = dadosPromotores[0].map(h => h.toString().trim().toUpperCase());
    
    let idxChaveP = headersPromotores.indexOf("CHAVE J");
    if (idxChaveP === -1) idxChaveP = headersPromotores.indexOf("CHAVE_J");
    if (idxChaveP === -1) idxChaveP = headersPromotores.indexOf("CHAVE");

    const idxNomeP = headersPromotores.indexOf("NOME");
    const idxPerfil = headersPromotores.indexOf("PERFIL");
    const idxSituacao = headersPromotores.indexOf("SITUAÇÃO");

    let listaPromotores = [];
    for (let i = 1; i < dadosPromotores.length; i++) {
      let row = dadosPromotores[i];
      if (!row[idxChaveP] || row[idxChaveP].toString().trim() === "") continue;

      if (row[idxSituacao] && row[idxSituacao].toString().trim().toUpperCase() === "ATIVO") {
        listaPromotores.push({
          chave: row[idxChaveP],
          nome: row[idxNomeP],
          perfil: row[idxPerfil]
        });
      }
    }

    // --- 2. BUSCAR LEADS LIVRES (Com busca dinâmica pelo novo layout de colunas) ---
    const abaLeads = ss.getSheetByName("Leads");
    const dadosLeads = abaLeads.getDataRange().getValues();
    const headersLeads = dadosLeads[0].map(h => h.toString().trim().toUpperCase());
    
    const idxCpf = headersLeads.indexOf("CPF");
    const idxNomeL = headersLeads.indexOf("NOME");
    const idxRenda = headersLeads.indexOf("RENDA");
    const idxPromotorLead = headersLeads.indexOf("PROMOTOR"); 

    let leadsLivres = [];
    for (let i = 1; i < dadosLeads.length; i++) {
      let row = dadosLeads[i];
      
      if (!row[idxCpf] || row[idxCpf].toString().trim() === "") continue;
      
      // Se a coluna PROMOTOR estiver vazia, o lead está disponível para atribuição
      if (idxPromotorLead === -1 || !row[idxPromotorLead] || row[idxPromotorLead].toString().trim() === "") {
        leadsLivres.push({
          cpf: row[idxCpf].toString().trim(),
          nome: row[idxNomeL] || "SEM NOME",
          renda: idxRenda !== -1 && row[idxRenda] ? row[idxRenda].toString().trim() : "N/I"
        });
      }
    }

    return { promotores: listaPromotores, leadsLivres: leadsLivres };

  } catch (e) {
    throw new Error("Erro ao buscar dados do painel admin: " + e.message);
  }
}

/**
 * Função 2: Grava a atribuição na Planilha
 * Recebe a lista de CPFs selecionados e a chave J do Promotor escolhido.
 */
function atribuirLeadsEmMassa(arrayCpfs, chavePromotor) {
  try {
    const ss = getDatabaseConnection();
    const abaLeads = ss.getSheetByName("Leads");
    const dadosLeads = abaLeads.getDataRange().getValues();
    const headersLeads = dadosLeads[0].map(h => h.toString().trim().toUpperCase());
    
    const idxCpf = headersLeads.indexOf("CPF");
    const idxPromotorLead = headersLeads.indexOf("PROMOTOR");
    const idxStatus = headersLeads.indexOf("STATUS");

    // Passa linha por linha na aba de Leads
    for (let i = 1; i < dadosLeads.length; i++) {
      let valorCpf = dadosLeads[i][idxCpf];
      
      // ESCUDO: Se não houver CPF, pula a linha para evitar erro de .toString()
      if (!valorCpf || valorCpf.toString().trim() === "") continue;
      
      let cpfLinha = valorCpf.toString().trim();
      
      // Se o CPF desta linha estiver na lista de selecionados
      if (arrayCpfs.includes(cpfLinha)) {
        // Atualiza a coluna Promotor com a Chave J (i + 1 para compensar o índice)
        abaLeads.getRange(i + 1, idxPromotorLead + 1).setValue(chavePromotor);
        
        // Coloca o status como "NOVO" para destacar na tela do promotor
        if(idxStatus !== -1) {
          abaLeads.getRange(i + 1, idxStatus + 1).setValue("NOVO");
        }
      }
    }
    
    return true; 
  } catch (e) {
    throw new Error("Erro ao salvar atribuições: " + e.message);
  }
}