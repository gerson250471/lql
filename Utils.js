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
 * Função 1: Carrega os dados para a tela do Administrador
 * Retorna os promotores ativos e os leads que estão sem promotor.
 */
function getDadosPainelAdmin() {
  try {
    const ss = getDatabaseConnection(); // Usando sua função de conexão existente

    // --- 1. BUSCAR PROMOTORES ATIVOS ---
    const abaPromotores = ss.getSheetByName("Promotores");
    const dadosPromotores = abaPromotores.getDataRange().getValues();
    const headersPromotores = dadosPromotores[0].map(h => h.toString().trim().toUpperCase());
    
    const idxChave = headersPromotores.indexOf("CHAVE J");
    const idxNomeP = headersPromotores.indexOf("NOME");
    const idxPerfil = headersPromotores.indexOf("PERFIL");
    const idxSituacao = headersPromotores.indexOf("SITUAÇÃO");

    let listaPromotores = [];
    for (let i = 1; i < dadosPromotores.length; i++) {
      let row = dadosPromotores[i];
      
      // ESCUDO: Se a linha for completamente vazia (sem chave), pula.
      if (!row[idxChave] || row[idxChave].toString().trim() === "") continue;

      // Pega apenas quem está ATIVO
      if (row[idxSituacao] && row[idxSituacao].toString().trim().toUpperCase() === "ATIVO") {
        listaPromotores.push({
          chave: row[idxChave],
          nome: row[idxNomeP],
          perfil: row[idxPerfil]
        });
      }
    }

    // --- 2. BUSCAR LEADS LIVRES ---
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
      
      // ESCUDO: Se a linha for fantasma (sem CPF), pula imediatamente.
      if (!row[idxCpf] || row[idxCpf].toString().trim() === "") continue;
      
      // Verifica se a coluna Promotor está vazia (lead livre)
      if (!row[idxPromotorLead] || row[idxPromotorLead].toString().trim() === "") {
        leadsLivres.push({
          cpf: row[idxCpf].toString().trim(),
          nome: row[idxNomeL],
          // Apenas lê o texto da faixa de renda como está na planilha
          renda: row[idxRenda] ? row[idxRenda].toString().trim() : "N/I"
        });
      }
    }

    return { promotores: listaPromotores, leadsLivres: leadsLivres };

  } catch (e) {
    throw new Error("Erro ao buscar dados do painel: " + e.message);
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