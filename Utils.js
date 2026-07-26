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
 * Retorna os promotores ativos e os leads que estão sem promotor.
 */
function getDadosPainelAdmin() {
  try {
    const ss = getDatabaseConnection();

    // Função de apoio para ignorar acentos e maiúsculas/minúsculas
    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    // --- 1. BUSCAR PROMOTORES ATIVOS ---
    const abaPromotores = ss.getSheetByName("Promotores");
    if (!abaPromotores) throw new Error("Aba 'Promotores' não encontrada.");

    const dadosPromotores = abaPromotores.getDataRange().getValues();
    if (dadosPromotores.length <= 1) return { promotores: [], leadsLivres: [] };

    const headersPromotores = dadosPromotores[0].map(h => normalizarTexto(h));

    const getColIndexP = (nomesPossiveis) => {
      for (let nome of nomesPossiveis) {
        let idx = headersPromotores.indexOf(normalizarTexto(nome));
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const idxChaveP = getColIndexP(["CHAVE J", "CHAVE_J", "CHAVE"]);
    const idxNomeP = getColIndexP(["NOME"]);
    const idxPerfil = getColIndexP(["PERFIL"]);
    const idxSituacao = getColIndexP(["SITUACAO", "STATUS", "SITUACÃO"]);

    if (idxChaveP === -1 || idxNomeP === -1) {
      throw new Error("Colunas 'CHAVE' ou 'NOME' não encontradas na aba Promotores.");
    }

    let listaPromotores = [];
    for (let i = 1; i < dadosPromotores.length; i++) {
      let row = dadosPromotores[i];
      let chaveVal = row[idxChaveP] ? String(row[idxChaveP]).trim() : "";
      
      // Escudo anti-linha vazia
      if (!chaveVal) continue;

      let situacaoVal = idxSituacao !== -1 && row[idxSituacao] 
        ? normalizarTexto(row[idxSituacao]) 
        : "ATIVO";

      // Aceita apenas quem estiver com Situação ATIVO
      if (situacaoVal === "ATIVO") {
        listaPromotores.push({
          chave: chaveVal,
          nome: row[idxNomeP] ? String(row[idxNomeP]).trim() : "Promotor sem Nome",
          perfil: idxPerfil !== -1 && row[idxPerfil] ? String(row[idxPerfil]).trim() : "BLACK"
        });
      }
    }

    // --- 2. BUSCAR LEADS LIVRES ---
    const abaLeads = ss.getSheetByName("Leads");
    if (!abaLeads) throw new Error("Aba 'Leads' não encontrada.");

    const dadosLeads = abaLeads.getDataRange().getValues();
    const headersLeads = dadosLeads[0].map(h => normalizarTexto(h));

    const getColIndexL = (nomesPossiveis) => {
      for (let nome of nomesPossiveis) {
        let idx = headersLeads.indexOf(normalizarTexto(nome));
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const idxCpf = getColIndexL(["CPF"]);
    const idxNomeL = getColIndexL(["NOME"]);
    const idxRenda = getColIndexL(["RENDA"]);
    const idxPromotorLead = getColIndexL(["PROMOTOR"]); 

    let leadsLivres = [];
    for (let i = 1; i < dadosLeads.length; i++) {
      let row = dadosLeads[i];
      let cpfVal = row[idxCpf] ? String(row[idxCpf]).trim() : "";
      
      if (!cpfVal) continue;
      
      // Se a coluna PROMOTOR estiver vazia, o lead está livre
      let promotorAssociado = idxPromotorLead !== -1 && row[idxPromotorLead] 
        ? String(row[idxPromotorLead]).trim() 
        : "";

      if (!promotorAssociado) {
        leadsLivres.push({
          cpf: cpfVal,
          nome: idxNomeL !== -1 && row[idxNomeL] ? String(row[idxNomeL]).trim() : "SEM NOME",
          renda: idxRenda !== -1 && row[idxRenda] ? String(row[idxRenda]).trim() : "N/I"
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