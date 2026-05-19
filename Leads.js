/**
 * FICHEIRO: Leads.js
 * Lógica do form/subform para higienização e histórico de base
 */

function getLeadsDoPromotor(chaveJ) {
  const ss = getDatabaseConnection(); 
  const sheet = ss.getSheetByName("Leads");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  
  const idxCpf = headers.indexOf("CPF");
  const idxNome = headers.indexOf("NOME");
  const idxCel1 = headers.indexOf("CEL1");
  const idxCel2 = headers.indexOf("CEL2");
  const idxRenda = headers.indexOf("RENDA");
  const idxPromotor = headers.indexOf("PROMOTOR");
  const idxStatus = headers.indexOf("STATUS");

  const leads = [];
  for(let i = 1; i < data.length; i++) {
    let row = data[i];
    if(row[idxPromotor] && row[idxPromotor].toString().trim().toUpperCase() === chaveJ.trim().toUpperCase()) {
      leads.push({
        cpf: row[idxCpf],
        nome: row[idxNome],
        cel1: row[idxCel1],
        cel2: row[idxCel2],
        renda: row[idxRenda],
        status: row[idxStatus] || "NOVO"
      });
    }
  }
  return leads;
}

/**
 * Salva a interação (histórico) na aba HistoricoLeads
 */
function registrarInteracaoLead(payload) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("HistoricoLeads");
    if (!sheet) throw new Error("Aba 'HistoricoLeads' não encontrada.");

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toUpperCase());

    // Se a aba estiver vazia (só com cabeçalhos), prepara uma nova linha vazia
    let novaLinha = new Array(headers.length).fill("");

    const idxData = headers.indexOf("DATA");
    const idxCpf = headers.indexOf("CPF");
    const idxPromotor = headers.indexOf("PROMOTOR");
    const idxAcao = headers.indexOf("AÇÃO");
    const idxObs = headers.indexOf("OBS");

    // Preenche os dados nos lugares certos
    if (idxData !== -1) novaLinha[idxData] = new Date();
    if (idxCpf !== -1) novaLinha[idxCpf] = payload.cpf;
    if (idxPromotor !== -1) novaLinha[idxPromotor] = payload.promotor;
    if (idxAcao !== -1) novaLinha[idxAcao] = payload.acao;
    if (idxObs !== -1) novaLinha[idxObs] = payload.obs;

    // Adiciona a linha na planilha
    sheet.appendRow(novaLinha);

    // BÔNUS: Atualiza o status geral do Lead na aba 'Leads'
    atualizarStatusLead(payload.cpf, payload.acao);

    return true;
  } catch (e) {
    throw new Error("Erro ao gravar histórico: " + e.message);
  }
}

/**
 * Função Bônus: Muda o status na aba Leads de acordo com a ação
 */
function atualizarStatusLead(cpfAlvo, novaAcao) {
  const ss = getDatabaseConnection();
  const abaLeads = ss.getSheetByName("Leads");
  const dados = abaLeads.getDataRange().getValues();
  const headers = dados[0].map(h => h.toString().trim().toUpperCase());
  
  const idxCpf = headers.indexOf("CPF");
  const idxStatus = headers.indexOf("STATUS");

  if (idxCpf === -1 || idxStatus === -1) return;

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][idxCpf] && dados[i][idxCpf].toString().trim() === cpfAlvo.toString().trim()) {
      abaLeads.getRange(i + 1, idxStatus + 1).setValue(novaAcao);
      break;
    }
  }
}

/**
 * Retorna os detalhes do Lead + Histórico dele para a tela
 */
function getDetalhesEHistoricoLead(cpfBusca, chavePromotor) {
  try {
    const ss = getDatabaseConnection();
    
    // 1. Busca os detalhes do Lead
    const abaLeads = ss.getSheetByName("Leads");
    const dadosLeads = abaLeads.getDataRange().getValues();
    const headLeads = dadosLeads[0].map(h => h.toString().trim().toUpperCase());
    
    let leadEncontrado = {};
    const idxCpf = headLeads.indexOf("CPF");
    const idxNome = headLeads.indexOf("NOME");
    const idxNasc = headLeads.indexOf("NASC");
    const idxMae = headLeads.indexOf("NOME_MAE");
    const idxRenda = headLeads.indexOf("RENDA");
    const idxCel1 = headLeads.indexOf("CEL1");
    const idxCel2 = headLeads.indexOf("CEL2");

    for (let i = 1; i < dadosLeads.length; i++) {
      if (dadosLeads[i][idxCpf] && dadosLeads[i][idxCpf].toString().trim() === cpfBusca.toString().trim()) {
        leadEncontrado = {
          cpf: dadosLeads[i][idxCpf].toString(),
          nome: dadosLeads[i][idxNome],
          // Formatação de data simples caso venha como objeto Date do Sheets
          nasc: (dadosLeads[i][idxNasc] instanceof Date) ? Utilities.formatDate(dadosLeads[i][idxNasc], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosLeads[i][idxNasc],
          nomeMae: dadosLeads[i][idxMae] || "-",
          renda: dadosLeads[i][idxRenda] || "N/I",
          cel1: dadosLeads[i][idxCel1] || "-",
          cel2: dadosLeads[i][idxCel2] || ""
        };
        break;
      }
    }

    // 2. Busca o Histórico
    const abaHist = ss.getSheetByName("HistoricoLeads");
    const dadosHist = abaHist.getDataRange().getValues();
    const headHist = dadosHist[0].map(h => h.toString().trim().toUpperCase());
    
    let historicoArray = [];
    const idH_Data = headHist.indexOf("DATA");
    const idH_Cpf = headHist.indexOf("CPF");
    const idH_Acao = headHist.indexOf("AÇÃO");
    const idH_Obs = headHist.indexOf("OBS");

    for (let i = 1; i < dadosHist.length; i++) {
      if (dadosHist[i][idH_Cpf] && dadosHist[i][idH_Cpf].toString().trim() === cpfBusca.toString().trim()) {
        
        let dataFormatada = "-";
        if (dadosHist[i][idH_Data]) {
          let dataBruta = dadosHist[i][idH_Data];
          dataFormatada = (dataBruta instanceof Date) 
            ? Utilities.formatDate(dataBruta, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") 
            : dataBruta.toString();
        }

        historicoArray.push({
          data: dataFormatada,
          acao: dadosHist[i][idH_Acao] ? dadosHist[i][idH_Acao].toString() : "Ação Desconhecida",
          obs: dadosHist[i][idH_Obs] ? dadosHist[i][idH_Obs].toString() : "-"
        });
      }
    }

    // Retorna a fusão dos dois para a tela (em ordem cronológica inversa)
    return { lead: leadEncontrado, historico: historicoArray.reverse() };

  } catch (e) {
    throw new Error("Erro ao buscar detalhes: " + e.message);
  }
}