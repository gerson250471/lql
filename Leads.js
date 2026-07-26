/**
 * FICHEIRO: Leads.js
 * Lógica do form/subform para higienização, cadastro completo do Back-Office e histórico de base
 */

/**
 * Retorna a lista simples de Leads atribuídos a um promotor específico
 */
function getLeadsDoPromotor(chaveJ) {
  const ss = getDatabaseConnection(); 
  const sheet = ss.getSheetByName("Leads");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  
  const idxCpf = headers.indexOf("CPF");
  const idxNome = headers.indexOf("NOME");
  const idxCel1 = headers.indexOf("CEL1");
  const idxRenda = headers.indexOf("RENDA");
  const idxPromotor = headers.indexOf("PROMOTOR");
  const idxStatus = headers.indexOf("STATUS");

  const leads = [];
  const buscaPromotor = chaveJ.trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (row[idxPromotor] && row[idxPromotor].toString().trim().toUpperCase() === buscaPromotor) {
      leads.push({
        cpf: row[idxCpf] ? row[idxCpf].toString().trim() : "-",
        nome: row[idxNome] || "N/I",
        cel1: row[idxCel1] || "-",
        renda: row[idxRenda] || "N/I",
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

    let novaLinha = new Array(headers.length).fill("");

    const idxData = headers.indexOf("DATA");
    const idxCpf = headers.indexOf("CPF");
    const idxPromotor = headers.indexOf("PROMOTOR");
    const idxAcao = headers.indexOf("AÇÃO");
    let idxObs = headers.indexOf("OBS");
    if (idxObs === -1) idxObs = headers.indexOf("OBSERVACAO");

    if (idxData !== -1) novaLinha[idxData] = new Date();
    if (idxCpf !== -1) novaLinha[idxCpf] = payload.cpf;
    if (idxPromotor !== -1) novaLinha[idxPromotor] = payload.promotor;
    if (idxAcao !== -1) novaLinha[idxAcao] = payload.acao;
    if (idxObs !== -1) novaLinha[idxObs] = payload.obs;

    sheet.appendRow(novaLinha);

    // Atualiza o status geral do Lead na aba 'Leads'
    atualizarStatusLead(payload.cpf, payload.acao);

    return true;
  } catch (e) {
    throw new Error("Erro ao gravar histórico: " + e.message);
  }
}

/**
 * Muda o status na aba Leads de acordo com a ação
 */
function atualizarStatusLead(cpfAlvo, novaAcao) {
  const ss = getDatabaseConnection();
  const abaLeads = ss.getSheetByName("Leads");
  if (!abaLeads) return;

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
 * Retorna os detalhes COMPLETOS do Lead (com a nova estrutura de Back-Office) + Histórico
 */
function getDetalhesEHistoricoLead(cpfBusca, chavePromotor) {
  try {
    const ss = getDatabaseConnection();
    
    // 1. Busca os detalhes estendidos do Lead na aba 'Leads'
    const abaLeads = ss.getSheetByName("Leads");
    const dadosLeads = abaLeads.getDataRange().getValues();
    const headLeads = dadosLeads[0].map(h => h.toString().trim().toUpperCase());
    
    let leadEncontrado = {};
    
    const getIdx = (nome) => headLeads.indexOf(nome);

    const idxFonte = getIdx("FONTE");
    const idxCpf = getIdx("CPF");
    const idxNome = getIdx("NOME");
    const idxMae = getIdx("NOME_MAE");
    const idxSexo = getIdx("SEXO");
    const idxNasc = getIdx("NASC");
    const idxRenda = getIdx("RENDA");
    const idxProduto = getIdx("PRODUTO");
    const idxLogradouro = getIdx("LOGRADOURO");
    const idxNumero = getIdx("NUMERO");
    const idxBairro = getIdx("BAIRRO");
    const idxCidade = getIdx("CIDADE");
    const idxUf = getIdx("UF");
    const idxCep = getIdx("CEP");
    const idxCel1 = getIdx("CEL1");
    const idxEmail1 = getIdx("EMAIL1");

    for (let i = 1; i < dadosLeads.length; i++) {
      let cpfLinha = dadosLeads[i][idxCpf] ? dadosLeads[i][idxCpf].toString().trim() : "";
      
      if (cpfLinha === cpfBusca.toString().trim()) {
        let rawNasc = dadosLeads[i][idxNasc];
        let nascFormatado = "-";
        if (rawNasc) {
          nascFormatado = (rawNasc instanceof Date) 
            ? Utilities.formatDate(rawNasc, Session.getScriptTimeZone(), "dd/MM/yyyy") 
            : rawNasc.toString();
        }

        leadEncontrado = {
          fonte: idxFonte !== -1 ? (dadosLeads[i][idxFonte] || "-") : "-",
          cpf: cpfLinha,
          nome: idxNome !== -1 ? (dadosLeads[i][idxNome] || "-") : "-",
          nomeMae: idxMae !== -1 ? (dadosLeads[i][idxMae] || "-") : "-",
          sexo: idxSexo !== -1 ? (dadosLeads[i][idxSexo] || "-") : "-",
          nasc: nascFormatado,
          renda: idxRenda !== -1 ? (dadosLeads[i][idxRenda] || "N/I") : "N/I",
          produto: idxProduto !== -1 ? (dadosLeads[i][idxProduto] || "-") : "-",
          logradouro: idxLogradouro !== -1 ? (dadosLeads[i][idxLogradouro] || "-") : "-",
          numero: idxNumero !== -1 ? (dadosLeads[i][idxNumero] || "-") : "-",
          bairro: idxBairro !== -1 ? (dadosLeads[i][idxBairro] || "-") : "-",
          cidade: idxCidade !== -1 ? (dadosLeads[i][idxCidade] || "-") : "-",
          uf: idxUf !== -1 ? (dadosLeads[i][idxUf] || "-") : "-",
          cep: idxCep !== -1 ? (dadosLeads[i][idxCep] || "-") : "-",
          cel1: idxCel1 !== -1 ? (dadosLeads[i][idxCel1] || "-") : "-",
          email1: idxEmail1 !== -1 ? (dadosLeads[i][idxEmail1] || "-") : "-"
        };
        break;
      }
    }

    // 2. Busca o Histórico de interações
    const abaHist = ss.getSheetByName("HistoricoLeads");
    let historicoArray = [];

    if (abaHist) {
      const dadosHist = abaHist.getDataRange().getValues();
      const headHist = dadosHist[0].map(h => h.toString().trim().toUpperCase());
      
      const idH_Data = headHist.indexOf("DATA");
      const idH_Cpf = headHist.indexOf("CPF");
      const idH_Acao = headHist.indexOf("AÇÃO");
      let idH_Obs = headHist.indexOf("OBS");
      if (idH_Obs === -1) idH_Obs = headHist.indexOf("OBSERVACAO");

      for (let i = 1; i < dadosHist.length; i++) {
        let cpfHist = dadosHist[i][idH_Cpf] ? dadosHist[i][idH_Cpf].toString().trim() : "";
        if (cpfHist === cpfBusca.toString().trim()) {
          let dataFormatada = "-";
          if (dadosHist[i][idH_Data]) {
            let dataBruta = dadosHist[i][idH_Data];
            dataFormatada = (dataBruta instanceof Date) 
              ? Utilities.formatDate(dataBruta, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") 
              : dataBruta.toString();
          }

          historicoArray.push({
            data: dataFormatada,
            acao: idH_Acao !== -1 ? (dadosHist[i][idH_Acao] || "Ação Desconhecida") : "Ação Desconhecida",
            obs: idH_Obs !== -1 ? (dadosHist[i][idH_Obs] || "-") : "-"
          });
        }
      }
    }

    return { lead: leadEncontrado, historico: historicoArray.reverse() };

  } catch (e) {
    throw new Error("Erro ao buscar detalhes do lead: " + e.message);
  }
}