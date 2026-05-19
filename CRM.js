/**
 * FICHEIRO: CRM.js
 * Gestão de propostas avulsas e funil de vendas
 */

function getPropostas(chaveJ) {
  const ss = getDatabaseConnection();
  const sheet = ss.getSheetByName("Propostas");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  
  const idxId = headers.indexOf("ID");
  const idxData = headers.indexOf("DATA");
  const idxPromotor = headers.indexOf("PROMOTOR");
  const idxCliente = headers.indexOf("CLIENTE");
  const idxValor = headers.indexOf("VALOR");
  const idxStatus = headers.indexOf("STATUS");

  const propostas = [];
  for(let i = 1; i < data.length; i++) {
    let row = data[i];
    if(row[idxPromotor] && row[idxPromotor].toString().trim().toUpperCase() === chaveJ.trim().toUpperCase()) {
      propostas.push({
        id: row[idxId],
        data: row[idxData] ? Utilities.formatDate(new Date(row[idxData]), "GMT-3", "dd/MM/yyyy") : "-",
        cliente: row[idxCliente],
        valor: row[idxValor],
        status: row[idxStatus] || "NOVA"
      });
    }
  }
  
  return propostas.reverse();
}

function salvarNovaProposta(dadosProposta) {
  try {
    const ss = getDatabaseConnection(); 
    let sheet = ss.getSheetByName("Propostas");
    if (!sheet) {
      sheet = ss.insertSheet("Propostas");
      sheet.appendRow(["ID", "DATA", "PROMOTOR", "CLIENTE", "CPF", "VALOR", "STATUS", "OBSERVACOES"]);
    }

    const novoId = Utilities.getUuid();
    const dataAtual = new Date();
    
    sheet.appendRow([
      novoId,
      dataAtual,
      dadosProposta.promotor,
      dadosProposta.cliente.toUpperCase(),
      dadosProposta.cpf,
      dadosProposta.valor,
      "NOVA",
      dadosProposta.obs
    ]);

    saveSystemLog({
      userKey: dadosProposta.promotor,
      userProfile: dadosProposta.perfil,
      operationType: "Nova Proposta CRM",
      operationResult: "Cliente: " + dadosProposta.cliente
    });

    return { sucesso: true, mensagem: "Proposta registada com sucesso!" };
  } catch (e) {
    return { sucesso: false, erro: "Erro ao guardar proposta: " + e.message };
  }
}

/**
 * Salva uma nova proposta na aba "Propostas"
 */
function salvarNovaProposta(dados) {
  try {
    const ss = getDatabaseConnection();
    let abaPropostas = ss.getSheetByName("Propostas");
    
    if (!abaPropostas) {
      throw new Error("Aba 'Propostas' não foi encontrada na planilha.");
    }

    const dataRows = abaPropostas.getDataRange().getValues();
    const headers = dataRows[0].map(h => h.toString().trim().toUpperCase());

    let novaLinha = new Array(headers.length).fill("");

    const idxData = headers.indexOf("DATA");
    const idxPromotor = headers.indexOf("PROMOTOR");
    const idxPerfil = headers.indexOf("PERFIL");
    const idxCliente = headers.indexOf("CLIENTE");
    const idxCpf = headers.indexOf("CPF");
    const idxValor = headers.indexOf("VALOR");
    const idxObs = headers.indexOf("OBS");
    const idxBanco = headers.indexOf("BANCO");
    const idxStatus = headers.indexOf("STATUS");

    // Preenche dinamicamente
    if(idxData !== -1) novaLinha[idxData] = new Date();
    if(idxPromotor !== -1) novaLinha[idxPromotor] = dados.promotor;
    if(idxPerfil !== -1) novaLinha[idxPerfil] = dados.perfil;
    if(idxCliente !== -1) novaLinha[idxCliente] = dados.cliente;
    if(idxCpf !== -1) novaLinha[idxCpf] = dados.cpf;
    if(idxValor !== -1) novaLinha[idxValor] = dados.valor;
    if(idxObs !== -1) novaLinha[idxObs] = dados.obs;
    if(idxBanco !== -1) novaLinha[idxBanco] = dados.banco;
    
    // Toda proposta nasce com status "NOVA"
    if(idxStatus !== -1) novaLinha[idxStatus] = "NOVA"; 

    abaPropostas.appendRow(novaLinha);

    return { sucesso: true };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}