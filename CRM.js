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
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (row[idxPromotor] && row[idxPromotor].toString().trim().toUpperCase() === chaveJ.trim().toUpperCase()) {
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
    const idxConvenio = headers.indexOf("CONVENIO");
    const idxOrgao = headers.indexOf("ORGAO"); 
    const idxProduto = headers.indexOf("PRODUTO"); // <-- MAPEANDO COLUNA PRODUTO      
    const idxStatus = headers.indexOf("STATUS");

    // Preenche dinamicamente
    if (idxData !== -1) novaLinha[idxData] = new Date();
    if (idxPromotor !== -1) novaLinha[idxPromotor] = dados.promotor;
    if (idxPerfil !== -1) novaLinha[idxPerfil] = dados.perfil;
    if (idxCliente !== -1) novaLinha[idxCliente] = dados.cliente.toUpperCase();
    if (idxCpf !== -1) novaLinha[idxCpf] = dados.cpf;
    if (idxValor !== -1) novaLinha[idxValor] = dados.valor;
    if (idxObs !== -1) novaLinha[idxObs] = dados.obs;
    if (idxBanco !== -1) novaLinha[idxBanco] = dados.banco;
    if (idxConvenio !== -1) novaLinha[idxConvenio] = dados.convenio; 
    if (idxOrgao !== -1) novaLinha[idxOrgao] = dados.orgao;          
    if (idxProduto !== -1) novaLinha[idxProduto] = dados.produto; // <-- GRAVANDO PRODUTO
    
    // Toda proposta nasce com status "NOVA"
    if (idxStatus !== -1) novaLinha[idxStatus] = "NOVA";

    abaPropostas.appendRow(novaLinha);

    // Grava no Log de Auditoria
    saveSystemLog({
      userKey: dados.promotor,
      userProfile: dados.perfil,
      operationType: "Nova Proposta CRM",
      operationResult: "Cliente: " + dados.cliente
    });

    return { sucesso: true, mensagem: "Proposta registada com sucesso!" };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Busca as opções de Convênio, Órgãos e Produtos a partir da aba "Padrao"
 */
function getOpcoesConvenioOrgao() {
  const ss = getDatabaseConnection();
  const sheet = ss.getSheetByName("Padrao");
  if (!sheet) return { convenios: [], mapa: {}, produtos: [] };

  const data = sheet.getDataRange().getValues();
  const convenios = [];
  const mapa = {};
  const produtos = []; // <-- ADICIONADO: Array para armazenar os produtos

  for (let i = 1; i < data.length; i++) {
    let conv = data[i][5]; // Coluna F
    if (conv && conv.toString().trim() !== "") {
      let c = conv.toString().trim().toUpperCase();
      if (!convenios.includes(c)) convenios.push(c);
    }
  }

  for (let i = 1; i < data.length; i++) {
    let convKey = data[i][7]; // Coluna H
    let orgVal = data[i][8];  // Coluna I
    
    if (convKey && orgVal) {
      let cKey = convKey.toString().trim().toUpperCase();
      let oVal = orgVal.toString().trim().toUpperCase();
      
      if (!mapa[cKey]) {
        mapa[cKey] = [];
      }
      if (!mapa[cKey].includes(oVal)) {
        mapa[cKey].push(oVal);
      }
    }
  }

  // --- ADICIONADO: Captura os Produtos únicos da Coluna K (Índice 10) ---
  for (let i = 1; i < data.length; i++) {
    let prod = data[i][10]; // Coluna K
    if (prod && prod.toString().trim() !== "") {
      let p = prod.toString().trim().toUpperCase();
      if (!produtos.includes(p)) produtos.push(p);
    }
  }

  return { convenios: convenios, mapa: mapa, produtos: produtos };
}