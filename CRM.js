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
 * Salva uma nova proposta na aba "Propostas" com ID Automático
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

    // MAPEAMENTO DE COLUNAS
    const idxId = headers.indexOf("ID");             // <-- ADICIONADO MAPEAMENTO DO ID
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
    const idxProduto = headers.indexOf("PRODUTO");       
    const idxStatus = headers.indexOf("STATUS");

    // PREENCHIMENTO DINÂMICO
    // Gera um código único baseado na data e hora (Ex: PRP-17189345)
    if (idxId !== -1) novaLinha[idxId] = "PRP-" + new Date().getTime(); 
    
    if (idxData !== -1) novaLinha[idxData] = new Date();
    if (idxPromotor !== -1) novaLinha[idxPromotor] = dados.promotor;
    if (idxPerfil !== -1) novaLinha[idxPerfil] = dados.perfil;
    if (idxCliente !== -1) novaLinha[idxCliente] = dados.cliente.toUpperCase();
    if (idxCpf !== -1) novaLinha[idxCpf] = dados.cpf;
    if (idxValor !== -1) novaLinha[idxValor] = dados.valor;
    if (idxObs !== -1) novaLinha[idxObs] = dados.obs;
    if (idxBanco !== -1) novaLinha[idxBanco] = dados.banco.toUpperCase();
    if (idxConvenio !== -1) novaLinha[idxConvenio] = dados.convenio; 
    if (idxOrgao !== -1) novaLinha[idxOrgao] = dados.orgao;          
    if (idxProduto !== -1) novaLinha[idxProduto] = dados.produto;
    
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

/**
 * Busca os dados completos de uma proposta específica pelo ID para edição
 */
function obterPropostaPorId(idProposta) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Propostas");
    if (!sheet) throw new Error("Aba 'Propostas' não encontrada.");

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toUpperCase());

    const idxId = headers.indexOf("ID");
    const idxCliente = headers.indexOf("CLIENTE");
    const idxCpf = headers.indexOf("CPF");
    const idxValor = headers.indexOf("VALOR");
    const idxBanco = headers.indexOf("BANCO");
    const idxConvenio = headers.indexOf("CONVENIO");
    const idxOrgao = headers.indexOf("ORGAO");
    const idxProduto = headers.indexOf("PRODUTO");
    const idxObs = headers.indexOf("OBS");

    // Procura a linha correspondente ao ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][idxId] && data[i][idxId].toString().trim() === idProposta.toString().trim()) {
        return {
          sucesso: true,
          proposta: {
            id: data[i][idxId],
            cliente: data[i][idxCliente],
            cpf: data[i][idxCpf],
            valor: data[i][idxValor],
            banco: data[i][idxBanco],
            convenio: data[i][idxConvenio],
            orgao: data[i][idxOrgao],
            produto: data[i][idxProduto],
            obs: data[i][idxObs]
          }
        };
      }
    }
    return { sucesso: false, erro: "Proposta não encontrada." };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Atualiza os dados de uma proposta existente na aba "Propostas"
 */
function atualizarPropostaExistente(dados) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Propostas");
    if (!sheet) throw new Error("Aba 'Propostas' não encontrada.");

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0].map(h => h.toString().trim().toUpperCase());

    const idxId = headers.indexOf("ID");
    const idxCliente = headers.indexOf("CLIENTE");
    const idxCpf = headers.indexOf("CPF");
    const idxValor = headers.indexOf("VALOR");
    const idxBanco = headers.indexOf("BANCO");
    const idxConvenio = headers.indexOf("CONVENIO");
    const idxOrgao = headers.indexOf("ORGAO");
    const idxProduto = headers.indexOf("PRODUTO");
    const idxObs = headers.indexOf("OBS");

    let linhaIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idxId] && data[i][idxId].toString().trim() === dados.id.toString().trim()) {
        linhaIndex = i + 1; // Ajusta para o índice real da planilha (base 1)
        break;
      }
    }

    if (linhaIndex === -1) throw new Error("Proposta não localizada para atualização.");

    // Atualiza apenas as colunas permitidas na linha correspondente
    if (idxCliente !== -1) sheet.getCell(linhaIndex, idxCliente + 1).setValue(dados.cliente.toUpperCase());
    if (idxCpf !== -1) sheet.getCell(linhaIndex, idxCpf + 1).setValue(dados.cpf);
    if (idxValor !== -1) sheet.getCell(linhaIndex, idxValor + 1).setValue(dados.valor);
    if (idxBanco !== -1) sheet.getCell(linhaIndex, idxBanco + 1).setValue(dados.banco.toUpperCase());
    if (idxConvenio !== -1) sheet.getCell(linhaIndex, idxConvenio + 1).setValue(dados.convenio);
    if (idxOrgao !== -1) sheet.getCell(linhaIndex, idxOrgao + 1).setValue(dados.orgao);
    if (idxProduto !== -1) sheet.getCell(linhaIndex, idxProduto + 1).setValue(dados.produto);
    if (idxObs !== -1) sheet.getCell(linhaIndex, idxObs + 1).setValue(dados.obs);

    // Grava no Log de Auditoria
    saveSystemLog({
      userKey: dados.promotor,
      userProfile: dados.perfil,
      operationType: "Edição Proposta CRM",
      operationResult: "ID: " + dados.id + " | Cliente: " + dados.cliente
    });

    return { sucesso: true, mensagem: "Proposta atualizada com sucesso!" };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}