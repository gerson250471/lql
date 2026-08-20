/**
 * Salva ou Atualiza um Promotor na base de dados
 */
function salvarPromotorBase(dados) {
  try {
    const ss = getDatabaseConnection();
    let aba = ss.getSheetByName("Promotores");
    
    if (!aba) {
      throw new Error("Aba 'Promotores' não encontrada no Sheets.");
    }

    const dataRows = aba.getDataRange().getValues();
    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = dataRows[0].map(h => normalizarTexto(h));

    const getIdx = (nome) => headers.indexOf(normalizarTexto(nome));

    const idxChave = getIdx("CHAVE");
    const idxNome = getIdx("NOME");
    const idxEmail = getIdx("EMAIL");
    const idxSenha = getIdx("SENHA");
    const idxNivel = getIdx("NIVEL_ACESSO");
    const idxPerfil = getIdx("PERFIL");
    const idxMeta = getIdx("META");
    const idxStatus = getIdx("STATUS");

    if (idxChave === -1 || idxNome === -1) {
      throw new Error("Colunas CHAVE ou NOME não encontradas na aba Promotores.");
    }

    let isEdicao = false;
    let rowIndex = -1;

    // Verifica se a chave (Chave J) já existe
    for (let i = 1; i < dataRows.length; i++) {
      if (String(dataRows[i][idxChave]).toUpperCase() === String(dados.chave).toUpperCase()) {
        isEdicao = true;
        rowIndex = i + 1; // +1 porque a planilha começa no índice 1
        break;
      }
    }

    let novaLinha = new Array(headers.length).fill("");

    // Preenche a linha com os dados do formulário
    if (idxChave !== -1) novaLinha[idxChave] = String(dados.chave).toUpperCase();
    if (idxNome !== -1) novaLinha[idxNome] = String(dados.nome).toUpperCase();
    if (idxEmail !== -1) novaLinha[idxEmail] = String(dados.email).toLowerCase();
    
    // Se for um novo cadastro, define senha inicial. Se for edição, mantém a antiga.
    if (idxSenha !== -1) {
      novaLinha[idxSenha] = isEdicao ? dataRows[rowIndex - 1][idxSenha] : "123456"; 
    }
    
    // Coluna para controlar se a pessoa precisa trocar a senha no 1º acesso (opcional)
    const idxTrocarSenha = getIdx("TROCAR_SENHA");
    if (idxTrocarSenha !== -1) {
        novaLinha[idxTrocarSenha] = isEdicao ? dataRows[rowIndex - 1][idxTrocarSenha] : "SIM";
    }

    if (idxNivel !== -1) novaLinha[idxNivel] = String(dados.nivel).toUpperCase();
    if (idxPerfil !== -1) novaLinha[idxPerfil] = String(dados.perfil).toUpperCase();
    if (idxMeta !== -1) novaLinha[idxMeta] = Number(dados.meta) || 0;
    if (idxStatus !== -1) novaLinha[idxStatus] = String(dados.status || "ATIVO").toUpperCase();

    // Grava na planilha
    if (isEdicao) {
      aba.getRange(rowIndex, 1, 1, headers.length).setValues([novaLinha]);
    } else {
      aba.appendRow(novaLinha);
    }

    return { sucesso: true, mensagem: isEdicao ? "Promotor atualizado com sucesso!" : "Promotor cadastrado com sucesso! Senha inicial: 123456" };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Busca a lista de todos os promotores cadastrados para a tabela de Gestão
 */
function getTodosPromotores() {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Promotores");
    if (!sheet) throw new Error("Aba 'Promotores' não encontrada no banco de dados.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: true, dados: [] };

    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));

    const getIdx = (nome) => headers.indexOf(normalizarTexto(nome));
    const idxChave = getIdx("CHAVE");
    const idxNome = getIdx("NOME");
    const idxEmail = getIdx("EMAIL");
    const idxPerfil = getIdx("PERFIL");
    const idxNivel = getIdx("NIVEL_ACESSO");
    const idxStatus = getIdx("STATUS");
    const idxMeta = getIdx("META");

    if (idxChave === -1) throw new Error("Coluna CHAVE não encontrada.");

    let promotores = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][idxChave]) continue; // Pula linhas vazias
      
      promotores.push({
        chave: data[i][idxChave],
        nome: idxNome !== -1 ? data[i][idxNome] : "-",
        email: idxEmail !== -1 ? data[i][idxEmail] : "",
        perfil: idxPerfil !== -1 ? data[i][idxPerfil] : "BLACK",
        nivel: idxNivel !== -1 ? data[i][idxNivel] : "PROMOTOR",
        status: idxStatus !== -1 ? data[i][idxStatus] : "ATIVO",
        meta: idxMeta !== -1 ? Number(data[i][idxMeta]) : 0
      });
    }

    return { sucesso: true, dados: promotores };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}