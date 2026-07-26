/**
 * Atualiza a senha do usuário e altera TROCAR_SENHA para "NÃO"
 */
function alterarSenhaUsuario(chaveUsuario, novaSenha) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Promotores");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toUpperCase());

    const idxChave = headers.indexOf("CHAVE");
    const idxSenha = headers.indexOf("SENHA");
    let idxTrocarSenha = headers.indexOf("TROCAR_SENHA");
    if (idxTrocarSenha === -1) idxTrocarSenha = headers.indexOf("TROCAR SENHA");

    // Converte a nova senha para Hash SHA-256 antes de salvar na planilha
    const novoHash = gerarHashSHA256(novaSenha);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxChave]).trim().toUpperCase() === String(chaveUsuario).trim().toUpperCase()) {

        // Atualiza a Senha em Hash SHA-256
        sheet.getRange(i + 1, idxSenha + 1).setValue(novoHash);

        // Atualiza TROCAR_SENHA para "NÃO"
        sheet.getRange(i + 1, idxTrocarSenha + 1).setValue("NÃO");

        return { sucesso: true, mensagem: "Senha alterada com sucesso!" };
      }
    }

    return { sucesso: false, erro: "Usuário não encontrado." };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Converte uma texto simples em Hash SHA-256 (Hexadecimal)
 */
function gerarHashSHA256(texto) {
  if (!texto) return "";
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(texto), Utilities.Charset.UTF_8);
  return rawHash.map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
}

/**
 * EXECUTE UMA VEZ: Criptografa todas as senhas em texto puro da aba 'Promotores'
 */
function criptografarSenhasExistentes() {
  const ss = getDatabaseConnection();
  const sheet = ss.getSheetByName("Promotores");
  if (!sheet) throw new Error("Aba 'Promotores' não encontrada.");

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  const idxSenha = headers.indexOf("SENHA");

  if (idxSenha === -1) {
    throw new Error("Coluna 'SENHA' não encontrada.");
  }

  let alterados = 0;

  for (let i = 1; i < data.length; i++) {
    const senhaAtual = String(data[i][idxSenha]).trim();
    
    // Verifica se a senha já não é um hash SHA-256 (Hashes SHA-256 possuem exatamente 64 caracteres em Hex)
    if (senhaAtual !== "" && senhaAtual.length !== 64) {
      const senhaHash = gerarHashSHA256(senhaAtual);
      sheet.getRange(i + 1, idxSenha + 1).setValue(senhaHash);
      alterados++;
    }
  }

  Logger.log(`🔒 Sucesso: ${alterados} senhas foram criptografadas com SHA-256!`);
}

/**
 * Valida o login aceitando E-mail ou CHAVE J, com busca flexível de cabeçalhos sem acentos
 */
function validarLogin(identificador, senhaDigitada) {
  try {
    if (!identificador || !senhaDigitada) {
      return { sucesso: false, erro: "Informe o usuário e a senha." };
    }

    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Promotores");
    if (!sheet) throw new Error("Aba 'Promotores' não encontrada.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: false, erro: "Nenhum promotor cadastrado." };

    // Normaliza os cabeçalhos removendo acentos para evitar falhas de comparação
    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));

    // Busca flexível de colunas (com ou sem acentos / com ou sem sublinhado)
    const getColIndex = (nomesPossiveis) => {
      for (let nome of nomesPossiveis) {
        let idx = headers.indexOf(normalizarTexto(nome));
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const idxChave = getColIndex(["CHAVE J", "CHAVE_J", "CHAVE"]);
    const idxEmail = getColIndex(["EMAIL", "E-MAIL"]);
    const idxSenha = getColIndex(["SENHA"]);
    const idxTrocarSenha = getColIndex(["TROCAR_SENHA", "TROCAR SENHA"]);
    const idxNome = getColIndex(["NOME"]);
    const idxNivel = getColIndex(["NIVEL DE ACESSO", "NIVEL_DE_ACESSO", "PERMISSAO"]);
    const idxSituacao = getColIndex(["SITUACAO", "SITUACÃO", "STATUS"]);
    const idxPerfil = getColIndex(["PERFIL"]);

    if (idxChave === -1 || idxEmail === -1 || idxSenha === -1) {
      throw new Error("Colunas obrigatórias (CHAVE, EMAIL, SENHA) não encontradas na aba Promotores.");
    }

    const termoBusca = String(identificador).trim().toLowerCase();
    const hashDigitado = gerarHashSHA256(senhaDigitada);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const chaveLinha = String(row[idxChave]).trim().toLowerCase();
      const emailLinha = String(row[idxEmail]).trim().toLowerCase();
      const situacao = idxSituacao !== -1 ? normalizarTexto(row[idxSituacao]) : "ATIVO";

      // Aceita CHAVE J ou E-MAIL
      if (termoBusca === chaveLinha || termoBusca === emailLinha) {
        
        if (situacao === "INATIVO" || situacao === "BLOQUEADO") {
          return { sucesso: false, erro: "Usuário inativo. Entre em contato com a gestão." };
        }

        const senhaSalva = String(row[idxSenha]).trim();

        // Comparação por Hash SHA-256
        if (senhaSalva === hashDigitado) {
          return {
            sucesso: true,
            usuario: {
              chave: row[idxChave],
              nome: idxNome !== -1 ? row[idxNome] : "Usuário",
              email: row[idxEmail],
              perfil: idxPerfil !== -1 ? row[idxPerfil] : "BLACK",
              nivelAcesso: idxNivel !== -1 ? row[idxNivel] : "USUARIO",
              trocarSenha: idxTrocarSenha !== -1 && normalizarTexto(row[idxTrocarSenha]) === "SIM"
            }
          };
        } else {
          return { sucesso: false, erro: "Senha incorreta." };
        }
      }
    }

    return { sucesso: false, erro: "Usuário ou e-mail não encontrado." };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}