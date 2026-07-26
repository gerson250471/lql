/**
 * Atualiza a senha do usuário e altera TROCAR_SENHA para "NÃO"
 */
function alterarSenhaUsuario(chaveUsuario, novaSenha) {
  try {
    if (!chaveUsuario || !novaSenha) {
      return { sucesso: false, erro: "Chave de usuário ou nova senha inválida." };
    }

    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Promotores");
    if (!sheet) throw new Error("Aba 'Promotores' não encontrada.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: false, erro: "Nenhum usuário cadastrado." };

    // Normalizador de cabeçalhos sem acentos
    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));

    const getColIndex = (nomesPossiveis) => {
      for (let nome of nomesPossiveis) {
        let idx = headers.indexOf(normalizarTexto(nome));
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const idxChave = getColIndex(["CHAVE J", "CHAVE_J", "CHAVE"]);
    const idxSenha = getColIndex(["SENHA"]);
    const idxTrocarSenha = getColIndex(["TROCAR_SENHA", "TROCAR SENHA"]);

    if (idxChave === -1 || idxSenha === -1) {
      throw new Error("Colunas obrigatórias (CHAVE e SENHA) não encontradas na aba Promotores.");
    }

    const novoHash = gerarHashSHA256(novaSenha);
    const chaveBusca = String(chaveUsuario).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const chaveLinha = String(data[i][idxChave]).trim().toLowerCase();

      if (chaveLinha === chaveBusca) {
        // Grava o Hash SHA-256 da nova senha na coluna SENHA
        sheet.getRange(i + 1, idxSenha + 1).setValue(novoHash);

        // Se a coluna TROCAR_SENHA existir, altera para "NÃO"
        if (idxTrocarSenha !== -1) {
          sheet.getRange(i + 1, idxTrocarSenha + 1).setValue("NÃO");
        }

        return { sucesso: true, mensagem: "Senha alterada com sucesso!" };
      }
    }

    return { sucesso: false, erro: "Usuário não encontrado para alteração." };

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
 * Valida o login trazendo todos os campos do perfil, inclusive a META
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

    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));

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
    const idxSituacao = getColIndex(["SITUACAO", "STATUS"]);
    const idxPerfil = getColIndex(["PERFIL"]);
    const idxMeta = getColIndex(["META"]); // <-- CAPTURA DA COLUNA META

    if (idxChave === -1 || idxEmail === -1 || idxSenha === -1) {
      throw new Error("Colunas obrigatórias não encontradas na aba Promotores.");
    }

    const termoBusca = String(identificador).trim().toLowerCase();
    const hashDigitado = gerarHashSHA256(senhaDigitada);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const chaveLinha = String(row[idxChave]).trim().toLowerCase();
      const emailLinha = String(row[idxEmail]).trim().toLowerCase();
      const situacao = idxSituacao !== -1 ? normalizarTexto(row[idxSituacao]) : "ATIVO";

      if (termoBusca === chaveLinha || termoBusca === emailLinha) {
        
        if (situacao === "INATIVO" || situacao === "BLOQUEADO") {
          return { sucesso: false, erro: "Usuário inativo. Entre em contato com a gestão." };
        }

        const senhaSalva = String(row[idxSenha]).trim();

        if (senhaSalva === hashDigitado) {
          return {
            sucesso: true,
            usuario: {
              chave: row[idxChave],
              nome: idxNome !== -1 ? row[idxNome] : "Usuário",
              email: row[idxEmail],
              perfil: idxPerfil !== -1 ? row[idxPerfil] : "BLACK",
              nivelAcesso: idxNivel !== -1 ? row[idxNivel] : "USUARIO",
              meta: idxMeta !== -1 ? Number(row[idxMeta]) || 0 : 0, // <-- RETORNA A META DO BANCO
              trocarSenha: idxTrocarSenha !== -1 && normalizarTexto(row[idxTrocarSenha]) === "SIM"
            }
          };
        } else {
          return { sucesso: false, erro: "Senha incorreta." };
        }
      }
    }

    return { sucesso: false, erro: "Usuário não encontrado." };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Solicitação de recuperação de senha: gera senha temporária, salva o Hash, marca TROCAR_SENHA="SIM" e envia e-mail.
 */
function solicitarRecuperacaoSenha(identificador) {
  try {
    if (!identificador) {
      return { sucesso: false, erro: "Informe a Chave J ou E-mail." };
    }

    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Promotores");
    if (!sheet) throw new Error("Aba 'Promotores' não encontrada.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: false, erro: "Nenhum usuário cadastrado." };

    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));

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

    if (idxChave === -1 || idxEmail === -1 || idxSenha === -1) {
      throw new Error("Colunas obrigatórias não encontradas na aba Promotores.");
    }

    const termoBusca = String(identificador).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const chaveLinha = String(row[idxChave]).trim().toLowerCase();
      const emailLinha = String(row[idxEmail]).trim().toLowerCase();

      if (termoBusca === chaveLinha || termoBusca === emailLinha) {
        const emailDestino = row[idxEmail];
        const nomeUsuario = idxNome !== -1 ? row[idxNome] : "Promotor";

        if (!emailDestino || !emailDestino.includes("@")) {
          return { sucesso: false, erro: "O usuário não possui um e-mail válido cadastrado." };
        }

        // Gera uma senha temporária única (Ex: LQL-8F2D1K)
        const senhaProvisoria = "LQL-" + Math.random().toString(36).substring(2, 8).toUpperCase();
        const hashProvisorio = gerarHashSHA256(senhaProvisoria);

        // Atualiza na planilha: Hash SHA-256 e marca para trocar no próximo acesso
        sheet.getRange(i + 1, idxSenha + 1).setValue(hashProvisorio);
        if (idxTrocarSenha !== -1) {
          sheet.getRange(i + 1, idxTrocarSenha + 1).setValue("SIM");
        }

        // Envia o e-mail
        const assunto = "LQL Soluções - Recuperação de Acesso";
        const corpoHtml = `
          <div style="font-family: Arial, sans-serif; color: #333; max-width: 500px; margin: 0 auto; border: 1px solid #e2e8f0; padding: 20px; rounded-radius: 10px;">
            <h2 style="color: #1e3a8a; text-align: center;">LQL SOLUÇÕES</h2>
            <p>Olá, <strong>${nomeUsuario}</strong>!</p>
            <p>Recebemos uma solicitação de recuperação de senha para a sua conta.</p>
            <p>A sua senha provisória de acesso é:</p>
            <div style="background-color: #f1f5f9; padding: 12px; text-align: center; font-size: 20px; font-weight: bold; letter-spacing: 2px; color: #1e40af; border-radius: 8px;">
              ${senhaProvisoria}
            </div>
            <p style="margin-top: 15px; font-size: 12px; color: #64748b;">
              <strong>Atenção:</strong> Por motivos de segurança, ao efetuar o login com esta senha, você será solicitado a cadastrar uma nova senha definitiva.
            </p>
          </div>
        `;

        MailApp.sendEmail({
          to: emailDestino,
          subject: assunto,
          htmlBody: corpoHtml
        });

        return { 
          sucesso: true, 
          mensagem: `Senha temporária enviada para o e-mail ${emailDestino.substring(0, 3)}***@***` 
        };
      }
    }

    return { sucesso: false, erro: "Usuário ou e-mail não encontrado." };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}