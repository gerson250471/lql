/**
 * FICHEIRO: Auth.js
 * Gestão de login e validação de promotores
 */

function autenticarUsuario(chaveJ) {
  try {
    const ss = getDatabaseConnection(); 
    const sheet = ss.getSheetByName("Promotores");
    if (!sheet) return { sucesso: false, erro: "Aba 'Promotores' não encontrada." };

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toUpperCase());
    
    const idxChave = headers.indexOf("CHAVE J");
    const idxNome = headers.indexOf("NOME");
    const idxPerfil = headers.indexOf("PERFIL");
    const idxSituacao = headers.indexOf("SITUAÇÃO");
    const idxMeta = headers.indexOf("META");
    const idxNivel = headers.indexOf("NÍVEL DE ACESSO"); // <-- Mapeamento Dinâmico

    const usuario = data.slice(1).find(row => 
      row[idxChave] && row[idxChave].toString().trim().toUpperCase() === chaveJ.trim().toUpperCase() && 
      row[idxSituacao] && row[idxSituacao].toString().trim().toUpperCase() === "ATIVO"
    );

    if (usuario) {
      return {
        sucesso: true,
        chave: usuario[idxChave], 
        nome: usuario[idxNome],  
        perfil: usuario[idxPerfil],
        meta: usuario[idxMeta],  
        nivelAcesso: idxNivel !== -1 ? usuario[idxNivel] : "PROMOTOR" // <-- Extração Segura com fallback
      };
    }
    return { sucesso: false, erro: "Chave J inválida ou utilizador inativo." };
  } catch (e) {
    return { sucesso: false, erro: "Erro no servidor: " + e.message };
  }
}