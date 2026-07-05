/**
 * Busca a tabela de parâmetros de comissão na aba 'Padrao' (Linhas 15 a 29)
 */
function getParametrosPadrao() {
  try {
    const ss = getDatabaseConnection(); 
    const sheet = ss.getSheetByName("Padrao");
    if (!sheet) throw new Error("Aba 'Padrao' não encontrada.");

    const intervalo = sheet.getRange("A15:E29").getValues();
    
    const cabecalhos = intervalo[0].map(h => h.toString().trim().toUpperCase());
    const dados = intervalo.slice(1).map(row => {
      return {
        grupo: row[0].toString().trim(),
        black: row[1],
        gold: row[2],
        gestor: row[3],
        top: row[4]
      };
    });

    return { sucesso: true, cabecalhos: cabecalhos, dados: dados };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Grava as alterações feitas pela Gestão na aba 'Padrao'
 * @param {Array} dados Array 2D contendo apenas as colunas editáveis
 */
function atualizarParametrosPadrao(dados) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Padrao");
    if (!sheet) throw new Error("Aba 'Padrao' não encontrada.");

    // A área de escrita começa na linha 16 (abaixo do cabeçalho)
    // e na coluna 2 (B - BLACK), ignorando a coluna A (Grupos)
    const linhaInicial = 16;
    const colunaInicial = 2; 
    const numLinhas = dados.length;
    const numColunas = 4; // São 4 perfis editáveis
    
    // Grava de uma só vez (muito mais rápido e seguro)
    sheet.getRange(linhaInicial, colunaInicial, numLinhas, numColunas).setValues(dados);
    
    // <-- CORREÇÃO VITAL: Força a planilha a recalcular todas as fórmulas instantaneamente!
    SpreadsheetApp.flush();
    
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}