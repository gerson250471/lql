/**
 * Busca todas as listas de opções financeiras da aba Parametros_Fin
 */
function getParametrosFinanceiros() {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Parametros_Fin");
    if (!sheet) throw new Error("Aba 'Parametros_Fin' não encontrada no Sheets.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { sucesso: true, dados: { tipos: [], categoriasEntrada: [], categoriasSaida: [], fontes: [], formasPagto: [], status: [] } };
    }

    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));

    let params = { tipos: [], categoriasEntrada: [], categoriasSaida: [], fontes: [], formasPagto: [], status: [] };

    const colTipo = headers.indexOf(normalizarTexto("TIPO"));
    const colCatEnt = headers.indexOf(normalizarTexto("CATEGORIA ENTRADA"));
    const colCatSai = headers.indexOf(normalizarTexto("CATEGORIA SAIDA"));
    const colFonte = headers.indexOf(normalizarTexto("FONTE"));
    const colForma = headers.indexOf(normalizarTexto("FORMA DE PAGTO"));
    const colStatus = headers.indexOf(normalizarTexto("STATUS"));

    // Coleta as opções ignorando células vazias
    for (let i = 1; i < data.length; i++) {
      if (colTipo !== -1 && data[i][colTipo]) params.tipos.push(String(data[i][colTipo]).trim().toUpperCase());
      if (colCatEnt !== -1 && data[i][colCatEnt]) params.categoriasEntrada.push(String(data[i][colCatEnt]).trim().toUpperCase());
      if (colCatSai !== -1 && data[i][colCatSai]) params.categoriasSaida.push(String(data[i][colCatSai]).trim().toUpperCase());
      if (colFonte !== -1 && data[i][colFonte]) params.fontes.push(String(data[i][colFonte]).trim().toUpperCase());
      if (colForma !== -1 && data[i][colForma]) params.formasPagto.push(String(data[i][colForma]).trim().toUpperCase());
      if (colStatus !== -1 && data[i][colStatus]) params.status.push(String(data[i][colStatus]).trim().toUpperCase());
    }

    return { sucesso: true, dados: params };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Salva as listas alteradas pela interface do utilizador
 */
function salvarParametrosFinanceiros(params) {
  try {
    const ss = getDatabaseConnection();
    let sheet = ss.getSheetByName("Parametros_Fin");
    if (!sheet) throw new Error("Aba 'Parametros_Fin' não encontrada no Sheets.");

    // Limpa os dados antigos (preservando o cabeçalho)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    // Encontra o tamanho da maior lista para saber quantas linhas criar
    const maxLen = Math.max(
      params.tipos.length, params.categoriasEntrada.length, params.categoriasSaida.length,
      params.fontes.length, params.formasPagto.length, params.status.length
    );

    if (maxLen === 0) return { sucesso: true }; // Tudo vazio

    let newData = [];
    for (let i = 0; i < maxLen; i++) {
      newData.push([
        params.tipos[i] || "",
        params.categoriasEntrada[i] || "",
        params.categoriasSaida[i] || "",
        params.fontes[i] || "",
        params.formasPagto[i] || "",
        params.status[i] || ""
      ]);
    }

    // Grava as novas listas de uma vez só na planilha
    sheet.getRange(2, 1, newData.length, 6).setValues(newData);
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}