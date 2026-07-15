/**
 * Busca a produção oficial filtrada por Promotor, Mês e Ano.
 */
function getProducaoPromotor(chavePromotor, mesFiltro, anoFiltro) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("bd_Producao");
    if (!sheet) throw new Error("Aba 'bd_Producao' não encontrada.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: true, dados: [] };

    const headers = data[0].map(h => h.toString().trim().toUpperCase());
    const idxChave = headers.indexOf("CHAVE J");
    const idxAno = headers.indexOf("ANO");
    const idxMes = headers.indexOf("MÊS"); // Conforme mapeado na sua lista

    if (idxChave === -1 || idxAno === -1 || idxMes === -1) {
      throw new Error("As colunas 'CHAVE J', 'ANO' ou 'MÊS' não foram encontradas.");
    }

    // Mapeamento dinâmico
    const idxDataMov = headers.indexOf("DATA MOVIMENTO");
    const idxConvenio = headers.indexOf("CONVENIO");
    const idxContrato = headers.indexOf("CONTRATO");
    const idxDataCont = headers.indexOf("DATA CONTRATO");
    const idxTaxa = headers.indexOf("TAXA");
    const idxParcela = headers.indexOf("PARCELA");
    const idxRestricao = headers.indexOf("RESTRICAO_RCC");
    const idxGrupo = headers.indexOf("GRUPO");
    const idxProduto = headers.indexOf("PRODUTO");
    const idxDesc = headers.indexOf("DESCRIÇÃO DO PRODUTO");
    const idxComissao = headers.indexOf("COMISSÃO");
    const idxValor = headers.indexOf("VALOR");
    const idxValorCons = headers.indexOf("VALOR CONSIDERADO");
    const idxPagoEm = headers.indexOf("PAGO EM");

    const formatData = (val) => val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy") : (val ? val.toString() : "-");
    const formatNum = (val) => typeof val === 'number' ? val : 0;

    const producaoLimpa = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Formata a célula do Excel/Sheets para garantir que comparações de texto funcionem
      const anoLinha = String(row[idxAno]).trim();
      const mesLinha = String(row[idxMes]).trim().padStart(2, '0'); // Garante que "7" vire "07"
      const mesBusca = String(mesFiltro).padStart(2, '0');

      // O Grande Filtro: Chave + Ano + Mês
      if (row[idxChave] === chavePromotor && anoLinha === String(anoFiltro) && mesLinha === mesBusca) {
        producaoLimpa.push({
          dataMovimento: formatData(row[idxDataMov]),
          convenio: row[idxConvenio] || "-",
          contrato: row[idxContrato] || "-",
          dataContrato: formatData(row[idxDataCont]),
          taxa: row[idxTaxa] || "-",
          parcela: formatNum(row[idxParcela]),
          restricao: row[idxRestricao] || "-",
          grupo: row[idxGrupo] || "-",
          produto: row[idxProduto] || "-",
          descricao: row[idxDesc] || "-",
          comissao: formatNum(row[idxComissao]),
          valor: formatNum(row[idxValor]),
          valorConsiderado: formatNum(row[idxValorCons]),
          pagoEm: formatData(row[idxPagoEm])
        });
      }
    }

    return { sucesso: true, dados: producaoLimpa };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}