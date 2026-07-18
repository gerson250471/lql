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

/**
 * Busca e consolida a produção de todos os promotores (Exclusivo para ADMIN)
 */
function getResumoProducaoAdmin(mesFiltro, anoFiltro) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("bd_Producao");
    if (!sheet) throw new Error("Aba 'bd_Producao' não encontrada.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: true, dados: [], totaisGlobais: {} };

    const headers = data[0].map(h => h.toString().trim().toUpperCase());
    
    // Mapeamento das colunas
    const idxChave = headers.indexOf("CHAVE J");
    const idxPromotor = headers.indexOf("PROMOTOR");
    const idxAno = headers.indexOf("ANO");
    const idxMes = headers.indexOf("MÊS");
    const idxComissao = headers.indexOf("COMISSÃO");
    const idxValorBruto = headers.indexOf("VALOR BRUTO");
    const idxValorCons = headers.indexOf("VALOR CONSIDERADO");
    const idxValorLiquido = headers.indexOf("VALOR LIQUIDO"); // Lucro da empresa

    if (idxChave === -1 || idxPromotor === -1 || idxMes === -1) {
      throw new Error("Colunas obrigatórias não encontradas na base de produção.");
    }

    const resumoPorPromotor = {};
    let globalVolume = 0;
    let globalBruto = 0;
    let globalComissao = 0;
    let globalLiquido = 0;
    let globalContratos = 0;

    const formatNum = (val) => typeof val === 'number' ? val : 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const anoLinha = String(row[idxAno]).trim();
      const mesLinha = String(row[idxMes]).trim().padStart(2, '0');
      const mesBusca = String(mesFiltro).padStart(2, '0');

      // Filtra pelo Mês e Ano
      if (anoLinha === String(anoFiltro) && mesLinha === mesBusca) {
        const chave = row[idxChave] || "SEM_CHAVE";
        const nomePromotor = row[idxPromotor] || "NÃO IDENTIFICADO";
        
        const vCons = formatNum(row[idxValorCons]);
        const vBruto = formatNum(row[idxValorBruto]);
        const vCom = formatNum(row[idxComissao]);
        const vLiq = formatNum(row[idxValorLiquido]);

        // Se o promotor ainda não existe no nosso objeto de resumo, criamos
        if (!resumoPorPromotor[chave]) {
          resumoPorPromotor[chave] = {
            chave: chave,
            nome: nomePromotor,
            qtdContratos: 0,
            volumeConsiderado: 0,
            valorBruto: 0,
            comissaoPaga: 0,
            valorLiquido: 0
          };
        }

        // Soma os valores para este promotor específico
        resumoPorPromotor[chave].qtdContratos += 1;
        resumoPorPromotor[chave].volumeConsiderado += vCons;
        resumoPorPromotor[chave].valorBruto += vBruto;
        resumoPorPromotor[chave].comissaoPaga += vCom;
        resumoPorPromotor[chave].valorLiquido += vLiq;

        // Soma os valores para a Empresa (Totais Globais)
        globalContratos += 1;
        globalVolume += vCons;
        globalBruto += vBruto;
        globalComissao += vCom;
        globalLiquido += vLiq;
      }
    }

    // Transforma o objeto em um Array e ordena do que vendeu mais para o que vendeu menos
    const arrayResumo = Object.values(resumoPorPromotor).sort((a, b) => b.volumeConsiderado - a.volumeConsiderado);

    const totaisGlobais = {
      contratos: globalContratos,
      volume: globalVolume,
      bruto: globalBruto,
      comissao: globalComissao,
      liquido: globalLiquido
    };

    return { sucesso: true, dados: arrayResumo, totaisGlobais: totaisGlobais };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}