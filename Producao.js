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
    const idxMes = headers.indexOf("MÊS");

    if (idxChave === -1 || idxAno === -1 || idxMes === -1) {
      throw new Error("As colunas 'CHAVE J', 'ANO' ou 'MÊS' não foram encontradas.");
    }

    const idxConvenio = headers.indexOf("CONVENIO");
    const idxContrato = headers.indexOf("CONTRATO");
    const idxDataCont = headers.indexOf("DATA CONTRATO");
    const idxTaxa = headers.indexOf("TAXA");
    
    // CORREÇÃO: Procura por "PARCELAS", "PARCELA" ou "PRAZO"
    let idxParcela = headers.indexOf("PARCELAS");
    if (idxParcela === -1) idxParcela = headers.indexOf("PARCELA");
    if (idxParcela === -1) idxParcela = headers.indexOf("PRAZO");

    const idxRestricao = headers.indexOf("RESTRICAO_RCC");
    const idxGrupo = headers.indexOf("GRUPO");
    const idxProduto = headers.indexOf("PRODUTO");
    const idxDesc = headers.indexOf("DESCRIÇÃO DO PRODUTO");
    const idxPagoEm = headers.indexOf("PAGO EM");
    
    // MAPEAMENTO DA OBSERVAÇÃO
    let idxObs = headers.indexOf("OBSERVAÇÃO");
    if (idxObs === -1) idxObs = headers.indexOf("OBSERVACAO");

    // MAPEAMENTO FINANCEIRO BLINDADO
    let idxProducao = headers.indexOf("PRODUÇÃO");
    if (idxProducao === -1) idxProducao = headers.indexOf("PRODUCAO");
    if (idxProducao === -1) idxProducao = headers.indexOf("VALOR CONSIDERADO");
    
    const idxValorBruto = headers.indexOf("VALOR BRUTO");
    
    let idxValorComissao = headers.indexOf("VALOR COMISSÃO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR COMISSAO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR");

    const formatData = (val) => val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy") : (val ? val.toString() : "-");
    const formatNum = (val) => typeof val === 'number' ? val : 0;

    const producaoLimpa = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const anoLinha = String(row[idxAno]).trim();
      const mesLinha = String(row[idxMes]).trim().padStart(2, '0');
      const mesBusca = String(mesFiltro).padStart(2, '0');

      if (row[idxChave] === chavePromotor && anoLinha === String(anoFiltro) && mesLinha === mesBusca) {
        producaoLimpa.push({
          convenio: row[idxConvenio] || "-",
          contrato: row[idxContrato] || "-",
          dataContrato: formatData(row[idxDataCont]),
          taxa: row[idxTaxa] || "-",
          prazo: idxParcela !== -1 ? (Number(row[idxParcela]) || 0) : 0, // Extrai o número do prazo limpo
          restricao: row[idxRestricao] || "-",
          grupo: row[idxGrupo] || "-",
          produto: row[idxProduto] || "-",
          descricao: row[idxDesc] || "-",
          producao: formatNum(row[idxProducao]),
          valorBruto: formatNum(row[idxValorBruto]),
          comissao: formatNum(row[idxValorComissao]),
          observacao: (idxObs !== -1 ? row[idxObs] : "") || "-",
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
    
    const idxChave = headers.indexOf("CHAVE J");
    const idxPromotor = headers.indexOf("PROMOTOR");
    const idxAno = headers.indexOf("ANO");
    const idxMes = headers.indexOf("MÊS");

    // MAPEAMENTO FINANCEIRO BLINDADO ADMIN
    let idxProducao = headers.indexOf("PRODUÇÃO");
    if (idxProducao === -1) idxProducao = headers.indexOf("PRODUCAO");
    if (idxProducao === -1) idxProducao = headers.indexOf("VALOR CONSIDERADO");

    const idxValorBruto = headers.indexOf("VALOR BRUTO");
    
    // CORREÇÃO AQUI
    let idxValorComissao = headers.indexOf("VALOR COMISSÃO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR COMISSAO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR");

    const idxValorLiquido = headers.indexOf("VALOR LIQUIDO");

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

      if (anoLinha === String(anoFiltro) && mesLinha === mesBusca) {
        const chave = row[idxChave] || "SEM_CHAVE";
        const nomePromotor = row[idxPromotor] || "NÃO IDENTIFICADO";
        
        const vCons = formatNum(row[idxProducao]);
        const vBruto = formatNum(row[idxValorBruto]);
        const vCom = formatNum(row[idxValorComissao]);
        const vLiq = formatNum(row[idxValorLiquido]);

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

        resumoPorPromotor[chave].qtdContratos += 1;
        resumoPorPromotor[chave].volumeConsiderado += vCons;
        resumoPorPromotor[chave].valorBruto += vBruto;
        resumoPorPromotor[chave].comissaoPaga += vCom;
        resumoPorPromotor[chave].valorLiquido += vLiq;

        globalContratos += 1;
        globalVolume += vCons;
        globalBruto += vBruto;
        globalComissao += vCom;
        globalLiquido += vLiq;
      }
    }

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