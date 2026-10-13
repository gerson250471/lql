/**
 * Busca a produção oficial filtrada por Promotor, Mês e Ano.
 */
function getProducaoPromotor(chavePromotor, mesFiltro, anoFiltro) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("bd_Producao");
    if (!sheet) throw new Error("Aba 'bd_Producao' não encontrada.");

    // 1. PRIMEIRO: Lê os dados da folha para a variável 'data' existir
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: true, dados: [] };

    // 2. SEGUNDO: Funções utilitárias que dependem ou não de 'data'
    const formatData = (val) => {
      if (!val) return "-";
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      let strVal = String(val).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(strVal)) {
        let partes = strVal.split("T")[0].split("-");
        if (partes.length === 3) {
          return `${partes[2]}/${partes[1]}/${partes[0]}`;
        }
      }
      return strVal;
    };

    const formatNum = (val) => typeof val === 'number' ? val : 0;

    // 3. TERCEIRO: Mapeamento de cabeçalhos utilizando a variável 'data' já carregada
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
    
    let idxParcela = headers.indexOf("PARCELAS");
    if (idxParcela === -1) idxParcela = headers.indexOf("PARCELA");
    if (idxParcela === -1) idxParcela = headers.indexOf("PRAZO");

    const idxRestricao = headers.indexOf("RESTRICAO_RCC");
    const idxGrupo = headers.indexOf("GRUPO");
    const idxProduto = headers.indexOf("PRODUTO");
    const idxDesc = headers.indexOf("DESCRIÇÃO DO PRODUTO");
    const idxPagoEm = headers.indexOf("PAGO EM");
    
    let idxObs = headers.indexOf("OBSERVAÇÃO");
    if (idxObs === -1) idxObs = headers.indexOf("OBSERVACAO");

    let idxProducao = headers.indexOf("PRODUÇÃO");
    if (idxProducao === -1) idxProducao = headers.indexOf("PRODUCAO");
    if (idxProducao === -1) idxProducao = headers.indexOf("VALOR CONSIDERADO");
    
    const idxValorBruto = headers.indexOf("VALOR BRUTO");
    
    let idxValorComissao = headers.indexOf("VALOR COMISSÃO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR COMISSAO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR");

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
          prazo: idxParcela !== -1 ? (Number(row[idxParcela]) || 0) : 0,
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

    // Normaliza os cabeçalhos (remove espaços extras e coloca em maiúsculas)
    const headers = data[0].map(h => h.toString().trim().toUpperCase());
    
    const idxChave = headers.indexOf("CHAVE J");
    const idxPromotor = headers.indexOf("PROMOTOR");
    const idxAno = headers.indexOf("ANO");
    let idxMes = headers.indexOf("MÊS");
    if (idxMes === -1) idxMes = headers.indexOf("MES");

    let idxProducao = headers.indexOf("PRODUÇÃO");
    if (idxProducao === -1) idxProducao = headers.indexOf("PRODUCAO");
    if (idxProducao === -1) idxProducao = headers.indexOf("VALOR CONSIDERADO");

    const idxValorBruto = headers.indexOf("VALOR BRUTO");
    
    let idxValorComissao = headers.indexOf("VALOR COMISSÃO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR COMISSAO");
    if (idxValorComissao === -1) idxValorComissao = headers.indexOf("VALOR");

    let idxValorLiquido = headers.indexOf("VALOR LIQUIDO");
    if (idxValorLiquido === -1) idxValorLiquido = headers.indexOf("VALOR LÍQUIDO");

    // NOVOS ÍNDICES: Para identificar se o contrato é Consórcio ou Consignado
    const idxProduto = headers.indexOf("PRODUTO");
    const idxGrupo = headers.indexOf("GRUPO");
    let idxConvenio = headers.indexOf("CONVÊNIO");
    if (idxConvenio === -1) idxConvenio = headers.indexOf("CONVENIO");
    let idxDescricao = headers.indexOf("DESCRIÇÃO");
    if (idxDescricao === -1) idxDescricao = headers.indexOf("DESCRICAO");

    if (idxChave === -1 || idxPromotor === -1 || idxMes === -1) {
      throw new Error("Colunas obrigatórias não encontradas na base de produção.");
    }

    const resumoPorPromotor = {};
    let globalVolume = 0;
    let globalVolumeConsig = 0;
    let globalVolumeConsorcio = 0;
    let globalBruto = 0;
    let globalComissao = 0;
    let globalLiquido = 0;
    let globalContratos = 0;

    // Garante leitura correta mesmo se algum número vier como texto formatado
    const formatNum = (val) => {
      if (typeof val === 'number') return val;
      if (!val) return 0;
      const limpo = String(val).replace(/[R$\s.]/g, '').replace(',', '.');
      return parseFloat(limpo) || 0;
    };

    const normalizarTexto = (txt) => String(txt || "")
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");

    const mesBusca = String(mesFiltro).padStart(2, '0');
    const anoBusca = String(anoFiltro).trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const anoLinha = String(row[idxAno]).trim();
      const mesLinha = String(row[idxMes]).trim().padStart(2, '0');

      if (anoLinha === anoBusca && mesLinha === mesBusca) {
        const chave = String(row[idxChave] || "SEM_CHAVE").trim();
        const nomePromotor = String(row[idxPromotor] || "NÃO IDENTIFICADO").trim();
        
        const vCons = formatNum(row[idxProducao]);
        const vBruto = idxValorBruto !== -1 ? formatNum(row[idxValorBruto]) : 0;
        const vCom = idxValorComissao !== -1 ? formatNum(row[idxValorComissao]) : 0;
        const vLiq = idxValorLiquido !== -1 ? formatNum(row[idxValorLiquido]) : 0;

        // Verifica se é CONSÓRCIO olhando Produto, Grupo, Convênio ou Descrição
        const txtClassificacao = normalizarTexto(
          (idxProduto !== -1 ? row[idxProduto] : "") + " " +
          (idxGrupo !== -1 ? row[idxGrupo] : "") + " " +
          (idxConvenio !== -1 ? row[idxConvenio] : "") + " " +
          (idxDescricao !== -1 ? row[idxDescricao] : "")
        );
        const ehConsorcio = txtClassificacao.includes("CONSORCIO");

        if (!resumoPorPromotor[chave]) {
          resumoPorPromotor[chave] = {
            chave: chave,
            nome: nomePromotor,
            qtdContratos: 0,
            volumeConsiderado: 0,
            valorBruto: 0,
            comissaoPaga: 0,
            valorLiquido: 0,
            // NOVAS GAVETAS SEPARADAS PARA O FRONT-END:
            consignado: { qtd: 0, volume: 0, bruto: 0, comissao: 0, liquido: 0 },
            consorcio: { qtd: 0, volume: 0, bruto: 0, comissao: 0, liquido: 0 }
          };
        }

        // 1. Soma no total geral do promotor
        resumoPorPromotor[chave].qtdContratos += 1;
        resumoPorPromotor[chave].volumeConsiderado += vCons;
        resumoPorPromotor[chave].valorBruto += vBruto;
        resumoPorPromotor[chave].comissaoPaga += vCom;
        resumoPorPromotor[chave].valorLiquido += vLiq;

        // 2. Soma na gaveta específica (Consórcio vs Consignado)
        const gaveta = ehConsorcio 
          ? resumoPorPromotor[chave].consorcio 
          : resumoPorPromotor[chave].consignado;

        gaveta.qtd += 1;
        gaveta.volume += vCons;
        gaveta.bruto += vBruto;
        gaveta.comissao += vCom;
        gaveta.liquido += vLiq;

        // 3. Soma nos totais globais da empresa
        globalContratos += 1;
        globalVolume += vCons;
        if (ehConsorcio) {
          globalVolumeConsorcio += vCons;
        } else {
          globalVolumeConsig += vCons;
        }
        globalBruto += vBruto;
        globalComissao += vCom;
        globalLiquido += vLiq;
      }
    }

    const arrayResumo = Object.values(resumoPorPromotor).sort((a, b) => b.volumeConsiderado - a.volumeConsiderado);

    const totaisGlobais = {
      contratos: globalContratos,
      volume: globalVolume,
      volumeConsignado: globalVolumeConsig,
      volumeConsorcio: globalVolumeConsorcio,
      bruto: globalBruto,
      comissao: globalComissao,
      liquido: globalLiquido
    };

    return { sucesso: true, dados: arrayResumo, totaisGlobais: totaisGlobais };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// NOVA FUNÇÃO: CALCULA O ESFORÇO DA EQUIPE NO CRM
function getMetricasCrmAdmin(mes, ano) {
  try {
    const ss = getDatabaseConnection();
    const abaLeads = ss.getSheetByName("Leads");
    const abaHistorico = ss.getSheetByName("HistoricoLeads");
    const abaPromotores = ss.getSheetByName("Promotores");

    let stats = {};

    // 1. Inicia os promotores
    if (abaPromotores) {
      const dadosP = abaPromotores.getDataRange().getValues();
      const idxChave = dadosP[0].findIndex(h => String(h).toUpperCase().includes("CHAVE"));
      const idxNome = dadosP[0].findIndex(h => String(h).toUpperCase() === "NOME");
      for (let i = 1; i < dadosP.length; i++) {
        let c = String(dadosP[i][idxChave]).toUpperCase().trim();
        if (c) stats[c] = { nome: dadosP[i][idxNome], leadsTotal: 0, leadsTrab: 0, interacoes: 0 };
      }
    }

    // 2. Conta a Carteira de Leads (Trabalhados vs Total)
    if (abaLeads) {
      const dadosL = abaLeads.getDataRange().getValues();
      const idxProm = dadosL[0].findIndex(h => String(h).toUpperCase() === "PROMOTOR");
      const idxStatus = dadosL[0].findIndex(h => String(h).toUpperCase() === "STATUS");
      if (idxProm !== -1) {
        for (let i = 1; i < dadosL.length; i++) {
          let p = String(dadosL[i][idxProm]).toUpperCase().trim();
          let st = String(dadosL[i][idxStatus]).toUpperCase().trim();
          if (p && stats[p]) {
            stats[p].leadsTotal++;
            if (st !== "NOVO") stats[p].leadsTrab++;
          }
        }
      }
    }

    // 3. Conta as Interações do Mês Específico
    if (abaHistorico) {
      const dadosH = abaHistorico.getDataRange().getValues();
      const idxPromH = dadosH[0].findIndex(h => String(h).toUpperCase() === "PROMOTOR");
      const idxDataH = dadosH[0].findIndex(h => String(h).toUpperCase() === "DATA");
      if (idxPromH !== -1 && idxDataH !== -1) {
        for (let i = 1; i < dadosH.length; i++) {
          let p = String(dadosH[i][idxPromH]).toUpperCase().trim();
          let rawData = dadosH[i][idxDataH];
          if (!p || !stats[p] || !rawData) continue;

          let dataVal = null;
          if (rawData instanceof Date) dataVal = rawData;
          else {
            let partes = String(rawData).split(" ")[0].split("/");
            if (partes.length === 3) dataVal = new Date(partes[2], partes[1] - 1, partes[0]);
          }

          if (dataVal) {
            let m = String(dataVal.getMonth() + 1).padStart(2, '0');
            let y = String(dataVal.getFullYear());
            if (m === String(mes).padStart(2, '0') && y === String(ano)) {
              stats[p].interacoes++;
            }
          }
        }
      }
    }

    // 4. Formata e faz a matemática final
    let arr = Object.keys(stats).map(k => {
      let s = stats[k];
      let penetracao = s.leadsTotal > 0 ? Math.round((s.leadsTrab / s.leadsTotal) * 100) : 0;
      let esforco = s.leadsTotal > 0 ? (s.interacoes / s.leadsTotal).toFixed(1) : 0;
      return {
        chave: k, nome: s.nome || k, leadsTotal: s.leadsTotal, leadsTrab: s.leadsTrab,
        interacoes: s.interacoes, penetracao: penetracao, esforco: esforco
      };
    }).filter(i => i.leadsTotal > 0 || i.interacoes > 0); // Oculta quem não tem leads nem trabalhou

    arr.sort((a, b) => b.interacoes - a.interacoes); // Ordena pelos mais esforçados

    return { sucesso: true, dados: arr };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}