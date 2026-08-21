/**
 * FICHEIRO: Financeiro.js
 * Módulo de Controle Financeiro do Sistema
 */

function salvarLancamentoFinanceiro(dados) {
  try {
    const ss = getDatabaseConnection();
    let aba = ss.getSheetByName("Financeiro");

    if (!aba) throw new Error("Aba 'Financeiro' não encontrada no banco de dados.");

    const dataRows = aba.getDataRange().getValues();
    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = dataRows[0].map(h => normalizarTexto(h));

    const getIdx = (nome) => headers.indexOf(normalizarTexto(nome));

    // NOVOS CABEÇALHOS DA IMAGEM
    const idxId = getIdx("ID");
    const idxDataVenc = getIdx("DATA_VENCIMENTO"); 
    const idxTipo = getIdx("TIPO");
    const idxCategoria = getIdx("CATEGORIA");
    const idxDescricao = getIdx("DESCRICAO");
    const idxValorEnt = getIdx("VALOR_ENTRADA"); // NOVO
    const idxValorSai = getIdx("VALOR_SAIDA");   // NOVO
    const idxForma = getIdx("FORMA_PAGTO");
    const idxPromotor = getIdx("PROMOTOR");
    const idxUsuario = getIdx("USUARIO_LANCAMENTO");
    const idxFonte = getIdx("FONTE"); 
    const idxObs = getIdx("OBSERVACAO");
    const idxDataPag = getIdx("DATA_PAGAMENTO");
    const idxDataLanc = getIdx("DATA_LANCAMENTO"); // NOVO
    const idxStatus = getIdx("STATUS");

    let isEdicao = false;
    let rowIndex = -1;

    if (dados.id && idxId !== -1) {
      for (let i = 1; i < dataRows.length; i++) {
        if (dataRows[i][idxId] === dados.id) {
          isEdicao = true;
          rowIndex = i + 1;
          break;
        }
      }
    }

    let novaLinha = new Array(headers.length).fill("");

    if (idxId !== -1) novaLinha[idxId] = isEdicao ? dados.id : "FIN-" + new Date().getTime();
    
    // Processa Data de Vencimento
    if (idxDataVenc !== -1) {
      if (dados.dataVencimento) {
        const partes = dados.dataVencimento.split("-"); 
        novaLinha[idxDataVenc] = new Date(partes[0], partes[1] - 1, partes[2]);
      } else {
        novaLinha[idxDataVenc] = new Date();
      }
    }

    // Regras de Status e Data de Pagamento
    const statusFormatado = String(dados.status).toUpperCase();
    if (idxStatus !== -1) novaLinha[idxStatus] = statusFormatado;

    if (idxDataPag !== -1) {
      if (statusFormatado === "PAGAMENTO REALIZADO") {
        const dataPagAtual = isEdicao ? dataRows[rowIndex - 1][idxDataPag] : "";
        novaLinha[idxDataPag] = (dataPagAtual && dataPagAtual instanceof Date) ? dataPagAtual : new Date();
      } else {
        novaLinha[idxDataPag] = "";
      }
    }

    // Regra da Data de Lançamento Automática
    if (idxDataLanc !== -1) {
        novaLinha[idxDataLanc] = (isEdicao && dataRows[rowIndex - 1][idxDataLanc]) ? dataRows[rowIndex - 1][idxDataLanc] : new Date();
    }

    // INTELIGÊNCIA DE DIVISÃO DE VALORES (Entrada vs Saída)
    const tipoUpper = String(dados.tipo).toUpperCase();
    const isReceita = tipoUpper.includes("RECEITA") || tipoUpper.includes("ENTRADA");
    let valorNum = Number(dados.valor) || 0;

    if (idxValorEnt !== -1) novaLinha[idxValorEnt] = isReceita ? valorNum : "";
    if (idxValorSai !== -1) novaLinha[idxValorSai] = !isReceita ? valorNum : "";

    // Preenchimento dos restantes campos
    if (idxTipo !== -1) novaLinha[idxTipo] = tipoUpper;
    if (idxCategoria !== -1) novaLinha[idxCategoria] = String(dados.categoria).toUpperCase();
    if (idxDescricao !== -1) novaLinha[idxDescricao] = String(dados.descricao).toUpperCase();
    if (idxForma !== -1) novaLinha[idxForma] = String(dados.formaPagto).toUpperCase();
    if (idxPromotor !== -1) novaLinha[idxPromotor] = String(dados.promotor || "").toUpperCase();
    if (idxUsuario !== -1) novaLinha[idxUsuario] = String(dados.usuarioLancamento || "").toUpperCase();
    if (idxFonte !== -1) novaLinha[idxFonte] = String(dados.fonte || "NÃO INFORMADA").toUpperCase(); 
    if (idxObs !== -1) novaLinha[idxObs] = String(dados.obs || "").toUpperCase(); 

    if (isEdicao) {
      aba.getRange(rowIndex, 1, 1, headers.length).setValues([novaLinha]);
    } else {
      aba.appendRow(novaLinha);
    }

    return { sucesso: true, mensagem: isEdicao ? "Lançamento atualizado com sucesso!" : "Lançamento registrado!" };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

function getLancamentosFinanceiros(mesFiltro, anoFiltro) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Financeiro");
    if (!sheet) return { sucesso: false, erro: "Aba Financeiro não encontrada." };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: true, dados: [], resumo: { receitas: 0, despesas: 0, saldo: 0 }, resumoFontes: {} };

    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));
    const getIdx = (nome) => headers.indexOf(normalizarTexto(nome));

    const idxId = getIdx("ID");
    const idxDataVenc = getIdx("DATA_VENCIMENTO");
    const idxDataPag = getIdx("DATA_PAGAMENTO");
    const idxTipo = getIdx("TIPO");
    const idxCategoria = getIdx("CATEGORIA");
    const idxDescricao = getIdx("DESCRICAO");
    // LER DAS NOVAS COLUNAS
    const idxValorEnt = getIdx("VALOR_ENTRADA");
    const idxValorSai = getIdx("VALOR_SAIDA");
    const idxForma = getIdx("FORMA_PAGTO");
    const idxStatus = getIdx("STATUS");
    const idxFonte = getIdx("FONTE"); 
    const idxPromotor = getIdx("PROMOTOR");
    const idxObs = getIdx("OBSERVACAO");

    const lancamentos = [];
    let totReceitas = 0;
    let totDespesas = 0;
    let resumoFontes = {}; 

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let rawData = row[idxDataVenc];
      if (!rawData) continue;

      let dVenc = new Date(rawData);
      if (isNaN(dVenc.getTime())) continue;

      let mesLinha = String(dVenc.getMonth() + 1).padStart(2, '0');
      let anoLinha = String(dVenc.getFullYear());

      if (mesLinha === String(mesFiltro).padStart(2, '0') && anoLinha === String(anoFiltro)) {
        let tipo = String(row[idxTipo] || "DESPESA").toUpperCase();
        
        // RECUPERA O VALOR CONSOLIDADO PARA A TELA
        let valorEnt = idxValorEnt !== -1 ? Number(row[idxValorEnt]) || 0 : 0;
        let valorSai = idxValorSai !== -1 ? Number(row[idxValorSai]) || 0 : 0;
        let valorAbsoluto = valorEnt > 0 ? valorEnt : valorSai;

        let status = String(row[idxStatus] || "PAGO").toUpperCase();
        let fonte = idxFonte !== -1 ? String(row[idxFonte] || "NÃO INFORMADA").toUpperCase() : "NÃO INFORMADA";

        if (!resumoFontes[fonte]) resumoFontes[fonte] = { saldo: 0 };

        if (status !== "CANCELADO") {
          if (tipo === "RECEITA" || tipo === "ENTRADA" || tipo.includes("ENTRADA")) {
            totReceitas += valorAbsoluto;
            resumoFontes[fonte].saldo += valorAbsoluto;
          } else if (tipo === "DESPESA" || tipo === "SAIDA" || tipo.includes("SAIDA") || tipo.includes("SAÍDA")) {
            totDespesas += valorAbsoluto;
            resumoFontes[fonte].saldo -= valorAbsoluto;
          }
        }

        let alerta = "none";
        if (status !== "PAGAMENTO REALIZADO" && status !== "CANCELADO") {
          dVenc.setHours(0, 0, 0, 0);
          let diffTime = dVenc.getTime() - hoje.getTime();
          let diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

          if (diffDays < 0) alerta = "red";
          else if (diffDays <= 5) alerta = "yellow";
        }

        let strDataPag = "";
        if (idxDataPag !== -1 && row[idxDataPag] && row[idxDataPag] instanceof Date) {
          strDataPag = Utilities.formatDate(row[idxDataPag], Session.getScriptTimeZone(), "dd/MM/yyyy");
        }

        lancamentos.push({
          id: row[idxId],
          dataFormatoInput: Utilities.formatDate(dVenc, Session.getScriptTimeZone(), "yyyy-MM-dd"), 
          dataVencimento: Utilities.formatDate(dVenc, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          dataPagamento: strDataPag,
          tipo: tipo,
          categoria: row[idxCategoria] || "-",
          descricao: row[idxDescricao] || "-",
          valor: valorAbsoluto, // Devolve o valor único para não quebrar a tela visual
          formaPagto: row[idxForma] || "-",
          status: status,
          fonte: fonte,
          promotor: idxPromotor !== -1 ? (row[idxPromotor] || "") : "",
          obs: idxObs !== -1 ? (row[idxObs] || "") : "",
          alerta: alerta
        });
      }
    }

    // NOVO: ORDENAÇÃO INTELIGENTE (1º Vermelho, 2º Amarelo, 3º Data)
    lancamentos.sort(function(a, b) {
      // Define os "pesos" das cores (quanto menor, mais no topo)
      const prioridade = { 'red': 1, 'yellow': 2, 'none': 3 };
      const pesoA = prioridade[a.alerta] || 3;
      const pesoB = prioridade[b.alerta] || 3;

      // 1º Nível de ordenação: Pela Cor
      if (pesoA !== pesoB) {
        return pesoA - pesoB;
      }

      // 2º Nível de ordenação: Se tiverem a mesma cor, ordena pela data 
      // (Data mais antiga primeiro)
      if (a.dataFormatoInput < b.dataFormatoInput) return -1;
      if (a.dataFormatoInput > b.dataFormatoInput) return 1;

      return 0; // Mantém a ordem caso sejam exatamente do mesmo dia
    });

    return {
      sucesso: true,
      dados: lancamentos, // <-- Retiramos o .reverse() porque já ordenamos acima
      resumo: { receitas: totReceitas, despesas: totDespesas, saldo: totReceitas - totDespesas },
      resumoFontes: resumoFontes
    };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}