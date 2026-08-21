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

    const idxId = getIdx("ID");
    const idxDataVenc = getIdx("DATA_VENCIMENTO"); 
    const idxTipo = getIdx("TIPO");
    const idxCategoria = getIdx("CATEGORIA");
    const idxDescricao = getIdx("DESCRICAO");
    const idxValorEnt = getIdx("VALOR_ENTRADA"); 
    const idxValorSai = getIdx("VALOR_SAIDA");   
    const idxForma = getIdx("FORMA_PAGTO");
    const idxPromotor = getIdx("PROMOTOR");
    const idxUsuario = getIdx("USUARIO_LANCAMENTO");
    const idxFonte = getIdx("FONTE"); 
    const idxObs = getIdx("OBSERVACAO");
    const idxDataPag = getIdx("DATA_PAGAMENTO");
    const idxDataLanc = getIdx("DATA_LANCAMENTO"); 
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

    // LÓGICA DE RECORRÊNCIA (Repetir X vezes)
    let parcelas = Number(dados.parcelas) || 1;
    if (isEdicao) parcelas = 1; // Se for edição, nunca repete

    let dataVencBase = new Date();
    if (dados.dataVencimento) {
      const partes = dados.dataVencimento.split("-"); 
      dataVencBase = new Date(partes[0], partes[1] - 1, partes[2]);
    }

    // LOOP PARA GERAR AS LINHAS
    for (let p = 0; p < parcelas; p++) {
      let novaLinha = new Array(headers.length).fill("");

      if (idxId !== -1) novaLinha[idxId] = isEdicao ? dados.id : "FIN-" + new Date().getTime() + "-" + p;
      
      // Incrementa os meses automaticamente na repetição
      let dataVencAtual = new Date(dataVencBase.getTime());
      dataVencAtual.setMonth(dataVencAtual.getMonth() + p);
      if (idxDataVenc !== -1) novaLinha[idxDataVenc] = dataVencAtual;

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

      if (idxDataLanc !== -1) {
          novaLinha[idxDataLanc] = (isEdicao && dataRows[rowIndex - 1][idxDataLanc]) ? dataRows[rowIndex - 1][idxDataLanc] : new Date();
      }

      const tipoUpper = String(dados.tipo).toUpperCase();
      const isReceita = tipoUpper.includes("RECEITA") || tipoUpper.includes("ENTRADA");
      let valorNum = Number(dados.valor) || 0;

      if (idxValorEnt !== -1) novaLinha[idxValorEnt] = isReceita ? valorNum : "";
      if (idxValorSai !== -1) novaLinha[idxValorSai] = !isReceita ? valorNum : "";

      if (idxTipo !== -1) novaLinha[idxTipo] = tipoUpper;
      if (idxCategoria !== -1) novaLinha[idxCategoria] = String(dados.categoria).toUpperCase();
      
      // Ajusta a descrição para mostrar a numeração das parcelas
      let descFinal = String(dados.descricao).toUpperCase();
      if (parcelas > 1) descFinal += ` (MÊS ${p + 1}/${parcelas})`;
      if (idxDescricao !== -1) novaLinha[idxDescricao] = descFinal;
      
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
    }

    return { 
      sucesso: true, 
      mensagem: isEdicao ? "Lançamento atualizado com sucesso!" : (parcelas > 1 ? `${parcelas} lançamentos gerados com sucesso!` : "Lançamento registrado!") 
    };

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
    if (data.length <= 1) return { sucesso: true, dados: [], resumo: { receitas: 0, despesas: 0, saldo: 0 }, resumoTipos: {} };

    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));
    const getIdx = (nome) => headers.indexOf(normalizarTexto(nome));

    const idxId = getIdx("ID");
    const idxDataVenc = getIdx("DATA_VENCIMENTO");
    const idxDataPag = getIdx("DATA_PAGAMENTO");
    const idxTipo = getIdx("TIPO");
    const idxCategoria = getIdx("CATEGORIA");
    const idxDescricao = getIdx("DESCRICAO");
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
    let resumoTipos = {}; // NOVO: Agrupamento por Tipo

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
      let status = String(row[idxStatus] || "PAGO").toUpperCase();

      let isMesFiltrado = (mesLinha === String(mesFiltro).padStart(2, '0') && anoLinha === String(anoFiltro));
      let dataInicioFiltro = new Date(anoFiltro, mesFiltro - 1, 1);
      let isPendenteAntigo = (dVenc < dataInicioFiltro && status !== "PAGAMENTO REALIZADO" && status !== "CANCELADO");

      if (isMesFiltrado || isPendenteAntigo) {
        let tipo = String(row[idxTipo] || "DESPESA").toUpperCase();
        let valorEnt = idxValorEnt !== -1 ? Number(row[idxValorEnt]) || 0 : 0;
        let valorSai = idxValorSai !== -1 ? Number(row[idxValorSai]) || 0 : 0;
        let valorAbsoluto = valorEnt > 0 ? valorEnt : valorSai;

        // CRIAÇÃO DO RESUMO POR TIPO
        if (!resumoTipos[tipo]) resumoTipos[tipo] = { total: 0, tipoCategoria: '' };

        if (status !== "CANCELADO") {
          if (tipo === "RECEITA" || tipo === "ENTRADA" || tipo.includes("ENTRADA")) {
            totReceitas += valorAbsoluto;
            resumoTipos[tipo].total += valorAbsoluto;
            resumoTipos[tipo].tipoCategoria = 'ENTRADA';
          } else if (tipo === "DESPESA" || tipo === "SAIDA" || tipo.includes("SAIDA") || tipo.includes("SAÍDA")) {
            totDespesas += valorAbsoluto;
            resumoTipos[tipo].total += valorAbsoluto;
            resumoTipos[tipo].tipoCategoria = 'SAIDA';
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
          valor: valorAbsoluto, 
          formaPagto: row[idxForma] || "-",
          status: status,
          fonte: idxFonte !== -1 ? String(row[idxFonte] || "-").toUpperCase() : "-",
          promotor: idxPromotor !== -1 ? (row[idxPromotor] || "") : "",
          obs: idxObs !== -1 ? (row[idxObs] || "") : "",
          alerta: alerta
        });
      }
    }

    lancamentos.sort(function(a, b) {
      const prioridade = { 'red': 1, 'yellow': 2, 'none': 3 };
      const pesoA = prioridade[a.alerta] || 3;
      const pesoB = prioridade[b.alerta] || 3;
      if (pesoA !== pesoB) return pesoA - pesoB;
      if (a.dataFormatoInput < b.dataFormatoInput) return -1;
      if (a.dataFormatoInput > b.dataFormatoInput) return 1;
      return 0; 
    });

    return {
      sucesso: true,
      dados: lancamentos,
      resumo: { receitas: totReceitas, despesas: totDespesas, saldo: totReceitas - totDespesas },
      resumoTipos: resumoTipos // Devolve agrupado por Tipo
    };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// NOVA FUNÇÃO: DELETAR LANÇAMENTO
function deletarLancamentoFinanceiro(id) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Financeiro");
    if (!sheet) throw new Error("Aba 'Financeiro' não encontrada.");

    const data = sheet.getDataRange().getValues();
    const idxId = data[0].findIndex(h => String(h).trim().toUpperCase() === "ID");
    
    if (idxId === -1) throw new Error("Coluna ID não encontrada.");

    for (let i = 1; i < data.length; i++) {
      if (data[i][idxId] === id) {
        sheet.deleteRow(i + 1); // +1 porque as planilhas começam no 1
        return { sucesso: true, mensagem: "Lançamento apagado com sucesso!" };
      }
    }
    return { sucesso: false, erro: "Lançamento não encontrado para exclusão." };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}