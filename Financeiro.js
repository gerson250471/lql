/**
 * FICHEIRO: Financeiro.js
 * Módulo de Controle Financeiro do Sistema
 */

function salvarLancamentoFinanceiro(dados) {
  try {
    const ss = getDatabaseConnection();
    let aba = ss.getSheetByName("Financeiro");

    if (!aba) {
      throw new Error("Aba 'Financeiro' não encontrada no banco de dados.");
    }

    const dataRows = aba.getDataRange().getValues();
    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = dataRows[0].map(h => normalizarTexto(h));

    const getIdx = (nome) => headers.indexOf(normalizarTexto(nome));

    const idxId = getIdx("ID");
    const idxData = getIdx("DATA");
    const idxTipo = getIdx("TIPO");
    const idxCategoria = getIdx("CATEGORIA");
    const idxDescricao = getIdx("DESCRICAO");
    const idxValor = getIdx("VALOR");
    const idxForma = getIdx("FORMA_PAGTO");
    const idxStatus = getIdx("STATUS");
    const idxPromotor = getIdx("PROMOTOR");
    const idxUsuario = getIdx("USUARIO_LANCAMENTO");
    const idxFonte = getIdx("FONTE"); 
    const idxObs = getIdx("OBSERVACAO"); // <-- NOVO CAMPO

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
    if (idxData !== -1) {
      if (dados.data) {
        const partes = dados.data.split("-"); 
        novaLinha[idxData] = new Date(partes[0], partes[1] - 1, partes[2]);
      } else {
        novaLinha[idxData] = new Date();
      }
    }
    if (idxTipo !== -1) novaLinha[idxTipo] = String(dados.tipo).toUpperCase();
    if (idxCategoria !== -1) novaLinha[idxCategoria] = String(dados.categoria).toUpperCase();
    if (idxDescricao !== -1) novaLinha[idxDescricao] = String(dados.descricao).toUpperCase();
    if (idxValor !== -1) novaLinha[idxValor] = Number(dados.valor) || 0;
    if (idxForma !== -1) novaLinha[idxForma] = String(dados.formaPagto).toUpperCase();
    if (idxStatus !== -1) novaLinha[idxStatus] = String(dados.status).toUpperCase();
    if (idxPromotor !== -1) novaLinha[idxPromotor] = String(dados.promotor || "").toUpperCase();
    if (idxUsuario !== -1) novaLinha[idxUsuario] = String(dados.usuarioLancamento || "").toUpperCase();
    if (idxFonte !== -1) novaLinha[idxFonte] = String(dados.fonte || "NÃO INFORMADA").toUpperCase(); 
    if (idxObs !== -1) novaLinha[idxObs] = String(dados.obs || "").toUpperCase(); // <-- NOVO CAMPO

    if (isEdicao) {
      aba.getRange(rowIndex, 1, 1, headers.length).setValues([novaLinha]);
    } else {
      aba.appendRow(novaLinha);
    }

    saveSystemLog({
      userKey: dados.usuarioLancamento,
      userProfile: "ADMIN",
      operationType: isEdicao ? "Atualização Financeira" : "Lançamento Financeiro",
      operationResult: "R$ " + dados.valor + " - " + dados.descricao
    });

    return { sucesso: true, mensagem: isEdicao ? "Lançamento atualizado com sucesso!" : "Lançamento registrado com sucesso!" };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

function getLancamentosFinanceiros(mesFiltro, anoFiltro) {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("Financeiro");
    if (!sheet) return { sucesso: true, dados: [], resumo: { receitas: 0, despesas: 0, saldo: 0 } };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sucesso: true, dados: [], resumo: { receitas: 0, despesas: 0, saldo: 0 } };

    const normalizarTexto = (txt) => String(txt || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const headers = data[0].map(h => normalizarTexto(h));

    const getIdx = (nome) => headers.indexOf(normalizarTexto(nome));

    const idxId = getIdx("ID");
    const idxData = getIdx("DATA");
    const idxTipo = getIdx("TIPO");
    const idxCategoria = getIdx("CATEGORIA");
    const idxDescricao = getIdx("DESCRICAO");
    const idxValor = getIdx("VALOR");
    const idxForma = getIdx("FORMA_PAGTO");
    const idxStatus = getIdx("STATUS");
    const idxFonte = getIdx("FONTE"); 
    const idxPromotor = getIdx("PROMOTOR");
    const idxObs = getIdx("OBSERVACAO"); // <-- NOVO CAMPO

    const lancamentos = [];
    let totReceitas = 0;
    let totDespesas = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let rawData = row[idxData];
      if (!rawData) continue;

      let d = new Date(rawData);
      if (isNaN(d.getTime())) continue;

      let mesLinha = String(d.getMonth() + 1).padStart(2, '0');
      let anoLinha = String(d.getFullYear());

      if (mesLinha === String(mesFiltro).padStart(2, '0') && anoLinha === String(anoFiltro)) {
        let tipo = String(row[idxTipo] || "DESPESA").toUpperCase();
        let valor = Number(row[idxValor]) || 0;
        let status = String(row[idxStatus] || "PAGO").toUpperCase();

        if (status !== "CANCELADO") {
          if (tipo === "RECEITA" || tipo === "ENTRADA") totReceitas += valor;
          else if (tipo === "DESPESA" || tipo === "SAIDA" || tipo === "SAÍDA") totDespesas += valor;
        }

        lancamentos.push({
          id: row[idxId],
          dataFormatoInput: Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"), 
          data: Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          tipo: tipo,
          categoria: row[idxCategoria] || "-",
          descricao: row[idxDescricao] || "-",
          valor: valor,
          formaPagto: row[idxForma] || "-",
          status: status,
          fonte: idxFonte !== -1 ? (row[idxFonte] || "-") : "-",
          promotor: idxPromotor !== -1 ? (row[idxPromotor] || "") : "",
          obs: idxObs !== -1 ? (row[idxObs] || "") : "" // <-- NOVO CAMPO
        });
      }
    }

    return {
      sucesso: true,
      dados: lancamentos.reverse(),
      resumo: {
        receitas: totReceitas,
        despesas: totDespesas,
        saldo: totReceitas - totDespesas
      }
    };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}