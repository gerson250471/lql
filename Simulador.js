/**
 * FICHEIRO: Simulador.js
 * Tabelas de cálculos e comissionamentos
 */

function getDadosComissao(perfil) {
  const ss = getDatabaseConnection();
  const sheet = ss.getSheetByName("bdComissao");
  if (!sheet) throw new Error("Aba 'bdComissao' não encontrada.");

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());

  const idxGrupo = headers.indexOf("GRUPO");
  const idxDesc = headers.indexOf("DESCRIÇÃO");
  const idxTini = headers.indexOf("TAXA INI");
  const idxTfin = headers.indexOf("TAXA FIN");
  const idxPini = headers.indexOf("PRAZO INI");
  const idxPfin = headers.indexOf("PRAZO FIM");
  const idxPerfil = headers.indexOf(perfil.trim().toUpperCase());

  if (idxPerfil === -1) throw new Error("Perfil '" + perfil + "' não encontrado no bdComissao.");

  const grupos = {};

  data.slice(1).forEach(row => {
    const nomeGrupo = row[idxGrupo];
    if (!nomeGrupo) return;

    const tituloTabela = `${nomeGrupo} ${row[idxDesc] || ""}`.trim();

    if (!grupos[tituloTabela]) {
      grupos[tituloTabela] = { titulo: tituloTabela, itens: [] };
    }

    const formatPercent = (val) => {
      if (typeof val === 'number') {
        return (val * 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%";
      }
      return val || "0,00%";
    };

    grupos[tituloTabela].itens.push({
      taxaIni: formatPercent(row[idxTini]),
      taxaFin: formatPercent(row[idxTfin]),
      prazoIni: row[idxPini] || "-",
      prazoFin: row[idxPfin] || "-",
      comissao: formatPercent(row[idxPerfil])
    });
  });

  return Object.values(grupos);
}

function getTabelasPorPerfil(perfil) {
  const ss = getDatabaseConnection();
  const sheet = ss.getSheetByName("Simulador");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    texto: `${row[1]} | ${row[4]} | ${row[2]}x`,
    fator: row[3]
  }));
}

//**
 * Busca o resumo consolidado de todas as comissões com os perfis lado a lado
  * @return { Object } Contém os cabeçalhos dos perfis e as linhas agrupadas
    */
function getResumoComissoesAdmin() {
  try {
    const ss = getDatabaseConnection();
    const sheet = ss.getSheetByName("bdComissao");
    if (!sheet) throw new Error("Aba 'bdComissao' não encontrada.");

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toUpperCase());

    const idxGrupo = headers.indexOf("GRUPO");
    const idxDesc = headers.indexOf("DESCRIÇÃO");
    const idxTini = headers.indexOf("TAXA INI");
    const idxTfin = headers.indexOf("TAXA FIN");
    const idxPini = headers.indexOf("PRAZO INI");
    const idxPfin = headers.indexOf("PRAZO FIM");

    if (idxGrupo === -1 || idxPfin === -1) {
      throw new Error("Colunas obrigatórias não encontradas na aba bdComissao. Verifique os cabeçalhos.");
    }

    const limiteFixo = Math.max(idxGrupo, idxDesc, idxTini, idxTfin, idxPini, idxPfin);
    const perfisDetectados = [];
    const indicesPerfis = [];

    // 🛡️ A CORREÇÃO: Lista de colunas a ignorar (que não são perfis)
    const colunasIgnoradas = ["ATUAL", "ULT ATUALIZAÇÃO", "ULT ATUALIZACAO", "ÚLTIMA ATUALIZAÇÃO"];

    for (let j = 0; j < headers.length; j++) {
      if (j > limiteFixo && headers[j] !== "") {
        let nomeColuna = headers[j].toString().trim().toUpperCase();
        // Só adiciona se o cabeçalho não estiver na lista de ignorados
        if (!colunasIgnoradas.includes(nomeColuna)) {
          perfisDetectados.push(headers[j]);
          indicesPerfis.push(j);
        }
      }
    }

    const grupos = {};

    const formatPercent = (val) => {
      if (typeof val === 'number') {
        return (val * 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%";
      }
      if (val instanceof Date) return "0,00%";
      return val ? String(val).trim() : "0,00%";
    };

    const sanitizeString = (val) => {
      if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
      return (val !== undefined && val !== null && val !== "") ? String(val).trim() : "-";
    };

    data.slice(1).forEach(row => {
      const nomeGrupo = row[idxGrupo];
      if (!nomeGrupo || String(nomeGrupo).trim() === "") return;

      const tituloTabela = `${String(nomeGrupo).trim()} ${row[idxDesc] ? String(row[idxDesc]).trim() : ""}`.trim();

      if (!grupos[tituloTabela]) {
        grupos[tituloTabela] = { titulo: tituloTabela, itens: [] };
      }

      const comissoesPorPerfil = {};
      perfisDetectados.forEach((perfil, index) => {
        const colunaI = indicesPerfis[index];
        comissoesPorPerfil[perfil] = formatPercent(row[colunaI]);
      });

      grupos[tituloTabela].itens.push({
        taxaIni: formatPercent(row[idxTini]),
        taxaFin: formatPercent(row[idxTfin]),
        prazoIni: sanitizeString(row[idxPini]),
        prazoFin: sanitizeString(row[idxPfin]),
        comissoes: comissoesPorPerfil
      });
    });

    return {
      sucesso: true,
      perfis: perfisDetectados,
      dados: Object.values(grupos)
    };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}