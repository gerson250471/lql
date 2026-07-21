/**
 * Importação otimizada com mapeamento dinâmico de cabeçalhos,
 * correção exata de datas e cruzamento preciso de taxas e prazos.
 */
function importarProducaoDoDrive() {
  try {
    const idPastaProducao = "1ivHA6X9Ku-1qyFTFkwbL-qNCz4PWSHji";
    const idPastaUsados = "1QTq9mXv_UyBqqgTmhPiSIJunR29WfEIp";
    
    const pastaProducao = DriveApp.getFolderById(idPastaProducao);
    const pastaUsados = DriveApp.getFolderById(idPastaUsados);
    
    const arquivos = pastaProducao.getFiles();
    
    if (!arquivos.hasNext()) {
      Logger.log("⚠️ A pasta 'Producao' está vazia.");
      return { sucesso: false, mensagem: "Nenhum arquivo encontrado na pasta de Produção." };
    }

    let arquivosProcessados = 0;
    const ss = getDatabaseConnection();
    
    const dadosPromotores = ss.getSheetByName("Promotores").getDataRange().getValues();
    const dadosComissao = ss.getSheetByName("bdComissao").getDataRange().getValues();
    const dadosProdutos = ss.getSheetByName("Produto").getDataRange().getValues();
    
    const sheetProd = ss.getSheetByName("bd_Producao");
    const dadosProducaoAtual = sheetProd.getLastRow() > 1 ? sheetProd.getRange(2, 1, sheetProd.getLastRow() - 1, sheetProd.getLastColumn()).getValues() : [];
    const contratosExistentes = new Set(dadosProducaoAtual.map(row => String(row[4]).trim()));

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nomeArquivo = arquivo.getName();
      Logger.log(`📁 Analisando: ${nomeArquivo}`);

      const resource = {
        title: "[TEMP_IMPORT] " + nomeArquivo,
        mimeType: MimeType.GOOGLE_SHEETS
      };
      const arquivoTemp = Drive.Files.insert(resource, arquivo, { convert: true });
      const planilhaTemp = SpreadsheetApp.openById(arquivoTemp.id);
      const abaDados = planilhaTemp.getSheets()[0];
      const linhasBrutas = abaDados.getDataRange().getValues();

      if (linhasBrutas.length <= 1) {
        Drive.Files.remove(arquivoTemp.id);
        continue;
      }

      // MAPEAMENTO DINÂMICO POR NOME DE CABEÇALHO
      const headers = linhasBrutas[0].map(h => h.toString().trim().toUpperCase());
      
      const getCol = (nomesPossiveis) => {
        for (let nome of nomesPossiveis) {
          let idx = headers.indexOf(nome);
          if (idx !== -1) return idx;
        }
        return -1;
      };

      // Mapeamento dinâmico com fallback estrito por índice (equivalente à coluna 8 do VBA)
      const idxDataMov = getCol(["DATA MOVIMENTO", "DATA_MOVIMENTO", "DATA MOV", "DATA"]);
      const idxChaveJ = getCol(["CHAVEJ", "CHAVE J", "CHAVE_J", "CHAVE", "PROMOTOR"]);
      const idxContrato = getCol(["NÚMERO PROPOSTA", "NUMERO PROPOSTA", "CONTRATO", "NUM CONTRATO", "NR CONTRATO"]);
      
      // Procura a coluna de data de contrato por nome ou assume a coluna padrão 7/8 do layout original
      let idxDataCont = getCol(["DATA CONTRATO", "DATA_CONTRATO", "DATA_PROPOSTA"]);
      
      const idxProduto = getCol(["CÓDIGO PRODUTO", "CODIGO PRODUTO", "PRODUTO", "COD PRODUTO"]);
      const idxConvenio = getCol(["CÓDIGO CONVÊNIO", "CODIGO CONVENIO", "CONVENIO", "CONVÊNIO"]);
      const idxPrazo = getCol(["PARCELAS", "PARCELA", "PRAZO"]);
      const idxBruto = getCol(["VALOR FINANCIADO", "VALOR BRUTO", "BRUTO"]);
      const idxLiquido = getCol(["VALOR FINANCIADO LÍQUIDO", "VALOR FINANCIADO LIQUIDO", "VALOR LIQUIDO", "LIQUIDO", "VALOR"]);
      const idxTaxa = getCol(["TAXA MENSAL DE JUROS", "TAXA"]);
      const idxCpf = getCol(["CPF", "CPF/CNPJ"]);
      const idxNomeCliente = getCol(["NOME CLIENTE", "CLIENTE", "NOME"]);
      const idxRestricao = getCol(["INDICADOR RESTRIÇÃO SRCC", "INDICADOR RESTRICAO SRCC", "RESTRICAO_RCC", "RESTRICAO"]);

      const novasLinhasParaInserir = [];

      for (let i = 1; i < linhasBrutas.length; i++) {
        const linha = linhasBrutas[i];
        
        const numContrato = idxContrato !== -1 ? String(linha[idxContrato] || "").trim() : "";
        if (!numContrato || contratosExistentes.has(numContrato)) continue;

        const chaveJ = idxChaveJ !== -1 ? String(linha[idxChaveJ] || "").trim() : "";

        // Conversor seguro de datas, extraindo ano e mês sem duplicidade de variáveis
        const processarDataEValores = (val) => {
          if (!val) return { dataFormatada: "", ano: "", mes: "" };
          let d = new Date(val);
          if (isNaN(d.getTime())) {
            const partes = String(val).split("/");
            if (partes.length === 3) {
              d = new Date(partes[2], partes[1] - 1, partes[0]);
            }
          }
          
          if (isNaN(d.getTime())) {
            return { dataFormatada: val, ano: "", mes: "" };
          }

          return {
            dataFormatada: d,
            ano: d.getFullYear(),
            mes: d.getMonth() + 1
          };
        };

        const resDataMov = processarDataEValores(idxDataMov !== -1 ? linha[idxDataMov] : "");
        const dataMovimento = resDataMov.dataFormatada;
        const ano = resDataMov.ano;
        const mes = resDataMov.mes;

       // Fallback robusto: se o índice dinâmico não achar, pega a coluna exata do layout original
        const brutoDataContrato = idxDataCont !== -1 ? linha[idxDataCont] : (linha[7] || linha[5]);
        const resDataCont = processarDataEValores(brutoDataContrato);
        const dataContrato = resDataCont.dataFormatada;

        // 1. Identifica Promotor e Perfil na aba Promotores
        let nomePromotor = "CADASTRO DESCONHECIDO";
        let perfilPromotor = "BLACK";
        
        for (let p = 1; p < dadosPromotores.length; p++) {
          if (String(dadosPromotores[p][0]).trim().toUpperCase() === chaveJ.toUpperCase()) {
            nomePromotor = dadosPromotores[p][1];
            perfilPromotor = String(dadosPromotores[p][2]).trim().toUpperCase();
            break;
          }
        }

        const codProduto = idxProduto !== -1 ? linha[idxProduto] : "";
        const convenio = idxConvenio !== -1 ? linha[idxConvenio] : "";
        const prazo = idxPrazo !== -1 ? Number(linha[idxPrazo]) || 0 : 0;
        const valorBruto = idxBruto !== -1 ? Number(linha[idxBruto]) || 0 : 0;
        const valorLiquido = idxLiquido !== -1 ? Number(linha[idxLiquido]) || 0 : 0;
        const taxa = idxTaxa !== -1 ? Number(linha[idxTaxa]) || 0 : 0; 
        const cpfCliente = idxCpf !== -1 ? linha[idxCpf] : "";
        const nomeCliente = idxNomeCliente !== -1 ? linha[idxNomeCliente] : "";
        const restricaoRcc = idxRestricao !== -1 ? (linha[idxRestricao] || "Não") : "Não";

        // 2. Resolve a descrição do produto baseado na aba Produto
        let grupoProduto = "CONSIGNADO INSS";
        let descProduto = "CONSIGNADO INSS CORRENTISTA NOVO";
        for (let pr = 1; pr < dadosProdutos.length; pr++) {
          if (Number(dadosProdutos[pr][0]) === Number(codProduto)) {
            grupoProduto = dadosProdutos[pr][1];
            descProduto = dadosProdutos[pr][2];
            break;
          }
        }

        // ====================================================================
        // 🚨 NOVA ARMADILHA DE DEBUG - FORA DOS PARÂMETROS PARA CORRIGIR 🚨
        if (nomePromotor.includes("ROBERTA") && codProduto == 2881 && valorLiquido == 550) {
          
          let debugGrupoProduto = grupoProduto;
          let debugDescProduto = descProduto;
          let debugTaxa = taxa;
          let debugPrazo = prazo;
          
          Logger.log(`🛑 PARADA DEBUG: ROBERTA | Produto: 2881 | Valor: 550`);
          Logger.log(`Procurando por -> Grupo: "${debugGrupoProduto}" | Descrição: "${debugDescProduto}"`);
          
          // 👉 COLOQUE O NOVO PONTO VERMELHO (BREAKPOINT) NA LINHA ABAIXO 👈
          let inspecionarVariaveis = true; 
        }
        // ====================================================================

        // 3. Busca a comissão na tabela bdComissao com Mapeamento Exato e Taxa Inteligente (Padrão VBA)
        let fatorComissao = 0;
        let observacaoComissao = "Fora dos Parâmetros"; // Fallback padrão
        let matchEncontrado = false;
        
        let colunaPerfilIdx = 8; 
        const cabecalhoComissao = dadosComissao[0];
        
        for (let c = 0; c < cabecalhoComissao.length; c++) {
          if (String(cabecalhoComissao[c]).trim().toUpperCase() === perfilPromotor) {
            colunaPerfilIdx = c;
            break;
          }
        }

        for (let c = 1; c < dadosComissao.length; c++) {
          const rowC = dadosComissao[c];
          const gFiltro = String(rowC[0]).trim().toUpperCase();
          const dFiltro = String(rowC[1]).trim().toUpperCase();

          // PESQUISA EXATA: Garante que o produto "CRÉDITO 13º" não cruze com a regra do "CRÉDITO SALÁRIO"
          if (grupoProduto.trim().toUpperCase() === gFiltro && descProduto.trim().toUpperCase() === dFiltro) {
            
            // NORMALIZAÇÃO INTELIGENTE DE TAXA (Resolve o bug do Google Sheets vs VBA)
            let rawIni = Number(rowC[2]) || 0;
            let rawFin = Number(rowC[3]) || 0;
            
            // Se o valor for < 1 (ex: 0.0164), é percentagem do Sheets, multiplica-se por 100.
            // Se for > 1 (ex: 1.64), o utilizador digitou o número direto, mantém-se.
            let taxaIni = (rawIni > 0 && rawIni <= 1) ? rawIni * 100 : rawIni;
            let taxaFin = (rawFin > 0 && rawFin <= 1) ? rawFin * 100 : rawFin;
            
            const prazoIni = Number(rowC[4]) || 0;
            const prazoFin = Number(rowC[5]) || 0;

            // Arredondamento para evitar falhas de ponto flutuante na comparação (ex: 1.82000000001)
            const tCheck = Math.round(taxa * 10000) / 10000;
            const tIniCheck = Math.round(taxaIni * 10000) / 10000;
            const tFinCheck = Math.round(taxaFin * 10000) / 10000;

            const matchTaxa = (tCheck >= tIniCheck - 0.001 && tCheck <= tFinCheck + 0.001);
            const matchPrazo = (prazo >= prazoIni && prazo <= prazoFin);

            if (matchTaxa && matchPrazo) {
              fatorComissao = Number(rowC[colunaPerfilIdx]) || 0;
              matchEncontrado = true;
              observacaoComissao = ""; // Limpa a observação pois validou com sucesso
              break; // Sai do loop, encontrou a comissão perfeita
            } else if (!matchTaxa) {
              observacaoComissao = "Abaixo da Taxa Minima"; // Memoriza a falha igual ao loop do VBA
            } else if (!matchPrazo) {
              observacaoComissao = "Fora do Prazo"; // Memoriza a falha de prazo
            }
          }
        }

        if (fatorComissao === 0 && observacaoComissao === "") {
             observacaoComissao = "Fora dos Parâmetros";
        }

        const valorComissaoReal = valorLiquido * fatorComissao;

        const novaLinha = [
          dataMovimento,         // A: Data Movimento
          cpfCliente,            // B: CPF
          "BANCO DO BRASIL",     // C: Banco
          convenio,              // D: Convênio
          numContrato,           // E: Contrato
          dataContrato,          // F: Data Contrato
          taxa,                  // G: Taxa
          prazo,                 // H: Parcela
          chaveJ,                // I: Chave J
          "",                    // J: Comissão PF
          restricaoRcc,          // K: Restrição RCC
          ano,                   // L: Ano
          mes,                   // M: Mês
          nomePromotor,          // N: Promotor
          codProduto,            // O: Produto
          fatorComissao,         // P: Comissão (%)
          perfilPromotor,        // Q: Perfil
          valorComissaoReal,     // R: Valor (Comissão em R$)
          descProduto,           // S: Descrição
          valorBruto,            // T: Valor Bruto
          valorLiquido,          // U: Valor Líquido
          valorLiquido,          // V: Valor Considerado / Produção
          "",                    // W: Agência
          "",                    // X: Empresa
          grupoProduto,          // Y: Desc. Convênio
          observacaoComissao,    // Z: Observação
          ""                     // AA: Pago Em
        ];

        novasLinhasParaInserir.push(novaLinha);
        contratosExistentes.add(numContrato); 
      }

      if (novasLinhasParaInserir.length > 0) {
        sheetProd.getRange(sheetProd.getLastRow() + 1, 1, novasLinhasParaInserir.length, novasLinhasParaInserir[0].length).setValues(novasLinhasParaInserir);
        Logger.log(`✅ Sucesso: ${novasLinhasParaInserir.length} contratos gravados com mapeamento dinâmico e taxas corrigidas.`);
      }

      Drive.Files.remove(arquivoTemp.id);
      arquivo.moveTo(pastaUsados);
      arquivosProcessados++;
    }

    return { sucesso: true, mensagem: "Importação concluída com sucesso!" };

  } catch (e) {
    Logger.log(`❌ ERRO: ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}