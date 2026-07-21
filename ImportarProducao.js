/**
 * Importação otimizada com mapeamento dinâmico de cabeçalhos,
 * integração da aba Convênio, correção exata de datas e cruzamento de taxas.
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
    
    // Carrega todas as tabelas de apoio para a memória
    const dadosPromotores = ss.getSheetByName("Promotores").getDataRange().getValues();
    const dadosComissao = ss.getSheetByName("bdComissao").getDataRange().getValues();
    const dadosProdutos = ss.getSheetByName("Produto").getDataRange().getValues();
    const dadosConvenio = ss.getSheetByName("Convenio").getDataRange().getValues(); // <-- NOVA ABA CARREGADA
    
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

      const idxDataMov = getCol(["DATA MOVIMENTO", "DATA_MOVIMENTO", "DATA MOV", "DATA"]);
      const idxChaveJ = getCol(["CHAVEJ", "CHAVE J", "CHAVE_J", "CHAVE", "PROMOTOR"]);
      const idxContrato = getCol(["NÚMERO PROPOSTA", "NUMERO PROPOSTA", "CONTRATO", "NUM CONTRATO", "NR CONTRATO"]);
      let idxDataCont = getCol(["DATA CONTRATO", "DATA_CONTRATO", "DATA_PROPOSTA"]);
      const codProduto = idxProduto !== -1 ? linha[idxProduto] : (linha[4] || "");
      const convenio = idxConvenio !== -1 ? linha[idxConvenio] : (linha[6] || ""); // Fallback para a coluna 6
      const idxPrazo = getCol(["PARCELAS", "PARCELA", "PRAZO"]); // Prioridade corrigida
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

        // Conversor seguro de datas
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

        const brutoDataContrato = idxDataCont !== -1 ? linha[idxDataCont] : (linha[7] || linha[5]);
        const resDataCont = processarDataEValores(brutoDataContrato);
        const dataContrato = resDataCont.dataFormatada;

        // 1. Identifica Promotor e Perfil
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

        // 2. Resolve o Convênio (Lê a nova aba para descobrir se é SP, INSS, etc)
        let descConvenio = "";
        for (let cv = 1; cv < dadosConvenio.length; cv++) {
          // Compara como Texto E como Número para não falhar na formatação
          if (String(dadosConvenio[cv][0]).trim() === String(convenio).trim() || 
              Number(dadosConvenio[cv][0]) === Number(convenio)) {
            descConvenio = String(dadosConvenio[cv][1]).trim().toUpperCase();
            break;
          }
        }

        // 3. Resolve a descrição do produto baseado na aba Produto
        let grupoProduto = "CONSIGNADO INSS";
        let descProduto = "CONSIGNADO INSS CORRENTISTA NOVO";
        for (let pr = 1; pr < dadosProdutos.length; pr++) {
          if (String(dadosProdutos[pr][0]).trim() === String(codProduto).trim()) {
            grupoProduto = String(dadosProdutos[pr][1]).trim().toUpperCase();
            descProduto = String(dadosProdutos[pr][2]).trim().toUpperCase();
            break;
          }
        }

        // 💥 O PULO DO GATO: Concatena o Convênio ao Grupo de Produto se for Consignado!
        if (grupoProduto === "CONSIGNADO" && descConvenio !== "") {
          grupoProduto = grupoProduto + " " + descConvenio; // Fica "CONSIGNADO SP" ou "CONSIGNADO INSS"
        }

        // ====================================================================
        // 🚨 ARMADILHA DE DEBUG - VALIDAÇÃO DA JUNÇÃO DO CONVÊNIO 🚨
        if (nomePromotor.includes("ROBERTA") && codProduto == 2881 && valorLiquido == 550) {
          
          let debugCodigoConvenioBruto = convenio; // O que o script leu do Excel
          let debugConvenioLido = descConvenio;
          let debugGrupoFinal = grupoProduto;
          
          Logger.log(`🛑 PARADA DEBUG: ROBERTA | Produto: 2881 | Valor Líquido: 550`);
          Logger.log(`👉 Código do Convênio lido do Excel: "${debugCodigoConvenioBruto}"`);
          Logger.log(`🔍 Resultado -> Sigla Encontrada: "${debugConvenioLido}" | Grupo Final: "${debugGrupoFinal}"`);
          
          // 👉 COLOQUE O PONTO VERMELHO (BREAKPOINT) NESTA LINHA ABAIXO 👈
          let inspecionarVariaveis = true; 
        }
        // ====================================================================

        // 4. Busca a comissão na tabela bdComissao com Mapeamento Exato
        let fatorComissao = 0;
        let observacaoComissao = "Fora dos Parâmetros";
        let matchEncontrado = false;
        
        let colunaPerfilIdx = 8; 
        const cabecalhoComissao = dadosComissao[0];
        
        for (let c = 0; c < cabecalhoComissao.length; c++) {
          if (String(cabecalhoComissao[c]).trim().toUpperCase() === perfilPromotor) {
            colunaPerfilIdx = c;
            break;
          }
        }

        let probTx = 0;
        let probParc = 0;

        for (let c = 1; c < dadosComissao.length; c++) {
          const rowC = dadosComissao[c];
          const gFiltro = String(rowC[0]).trim().toUpperCase();
          const dFiltro = String(rowC[1]).trim().toUpperCase();

          if (grupoProduto.toUpperCase() === gFiltro && descProduto.toUpperCase() === dFiltro) {
            
            let rawIni = Number(rowC[2]) || 0;
            let rawFin = Number(rowC[3]) || 0;
            
            let taxaIni = (rawIni > 0 && rawIni <= 1) ? rawIni * 100 : rawIni;
            let taxaFin = (rawFin > 0 && rawFin <= 1) ? rawFin * 100 : rawFin;
            
            const prazoIni = Number(rowC[4]) || 0;
            const prazoFin = Number(rowC[5]) || 0;

            const tCheck = Math.round(taxa * 10000) / 10000;
            const tIniCheck = Math.round(taxaIni * 10000) / 10000;
            const tFinCheck = Math.round(taxaFin * 10000) / 10000;

            const matchTaxa = (tCheck >= tIniCheck - 0.001 && tCheck <= tFinCheck + 0.001);
            const matchPrazo = (prazo >= prazoIni && prazo <= prazoFin);

            if (matchTaxa && matchPrazo) {
              fatorComissao = Number(rowC[colunaPerfilIdx]) || 0;
              matchEncontrado = true;
              observacaoComissao = ""; 
              break; 
            } else if (!matchTaxa) {
              observacaoComissao = "Abaixo da Taxa Minima";
            } else if (!matchPrazo) {
              observacaoComissao = "Fora do Prazo";
            }
          }
        }

        if (fatorComissao === 0 && observacaoComissao === "") {
             observacaoComissao = "Fora dos Parâmetros";
        }

        const valorComissaoReal = valorLiquido * fatorComissao;

        const novaLinha = [
          dataMovimento,         // A
          cpfCliente,            // B
          "BANCO DO BRASIL",     // C
          convenio,              // D
          numContrato,           // E
          dataContrato,          // F
          taxa,                  // G
          prazo,                 // H
          chaveJ,                // I
          "",                    // J
          restricaoRcc,          // K
          ano,                   // L
          mes,                   // M
          nomePromotor,          // N
          codProduto,            // O
          fatorComissao,         // P
          perfilPromotor,        // Q
          valorComissaoReal,     // R
          descProduto,           // S
          valorBruto,            // T
          valorLiquido,          // U
          valorLiquido,          // V
          "",                    // W
          "",                    // X
          grupoProduto,          // Y
          observacaoComissao,    // Z
          ""                     // AA
        ];

        novasLinhasParaInserir.push(novaLinha);
        contratosExistentes.add(numContrato); 
      }

      if (novasLinhasParaInserir.length > 0) {
        sheetProd.getRange(sheetProd.getLastRow() + 1, 1, novasLinhasParaInserir.length, novasLinhasParaInserir[0].length).setValues(novasLinhasParaInserir);
        Logger.log(`✅ Sucesso: ${novasLinhasParaInserir.length} contratos gravados.`);
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