/**
 * Automatiza a importação com logs de diagnóstico de cabeçalho.
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
      return { sucesso: false, mensagem: "Nenhum arquivo encontrado." };
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

      const headers = linhasBrutas[0].map(h => h.toString().trim().toUpperCase());
      Logger.log(`🔍 Cabeçalhos detetados no Excel: ${JSON.stringify(headers)}`);
      
      const getCol = (nomesPossiveis) => {
        for (let nome of nomesPossiveis) {
          let idx = headers.indexOf(nome);
          if (idx !== -1) return idx;
        }
        return -1;
      };

      const idxDataMov = getCol(["DATA MOVIMENTO", "DATA_MOVIMENTO", "DATA MOV", "DATA"]);
      const idxChaveJ = getCol(["CHAVE J", "CHAVE_J", "CHAVE", "PROMOTOR"]);
      const idxContrato = getCol(["CONTRATO", "NUM CONTRATO", "NR CONTRATO"]);
      const idxDataCont = getCol(["DATA CONTRATO", "DATA_CONTRATO"]);
      const idxProduto = getCol(["PRODUTO", "COD PRODUTO", "CÓDIGO PRODUTO"]);
      const idxConvenio = getCol(["CONVENIO", "CONVÊNIO"]);
      const idxPrazo = getCol(["PARCELA", "PRAZO", "PARCELAS"]);
      const idxBruto = getCol(["VALOR BRUTO", "BRUTO"]);
      const idxLiquido = getCol(["VALOR LIQUIDO", "LIQUIDO", "VALOR LÍQUIDO", "VALOR"]);
      const idxTaxa = getCol(["TAXA"]);
      const idxCpf = getCol(["CPF", "CPF/CNPJ"]);
      const idxNomeCliente = getCol(["CLIENTE", "NOME CLIENTE", "NOME"]);
      const idxRestricao = getCol(["RESTRICAO_RCC", "RESTRICAO", "RESTRIÇÃO RCC"]);

      const novasLinhasParaInserir = [];

      for (let i = 1; i < linhasBrutas.length; i++) {
        const linha = linhasBrutas[i];
        
        // Se a coluna de contrato foi encontrada, usa-a; senão, tenta pegar pelo índice fixo de segurança caso os cabeçalhos difiram
        let numContrato = "";
        if (idxContrato !== -1) {
          numContrato = String(linha[idxContrato] || "").trim();
        } else if (linha.length > 8) {
          numContrato = String(linha[8] || "").trim(); // Posição padrão anterior
        }

        if (!numContrato || contratosExistentes.has(numContrato)) continue;

        let chaveJ = "";
        if (idxChaveJ !== -1) {
          chaveJ = String(linha[idxChaveJ] || "").trim();
        } else if (linha.length > 3) {
          chaveJ = String(linha[3] || "").trim();
        }

        const parseData = (val) => {
          if (!val) return "";
          let d = new Date(val);
          return isNaN(d.getTime()) ? "" : d;
        };

        const dataMovimento = idxDataMov !== -1 ? parseData(linha[idxDataMov]) : parseData(linha[1]);
        const dataContrato = idxDataCont !== -1 ? parseData(linha[idxDataCont]) : parseData(linha[7]);

        let nomePromotor = "CADASTRO DESCONHECIDO";
        let perfilPromotor = "BLACK";
        
        for (let p = 1; p < dadosPromotores.length; p++) {
          if (String(dadosPromotores[p][0]).trim().toUpperCase() === chaveJ.toUpperCase()) {
            nomePromotor = dadosPromotores[p][1];
            perfilPromotor = String(dadosPromotores[p][2]).trim().toUpperCase();
            break;
          }
        }

        const codProduto = idxProduto !== -1 ? linha[idxProduto] : linha[4];
        const convenio = idxConvenio !== -1 ? linha[idxConvenio] : linha[6];
        const prazo = idxPrazo !== -1 ? Number(linha[idxPrazo]) || 0 : Number(linha[9]) || 0;
        const valorBruto = idxBruto !== -1 ? Number(linha[idxBruto]) || 0 : Number(linha[10]) || 0;
        const valorLiquido = idxLiquido !== -1 ? Number(linha[idxLiquido]) || 0 : Number(linha[11]) || 0;
        const taxa = idxTaxa !== -1 ? Number(linha[idxTaxa]) || 0 : Number(linha[16]) || 0;
        const cpfCliente = idxCpf !== -1 ? linha[idxCpf] : linha[22];
        const nomeCliente = idxNomeCliente !== -1 ? linha[idxNomeCliente] : linha[23];
        const restricaoRcc = idxRestricao !== -1 ? linha[idxRestricao] : (linha[27] || "Não");

        let grupoProduto = "CONSIGNADO INSS";
        let descProduto = "CONSIGNADO INSS CORRENTISTA NOVO";
        for (let pr = 1; pr < dadosProdutos.length; pr++) {
          if (Number(dadosProdutos[pr][0]) === Number(codProduto)) {
            grupoProduto = dadosProdutos[pr][1];
            descProduto = dadosProdutos[pr][2];
            break;
          }
        }

        let fatorComissao = 0;
        let observacaoComissao = "";
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
          const taxaIni = Number(rowC[2]) * 100;
          const taxaFin = Number(rowC[3]) * 100;
          const prazoIni = Number(rowC[4]);
          const prazoFin = Number(rowC[5]);

          if (grupoProduto.toUpperCase().includes(gFiltro) && descProduto.toUpperCase().includes(dFiltro)) {
            if (taxa >= taxaIni && taxa <= taxaFin && prazo >= prazoIni && prazo <= prazoFin) {
              fatorComissao = Number(rowC[colunaPerfilIdx]) || 0;
              break;
            }
          }
        }

        if (fatorComissao === 0) {
          observacaoComissao = "Abaixo da Taxa Minima / Fora do Prazo";
        }

        const ano = dataMovimento ? new Date(dataMovimento).getFullYear() : "";
        const mes = dataMovimento ? new Date(dataMovimento).getMonth() + 1 : "";
        const valorComissaoReal = valorLiquido * fatorComissao;

        const novaLinha = [
          dataMovimento, cpfCliente, "BANCO DO BRASIL", convenio, numContrato, 
          dataContrato, taxa, prazo, chaveJ, "", restricaoRcc, ano, mes, 
          nomePromotor, codProduto, fatorComissao, perfilPromotor, valorComissaoReal, 
          descProduto, valorBruto, valorLiquido, valorLiquido, "", "", 
          grupoProduto, observacaoComissao, ""
        ];

        novasLinhasParaInserir.push(novaLinha);
        contratosExistentes.add(numContrato); 
      }

      Logger.log(`✨ Linhas mapeadas prontas para inserção: ${novasLinhasParaInserir.length}`);

      if (novasLinhasParaInserir.length > 0) {
        sheetProd.getRange(sheetProd.getLastRow() + 1, 1, novasLinhasParaInserir.length, novasLinhasParaInserir[0].length).setValues(novasLinhasParaInserir);
        Logger.log(`✅ Sucesso absoluto: ${novasLinhasParaInserir.length} contratos gravados.`);
      } else {
        Logger.log(`⚠️ Nenhuma linha foi gravada (verifique se os contratos já existiam na base).`);
      }

      Drive.Files.remove(arquivoTemp.id);
      arquivo.moveTo(pastaUsados);
      arquivosProcessados++;
    }

    return { sucesso: true, mensagem: "Importação concluída!" };

  } catch (e) {
    Logger.log(`❌ ERRO: ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}