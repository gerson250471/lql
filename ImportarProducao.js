/**
 * Automatiza a importação dos arquivos de produção do Google Drive,
 * mapeando dinamicamente as colunas por nome e aplicando as regras de comissão.
 */
function importarProducaoDoDrive() {
  try {
    const idPastaProducao = "1ivHA6X9Ku-1qyFTFkwbL-qNCz4PWSHji";
    const idPastaUsados = "1QTq9mXv_UyBqqgTmhPiSIJunR29WfEIp";
    
    const pastaProducao = DriveApp.getFolderById(idPastaProducao);
    const pastaUsados = DriveApp.getFolderById(idPastaUsados);
    
    const arquivos = pastaProducao.getFiles();
    
    if (!arquivos.hasNext()) {
      Logger.log("⚠️ AVISO: A pasta 'Producao' está totalmente vazia.");
      return { sucesso: false, mensagem: "Nenhum arquivo encontrado na pasta de Produção." };
    }

    let arquivosProcessados = 0;
    const ss = getDatabaseConnection();
    
    // Carrega tabelas de apoio da base de dados em memória
    const dadosPromotores = ss.getSheetByName("Promotores").getDataRange().getValues();
    const dadosComissao = ss.getSheetByName("bdComissao").getDataRange().getValues();
    const dadosProdutos = ss.getSheetByName("Produto").getDataRange().getValues();
    
    const sheetProd = ss.getSheetByName("bd_Producao");
    const sheetCliente = ss.getSheetByName("bd_Cliente");
    const dadosProducaoAtual = sheetProd.getLastRow() > 1 ? sheetProd.getRange(2, 1, sheetProd.getLastRow() - 1, sheetProd.getLastColumn()).getValues() : [];
    
    const contratosExistentes = new Set(dadosProducaoAtual.map(row => String(row[4]).trim()));

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nomeArquivo = arquivo.getName();
      const mimeType = arquivo.getMimeType();
      
      Logger.log(`📁 Analisando arquivo: ${nomeArquivo}`);

      if (mimeType !== MimeType.MICROSOFT_EXCEL && 
          mimeType !== MimeType.GOOGLE_SHEETS && 
          !nomeArquivo.endsWith(".xlsx")) {
        continue;
      }

      let abaDados;
      let idArquivoTemp = null;

      if (mimeType === MimeType.MICROSOFT_EXCEL) {
        const resource = {
          title: "[TEMP_IMPORT] " + nomeArquivo,
          mimeType: MimeType.GOOGLE_SHEETS
        };
        const arquivoTemp = Drive.Files.insert(resource, arquivo, { convert: true });
        idArquivoTemp = arquivoTemp.id;
        const planilhaTemp = SpreadsheetApp.openById(idArquivoTemp);
        abaDados = planilhaTemp.getSheets()[0];
      } else {
        abaDados = SpreadsheetApp.openById(arquivo.getId()).getSheets()[0];
      }

      const linhasBrutas = abaDados.getDataRange().getValues();
      if (linhasBrutas.length <= 1) {
        if (idArquivoTemp) Drive.Files.remove(idArquivoTemp);
        continue;
      }

      // MAPEAMENTO DINÂMICO DOS CABEÇALHOS DO EXCEL IMPORTADO
      const headers = linhasBrutas[0].map(h => h.toString().trim().toUpperCase());
      
      const getCol = (nomesPossiveis) => {
        for (let nome of nomesPossiveis) {
          let idx = headers.indexOf(nome);
          if (idx !== -1) return idx;
        }
        return -1;
      };

      // Identifica os índices dinamicamente na planilha de origem
      const idxDataMov = getCol(["DATA MOVIMENTO", "DATA_MOVIMENTO", "DATA MOV"]);
      const idxChaveJ = getCol(["CHAVE J", "CHAVE_J", "CHAVE"]);
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
        
        const numContrato = idxContrato !== -1 ? String(linha[idxContrato] || "").trim() : "";
        if (!numContrato || contratosExistentes.has(numContrato)) continue;

        const chaveJ = idxChaveJ !== -1 ? String(linha[idxChaveJ] || "").trim() : "";
        
        // Formata datas com segurança para evitar bugs de 1969
        const parseData = (val) => {
          if (!val) return "";
          let d = new Date(val);
          return isNaN(d.getTime()) ? "" : d;
        };

        const dataMovimento = idxDataMov !== -1 ? parseData(linha[idxDataMov]) : "";
        const dataContrato = idxDataCont !== -1 ? parseData(linha[idxDataCont]) : "";

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
        const restricaoRcc = idxRestricao !== -1 ? linha[idxRestricao] : "Não";

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

        // 3. Busca a comissão na tabela bdComissao
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
          dataMovimento, 
          cpfCliente,                                    
          "BANCO DO BRASIL",                             
          convenio,                                      
          numContrato,                                   
          dataContrato,    
          taxa,                                          
          prazo,                                         
          chaveJ,                                        
          "",                                            
          restricaoRcc,                                  
          ano,                                           
          mes,                                           
          nomePromotor,                                  
          codProduto,                                    
          fatorComissao,                                 
          perfilPromotor,                                
          valorComissaoReal,                             
          descProduto,                                   
          valorBruto,                                    
          valorLiquido,                                  
          valorLiquido,                                  
          "",                                            
          "",                                            
          grupoProduto,                                  
          observacaoComissao,                            
          ""                                             
        ];

        novasLinhasParaInserir.push(novaLinha);
        contratosExistentes.add(numContrato); 
      }

      if (novasLinhasParaInserir.length > 0) {
        sheetProd.getRange(sheetProd.getLastRow() + 1, 1, novasLinhasParaInserir.length, novasLinhasParaInserir[0].length).setValues(novasLinhasParaInserir);
        Logger.log(`✅ Sucesso: ${novasLinhasParaInserir.length} contratos gravados com mapeamento dinâmico.`);
      }

      if (idArquivoTemp) Drive.Files.remove(idArquivoTemp);
      
      arquivo.moveTo(pastaUsados);
      arquivosProcessados++;
    }

    return { sucesso: true, mensagem: `${arquivosProcessados} arquivo(s) importado(s) com sucesso!` };

  } catch (e) {
    Logger.log(`❌ ERRO CRÍTICO: ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}