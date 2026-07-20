/**
 * Automatiza a importação dos arquivos de produção do Google Drive,
 * aplicando regras de comissão (bdComissao) e promotores (Promotores).
 */
function importarProducaoDoDrive() {
  try {
    const idPastaProducao = "1ivHA6X9Ku-1qyFTFkwbL-qNCz4PWSHji";
    const idPastaUsados = "1QTq9mXv_UyBqqgTmhPiSIJunR29WfEIp";
    
    const pastaProducao = DriveApp.getFolderById(idPastaProducao);
    const pastaUsados = DriveApp.getFolderById(idPastaUsados);
    
    // Lista todos os arquivos na pasta sem restrição estrita de MIME para diagnóstico
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
      
      Logger.log(`📁 Analisando arquivo: ${nomeArquivo} (Tipo: ${mimeType})`);

      // Aceita tanto ficheiros Excel quanto Google Sheets convertidos
      if (mimeType !== MimeType.MICROSOFT_EXCEL && 
          mimeType !== MimeType.GOOGLE_SHEETS && 
          !nomeArquivo.endsWith(".xlsx")) {
        Logger.log(`⏭️ Ignorado (não é Excel compatível): ${nomeArquivo}`);
        continue;
      }

      let abaDados;
      let idArquivoTemp = null;

      if (mimeType === MimeType.MICROSOFT_EXCEL) {
        // Converte o .xlsx para Google Sheets temporariamente
        const resource = {
          title: "[TEMP_IMPORT] " + nomeArquivo,
          mimeType: MimeType.GOOGLE_SHEETS
        };
        const arquivoTemp = Drive.Files.insert(resource, arquivo, { convert: true });
        idArquivoTemp = arquivoTemp.id;
        const planilhaTemp = SpreadsheetApp.openById(idArquivoTemp);
        abaDados = planilhaTemp.getSheets()[0];
      } else {
        // Se já for uma planilha nativa do Google
        abaDados = SpreadsheetApp.openById(arquivo.getId()).getSheets()[0];
      }

      const linhasBrutas = abaDados.getDataRange().getValues();
      Logger.log(`📊 Linhas encontradas no arquivo ${nomeArquivo}: ${linhasBrutas.length}`);

      if (linhasBrutas.length <= 1) {
        if (idArquivoTemp) Drive.Files.remove(idArquivoTemp);
        continue;
      }

      const novasLinhasParaInserir = [];
      const novosClientesParaInserir = [];

      for (let i = 1; i < linhasBrutas.length; i++) {
        const linha = linhasBrutas[i];
        
        const dataMovimento = linha[1]; 
        const chaveJ = String(linha[3] || "").trim(); 
        const numContrato = String(linha[8] || "").trim(); 

        if (!numContrato || contratosExistentes.has(numContrato)) continue;

        let nomePromotor = "CADASTRO DESCONHECIDO";
        let perfilPromotor = "BLACK";
        
        for (let p = 1; p < dadosPromotores.length; p++) {
          if (String(dadosPromotores[p][0]).trim().toUpperCase() === chaveJ.toUpperCase()) {
            nomePromotor = dadosPromotores[p][1];
            perfilPromotor = String(dadosPromotores[p][2]).trim().toUpperCase();
            break;
          }
        }

        const dataContrato = linha[7];
        const codProduto = linha[4];
        const convenio = linha[6];
        const prazo = Number(linha[9]) || 0;
        const valorBruto = Number(linha[10]) || 0;
        const valorLiquido = Number(linha[11]) || 0;
        const taxa = Number(linha[16]) || 0; 
        const cpfCliente = linha[22];
        const nomeCliente = linha[23];
        const restricaoRcc = linha[27];

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
          dataMovimento ? new Date(dataMovimento) : "", 
          cpfCliente,                                    
          "BANCO DO BRASIL",                             
          convenio,                                      
          numContrato,                                   
          dataContrato ? new Date(dataContrato) : "",    
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

        if (cpfCliente) {
          novosClientesParaInserir.push([cpfCliente, nomeCliente, "Banco do Brasil", new Date()]);
        }
      }

      if (novasLinhasParaInserir.length > 0) {
        sheetProd.getRange(sheetProd.getLastRow() + 1, 1, novasLinhasParaInserir.length, novasLinhasParaInserir[0].length).setValues(novasLinhasParaInserir);
        Logger.log(`✅ Sucesso: ${novasLinhasParaInserir.length} contratos gravados em 'bd_Producao'.`);
      } else {
        Logger.log(`⚠️ Nenhum contrato novo para inserir (podem já existir na base).`);
      }

      if (idArquivoTemp) Drive.Files.remove(idArquivoTemp);
      
      // Move o arquivo original para a pasta de usados
      arquivo.moveTo(pastaUsados);
      Logger.log(`🚚 Arquivo movido para 'Arquivos_Usados'.`);
      
      arquivosProcessados++;
    }

    return { sucesso: true, mensagem: `${arquivosProcessados} arquivo(s) importado(s) com sucesso!` };

  } catch (e) {
    Logger.log(`❌ ERRO CRÍTICO: ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}