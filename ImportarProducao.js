/**
 * Automatiza a importação dos arquivos .xlsx de produção do Google Drive,
 * aplicando regras de comissão (bdComissao) e promotores (Promotores).
 */
function importarProducaoDoDrive() {
  try {
    const idPastaProducao = "1ivHA6X9Ku-1qyFTFkwbL-qNCz4PWSHji";
    const idPastaUsados = "1QTq9mXv_UyBqqgTmhPiSIJunR29WfEIp";
    
    const pastaProducao = DriveApp.getFolderById(idPastaProducao);
    const pastaUsados = DriveApp.getFolderById(idPastaUsados);
    
    const arquivos = pastaProducao.getFilesByType(MimeType.MICROSOFT_EXCEL);
    
    if (!arquivos.hasNext()) {
      return { sucesso: false, mensagem: "Nenhum arquivo .xlsx encontrado na pasta de Produção." };
    }

    let arquivosProcessados = 0;
    const ss = getDatabaseConnection();
    
    // Carrega tabelas de apoio da base de dados em memória para alta performance
    const dadosPromotores = ss.getSheetByName("Promotores").getDataRange().getValues();
    const dadosComissao = ss.getSheetByName("bdComissao").getDataRange().getValues();
    const dadosProdutos = ss.getSheetByName("Produto").getDataRange().getValues();
    
    const sheetProd = ss.getSheetByName("bd_Producao");
    const sheetCliente = ss.getSheetByName("bd_Cliente");
    const dadosProducaoAtual = sheetProd.getLastRow() > 1 ? sheetProd.getRange(2, 1, sheetProd.getLastRow() - 1, sheetProd.getLastColumn()).getValues() : [];
    
    // Índices de contratos já existentes para evitar duplicados (Coluna E / Índice 4 = Contrato)
    const contratosExistentes = new Set(dadosProducaoAtual.map(row => String(row[4]).trim()));

    while (arquivos.hasNext()) {
      const arquivoExcel = arquivos.next();
      const nomeArquivo = arquivoExcel.getName();

      // Converte temporariamente para Google Sheets para leitura nativa
      const resource = {
        title: "[TEMP_IMPORT] " + nomeArquivo,
        mimeType: MimeType.GOOGLE_SHEETS
      };
      const arquivoTemp = Drive.Files.insert(resource, arquivoExcel, { convert: true });
      const planilhaTemp = SpreadsheetApp.openById(arquivoTemp.id);
      const abaDados = planilhaTemp.getSheets()[0];
      const linhasBrutas = abaDados.getDataRange().getValues();

      if (linhasBrutas.length <= 1) {
        Drive.Files.remove(arquivoTemp.id);
        continue;
      }

      const cabecalhos = linhasBrutas[0];
      const novasLinhasParaInserir = [];
      const novosClientesParaInserir = [];

      // Processa cada linha do arquivo importado (começando da linha 1)
      for (let i = 1; i < linhasBrutas.length; i++) {
        const linha = linhasBrutas[i];
        
        // Mapeamento padrão compatível com a estrutura bruta do relatório
        const dataMovimento = linha[1]; // Coluna B aproximada ou ajustada conforme layout
        const chaveJ = String(linha[3] || "").trim(); // Chave J
        const numContrato = String(linha[8] || "").trim(); // Contrato

        if (!numContrato || contratosExistentes.has(numContrato)) continue;

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

        // 2. Extrai dados financeiros e contratuais
        const dataContrato = linha[7];
        const codProduto = linha[4];
        const convenio = linha[6];
        const prazo = Number(linha[9]) || 0;
        const valorBruto = Number(linha[10]) || 0;
        const valorLiquido = Number(linha[11]) || 0;
        const taxa = Number(linha[16]) || 0; // Taxa
        const cpfCliente = linha[22];
        const nomeCliente = linha[23];
        const restricaoRcc = linha[27];

        // 3. Resolve a descrição do produto baseado na aba Produto
        let grupoProduto = "CONSIGNADO INSS";
        let descProduto = "CONSIGNADO INSS CORRENTISTA NOVO";
        for (let pr = 1; pr < dadosProdutos.length; pr++) {
          if (Number(dadosProdutos[pr][0]) === Number(codProduto)) {
            grupoProduto = dadosProdutos[pr][1];
            descProduto = dadosProdutos[pr][2];
            break;
          }
        }

        // 4. Busca a comissão na tabela bdComissao (Cruzando Grupo, Descrição, Taxa e Prazo)
        let fatorComissao = 0;
        let observacaoComissao = "";

        // Identifica a coluna correta do perfil na tabela bdComissao (Cabeçalho da linha 0)
        let colunaPerfilIdx = 8; // Default GESTOR/BLACK (ajustado conforme colunas da imagem)
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

        // Monta a linha alinhada exatamente à estrutura de colunas do bd_Producao
        const novaLinha = [
          dataMovimento ? new Date(dataMovimento) : "", // A: Data Movimento
          cpfCliente,                                    // B: CPF
          "BANCO DO BRASIL",                             // C: Banco
          convenio,                                      // D: Convênio
          numContrato,                                   // E: Contrato
          dataContrato ? new Date(dataContrato) : "",    // F: Data Contrato
          taxa,                                          // G: Taxa
          prazo,                                         // H: Parcela
          chaveJ,                                        // I: Chave J
          "",                                            // J: Comissão PF
          restricaoRcc,                                  // K: Restrição RCC
          ano,                                           // L: Ano
          mes,                                           // M: Mês
          nomePromotor,                                  // N: Promotor
          codProduto,                                    // O: Produto (Código)
          fatorComissao,                                 // P: Comissão (%)
          perfilPromotor,                                // Q: Perfil
          valorComissaoReal,                             // R: Valor (Comissão em R$)
          descProduto,                                   // S: Descrição
          valorBruto,                                    // T: Valor Bruto
          valorLiquido,                                  // U: Valor Líquido
          valorLiquido,                                  // V: Valor Considerado / Produção
          "",                                            // W: Agência
          "",                                            // X: Empresa
          grupoProduto,                                  // Y: Desc. Convênio
          observacaoComissao,                            // Z: Observação
          ""                                             // AA: Pago Em
        ];

        novasLinhasParaInserir.push(novaLinha);
        contratosExistentes.add(numContrato); // Evita duplicidade dentro do próprio lote

        // Prepara inclusão na base de clientes se não existir
        if (cpfCliente) {
          novosClientesParaInserir.push([cpfCliente, nomeCliente, "Banco do Brasil", new Date()]);
        }
      }

      // Grava em lote na aba bd_Producao se houver novos dados
      if (novasLinhasParaInserir.length > 0) {
        sheetProd.getRange(sheetProd.getLastRow() + 1, 1, novasLinhasParaInserir.length, novasLinhasParaInserir[0].setValues ? novasLinhasParaInserir[0].length : novasLinhasParaInserir[0].length).setValues(novasLinhasParaInserir);
      }

      Drive.Files.remove(arquivoTemp.id);
      arquivoExcel.moveTo(pastaUsados);
      arquivosProcessados++;
    }

    return { sucesso: true, mensagem: `${arquivosProcessados} arquivo(s) de produção importado(s) com sucesso!` };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}