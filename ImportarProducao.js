/**
 * Importa a produção do Drive utilizando o mapeamento posicional exato (padrão VBA),
 * calculando comissões, perfis e evitando duplicidade por contrato.
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
    
    // Carrega tabelas de apoio da base de dados em memória
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

      const novasLinhasParaInserir = [];
      const novosClientesParaInserir = [];

      // Mapeamento posicional exato baseado na estrutura do seu VBA (MOD_A_0_PRODUCAO.vba)
      for (let i = 1; i < linhasBrutas.length; i++) {
        const linha = linhasBrutas[i];
        
        // Coluna E do VBA (Índice 4 no array = Contrato)
        const numContrato = String(linha[8] || "").trim(); 
        if (!numContrato || contratosExistentes.has(numContrato)) continue;

        // Coluna I do VBA (Índice 8 no array = Chave J)
        const chaveJ = String(linha[3] || "").trim(); 

        const parseData = (val) => {
          if (!val) return "";
          let d = new Date(val);
          return isNaN(d.getTime()) ? "" : d;
        };

        const dataMovimento = parseData(linha[1]); // Data Movimento
        const dataContrato = parseData(linha[7]);  // Data Contrato

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

        const codProduto = Number(linha[4]) || 0;     // Produto
        const convenio = linha[6];                     // Convênio
        const prazo = Number(linha[9]) || 0;           // Parcela / Prazo
        const valorBruto = Number(linha[10]) || 0;     // Valor Bruto
        const valorLiquido = Number(linha[11]) || 0;   // Valor Líquido
        const taxa = Number(linha[16]) || 0;           // Taxa Mensal de Juros
        const cpfCliente = linha[22];                  // CPF
        const nomeCliente = linha[23];                 // Nome Cliente
        const restricaoRcc = linha[27] || "Não";       // Restrição RCC

        // 2. Resolve a descrição do produto baseado na aba Produto
        let grupoProduto = "CONSIGNADO INSS";
        let descProduto = "CONSIGNADO INSS CORRENTISTA NOVO";
        for (let pr = 1; pr < dadosProdutos.length; pr++) {
          if (Number(dadosProdutos[pr][0]) === codProduto) {
            grupoProduto = dadosProdutos[pr][1];
            descProduto = dadosProdutos[pr][2];
            break;
          }
        }

        // 3. Busca a comissão exata na tabela bdComissao (Cruzando faixas de taxa e prazo)
        let fatorComissao = 0;
        let observacaoComissao = "";
        let colunaPerfilIdx = 8; // Índice padrão para GESTOR/BLACK
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
          const taxaIni = Number(rowC[2]) * 100;
          const taxaFin = Number(rowC[3]) * 100;
          const prazoIni = Number(rowC[4]);
          const prazoFin = Number(rowC[5]);

          if (grupoProduto.toUpperCase().includes(gFiltro) && descProduto.toUpperCase().includes(dFiltro)) {
            if (taxa >= taxaIni && taxa <= taxaFin) probTx++;
            if (prazo >= prazoIni && prazo <= prazoFin) probParc++;

            if (taxa >= taxaIni && taxa <= taxaFin && prazo >= prazoIni && prazo <= prazoFin) {
              fatorComissao = Number(rowC[colunaPerfilIdx]) || 0;
              break;
            }
          }
        }

        // Regra idêntica ao VBA para preencher a observação de erro
        if (fatorComissao === 0) {
          if (probTx === 0) observacaoComissao = "Abaixo da Taxa Minima";
          else if (probParc === 0) observacaoComissao = "Fora do Prazo";
          else observacaoComissao = "Fora dos Parâmetros";
        }

        const ano = dataMovimento ? new Date(dataMovimento).getFullYear() : "";
        const mes = dataMovimento ? new Date(dataMovimento).getMonth() + 1 : "";
        const valorComissaoReal = valorLiquido * fatorComissao;

        // Estrutura exata das colunas da aba bd_Producao
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
        Logger.log(`✅ Sucesso absoluto: ${novasLinhasParaInserir.length} contratos gravados com precisão posicional.`);
      } else {
        Logger.log(`⚠️ Nenhuma linha nova gravada (verifique se os contratos já existiam).`);
      }

      Drive.Files.remove(arquivoTemp.id);
      
      // Move o ficheiro para a pasta de usados
      arquivo.moveTo(pastaUsados);
      arquivosProcessados++;
    }

    return { sucesso: true, mensagem: "Importação concluída com sucesso!" };

  } catch (e) {
    Logger.log(`❌ ERRO: ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}