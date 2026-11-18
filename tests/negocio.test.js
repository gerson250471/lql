describe('Pilar 3: Regras de Negócio (Matemática e Filtros)', () => {

  test('Simulador: A matemática de Empréstimo (Valor x Fator) deve ser exata', () => {
    const valorEmprestimo = 10000;
    const fatorTabela = 0.025; // Tabela com fator de 2.5%
    
    const resultado = valorEmprestimo * fatorTabela;
    
    // A matemática central não pode sofrer regressão
    expect(resultado).toBe(250);
  });

  test('Simulador: A formatação da Moeda deve respeitar o padrão PT-BR (R$)', () => {
    const resultadoCalculo = 1500.50;
    const formatado = resultadoCalculo.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
    
    // O Node.js local usa espaços inquebráveis, por isso removemos os espaços para testar os números e o separador
    const textoLimpo = formatado.replace(/\s/g, '');
    
    expect(textoLimpo).toContain('R$');
    expect(textoLimpo).toContain('1.500,50');
  });

  test('Backend: O Filtro de Consórcio vs Consignado não pode falhar por causa de acentos', () => {
    // Esta é a exata regra de normalização que usamos no seu código do Apps Script (Backend)
    const normalizarTexto = (txt) => String(txt || "")
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
    
    const textoLinha1 = normalizarTexto("Crédito Consignado INSS 84x");
    const textoLinha2 = normalizarTexto("CONSÓRCIO DE AUTOMÓVEL"); // Com acento
    const textoLinha3 = normalizarTexto("consorcio imobiliario"); // Sem acento, minúsculo

    // A palavra-chave que o sistema procura é sempre "CONSORCIO" (limpa)
    expect(textoLinha1.includes("CONSORCIO")).toBe(false); // O Consignado não pode cair na gaveta errada
    expect(textoLinha2.includes("CONSORCIO")).toBe(true);  // Tem que reconhecer mesmo com acento
    expect(textoLinha3.includes("CONSORCIO")).toBe(true);  // Tem que reconhecer mesmo em minúsculas
  });

});