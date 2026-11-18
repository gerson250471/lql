const fs = require('fs');
const path = require('path');

// Aponta para o seu arquivo Index.html na raiz do projeto
const htmlPath = path.resolve(__dirname, '../Index.html');
const htmlContent = fs.readFileSync(htmlPath, 'utf8');

describe('Pilar 1: Testes de Estrutura e Interface (Prevenção de Apagões)', () => {

  test('O arquivo Index.html deve existir e conter o escopo completo', () => {
    expect(htmlContent).toBeDefined();
    // Garante que o arquivo não foi acidentalmente esvaziado ou cortado pela metade
    expect(htmlContent.length).toBeGreaterThan(10000); 
  });

  test('Os Modais de segurança e operação DEVEM estar presentes', () => {
    const modaisObrigatorios = [
      'id="modalTrocaSenha"',
      'id="modalFinanceiro"',
      'id="modalNovaProposta"',
      'id="modalEsqueciSenha"',
      'id="modalLead"'
    ];

    modaisObrigatorios.forEach(modal => {
      // Se algum destes faltar, o Jest cancela o deploy imediatamente
      expect(htmlContent).toContain(modal); 
    });
  });

  test('Os contêineres e tabelas de injeção de dados DEVEM existir no DOM', () => {
    const tabelasObrigatorias = [
      'id="tabelaAdminProducaoBody"',
      'id="tabelaFinanceiroBody"',
      'id="commCont"',
      'id="tabelaPropostasBody"',
      'id="containerLeads"'
    ];

    tabelasObrigatorias.forEach(tabela => {
      expect(htmlContent).toContain(tabela);
    });
  });

  test('As funções vitais de JS não podem ser substituídas por placeholders vazios', () => {
    const funcoesObrigatorias = [
      'function efetuarLogin()',
      'function iniciarSessao(',
      'function logout()',
      'function carregarPainelAdmin()',
      'function salvarLancamentoFinanceiro()'
    ];

    funcoesObrigatorias.forEach(func => {
      expect(htmlContent).toContain(func);
    });
  });

});