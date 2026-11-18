const fs = require('fs');
const path = require('path');

describe('Pilar 5: Sintaxe e Integridade do Código (Prevenção de Erros Fatais)', () => {
  
  test('O JavaScript do Index.html DEVE ter a sintaxe perfeita', () => {
    const htmlPath = path.resolve(__dirname, '../Index.html');
    const htmlContent = fs.readFileSync(htmlPath, 'utf8');
    const scriptMatch = htmlContent.match(/<script>([\s\S]*?)<\/script>/i);
    
    expect(scriptMatch).not.toBeNull(); 
    
    const codigoJS = scriptMatch[1];
    expect(() => {
      new Function(codigoJS);
    }).not.toThrow(SyntaxError);
  });

});

describe('Pilar 5.1: Varredura Automática do Backend (Google Apps Script)', () => {
  
  const rootPath = path.resolve(__dirname, '../');
  const todosOsArquivos = fs.readdirSync(rootPath);
  
  // Filtra apenas os ficheiros .js que vão para a Google, ignorando ficheiros de configuração do Node
  const ficheirosBackend = todosOsArquivos.filter(arquivo => {
    return arquivo.endsWith('.js') && 
           arquivo !== 'atualizar-versao.js' && 
           arquivo !== 'jest.config.js';
  });

  // Cria um teste automatizado para CADA ficheiro .js encontrado na sua pasta
  ficheirosBackend.forEach(ficheiro => {
    test(`O ficheiro "${ficheiro}" NÃO PODE ter erros de sintaxe (chaves, vírgulas, textos soltos)`, () => {
      const filePath = path.join(rootPath, ficheiro);
      const codigo = fs.readFileSync(filePath, 'utf8');
      
      expect(() => {
        new Function(codigo);
      }).not.toThrow(SyntaxError);
    });
  });

});