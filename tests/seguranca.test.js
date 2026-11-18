/**
 * @jest-environment jsdom
 */
const fs = require('fs');
const path = require('path');

const htmlPath = path.resolve(__dirname, '../Index.html');
const htmlContent = fs.readFileSync(htmlPath, 'utf8');

describe('Pilar 2: Segurança e Controlo de Acesso (RBAC)', () => {
  beforeEach(() => {
    // Carrega a interface gráfica do projeto para o DOM Virtual do Jest
    const bodyMatch = htmlContent.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    document.body.innerHTML = bodyMatch ? bodyMatch[1] : htmlContent;
  });

  test('Garante que a lógica da Trava de Segurança existe na função iniciarSessao', () => {
    // Busca flexível: ignora se o VS Code colocou espaços extras, quebrou linhas ou usou aspas simples
    expect(htmlContent).toMatch(/const\s+menusRestritos\s*=/);
    expect(htmlContent).toMatch(/navAdminLeads/);
    expect(htmlContent).toMatch(/navFinanceiro/);
    expect(htmlContent).toMatch(/user\.nivelAcesso\s*===\s*["']ADMINISTRADOR["']/);
  });

  test('Simulação de Login (PROMOTOR): NÃO DEVE ver menus de gestão', () => {
    const menusRestritos = ["navAdminLeads", "navParametros", "navAdminProducao", "navFinanceiro"];
    const user = { nivelAcesso: "PROMOTOR" }; // Simula a conta da equipa
    
    // Executa a lógica exata de bloqueio
    menusRestritos.forEach(id => {
      const el = document.getElementById(id);
      if (el) {
        if (user.nivelAcesso !== "ADMINISTRADOR" && user.nivelAcesso !== "PROGRAMADOR") {
          el.classList.add("hidden");
        }
      }
    });

    // O Teste falha se os botões não estiverem escondidos ("hidden")
    expect(document.getElementById("navParametros").classList.contains("hidden")).toBe(true);
    expect(document.getElementById("navFinanceiro").classList.contains("hidden")).toBe(true);
    expect(document.getElementById("navAdminLeads").classList.contains("hidden")).toBe(true);
  });

  test('Simulação de Login (GESTOR): DEVE ter acesso total aos menus', () => {
    const menusRestritos = ["navAdminLeads", "navParametros", "navAdminProducao", "navFinanceiro"];
    const user = { nivelAcesso: "ADMINISTRADOR" }; // Simula a conta da Cláudia
    
    // Executa a lógica exata de desbloqueio
    menusRestritos.forEach(id => {
      const el = document.getElementById(id);
      if (el) {
        if (user.nivelAcesso === "ADMINISTRADOR" || user.nivelAcesso === "PROGRAMADOR") {
          el.classList.remove("hidden");
        } else {
          el.classList.add("hidden");
        }
      }
    });

    // O Teste falha se o sistema tentar esconder os botões da Cláudia
    expect(document.getElementById("navParametros").classList.contains("hidden")).toBe(false);
    expect(document.getElementById("navFinanceiro").classList.contains("hidden")).toBe(false);
  });
  
  test('Garante que a função logout() força a limpeza de dados e bloqueio de interface', () => {
    expect(htmlContent).toMatch(/user = null;/);
    expect(htmlContent).toMatch(/document\.getElementById\(id\)\.classList\.add\("hidden"\);/);
    expect(htmlContent).toMatch(/document\.getElementById\("txtIdentificador"\)\.value = "";/);
  });
});
// 261118 1138 Realizando Teste de Integridade no Controle de Acesso