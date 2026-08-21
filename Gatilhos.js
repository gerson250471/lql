/**
 * Cria os gatilhos diários para rodar as automações do sistema.
 * Execute esta função apenas UMA VEZ manualmente no ambiente de produção.
 */
function criarGatilhosDoSistema() {
  const gatilhos = ScriptApp.getProjectTriggers();
  
  // 1. Limpa os gatilhos antigos para evitar duplicação[cite: 2]
  for (let i = 0; i < gatilhos.length; i++) {
    let nomeFuncao = gatilhos[i].getHandlerFunction();
    if (nomeFuncao === "importarProducaoDoDrive" || nomeFuncao === "atualizarStatusFinanceiroAutomatizado") {
      ScriptApp.deleteTrigger(gatilhos[i]);
    }
  }

  // 2. Cria o gatilho da Produção (Final do expediente)[cite: 2]
  ScriptApp.newTrigger("importarProducaoDoDrive")
    .timeBased()
    .everyDays(1)
    .atHour(18) // Roda entre 18h e 19h[cite: 2]
    .create();

  // 3. Cria o novo gatilho do Financeiro (Madrugada)
  ScriptApp.newTrigger("atualizarStatusFinanceiroAutomatizado")
    .timeBased()
    .everyDays(1)
    .atHour(0) // Roda entre 00h e 01h da manhã, quando ninguém está usando o sistema
    .create();

  Logger.log("✅ Gatilhos configurados com sucesso: Produção (18h) e Financeiro (00h)!");
}