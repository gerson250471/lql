/**
 * Cria um gatilho diário para rodar a importação automaticamente.
 * Execute esta função apenas UMA VEZ manualmente para registrar o gatilho.
 */
function criarGatilhoDiario() {
  // Remove gatilhos anteriores com o mesmo nome para evitar duplicações
  const gatilhos = ScriptApp.getProjectTriggers();
  for (let i = 0; i < gatilhos.length; i++) {
    if (gatilhos[i].getHandlerFunction() === "importarProducaoDoDrive") {
      ScriptApp.deleteTrigger(gatilhos[i]);
    }
  }

  // Cria um novo gatilho configurado para rodar todos os dias
  // Exemplo: Entre 18h e 19h (hora do final do expediente)
  ScriptApp.newTrigger("importarProducaoDoDrive")
    .timeBased()
    .everyDays(1)
    .atHour(18) // Altere a hora desejada aqui (0 a 23)
    .create();

  Logger.log("✅ Gatilho diário configurado com sucesso para rodar às 18h!");
}