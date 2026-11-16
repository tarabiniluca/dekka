// ================================================================
// triggers_additions.gs — AGGIUNTE A triggers.gs del gestionale
//
// ISTRUZIONI:
//   Dopo `clasp pull`, apri triggers.gs e aggiungi la funzione
//   setupTriggerRiallineamento() in coda al file.  Poi `clasp push`.
//   Esegui setupTriggerRiallineamento() una sola volta dallo script
//   editor (non serve un trigger per questa funzione stessa).
// ================================================================

/**
 * Installa il trigger che esegue riallineaDoppioni() ogni 5 minuti.
 * Da eseguire una sola volta dopo il deploy.
 * Se chiamata di nuovo rimuove il trigger precedente ed evita duplicati.
 */
function setupTriggerRiallineamento() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'riallineaDoppioni')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('riallineaDoppioni')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Trigger riallineaDoppioni installato (ogni 5 min).');
}
