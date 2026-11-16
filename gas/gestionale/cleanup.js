// ─────────────────────────────────────────────────────────────
// Pulizia automatica righe con anno ricezione > 3 anni fa
// ─────────────────────────────────────────────────────────────
function cleanOldRecords() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sheets      = ss.getSheets();
  const annoLimite  = new Date().getFullYear() - 3;

  for (const sheet of sheets) {
    if (sheet.getName() === "Dashboard") continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    const dates = sheet.getRange(2, G_COL.DATA_RICEZIONE, lastRow - 1, 1).getValues();

    for (let i = dates.length - 1; i >= 0; i--) {
      const d = dates[i][0];
      if (!d) continue;
      if (new Date(d).getFullYear() <= annoLimite) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  Logger.log("Pulizia completata — " + new Date());
}

// Installa trigger mensile — esegui una volta sola
function installCleanupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "cleanOldRecords")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("cleanOldRecords")
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();

  Logger.log("Trigger pulizia mensile installato.");
}