// ================================================================
// constants_additions.gs — AGGIUNTE A constants.gs del gestionale
//
// ISTRUZIONI:
//   Dopo `clasp pull`, apri constants.gs e aggiungi le righe
//   qui sotto in coda al file.  Poi `clasp push`.
//
// Se REG_SHEET_NAME, NOTE_PREFIX, BG_PRIORITA e BG_NO_ACQUISTO
// sono già definiti in registro_doppioni.gs puoi saltare questo
// file — sono già lì come costanti locali.  Aggiungile qui solo
// se vuoi centralizzare tutte le costanti in constants.gs.
// ================================================================

// ID del gestionale — usato dai Mangia Link per scrivere nel registro
// via SpreadsheetApp.openById().  Copialo dalla barra degli indirizzi
// del gestionale (la stringa lunga tra /d/ e /edit).
const GESTIONALE_SPREADSHEET_ID = 'INSERISCI_QUI_L_ID_DEL_GESTIONALE';
