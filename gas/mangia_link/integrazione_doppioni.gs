// ================================================================
// integrazione_doppioni.gs — AGGIUNTE agli script Mangia Link
//
// File IDENTICO per tutti i clienti (differisce solo CLIENT_NAME
// già definito in constants.gs di ciascun Mangia Link).
//
// ISTRUZIONI DI DEPLOY (per ogni Mangia Link):
//   1. clasp pull nel progetto Mangia Link del cliente
//   2. Copia questo file nella directory clasp del Mangia Link
//   3. Aggiungi GESTIONALE_SPREADSHEET_ID in constants.gs
//      (vedi sotto)
//   4. Nel codice del processamento link (dove oggi viene
//      registrato il link / scraping del prezzo), aggiungi
//      DOPO la logica esistente:
//
//        const reg = registraInGestionale_(urlAnnuncio);
//        // Facoltativo: puoi loggare reg.isDoppio per debug
//
//   5. Se lo script ha un onEdit che cancella/sostituisce un link,
//      aggiungi PRIMA della sostituzione:
//
//        rimuoviDaGestionale_(vecchioUrl);
//
//   6. clasp push
//
// AGGIUNTA A constants.gs DEL MANGIA LINK:
//   const GESTIONALE_SPREADSHEET_ID = 'ID_DEL_GESTIONALE';
//   (stesso ID usato in constants_additions.gs del gestionale)
// ================================================================

/**
 * Registra un link nel registro doppioni del gestionale.
 * Da chiamare dopo che il Mangia Link ha validato e processato il link.
 *
 * @param  {string} url  URL grezzo dell'annuncio
 * @return {Object|null} { isDoppio, priorityClient, priorityTs } oppure null
 */
function registraInGestionale_(url) {
  if (!url) return null;
  try {
    const gestionale = SpreadsheetApp.openById(GESTIONALE_SPREADSHEET_ID);
    const reg = gestionale.getSheetByName('DEKKA_Registro');
    if (!reg) {
      Logger.log('registraInGestionale_: scheda DEKKA_Registro non trovata nel gestionale.');
      return null;
    }

    const codice = normalizzaCodiceAnnuncioML_(url);
    if (!codice) return { isDoppio: false };

    const timestamp = new Date().toISOString();
    const data = reg.getDataRange().getValues();

    let rigaEsistente = -1;
    let priorityClient = null;
    let priorityTs = null;

    for (let i = 1; i < data.length; i++) {
      const [cn, cod, , ts] = data[i];
      if (cod !== codice) continue;
      if (cn === CLIENT_NAME) {
        rigaEsistente = i + 1;
      } else {
        if (!priorityTs || ts < priorityTs) {
          priorityClient = cn;
          priorityTs = ts;
        }
      }
    }

    if (rigaEsistente > 0) {
      reg.getRange(rigaEsistente, 1, 1, 4).setValues([[CLIENT_NAME, codice, url, timestamp]]);
    } else {
      reg.appendRow([CLIENT_NAME, codice, url, timestamp]);
    }

    return { isDoppio: !!priorityClient, priorityClient, priorityTs };
  } catch (e) {
    Logger.log('registraInGestionale_: ' + e.message);
    return null;
  }
}

/**
 * Rimuove dal registro la voce corrispondente a un URL.
 * Da chiamare quando un link viene cancellato o sostituito
 * nel Mangia Link (es. all'inizio dell'onEdit sul vecchio valore).
 *
 * @param {string} url  URL da rimuovere
 */
function rimuoviDaGestionale_(url) {
  if (!url) return;
  try {
    const gestionale = SpreadsheetApp.openById(GESTIONALE_SPREADSHEET_ID);
    const reg = gestionale.getSheetByName('DEKKA_Registro');
    if (!reg) return;

    const codice = normalizzaCodiceAnnuncioML_(url);
    if (!codice) return;

    const data = reg.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === CLIENT_NAME && data[i][1] === codice) {
        reg.deleteRow(i + 1);
      }
    }
  } catch (e) {
    Logger.log('rimuoviDaGestionale_: ' + e.message);
  }
}

// -----------------------------------------------------------------
// Copia locale di normalizzaCodiceAnnuncio per non dipendere dal
// gestionale (il Mangia Link è uno script separato).
// Deve restare sincronizzata con la versione in registro_doppioni.gs.
// -----------------------------------------------------------------
function normalizzaCodiceAnnuncioML_(url) {
  if (!url || typeof url !== 'string') return null;
  url = url.trim().split('?')[0].replace(/\/+$/, '');

  let m = url.match(/paypay(?:fleamarket)?\.(?:yahoo\.co\.jp|me)\/(?:.*\/)?items?\/(z[A-Za-z0-9]+)/i);
  if (m) return 'ppf_' + m[1].toLowerCase();

  m = url.match(/mercari\.com\/(?:.*\/)?item\/(m[A-Za-z0-9]+)/i);
  if (m) return 'mer_' + m[1].toLowerCase();

  m = url.match(/snkrdunk\.com\/(?:.*\/)?apparels\/([^/?#]+)\/used\/([A-Za-z0-9]+)/i);
  if (m) return 'snkr_' + m[1].toLowerCase() + '_' + m[2].toLowerCase();

  if (url.match(/snkrdunk\.com\/(?:.*\/)?apparels\/[^/?#]+\/?$/i)) return null;

  return null;
}
