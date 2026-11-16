// ================================================================
// integrazione_doppioni.js 窶・DEKKA Mangia Link ﾂｷ Dennys Marzola
// Registra e rimuove link nel registro doppioni del gestionale.
//
// Usa GESTIONALE_ID giﾃ definito in questo script (stessa costante
// usata per le altre operazioni sul gestionale).
// ================================================================

/**
 * Registra un link nel registro doppioni del gestionale.
 * Chiamare dopo pushToGestionale_().
 *
 * @param  {string} url  URL grezzo dell'annuncio
 */
function registraInGestionale_(url) {
  if (!url) return;
  try {
    const gestionale = SpreadsheetApp.openById(GESTIONALE_ID);
    const reg = gestionale.getSheetByName('DEKKA_Registro');
    if (!reg) {
      Logger.log('registraInGestionale_: scheda DEKKA_Registro non trovata nel gestionale.');
      return;
    }

    const codice = normalizzaCodiceAnnuncioML_(url);
    if (!codice) return;

    const timestamp = new Date().toISOString();
    const data = reg.getDataRange().getValues();

    let rigaEsistente = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === CLIENT_NAME && data[i][1] === codice) {
        rigaEsistente = i + 1;
        break;
      }
    }

    if (rigaEsistente > 0) {
      reg.getRange(rigaEsistente, 1, 1, 4).setValues([[CLIENT_NAME, codice, url, timestamp]]);
    } else {
      reg.appendRow([CLIENT_NAME, codice, url, timestamp]);
    }

    Logger.log('registraInGestionale_: registrato codice=' + codice + ' client=' + CLIENT_NAME);
  } catch (e) {
    Logger.log('registraInGestionale_: ' + e.message);
  }
}

/**
 * Rimuove dal registro la voce corrispondente a un URL.
 * Chiamare prima di rimuovere o sostituire un link.
 *
 * @param {string} url  URL da rimuovere
 */
function rimuoviDaGestionale_(url) {
  if (!url) return;
  try {
    const gestionale = SpreadsheetApp.openById(GESTIONALE_ID);
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
// gestionale (il Mangia Link ﾃｨ uno script separato).
// Deve restare sincronizzata con la versione in registro_doppioni.js.
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

