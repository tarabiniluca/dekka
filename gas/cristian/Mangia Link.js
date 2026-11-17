// ═══════════════════════════════════════════════════════════════
// DEKKA — Mangia Link · Cristian Bertolini
// ═══════════════════════════════════════════════════════════════

// ─── REGEX PIATTAFORME SUPPORTATE ─────────────────────────────
// SNKRDUNK: accetta sia l'annuncio usato (/apparels/ID/used/ID)
// sia la pagina di catalogo (/apparels/ID), per cui il prezzo
// viene inserito manualmente dal cliente.
const LINK_REGEX = /^https:\/\/(paypayfleamarket\.yahoo\.co\.jp\/item\/[a-z0-9]+|snkrdunk\.com\/[a-z]+\/\d+(\/[a-z]+\/\d+)?|jp\.mercari\.com\/item\/m[0-9]+)$/i;

// Link di catalogo SNKRDUNK (senza annuncio specifico)
const SNKRDUNK_CATALOG_REGEX = /^https:\/\/snkrdunk\.com\/[a-z]+\/\d+$/i;

// ─── CONFIGURAZIONE ───────────────────────────────────────────
const GESTIONALE_ID = "11epk8d4GAnNcTOcL8xOushWKewge5oWNcrgESXarOHw";
const CLIENT_NAME   = "Cristian Bertolini";
const NOTIFY_EMAIL  = "tarabini.luca@gmail.com";

// ─── SPESE SPEDIZIONE SNKRDUNK ────────────────────────────────
const SHIPPING_SINGLE_CARD = 300;  // Spese per carta singola
const SHIPPING_OTHER       = 990;  // Spese per altri prodotti

// ─── CONFIGURAZIONE RETRY (riduce i fallimenti di recupero) ───
const FETCH_MAX_RETRIES    = 3;
const FETCH_RETRY_DELAY_MS = 800; // base delay, raddoppia ad ogni tentativo

const FETCH_HEADERS = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
  "Accept-Language": "ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7"
};

// ─── COLONNE schede mese ──────────────────────────────────────
// B=STATO | C=SPEDIZIONE | D=LINK | E=Q.TA | F=IMPORTI
const COL_STATO   = 2; // B
const COL_LINK    = 4; // D
const COL_QTA     = 5; // E
const COL_IMPORTI = 6; // F
// ─────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────
// Trigger principale — smista tra invio link e correzione manuale prezzo
// ─────────────────────────────────────────────────────────────
function onEditInstallable(e) {
  if (!e || !e.range) return;

  const sheet     = e.range.getSheet();
  const sheetName = sheet.getName();

  const months = [
    "Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
    "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"
  ];

  if (sheetName === "MANGIA LINK") {
    handleLinkSubmission_(e);
    return;
  }

  if (months.includes(sheetName)) {
    handleManualPriceEdit_(e, sheet);
    return;
  }
}

// ─────────────────────────────────────────────────────────────
// Verifica se il link è una pagina di catalogo SNKRDUNK
// (es. https://snkrdunk.com/apparels/881423): nessun annuncio
// specifico, quindi nessun prezzo recuperabile in automatico.
// ─────────────────────────────────────────────────────────────
function isSnkrdunkCatalogLink_(link) {
  return SNKRDUNK_CATALOG_REGEX.test(link);
}

// ─────────────────────────────────────────────────────────────
// Gestione invio link da MANGIA LINK
// ─────────────────────────────────────────────────────────────
function handleLinkSubmission_(e) {

  const inputSheet = "MANGIA LINK";
  const inputCell  = "B2";
  const startRow   = 3;

  const ss    = e.source;
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== inputSheet) return;
  if (e.range.getA1Notation() !== inputCell) return;

  const link = e.range.getValue();
  if (!link) return;

  const cache    = CacheService.getScriptCache();
  const lockKey  = "editLock";
  const lastLink = cache.get(lockKey);

  if (lastLink === link) return;

  cache.put(lockKey, link, 60);

  const ui = SpreadsheetApp.getUi();

  // 1. Valida il link
  if (!LINK_REGEX.test(link)) {
    ui.alert("Link non valido.\n\nPiattaforme supportate:\n• PayPay フリマ\n• SNKRDUNK\n• Mercari");
    e.range.clearContent();
    cache.remove(lockKey);
    return;
  }

  const isCatalog = isSnkrdunkCatalogLink_(link);

  const months = [
    "Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
    "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"
  ];

  const monthIndex  = parseInt(Utilities.formatDate(new Date(), "Asia/Tokyo", "M")) - 1;
  const monthName   = months[monthIndex];
  const targetSheet = ss.getSheetByName(monthName);
  if (!targetSheet) {
    cache.remove(lockKey);
    return;
  }

  const totalRows    = targetSheet.getLastRow();
  const columnValues = totalRows >= startRow
    ? targetSheet.getRange(startRow, COL_LINK, totalRows - startRow + 1, 1).getValues()
    : [];

  // 2. Controlla duplicati
  let lastUsedRow = startRow - 1;
  for (let i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] === link) {
      ui.alert("Link già esistente.");
      e.range.clearContent();
      cache.remove(lockKey);
      return;
    }
    if (columnValues[i][0] !== "") lastUsedRow = startRow + i;
  }

  // 2b. Piano Trial: blocca se scaduto o se il massimale (100 link) è esaurito
  const trialCheck = checkTrialLimit_();
  if (trialCheck.blocked) {
    ui.alert("Invio non disponibile", trialCheck.message, ui.ButtonSet.OK);
    e.range.clearContent();
    cache.remove(lockKey);
    return;
  }

  // 3. Chiede conferma (messaggio dedicato per i link di catalogo)
  const confirmMessage = isCatalog
    ? "Questo è un link di catalogo SNKRDUNK: il prezzo non viene recuperato automaticamente " +
      "e dovrà essere inserito a mano nella colonna IMPORTI.\n\n" +
      "Inserisci il solo prezzo del prodotto, senza spese di spedizione. " +
      "L'importo inserito verrà poi aggiornato dall'amministratore aggiungendo le spese di spedizione " +
      "SNKRDUNK (¥" + SHIPPING_SINGLE_CARD + " per carta singola, ¥" + SHIPPING_OTHER + " per tutto il resto): " +
      "l'importo definitivo comparirà nella stessa cella.\n\n" +
      "Vuoi inserire questo link?"
    : "Vuoi inserire questo link?";

  const response = ui.alert(
    "Conferma inserimento",
    confirmMessage,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    e.range.clearContent();
    cache.remove(lockKey);
    return;
  }

  const row = lastUsedRow < startRow ? startRow : lastUsedRow + 1;

  // 4. Recupera il prezzo (link catalogo: nessun prezzo, inserimento manuale)
  const productData = isCatalog
    ? buildSnkrdunkCatalogData_(link)
    : fetchProductData(link);
  const price = productData?.price ?? null;

  // 5. Scrive link, prezzo e stato iniziale sul foglio cliente
  targetSheet.getRange(row, COL_STATO).setValue("NUOVO");
  targetSheet.getRange(row, COL_LINK).setValue(link);
  targetSheet.getRange(row, COL_IMPORTI).setValue(price !== null ? price : "N/D");

  if (isCatalog) {
    targetSheet.getRange(row, COL_IMPORTI)
      .setNote("Link di catalogo SNKRDUNK: inserire il prezzo del prodotto (senza spese di spedizione). " +
               "L'amministratore aggiornerà l'importo aggiungendo le spese di spedizione.");
  }

  // 6. Push al gestionale
  pushToGestionale_(link, price, monthName);

  // 6b. Registra nel controllo doppioni (asincrono — errori solo in log)
  registraInGestionale_(link);

  // 7. Invia notifica email
  sendOrderNotification_(link, productData, monthName);

  // 8. Pulizia finale
  e.range.clearContent();
  cache.remove(lockKey);
}

// ─────────────────────────────────────────────────────────────
// PIANO TRIAL — limite non sforabile (pagina "Piani e Tariffe")
//
// Legge dalla Dashboard del gestionale il piano del cliente e la
// scadenza (colonna SCADENZA PIANO, formato "gg/mm/aaaa ..."): la
// finestra del piano va da scadenza − 30 giorni a scadenza, come nel
// gestionale. Se la scadenza non è impostata, si usa il mese corrente.
//
// Conta i link inviati nella finestra in qualsiasi stato tranne quelli
// annullati (CANCELLATO, RESO, SOLD OUT, SMARRITO): anche i link NUOVO
// contano, altrimenti il cliente potrebbe inviarne più di 100 prima
// che vengano lavorati. Ogni link vale ceil(Q.TA / 25) unità
// (Q.TA vuota = 1), come nel gestionale.
//
// Blocca l'invio se:
//   • il Trial è scaduto (30 giorni trascorsi), oppure
//   • le unità già usate hanno raggiunto il massimale.
// Per gli altri piani non blocca mai (l'eccedenza è consentita).
// In caso di errore di lettura non blocca (fail-open) e registra il log.
// ─────────────────────────────────────────────────────────────
const TRIAL_KEY       = "TRIAL";
const TRIAL_MASSIMALE = 100;

function checkTrialLimit_() {
  const ok = { blocked: false, message: "" };

  try {
    const gestSS    = SpreadsheetApp.openById(GESTIONALE_ID);
    const dashboard = gestSS.getSheetByName("Dashboard");
    const gestSheet = gestSS.getSheetByName(CLIENT_NAME);
    if (!dashboard || !gestSheet) return ok;

    const COL_DASH_CLIENTE  = 2;  // B
    const COL_DASH_PIANO    = 3;  // C
    const COL_DASH_SCADENZA = 10; // J

    // Riga del cliente in Dashboard (per nome)
    const dashLastRow = dashboard.getLastRow();
    if (dashLastRow < 2) return ok;
    const righe = dashboard.getRange(2, COL_DASH_CLIENTE, dashLastRow - 1,
                                     COL_DASH_SCADENZA - COL_DASH_CLIENTE + 1).getValues();
    const riga = righe.find(r => r[0] === CLIENT_NAME);
    if (!riga) return ok;

    const piano = (riga[COL_DASH_PIANO - COL_DASH_CLIENTE] || "").toString().trim().toUpperCase();
    if (piano !== TRIAL_KEY) return ok;

    // Finestra del piano
    const oggi = new Date();
    let inizio = null, fine = null;
    const scadStr = (riga[COL_DASH_SCADENZA - COL_DASH_CLIENTE] || "").toString();
    const m = scadStr.match(/^(\d{2})\/(\d{2})\/(\d{4})/);

    if (m) {
      fine = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 23, 59, 59, 999);
      inizio = new Date(fine);
      inizio.setDate(inizio.getDate() - 30);
      inizio.setHours(0, 0, 0, 0);

      if (oggi > fine) {
        return {
          blocked: true,
          message: "Il periodo Trial è terminato il " + m[1] + "/" + m[2] + "/" + m[3] + ".\n\n" +
                   "Per continuare a inviare link è necessario attivare uno dei piani mensili " +
                   "(Starter, Basic o Pro). Contatta DEKKA per l'attivazione."
        };
      }
    } else {
      // Scadenza non impostata: fallback al mese corrente
      inizio = new Date(oggi.getFullYear(), oggi.getMonth(), 1, 0, 0, 0, 0);
      fine   = new Date(oggi.getFullYear(), oggi.getMonth() + 1, 0, 23, 59, 59, 999);
    }

    // Unità già utilizzate nella finestra
    const statiAnnullati = ["CANCELLATO", "RESO", "SOLD OUT", "SMARRITO"];
    let unitaUsate = 0;

    const lastRow = gestSheet.getLastRow();
    if (lastRow >= 3) {
      const valori = gestSheet.getRange(3, 2, lastRow - 2, 5).getValues(); // B:F
      for (const r of valori) {
        const data  = r[0];               // B = DATA RICEZIONE
        const stato = r[1];               // C = STATO
        const link  = r[3];               // E = LINK
        const qta   = r[4];               // F = Q.TA
        if (!data || !link) continue;
        if (statiAnnullati.includes(stato)) continue;
        const d = new Date(data);
        if (isNaN(d.getTime()) || d < inizio || d > fine) continue;
        unitaUsate += Math.ceil((parseInt(qta) || 1) / 25);
      }
    }

    Logger.log("Trial check — " + CLIENT_NAME + ": " + unitaUsate + " / " + TRIAL_MASSIMALE);

    if (unitaUsate >= TRIAL_MASSIMALE) {
      return {
        blocked: true,
        message: "Hai raggiunto il massimale del piano Trial (" + TRIAL_MASSIMALE + " link).\n\n" +
                 "Il Trial non è sforabile: per inviare altri link è necessario attivare uno dei " +
                 "piani mensili (Starter, Basic o Pro). Contatta DEKKA per l'attivazione."
      };
    }

    return ok;

  } catch (err) {
    Logger.log("Errore controllo limite Trial (invio consentito): " + err.message);
    return ok;
  }
}

// ─────────────────────────────────────────────────────────────
// Dati per un link di catalogo SNKRDUNK: nessun prezzo
// automatico. Legge solo il nome dal <title> (best effort,
// utile per l'email di notifica).
// ─────────────────────────────────────────────────────────────
function buildSnkrdunkCatalogData_(link) {
  let name = "N/D";

  try {
    const fetchResult = fetchWithRetry_(link);
    if (fetchResult.success && fetchResult.html) {
      const titleMatch = fetchResult.html.match(/<title[^>]*>([^<]+)<\/title>/);
      if (titleMatch) {
        name = titleMatch[1]
          .replace(/通販・買取・相場.*$/, "")
          .replace(/｜スニダン\s*$/, "")
          .trim() || "N/D";
      }
    }
  } catch (err) {
    Logger.log("SNKRDUNK catalogo - lettura nome fallita: " + err.message);
  }

  Logger.log("SNKRDUNK catalogo - prezzo da inserire manualmente per: " + link);

  return {
    price:     null,
    name:      name,
    basePrice: null,
    platform:  "SNKRDUNK (catalogo)",
    catalog:   true
  };
}

// ─────────────────────────────────────────────────────────────
// Gestione correzione manuale prezzo su scheda mese
// Sincronizza col Gestionale quando il cliente inserisce a mano
// un prezzo (tipicamente per correggere un N/D o per i link
// di catalogo SNKRDUNK)
// ─────────────────────────────────────────────────────────────
function handleManualPriceEdit_(e, sheet) {
  const startRow = 3;
  const col      = e.range.getColumn();
  const row      = e.range.getRow();
  const numRows  = e.range.getNumRows();

  // Interessa solo la colonna IMPORTI, righe dati (non intestazione)
  if (col !== COL_IMPORTI || row < startRow) return;

  for (let i = 0; i < numRows; i++) {
    const currentRow = row + i;

    const link = sheet.getRange(currentRow, COL_LINK).getValue();
    if (!link) continue; // riga senza link, niente da sincronizzare

    const priceCell = sheet.getRange(currentRow, COL_IMPORTI);
    const newValue  = priceCell.getValue();

    // Ignora celle svuotate o ancora testuali tipo "N/D"
    if (newValue === "" || newValue === null) continue;
    if (typeof newValue !== "number") continue;

    // Nota sulla cella del cliente: segnala che è stato inserito a mano
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "dd/MM/yyyy HH:mm");
    priceCell.setNote("Prezzo inserito manualmente il " + now);

    pushManualPriceToGestionale_(link, newValue);
  }
}

// ─────────────────────────────────────────────────────────────
// Cerca il link nella scheda cliente del Gestionale e aggiorna
// la colonna IMPORTI con il prezzo inserito manualmente
// ─────────────────────────────────────────────────────────────
function pushManualPriceToGestionale_(link, newPrice) {
  try {
    const gestSS    = SpreadsheetApp.openById(GESTIONALE_ID);
    const gestSheet = gestSS.getSheetByName(CLIENT_NAME);
    if (!gestSheet) {
      Logger.log("Scheda cliente non trovata nel gestionale: " + CLIENT_NAME);
      return;
    }

    const lastRow = gestSheet.getLastRow();
    if (lastRow < 3) return;

    const links = gestSheet.getRange(3, 5, lastRow - 2, 1).getValues(); // colonna E = LINK

    for (let i = 0; i < links.length; i++) {
      if (links[i][0] === link) {
        const gestRow   = i + 3;
        const gestCell  = gestSheet.getRange(gestRow, 7); // colonna G = IMPORTI
        const now       = Utilities.formatDate(new Date(), "Asia/Tokyo", "dd/MM/yyyy HH:mm");

        gestCell.setValue(newPrice);
        gestCell.setNote("Prezzo inserito manualmente dal cliente il " + now);

        Logger.log("Prezzo manuale sincronizzato al gestionale — riga " + gestRow + " — valore: ¥" + newPrice);
        return;
      }
    }

    Logger.log("Link non trovato nel gestionale per sincronizzazione manuale: " + link);

  } catch (err) {
    Logger.log("Errore sincronizzazione prezzo manuale al gestionale: " + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Fetch con retry ed header realistici
// ─────────────────────────────────────────────────────────────
function fetchWithRetry_(url) {
  let lastError = null;
  let lastCode = null;

  for (let attempt = 1; attempt <= FETCH_MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: FETCH_HEADERS
      });

      const code = response.getResponseCode();
      lastCode = code;

      if (code === 200) {
        return { success: true, code: code, html: response.getContentText() };
      }

      // 403/429/5xx: probabile blocco temporaneo o rate limit → ritenta
      if ([403, 429, 500, 502, 503, 504].includes(code) && attempt < FETCH_MAX_RETRIES) {
        Utilities.sleep(FETCH_RETRY_DELAY_MS * attempt);
        continue;
      }

      return { success: false, code: code, html: response.getContentText() };

    } catch (e) {
      lastError = e.message;
      if (attempt < FETCH_MAX_RETRIES) {
        Utilities.sleep(FETCH_RETRY_DELAY_MS * attempt);
        continue;
      }
    }
  }

  return { success: false, code: lastCode, html: null, error: lastError };
}

// ─────────────────────────────────────────────────────────────
// Estrazione prezzo con fallback multipli
// ─────────────────────────────────────────────────────────────
function fetchProductData(link) {
  const fetchResult = fetchWithRetry_(link);

  if (!fetchResult.success || !fetchResult.html) {
    Logger.log("Fetch fallito per: " + link + " (code: " + fetchResult.code + ")");
    return null;
  }

  const html = fetchResult.html;

  // ── Branch Mercari: legge meta tag product:price:amount ──
  if (link.includes("jp.mercari.com")) {
    const priceMatch = html.match(/<meta\s+name="product:price:amount"\s+content="([0-9]+)"/);
    const nameMatch  = html.match(/<title[^>]*>([^<]+)<\/title>/);

    if (priceMatch) {
      return {
        price:     parseInt(priceMatch[1], 10),
        name:      nameMatch ? nameMatch[1].replace(/ - メルカリ$/, "").trim() : "N/D",
        basePrice: parseInt(priceMatch[1], 10),
        platform:  "Mercari"
      };
    }
    Logger.log("Parsing Mercari fallito per: " + link);
    return null;
  }

  // ── Branch SNKRDUNK (annunci usati): estrazione dedicata ──
  if (link.includes("snkrdunk.com")) {
    return fetchSnkrdunkProductData_(html, link);
  }

  // ── Branch PayPay フリマ: legge JSON-LD ──────
  const jsonMatches = html.match(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/g);

  if (jsonMatches) {
    for (const match of jsonMatches) {
      try {
        const jsonContent = match.replace(/<script type="application\/ld\+json">|<\/script>/g, '').trim();
        const jsonData = JSON.parse(jsonContent);

        let product = null;
        if (Array.isArray(jsonData)) {
          product = jsonData.find(item => item["@type"] === "Product");
        } else if (jsonData["@type"] === "Product") {
          product = jsonData;
        } else if (jsonData["@graph"]) {
          product = jsonData["@graph"].find(item => item["@type"] === "Product");
        }

        if (product && product.offers) {
          const basePrice = product.offers.price ?? product.offers[0]?.price ?? null;

          if (basePrice !== null) {
            return {
              price:     basePrice,
              name:      product.name || "N/D",
              basePrice: basePrice,
              platform:  "PayPay フリマ"
            };
          }
        }
      } catch (parseErr) {
        continue;
      }
    }
  }

  // ── Fallback generico: meta tag Open Graph / itemprop ──
  const ogPriceMatch = html.match(/<meta\s+property="og:price:amount"\s+content="([0-9]+)"/) ||
                       html.match(/<meta\s+itemprop="price"\s+content="([0-9]+)"/);

  if (ogPriceMatch) {
    const basePrice = parseInt(ogPriceMatch[1], 10);

    Logger.log("Prezzo recuperato via fallback OG per: " + link);

    return {
      price:     basePrice,
      name:      "N/D (fallback OG)",
      basePrice: basePrice,
      platform:  "PayPay フリマ"
    };
  }

  Logger.log("Nessun prezzo trovato per: " + link);
  return null;
}

// ─────────────────────────────────────────────────────────────
// Estrazione dedicata SNKRDUNK (annunci usati)
//
// Dal redesign Next.js del sito (2025), il prezzo non è più
// affidabilmente presente in un blocco JSON-LD nel body: viene
// letto invece dal meta description SEO, sempre presente lato
// server nel formato "...¥120,000で買える...".
// JSON-LD e meta OG restano come fallback, nel caso in cui
// SNKRDUNK ripristini quella struttura in futuro.
// ─────────────────────────────────────────────────────────────
function fetchSnkrdunkProductData_(html, link) {
  let basePrice = null;

  // 1. Metodo primario: meta description SEO
  const descPriceMatch = html.match(/¥([\d,]+)で買える/);
  if (descPriceMatch) {
    basePrice = parseInt(descPriceMatch[1].replace(/,/g, ''), 10);
  }

  // 2. Fallback: vecchio metodo JSON-LD (nel caso torni disponibile)
  if (basePrice === null) {
    basePrice = extractPriceFromJsonLd_(html);
    if (basePrice !== null) {
      Logger.log("SNKRDUNK - prezzo recuperato via JSON-LD (fallback) per: " + link);
    }
  }

  // 3. Fallback finale: meta Open Graph / itemprop generico
  if (basePrice === null) {
    const ogPriceMatch = html.match(/<meta\s+property="og:price:amount"\s+content="([0-9]+)"/) ||
                         html.match(/<meta\s+itemprop="price"\s+content="([0-9]+)"/);
    if (ogPriceMatch) {
      basePrice = parseInt(ogPriceMatch[1], 10);
      Logger.log("SNKRDUNK - prezzo recuperato via fallback OG per: " + link);
    }
  }

  if (basePrice === null) {
    Logger.log("Nessun prezzo trovato per SNKRDUNK: " + link);
    return null;
  }

  // 4. Nome prodotto, letto dal tag <title>
  const titleMatch = html.match(/<title[^>]*>([^<]+)<\/title>/);
  const titleText  = titleMatch ? titleMatch[1] : "";
  const name = titleText
    ? titleText
        .replace(/\s*\d+枚のシングルトレカ通販.*$/, "")
        .replace(/｜スニダン\s*$/, "")
        .trim()
    : "N/D";

  // 5. Determina se è una single card e applica la spedizione corretta
  const isSingleCard = detectSingleCardSNKRDUNK_(titleText, html);
  const shippingCost = isSingleCard ? SHIPPING_SINGLE_CARD : SHIPPING_OTHER;
  const finalPrice   = basePrice + shippingCost;

  Logger.log("SNKRDUNK - Base: ¥" + basePrice + " | Single Card: " + isSingleCard + " | Shipping: ¥" + shippingCost + " | Total: ¥" + finalPrice);

  return {
    price:     finalPrice,
    name:      name || "N/D",
    basePrice: basePrice,
    platform:  "SNKRDUNK"
  };
}

// ─────────────────────────────────────────────────────────────
// Estrae il prezzo da un eventuale blocco JSON-LD nella pagina
// (usato come fallback per SNKRDUNK)
// ─────────────────────────────────────────────────────────────
function extractPriceFromJsonLd_(html) {
  const jsonMatches = html.match(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/g);
  if (!jsonMatches) return null;

  for (const match of jsonMatches) {
    try {
      const jsonContent = match.replace(/<script type="application\/ld\+json">|<\/script>/g, '').trim();
      const jsonData = JSON.parse(jsonContent);

      let product = null;
      if (Array.isArray(jsonData)) {
        product = jsonData.find(item => item["@type"] === "Product");
      } else if (jsonData["@type"] === "Product") {
        product = jsonData;
      } else if (jsonData["@graph"]) {
        product = jsonData["@graph"].find(item => item["@type"] === "Product");
      }

      if (product && product.offers) {
        const basePrice = product.offers.price ?? product.offers[0]?.price ?? null;
        if (basePrice !== null) return parseFloat(basePrice);
      }
    } catch (parseErr) {
      continue;
    }
  }

  return null;
}

// ─────────────────────────────────────────────────────────────
// Determina se un prodotto SNKRDUNK è una single card
//
// Applica la logica quantità 1 + categoria single card leggendo
// il tag <title> della pagina, con fallback sull'intero HTML.
// Se in futuro SNKRDUNK cambia il formato del <title>, questa
// funzione andrà rivista.
// ─────────────────────────────────────────────────────────────
function detectSingleCardSNKRDUNK_(titleText, html) {
  const hasSize1Mai = /1枚/.test(titleText) || /1枚/.test(html);

  const hasSingleCardCategory = /シングル(トレカ|カード)/.test(titleText) ||
                                 /Single Card|シングルカード/i.test(html);

  const isSingleCard = hasSize1Mai && hasSingleCardCategory;

  Logger.log("Single Card Detection - Size:1枚=" + hasSize1Mai + " | Category:SingleCard=" + hasSingleCardCategory + " | Result=" + isSingleCard);

  return isSingleCard;
}

// ─────────────────────────────────────────────────────────────
// Invia email di notifica per nuovo ordine (CON THREADING)
// ─────────────────────────────────────────────────────────────
function sendOrderNotification_(link, productData, monthName) {
  try {
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "dd/MM/yyyy HH:mm");

    let orderInfo =
      "\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      "Data/Ora: " + now + "\n" +
      "Mese: " + monthName + "\n";

    if (productData) {
      orderInfo += "Piattaforma: " + productData.platform + "\n";
      orderInfo += "Prodotto: " + productData.name + "\n";

      if (productData.catalog) {
        orderInfo += "Importo: da inserire manualmente dal cliente (link di catalogo)\n";
        orderInfo += "Azione admin: dopo l'inserimento, aggiornare l'importo nel gestionale " +
                     "aggiungendo le spese di spedizione (¥" + SHIPPING_SINGLE_CARD + " carta singola, ¥" +
                     SHIPPING_OTHER + " altri prodotti)\n";
      } else {
        if (productData.platform === "SNKRDUNK" && productData.basePrice !== null) {
          orderInfo += "Prezzo base: ¥" + formatNumber_(productData.basePrice) + "\n";
          const shipping = productData.price - productData.basePrice;
          orderInfo += "Spese spedizione: ¥" + formatNumber_(shipping) + "\n";
        }

        orderInfo += "Importo totale: ¥" + formatNumber_(productData.price) + "\n";
      }
    } else {
      orderInfo += "Importo: N/D\n";
    }

    orderInfo += "Link: " + link + "\n";

    // PropertiesService per memorizzare ID della prima email
    const props = PropertiesService.getScriptProperties();
    const messageIdKey = "DEKKA_MSG_ID_" + CLIENT_NAME.replace(/\s+/g, "_").toUpperCase();
    const storedMessageId = props.getProperty(messageIdKey);

    if (storedMessageId) {
      // Email esiste già — MODIFICA quella esistente
      Logger.log("DEBUG - Modificando messaggio esistente: " + storedMessageId);

      try {
        // Leggi il messaggio corrente
        const message = Gmail.Users.Messages.get("me", storedMessageId, { format: "raw" });
        const rawMessage = Utilities.newBlob(Utilities.base64Decode(message.raw)).getDataAsString();

        // Trova dove finisce l'header (doppia newline)
        const headerEnd = rawMessage.indexOf("\n\n");
        const headers = rawMessage.substring(0, headerEnd);
        let body = rawMessage.substring(headerEnd + 2);

        // Aggiungi il nuovo ordine al corpo
        body = body + orderInfo;

        // Ricrea l'email completa
        const updatedRawEmail = headers + "\n\n" + body;
        const encodedEmail = Utilities.base64Encode(updatedRawEmail, Utilities.Charset.UTF_8);

        // Elimina il vecchio messaggio
        Gmail.Users.Messages.trash("me", storedMessageId);
        Logger.log("✓ Messaggio vecchio eliminato");

        // Invia il messaggio aggiornato
        const newMessage = Gmail.Users.Messages.send(
          { raw: encodedEmail },
          "me"
        );

        // Aggiorna l'ID salvato
        props.setProperty(messageIdKey, newMessage.id);

        Logger.log("✓ Messaggio aggiornato - nuovo ID: " + newMessage.id);
        Logger.log("✓ Ordine aggiunto al messaggio unico");

      } catch (modifyErr) {
        // Se la modifica fallisce, invia un nuovo messaggio e usa quello come nuovo root
        Logger.log("⚠ Impossibile modificare, inviando nuovo messaggio: " + modifyErr.message);
        sendNewOrderEmail_(link, productData, monthName, orderInfo, messageIdKey, props);
      }

    } else {
      // Nessun messaggio salvato — CREA il primo
      Logger.log("DEBUG - Creando primo messaggio");
      sendNewOrderEmail_(link, productData, monthName, orderInfo, messageIdKey, props);
    }

  } catch (err) {
    Logger.log("✗ Errore principale: " + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Helper: invia nuovo messaggio e salva l'ID
// ─────────────────────────────────────────────────────────────
function sendNewOrderEmail_(link, productData, monthName, orderInfo, messageIdKey, props) {
  try {
    const subject = "DEKKA - Ordini da " + CLIENT_NAME;

    let body = "Ordini ricevuti da " + CLIENT_NAME + "\n\n";
    body += "═══════════════════════════════════════════════════\n";
    body += "RIEPILOGO ORDINI\n";
    body += "═══════════════════════════════════════════════════";
    body += orderInfo;
    body += "\n═══════════════════════════════════════════════════\n";

    const rawEmail =
      "To: " + NOTIFY_EMAIL + "\n" +
      "Subject: " + subject + "\n" +
      "Content-Type: text/plain; charset=UTF-8\n" +
      "\n" +
      body;

    const encodedEmail = Utilities.base64Encode(rawEmail, Utilities.Charset.UTF_8);

    const sentMessage = Gmail.Users.Messages.send(
      { raw: encodedEmail },
      "me"
    );

    // Salva l'ID per future modifiche
    props.setProperty(messageIdKey, sentMessage.id);

    Logger.log("✓ Primo messaggio creato - ID: " + sentMessage.id);
    Logger.log("✓ Ordine aggiunto");

    // Aggiungi label
    Utilities.sleep(500);

    const labelName = "DEKKA/" + CLIENT_NAME;
    const labels = Gmail.Users.Labels.list("me").labels || [];
    let labelId = null;

    for (let i = 0; i < labels.length; i++) {
      if (labels[i].name === labelName) {
        labelId = labels[i].id;
        break;
      }
    }

    if (!labelId) {
      const newLabel = Gmail.Users.Labels.create(
        {
          name: labelName,
          labelListVisibility: "labelShow",
          messageListVisibility: "show"
        },
        "me"
      );
      labelId = newLabel.id;
    }

    Gmail.Users.Messages.modify(
      { addLabelIds: [labelId] },
      "me",
      sentMessage.id
    );

    Logger.log("✓ Label applicato: " + labelName);

  } catch (err) {
    Logger.log("✗ Errore creazione messaggio: " + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Formatta numero con separatore migliaia
// ─────────────────────────────────────────────────────────────
function formatNumber_(num) {
  if (num === null || num === undefined) return "N/D";
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// ─────────────────────────────────────────────────────────────
// Scrive il nuovo ordine sulla scheda cliente del gestionale
// B=DATA RIC | C=STATO | D=SPED | E=LINK | F=Q.TA | G=IMPORTI
// ─────────────────────────────────────────────────────────────
function pushToGestionale_(link, price, monthName) {
  try {
    const gestSS    = SpreadsheetApp.openById(GESTIONALE_ID);
    const gestSheet = gestSS.getSheetByName(CLIENT_NAME);
    if (!gestSheet) {
      Logger.log("Scheda cliente non trovata nel gestionale: " + CLIENT_NAME);
      return;
    }

    const lastRow  = gestSheet.getLastRow();
    const newRow   = lastRow < 3 ? 3 : lastRow + 1;
    const now      = new Date();
    const dataCell = gestSheet.getRange(newRow, 2);

    dataCell.setValue(now);
    dataCell.setNumberFormat("dd/mm/yyyy");
    gestSheet.getRange(3, 2, newRow - 2, 1).setNumberFormat("dd/mm/yyyy");
    gestSheet.getRange(newRow, 3).setValue("NUOVO");          // C = STATO
    gestSheet.getRange(newRow, 5).setValue(link);             // E = LINK

    // F (Q.TA) lasciata vuota — inserita manualmente nel gestionale
    gestSheet.getRange(newRow, 7).setValue(                   // G = IMPORTI
      price !== null ? price : "N/D"
    );

    Logger.log("Ordine pushato al gestionale — riga " + newRow);

    refreshDashboardOnGestionale_(gestSS);

  } catch (err) {
    Logger.log("Errore push gestionale: " + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Aggiorna la Dashboard del gestionale dopo l'invio di un link
//
// Aggiorna SOLO la colonna ORDINI ATTIVI della riga di questo
// cliente (abbinamento per NOME, non per posizione).
//
// PIANO, UTILIZZO PIANO e massimale NON vengono toccati: sono
// calcolati esclusivamente dal gestionale (refreshDashboard_),
// che conosce piano selezionato, massimali ufficiali e finestra
// temporale del piano. Un nuovo link entra in stato NUOVO, che
// non consuma il piano: l'utilizzo cambia solo quando lo stato
// passa a PAGATO/RICEVUTO/SPEDITO nel gestionale, e in quel
// momento è il gestionale stesso ad aggiornare Dashboard e client.
// ─────────────────────────────────────────────────────────────
function refreshDashboardOnGestionale_(gestSS) {
  try {
    const dashboard = gestSS.getSheetByName("Dashboard");
    const gestSheet = gestSS.getSheetByName(CLIENT_NAME);
    if (!dashboard || !gestSheet) return;

    const COL_DASH_CLIENTE = 2; // B
    const COL_DASH_ATTIVI  = 5; // E

    // Riga del cliente in Dashboard (per nome)
    const dashLastRow = dashboard.getLastRow();
    if (dashLastRow < 2) return;

    const nomi = dashboard.getRange(2, COL_DASH_CLIENTE, dashLastRow - 1, 1).getValues();
    let dashRow = null;
    for (let i = 0; i < nomi.length; i++) {
      if (nomi[i][0] === CLIENT_NAME) { dashRow = i + 2; break; }
    }
    if (!dashRow) {
      Logger.log("Cliente non presente in Dashboard: " + CLIENT_NAME + " — nessun aggiornamento");
      return;
    }

    // Ordini attivi = data presente, stato non vuoto, non annullato e non spedito
    // (stessa logica di refreshDashboard_ nel gestionale)
    const statiNonAttivi = ["SPEDITO", "CANCELLATO", "RESO", "SOLD OUT", "SMARRITO"];
    let ordiniAttivi = 0;

    const lastRow = gestSheet.getLastRow();
    if (lastRow >= 3) {
      const valori = gestSheet.getRange(3, 2, lastRow - 2, 2).getValues(); // B:C
      ordiniAttivi = valori.filter(r =>
        r[0] !== "" && r[0] !== null && r[0] !== false &&
        r[1] !== "" && r[1] !== null && r[1] !== false &&
        !statiNonAttivi.includes(r[1])
      ).length;
    }

    dashboard.getRange(dashRow, COL_DASH_ATTIVI).setValue(ordiniAttivi);
    Logger.log("Ordini attivi aggiornati in Dashboard — " + CLIENT_NAME + ": " + ordiniAttivi);

  } catch (err) {
    Logger.log("Errore refresh dashboard: " + err.message);
  }
}