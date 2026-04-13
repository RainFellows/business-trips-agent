/**
 * @version 2.4.0
 * @author Antigravity AI
 * @description Automatizace cestovních příkazů z Kalendáře do Tabulky.
 * Feature 2.4.0: Standardizace měst, oddělení dopravy do sloupce, vylepšená filtrace.
 */

// --- KONFIGURACE ---
const CONFIG = {
  DOMOVSKE_MESTO: "Ostrava",
  HLEDANY_TEXT: "vlakem OR autem", 
  HODINY_BUFFER: 1,       
  EMAIL_PREDMET: "Cestovní report připraven: ",
  EMAIL_PRIJEMCE: Session.getActiveUser().getEmail(), 
  SHEET_HEADER: ["Popis cesty", "Odjezd", "Příjezd", "Destinace", "Doprava", "km autem", "Klient"],
  IGNOROVANE_DOMENY: ["rainfellows.cz", "gmail.com", "seznam.cz", "outlook.com", "email.cz", "milovkynekrasy.cz", "rfconsultants.cz", "resource.calendar.google.com"],
  COLORS: {
    HEADER_BG: "#4c1130",
    HEADER_TEXT: "#ffffff",
    ROW_BANDING: "#f3f3f3",
    BORDER: "#000000"
  }
};

// --- HLAVNÍ SPOUŠTĚCÍ FUNKCE ---

function main_generovatCestovniPrikazy() {
  Logger.log("--- ZAČÁTEK SKRIPTU ---");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Skript neběží v kontejneru tabulky.");

    const perioda = ziskatObdobiMinulehoMesice();
    Logger.log(`Zpracovávám období: ${perioda.nazevListu}`);

    const cesty = ziskatCestyZKalendare(perioda.start, perioda.end);
    
    if (cesty.length === 0) {
      Logger.log("Nebyly nalezeny žádné cesty.");
      return; 
    }

    const vysledekZapisu = zapsatDoTabulky(ss, cesty, perioda.nazevListu);
    odeslatNotifikaci(vysledekZapisu.listUrl, perioda.nazevListu, cesty.length);

  } catch (e) {
    Logger.log("FATAL ERROR: " + e.message);
    MailApp.sendEmail(CONFIG.EMAIL_PRIJEMCE, "CHYBA: Generování cestovních příkazů", e.message);
  }
  
  Logger.log("--- KONEC SKRIPTU ---");
}

// --- LOGIKA ---

function ziskatObdobiMinulehoMesice() {
  const dnes = new Date();
  const prvniDenMinuleho = new Date(dnes.getFullYear(), dnes.getMonth() - 1, 1);
  const prvniDenTohoto = new Date(dnes.getFullYear(), dnes.getMonth(), 1);
  
  let nazevMesice = prvniDenMinuleho.toLocaleString('cs-CZ', { month: 'long' });
  nazevMesice = nazevMesice.charAt(0).toUpperCase() + nazevMesice.slice(1);

  return {
    start: prvniDenMinuleho,
    end: prvniDenTohoto,
    nazevListu: nazevMesice
  };
}

function ziskatCestyZKalendare(start, end) {
  const calendar = CalendarApp.getDefaultCalendar();
  const udalosti = calendar.getEvents(start, end, { q: CONFIG.HLEDANY_TEXT });
  const myEmail = Session.getActiveUser().getEmail();
  
  const cesty = [];
  const zpracovaneIDs = new Set();
  const homeCityLower = CONFIG.DOMOVSKE_MESTO.toLowerCase();

  udalosti.sort((a, b) => a.getStartTime() - b.getStartTime());

  for (let i = 0; i < udalosti.length; i++) {
    const udalost = udalosti[i];
    const creatory = udalost.getCreators();
    
    if (!creatory.includes(myEmail)) continue;

    const id = udalost.getId();
    const titleLower = udalost.getTitle().toLowerCase();
    const jeAuto = titleLower.includes("autem");
    
    if (zpracovaneIDs.has(id)) continue;

    if (titleLower.includes(homeCityLower)) {
      const jeOdjezd = titleLower.includes("z " + homeCityLower);
      const jePrijezd = titleLower.includes("do " + homeCityLower);

      if (jeOdjezd) {
        zpracovatOdjezd(udalost, titleLower, cesty, zpracovaneIDs, jeAuto);
      } else if (jePrijezd) {
        zpracovatPrijezd(udalost, titleLower, cesty, zpracovaneIDs, jeAuto);
      }
    } else if (titleLower.includes("z ") && titleLower.includes("do ")) {
      zpracovatCestuMeziMesty(udalost, titleLower, cesty, zpracovaneIDs, jeAuto);
    }
  }

  return cesty;
}

function zpracovatOdjezd(udalost, title, cestyRef, idsRef, jeAuto) {
  const zacatek = udalost.getStartTime();
  const buffer = jeAuto ? 0 : CONFIG.HODINY_BUFFER;
  const startCesty = new Date(zacatek.getTime() - buffer * 3600000);
  
  let cil = title.match(/do ([^,]+)/i);
  cil = cil ? formatovatMesto(cil[1]) : "Neznámá destinace";

  const km = jeAuto ? ziskatKm(CONFIG.DOMOVSKE_MESTO, cil) : "";
  const klient = ziskatKlientaZPrekryvu(udalost.getStartTime(), udalost.getEndTime(), udalost.getId());

  cestyRef.push({
    typ: `${cil}`,
    doprava: jeAuto ? 'Auto' : 'Vlak',
    start: startCesty,
    konec: "",
    cil: cil,
    jeDoma: false,
    jeAuto: jeAuto,
    km: km,
    klient: klient,
    udalostId: udalost.getId()
  });
  
  idsRef.add(udalost.getId());
}

function zpracovatPrijezd(udalost, title, cestyRef, idsRef, jeAuto) {
  const konec = udalost.getEndTime();
  const buffer = jeAuto ? 0 : CONFIG.HODINY_BUFFER;
  const konecCesty = new Date(konec.getTime() + buffer * 3600000);
  
  let sparovano = false;
  
  for (let j = cestyRef.length - 1; j >= 0; j--) {
    const cesta = cestyRef[j];
    
    if (cesta.konec === "" && !cesta.jeDoma) {
      cesta.konec = konecCesty; 
      cesta.typ = `${CONFIG.DOMOVSKE_MESTO} -> ${cesta.cil} a zpět`;
      if (jeAuto && cesta.km) cesta.km = cesta.km * 2; 
      
      // HLEDÁME KLIENTA V CELÉ DOBĚ PRACOVNÍ CESTY (od startu do konce)
      cesta.klient = ziskatKlientaZPrekryvu(cesta.start, cesta.konec, "");
      
      cesta.jeDoma = true;
      idsRef.add(udalost.getId());
      sparovano = true;
      break;
    }
  }

  if (!sparovano) {
    const startMesto = title.match(/z ([^,]+)/i);
    const startMestoText = startMesto ? formatovatMesto(startMesto[1]) : "";
    const km = jeAuto ? ziskatKm(startMestoText, CONFIG.DOMOVSKE_MESTO) : "";
    const klient = ziskatKlientaZPrekryvu(udalost.getStartTime(), udalost.getEndTime(), udalost.getId());

    cestyRef.push({
      typ: `Příjezd domů (chybí odjezd)`,
      doprava: jeAuto ? 'Auto' : 'Vlak',
      start: "",
      konec: konecCesty,
      cil: CONFIG.DOMOVSKE_MESTO,
      jeDoma: true,
      jeAuto: jeAuto,
      km: km,
      klient: klient,
      udalostId: udalost.getId()
    });
    idsRef.add(udalost.getId());
  }
}

function zpracovatCestuMeziMesty(udalost, title, cestyRef, idsRef, jeAuto) {
  let startMesto = title.match(/z ([^,]+)/i);
  startMesto = startMesto ? formatovatMesto(startMesto[1]) : "";
  let cilMesto = title.match(/do ([^,]+)/i);
  cilMesto = cilMesto ? formatovatMesto(cilMesto[1]) : "";

  const km = jeAuto ? ziskatKm(startMesto, cilMesto) : "";
  const klient = ziskatKlientaZPrekryvu(udalost.getStartTime(), udalost.getEndTime(), udalost.getId());

  cestyRef.push({
    typ: `${startMesto} -> ${cilMesto}`,
    doprava: jeAuto ? 'Auto' : 'Vlak',
    start: udalost.getStartTime(),
    konec: udalost.getEndTime(),
    cil: `${startMesto} -> ${cilMesto}`,
    jeDoma: false,
    jeAuto: jeAuto,
    km: km,
    klient: klient,
    udalostId: udalost.getId()
  });
  
  idsRef.add(udalost.getId());
}

function zapsatDoTabulky(ss, data, nazevZaklad) {
  const nazevListu = `${nazevZaklad} - Vlaky`;
  let list = ss.getSheetByName(nazevListu);

  if (list) {
    list.clear();
  } else {
    list = ss.insertSheet(nazevListu);
  }

  list.appendRow(CONFIG.SHEET_HEADER);
  const headerRange = list.getRange(1, 1, 1, CONFIG.SHEET_HEADER.length);
  const rows = data.map(c => [c.typ, c.start, c.konec, c.cil, c.doprava, c.km, c.klient]);

  if (rows.length > 0) {
    const dataRange = list.getRange(2, 1, rows.length, rows[0].length);
    dataRange.setValues(rows);

    const fullTableRange = list.getRange(1, 1, rows.length + 1, CONFIG.SHEET_HEADER.length);
    fullTableRange.setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    
    headerRange.setFontWeight("bold")
      .setBackground(CONFIG.COLORS.HEADER_BG)
      .setFontColor(CONFIG.COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");

    list.getRange(2, 2, rows.length, 2)
      .setNumberFormat("dd.MM.yyyy HH:mm")
      .setHorizontalAlignment("center");

    list.getRange(2, 1, rows.length, 1).setHorizontalAlignment("left"); 
    list.getRange(2, 4, rows.length, 1).setHorizontalAlignment("center"); 
    list.getRange(2, 5, rows.length, 1).setHorizontalAlignment("center");
    list.getRange(2, 6, rows.length, 1).setHorizontalAlignment("center");
    list.getRange(2, 7, rows.length, 1).setHorizontalAlignment("left"); 

    if (fullTableRange.getBandings().length === 0) {
       fullTableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    }

    list.autoResizeColumns(1, CONFIG.SHEET_HEADER.length);
    if (list.getColumnWidth(1) < 200) list.setColumnWidth(1, 260); 
    if (list.getColumnWidth(7) < 150) list.setColumnWidth(7, 200); 
  }

  return { listUrl: `${ss.getUrl()}#gid=${list.getSheetId()}` };
}

function odeslatNotifikaci(url, mesic, pocetCest) {
  MailApp.sendEmail({
    to: CONFIG.EMAIL_PRIJEMCE,
    subject: `${CONFIG.EMAIL_PREDMET}${mesic}`,
    htmlBody: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; color: #333;">
        <h2 style="color: ${CONFIG.COLORS.HEADER_BG};">🚄 Cestovní report: ${mesic}</h2>
        <p>Report byl úspěšně vygenerován.</p>
        <a href="${url}" style="background-color: ${CONFIG.COLORS.HEADER_BG}; color: white; padding: 12px 25px; text-decoration: none; border-radius: 4px; display: inline-block;">
          Otevřít tabulku
        </a>
      </div>
    `
  });
}

function dev_dryRunTest() {
  Logger.log("--- ZAČÁTEK TESTU (DRY RUN) ---");
  const perioda = ziskatObdobiMinulehoMesice();
  const cesty = ziskatCestyZKalendare(perioda.start, perioda.end);
  
  cesty.forEach((c, i) => {
    Logger.log(`${i+1}. [${c.typ}] | Doprava: ${c.doprava} | Cíl: ${c.cil} | KM: ${c.km || '-'} | Klient: ${c.klient || '???'}`);
  });
  Logger.log("--- KONEC TESTU ---");
}

function ziskatKm(startMesto, cilMesto) {
  if (!startMesto || !cilMesto || startMesto === cilMesto) return 0;
  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(startMesto).setDestination(cilMesto)
      .setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
    if (directions.routes && directions.routes.length > 0) {
      return Math.ceil(directions.routes[0].legs[0].distance.value / 1000);
    }
  } catch (e) {}
  return 0;
}

function ziskatKlientaZPrekryvu(start, end, transportEventId) {
  if (!start || !end) return "";
  const udalosti = CalendarApp.getDefaultCalendar().getEvents(start, end);
  const domeny = new Set();
  
  udalosti.forEach(ev => {
    // Přeskočíme samotnou cestovní událost
    if (ev.getId() === transportEventId) return;
    
    const titleLower = ev.getTitle().toLowerCase();
    const myEmail = Session.getActiveUser().getEmail().toLowerCase();

    // 1. FILTRACE: Podle klíčových slov v názvu (Zrušeno / Canceled / Declined)
    if (titleLower.includes("zrušeno") || 
        titleLower.includes("zrušená") || 
        titleLower.includes("canceled") || 
        titleLower.includes("cancelled") ||
        titleLower.includes("declined")) {
      return;
    }

    // 2. FILTRACE: Pouze schůzky, které jsem přijal (OWNER je brán jako přijatý, pokud není výslovně NO)
    const myStatus = ev.getMyStatus();
    
    // Pokud je status NO (odmítnuto), ignorujeme vždy (i u OWNER)
    if (myStatus === CalendarApp.GuestStatus.NO) return;
    
    // Extra kontrola pro OWNER: v některých případech getMyStatus() vrací OWNER, i když je v seznamu hostů NO
    const myGuestRecord = ev.getGuestByEmail(myEmail);
    if (myGuestRecord && myGuestRecord.getGuestStatus() === CalendarApp.GuestStatus.NO) {
      return;
    }
    
    // Musí to být buď OWNER, YES nebo MAYBE (pokud chceme být benevolentnější), 
    // ale pro jistotu držíme YES / OWNER
    if (myStatus !== CalendarApp.GuestStatus.YES && myStatus !== CalendarApp.GuestStatus.OWNER) {
      return;
    }
    
    // 3. Získáme hosty (včetně nepotvrzených)
    const guestEmails = ev.getGuestList(true).map(g => g.getEmail().toLowerCase());
    const creators = ev.getCreators().map(c => c.toLowerCase());
    
    let nalezenaNovaDomena = false;
    [...guestEmails, ...creators].forEach(email => {
      const match = email.match(/@([^@]+)$/);
      if (match && !CONFIG.IGNOROVANE_DOMENY.includes(match[1])) {
        if (!domeny.has(match[1])) {
           domeny.add(match[1]);
           nalezenaNovaDomena = true;
        }
      }
    });

    if (nalezenaNovaDomena) {
      Logger.log(`   [Identifikace] Událost: "${ev.getTitle()}" | Čas: ${ev.getStartTime().getHours()}:${ev.getStartTime().getMinutes()}`);
    }
  });
  return Array.from(domeny).join(", ");
}

/**
 * Zformátuje název města: Velké první písmeno, odstranění "hl. n." apod.
 */
function formatovatMesto(text) {
  if (!text) return "";
  let mesto = text.toLowerCase()
    .replace(/ hl\. ?n\./g, "")
    .replace(/ hl n/g, "")
    .replace(/ hl\.n\./g, "")
    .replace(/ nádraží/g, "")
    .replace(/ - centrum/g, "")
    .trim();
  
  if (mesto.length === 0) return text;
  return mesto.charAt(0).toUpperCase() + mesto.slice(1);
}

/**
 * DIAGNOSTIKA: Vypíše detaily o všech událostech v aktuálně zpracovávaném období.
 */
function dev_diagVypisDetailyUdalosti() {
  Logger.log("--- START DIAGNOSTIKY ---");
  const perioda = ziskatObdobiMinulehoMesice();
  const udalosti = CalendarApp.getDefaultCalendar().getEvents(perioda.start, perioda.end);
  
  udalosti.forEach((ev, i) => {
    Logger.log(`${i + 1}. [${ev.getTitle()}] (${ev.getStartTime().toLocaleString()} - ${ev.getEndTime().toLocaleString()})`);
    Logger.log(`   Tvůrci: ${ev.getCreators().join(", ")}`);
    Logger.log(`   Hosté: ${ev.getGuestList(true).map(g => g.getEmail()).join(", ") || "žádní"}`);
    Logger.log("------------------------------------------");
  });
  Logger.log("--- KONEC DIAGNOSTIKY ---");
}
