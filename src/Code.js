/**
 * @version 2.2.0
 * @author Antigravity AI
 * @description Automatizace cestovních příkazů z Kalendáře do Tabulky.
 * Fix 2.2.0: Oprava filtrace (vlastní události) a rozsahu dat (zahrnutí posledního dne).
 */

// --- KONFIGURACE ---
const CONFIG = {
  DOMOVSKE_MESTO: "Ostrava",
  HLEDANY_TEXT: "vlakem", 
  HODINY_BUFFER: 1,       
  EMAIL_PREDMET: "Cestovní report připraven: ",
  EMAIL_PRIJEMCE: Session.getActiveUser().getEmail(), 
  SHEET_HEADER: ["Popis cesty", "Odjezd (Datum a čas)", "Příjezd (Datum a čas)", "Destinace"],
  COLORS: {
    HEADER_BG: "#4c1130", // Tmavě vínová (profi vzhled)
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
  
  // První den minulého měsíce 00:00:00
  const prvniDenMinuleho = new Date(dnes.getFullYear(), dnes.getMonth() - 1, 1);
  
  // První den AKTUÁLNÍHO měsíce 00:00:00 
  // To zajistí, že getEvents(start, end) zahrne i události z celého posledního dne minulého měsíce
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
  
  Logger.log(`Nalezeno hrubých událostí: ${udalosti.length}`);

  const cesty = [];
  const zpracovaneIDs = new Set();
  const homeCityLower = CONFIG.DOMOVSKE_MESTO.toLowerCase();

  // Seřadíme události podle času, abychom zajistili správné párování
  udalosti.sort((a, b) => a.getStartTime() - b.getStartTime());

  for (let i = 0; i < udalosti.length; i++) {
    const udalost = udalosti[i];
    
    // FILTRACE: Pouze mé události (kontrola tvůrce)
    //getCreators vrací pole emailů, které událost vytvořily (u sdílených kalendářů tam bývá ten, kdo ji tam dal)
    const creatory = udalost.getCreators();
    if (!creatory.includes(myEmail)) {
      Logger.log(`Přeskakuji událost - nejsem tvůrce: ${udalost.getTitle()} (Vytvořil: ${creatory.join(', ')})`);
      continue;
    }

    const id = udalost.getId();
    const titleLower = udalost.getTitle().toLowerCase();
    
    if (zpracovaneIDs.has(id)) continue;

    if (titleLower.includes(homeCityLower)) {
      const jeOdjezd = titleLower.includes("z " + homeCityLower);
      const jePrijezd = titleLower.includes("do " + homeCityLower);

      if (jeOdjezd) {
        zpracovatOdjezd(udalost, titleLower, cesty, zpracovaneIDs);
      } else if (jePrijezd) {
        zpracovatPrijezd(udalost, titleLower, cesty, zpracovaneIDs);
      }
    } else if (titleLower.includes("z ") && titleLower.includes("do ")) {
      zpracovatCestuMeziMesty(udalost, titleLower, cesty, zpracovaneIDs);
    }
  }

  return cesty;
}

function zpracovatOdjezd(udalost, title, cestyRef, idsRef) {
  const zacatek = udalost.getStartTime();
  const startCesty = new Date(zacatek.getTime() - CONFIG.HODINY_BUFFER * 3600000);
  
  let cil = title.match(/do ([^,]+)/i);
  cil = cil ? cil[1].trim() : "Neznámá destinace";

  cestyRef.push({
    typ: `Cesta do: ${cil}`, // Upravený popis pro lepší čitelnost
    start: startCesty,       // Ukládáme jako Date object pro formátování v Sheets
    konec: "",
    cil: cil,
    jeDoma: false,           // Příznak, že jsme na cestě
    udalostId: udalost.getId()
  });
  
  idsRef.add(udalost.getId());
}

function zpracovatPrijezd(udalost, title, cestyRef, idsRef) {
  const konec = udalost.getEndTime();
  const konecCesty = new Date(konec.getTime() + CONFIG.HODINY_BUFFER * 3600000);
  
  let startMesto = title.match(/z ([^,]+)/i);
  startMesto = startMesto ? startMesto[1].trim() : "";

  // INTELIGENTNÍ PÁROVÁNÍ
  // 1. Zkusíme najít poslední otevřenou cestu (kde není vyplněný konec)
  // 2. Nebudeme striktně porovnávat město (protože "Praha hl.n." != "Praha")
  // 3. Předpokládáme, že pokud mám otevřenou cestu z domova, tento příjezd domů k ní patří.
  
  let sparovano = false;
  
  // Procházíme odzadu, hledáme poslední nedokončenou cestu
  for (let j = cestyRef.length - 1; j >= 0; j--) {
    const cesta = cestyRef[j];
    
    // Podmínka párování: Je to cesta pryč a ještě nemá konec (návrat)
    if (cesta.konec === "" && !cesta.jeDoma) {
      cesta.konec = konecCesty; // Doplníme čas návratu
      cesta.typ = `Cesta: ${CONFIG.DOMOVSKE_MESTO} -> ${cesta.cil} a zpět`; // Aktualizujeme název
      cesta.jeDoma = true; // Uzavřeme cestu
      
      idsRef.add(udalost.getId());
      sparovano = true;
      Logger.log(`Spárováno fuzzy logikou: ${cesta.cil} <-> Návrat domů`);
      break;
    }
  }

  // Pokud se nespárovalo (např. chybí odjezd v kalendáři), zapíšeme samostatně
  if (!sparovano) {
    cestyRef.push({
      typ: "Příjezd domů (chybí odjezd)",
      start: "",
      konec: konecCesty,
      cil: CONFIG.DOMOVSKE_MESTO,
      jeDoma: true,
      udalostId: udalost.getId()
    });
    idsRef.add(udalost.getId());
  }
}

function zpracovatCestuMeziMesty(udalost, title, cestyRef, idsRef) {
  let startMesto = title.match(/z ([^,]+)/i);
  startMesto = startMesto ? startMesto[1].trim() : "";
  let cilMesto = title.match(/do ([^,]+)/i);
  cilMesto = cilMesto ? cilMesto[1].trim() : "";

  cestyRef.push({
    typ: "Cesta vlakem (mimo domov)",
    start: udalost.getStartTime(),
    konec: udalost.getEndTime(),
    cil: `${startMesto} -> ${cilMesto}`,
    jeDoma: false,
    udalostId: udalost.getId()
  });
  
  idsRef.add(udalost.getId());
}

function zapsatDoTabulky(ss, data, nazevZaklad) {
  const nazevListu = `${nazevZaklad} - Vlaky`;
  let list = ss.getSheetByName(nazevListu);

  // Pokud list existuje, smažeme ho a vytvoříme nový pro čistý start (nebo smažeme obsah)
  // Zde volím smazání obsahu pro zachování ID listu, pokud by na něj vedly odkazy
  if (list) {
    list.clear();
  } else {
    list = ss.insertSheet(nazevListu);
  }

  // 1. Zápis hlavičky
  list.appendRow(CONFIG.SHEET_HEADER);
  const headerRange = list.getRange(1, 1, 1, CONFIG.SHEET_HEADER.length);
  
  // 2. Příprava dat
  const rows = data.map(c => [c.typ, c.start, c.konec, c.cil]);

  // 3. Hromadný zápis dat
  if (rows.length > 0) {
    const dataRange = list.getRange(2, 1, rows.length, rows[0].length);
    dataRange.setValues(rows);

    // --- FORMÁTOVÁNÍ (DESIGN) ---
    
    // A) Ohraničení pro celou tabulku
    const fullTableRange = list.getRange(1, 1, rows.length + 1, CONFIG.SHEET_HEADER.length);
    fullTableRange.setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    
    // B) Hlavička - Styl
    headerRange.setFontWeight("bold")
      .setBackground(CONFIG.COLORS.HEADER_BG)
      .setFontColor(CONFIG.COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center");

    // C) Formátování datumu a času (Sloupec 2 a 3)
    // Formát: den.měsíc.rok hodina:minuta (např. 01.12.2025 14:30)
    list.getRange(2, 2, rows.length, 2)
      .setNumberFormat("dd.MM.yyyy HH:mm")
      .setHorizontalAlignment("center"); // Časy na střed

    // D) Zarovnání textů (Sloupec 1 a 4)
    list.getRange(2, 1, rows.length, 1).setHorizontalAlignment("left");
    list.getRange(2, 4, rows.length, 1).setHorizontalAlignment("left");

    // E) Střídavé barvy řádků (Zebra striping) - pro lepší čitelnost
    // Pokud je hodně řádků, uděláme to jednoduše přes loop nebo conditional formatting,
    // ale nejčistší je applyRowBanding (pokud ho metoda podporuje, jinak manuálně)
    // GAS má applyRowBanding na range.
    if (fullTableRange.getBandings().length === 0) {
       fullTableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    }

    // F) Auto-resize
    list.autoResizeColumns(1, 4);
    
    // Trochu rozšířit sloupce pro vzdušnost (optional)
    const currentWidth1 = list.getColumnWidth(1);
    if (currentWidth1 < 200) list.setColumnWidth(1, 250); // Sloupec Popis
  }

  Logger.log(`Zapsáno a naformátováno ${rows.length} řádků.`);
  
  return { listUrl: `${ss.getUrl()}#gid=${list.getSheetId()}` };
}

function odeslatNotifikaci(url, mesic, pocetCest) {
  MailApp.sendEmail({
    to: CONFIG.EMAIL_PRIJEMCE,
    subject: `${CONFIG.EMAIL_PREDMET}${mesic}`,
    htmlBody: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; color: #333;">
        <h2 style="color: ${CONFIG.COLORS.HEADER_BG};">🚄 Cestovní report: ${mesic}</h2>
        <p>Report byl úspěšně vygenerován a naformátován.</p>
        <table style="border-collapse: collapse; width: 100%; max-width: 400px; margin-bottom: 20px;">
          <tr style="background-color: #f8f9fa;">
            <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>Počet cest:</strong></td>
            <td style="padding: 10px; border-bottom: 1px solid #ddd;">${pocetCest}</td>
          </tr>
        </table>
        <a href="${url}" style="background-color: ${CONFIG.COLORS.HEADER_BG}; color: white; padding: 12px 25px; text-decoration: none; border-radius: 4px; display: inline-block;">
          Otevřít tabulku
        </a>
      </div>
    `
  });
}

/**
 * POMOCNÁ FUNKCE PRO TESTOVÁNÍ (DRY RUN)
 * Spusťte tuto funkci v Apps Script editoru pro ověření fixu bez zápisu do ostré tabulky.
 */
function dev_dryRunTest() {
  Logger.log("--- ZAČÁTEK TESTU (DRY RUN) ---");
  const perioda = ziskatObdobiMinulehoMesice();
  Logger.log(`Testované období: ${perioda.start.toLocaleString()} - ${perioda.end.toLocaleString()}`);
  
  const cesty = ziskatCestyZKalendare(perioda.start, perioda.end);
  
  Logger.log(`VÝSLEDEK: Nalezeno ${cesty.length} cest k zapsání.`);
  cesty.forEach((c, i) => {
    Logger.log(`${i+1}. [${c.typ}] | Cíl: ${c.cil} | Odjezd: ${c.start.toLocaleString()} | Příjezd: ${c.konec ? c.konec.toLocaleString() : "???"}`);
  });
  
  Logger.log("--- KONEC TESTU ---");
}