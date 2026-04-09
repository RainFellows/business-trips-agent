/**
 * # MASTER CHECKLIST: Migrace a vylepšení Cestovních příkazů
 *
 * ## STATUS: [READY_TO_DEPLOY]
 *
 * ## 1. Migrace do kontejneru
 * - [x] Odstranit hardcoded SPREADSHEET_ID.
 * - [x] Implementovat `SpreadsheetApp.getActiveSpreadsheet()`.
 * - [x] Nastavit automatickou detekci časového pásma.
 *
 * ## 2. Core Logika (Refactoring)
 * - [x] Přepsat `var` na `const`/`let` (ES6 standard).
 * - [x] Zapouzdřit logiku do konfiguračního objektu `CONFIG`.
 * - [x] Oddělit logiku získávání dat z kalendáře od zápisu do tabulky.
 * - [x] Zachovat logiku párování cest (Odjezd -> Příjezd).
 *
 * ## 3. Notifikace
 * - [x] Implementovat `MailApp.sendEmail`.
 * - [x] Generovat dynamický odkaz na vytvořený list (gid).
 * - [x] Získat email aktuálního uživatele jako příjemce (fallback).
 *
 * ## 4. Bezpečnost a Error Handling
 * - [x] Přidat `try-catch` blok pro hlavní exekuci.
 * - [x] Validovat existenci kalendáře.
 * - [x] Ošetřit případ, kdy nejsou nalezena žádná data (aby nechodil prázdný email, nebo chodil info email).
 *
 * ## POZNÁMKY K NASAZENÍ
 * 1. Otevři editor skriptů v Google Sheets.
 * 2. Vlož obsah `Code.gs`.
 * 3. Nastav Trigger: Funkce `main_generovatCestovniPrikazy` -> Time-driven -> Monthly -> 1st day -> Midnight.
 */
function _readme() {}