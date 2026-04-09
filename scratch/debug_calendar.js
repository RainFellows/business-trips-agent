function debug_checkEventsOnMarch20() {
  const start = new Date(2026, 2, 20, 0, 0, 0); // 20. březen 2026
  const end = new Date(2026, 2, 21, 0, 0, 0);
  
  const events = CalendarApp.getDefaultCalendar().getEvents(start, end);
  Logger.log(`Nalezeno událostí pro 20. 3.: ${events.length}`);
  
  events.forEach(ev => {
    Logger.log(`--- Událost: ${ev.getTitle()} ---`);
    Logger.log(`Čas: ${ev.getStartTime()} - ${ev.getEndTime()}`);
    Logger.log(`ID: ${ev.getId()}`);
    Logger.log(`Tvůrci: ${ev.getCreators().join(', ')}`);
    
    const guests = ev.getGuestList(true);
    if (guests.length > 0) {
      Logger.log(`Hosté: ${guests.map(g => g.getEmail()).join(', ')}`);
    } else {
      Logger.log("Hosté: ŽÁDNÍ");
    }
  });
}
