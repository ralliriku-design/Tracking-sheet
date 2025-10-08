/** NF_Menu_Triggers.gs — Yhdistetty, turvallinen menu + ajastimet (additiivinen) */
function NF_onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('NewFlow');

    // Päivittäinen ja viikkoraportit (PR #4:n NF_* -toiminnot)
    if (typeof NF_RunDailyFlow === 'function') {
      menu.addItem('Run daily flow now', 'NF_RunDailyFlow');
    }
    if (typeof NF_BuildWeeklyReports === 'function') {
      menu.addItem('Build SOK & Kärkkäinen (last week)', 'NF_BuildWeeklyReports');
    }
    if (typeof NF_BuildDeliveryTimes === 'function') {
      menu.addItem('Delivery Times list', 'NF_BuildDeliveryTimes');
    }
    if (typeof NF_MakeCountryWeekLeadtime === 'function') {
      menu.addItem('Country Week Leadtime', 'NF_MakeCountryWeekLeadtime');
    }

    // Viikkopalvelutaso (PR #6), lisätään vain jos funktio on olemassa
    if (typeof NF_buildWeeklyServiceLevels === 'function') {
      menu.addItem('Rakenna viikkopalvelutaso (ALL/SOK/KRK)', 'NF_buildWeeklyServiceLevels');
    }

    menu.addSeparator();

    // Ajastimet (asennus/poisto) — varmistetaan että handlerit ovat olemassa
    menu.addItem('Install weekday 12:00', 'NF_setupWeekday1200');
    menu.addItem('Install weekly Mon 02:00', 'NF_setupWeeklyMon0200');
    menu.addItem('Remove NF triggers', 'NF_clearNFTriggers');

    menu.addSeparator();

    // Suorat käynnistyskomennot, jos handlerit ovat olemassa
    if (typeof NF_runDaily === 'function') {
      menu.addItem('Aja: Päivittäinen NewFlow', 'NF_runDaily');
    }
    if (typeof NF_runWeekly === 'function') {
      menu.addItem('Aja: Viikkoraportit', 'NF_runWeekly');
    }

    menu.addToUi();
  } catch (err) {
    Logger.log('NF_onOpen error: ' + err);
  }
}

// Ajastimien asennus — luodaan vain, jos handleri on olemassa
function NF_setupWeekday1200() {
  const ss = SpreadsheetApp.getActive();
  if (typeof NF_runDaily !== 'function') {
    ss.toast('NF_runDaily puuttuu — varmista, että PR #4 NF_Main.gs on asennettu.');
    return;
  }
  NF_clearTriggersBy_(['NF_runDaily']);
  ScriptApp.newTrigger('NF_runDaily').timeBased().everyDays(1).atHour(12).nearMinute(0).create();
  ss.toast('NewFlow: päivittäinen ajo ajastettu klo 12:00');
}

function NF_setupWeeklyMon0200() {
  const ss = SpreadsheetApp.getActive();
  if (typeof NF_runWeekly !== 'function') {
    ss.toast('NF_runWeekly puuttuu — varmista, että PR #4 NF_Main.gs on asennettu.');
    return;
  }
  NF_clearTriggersBy_(['NF_runWeekly']);
  ScriptApp.newTrigger('NF_runWeekly').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(2).nearMinute(0).create();
  ss.toast('NewFlow: viikkoajo ajastettu ma 02:00');
}

function NF_clearNFTriggers() {
  const names = new Set(['NF_runDaily', 'NF_runWeekly']);
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => {
    if (names.has(tr.getHandlerFunction())) {
      ScriptApp.deleteTrigger(tr);
    }
  });
  SpreadsheetApp.getActive().toast('NewFlow: NF-ajastukset poistettu');
}

function NF_clearTriggersBy_(names) {
  const set = new Set(names);
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (set.has(tr.getHandlerFunction())) ScriptApp.deleteTrigger(tr);
  });
}