/**
 * 04_TestScript.gs — Paikalliset itse-testit (ei ulkoisia API-kutsuja)
 * Voit poistaa tämän tuotantokäytössä.
 */

function seedDemoData_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Adhoc_Tracking') || ss.insertSheet('Adhoc_Tracking');
  sh.clear();
  sh.getRange(1,1,1,7).setValues([[
    'Carrier','Tracking number','RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw'
  ]]);
  sh.getRange(2,1,3,2).setValues([
    ['Kaukokiito','ABC123456'],
    ['Kaukokiito','DEF987654'],
    ['Posti','JJFI000000000000000000'] // dummy
  ]);
  sh.setFrozenRows(1);
}

function test_LocalParsers_(){
  // Päiväys
  const d1 = parseDateFlexible_('2025-09-17 13:44:22');
  const d2 = parseDateFlexible_('17.09.2025 13:44');
  Logger.log('parseDateFlexible_: %s | %s', d1, d2);

  // Kaukokiito JSON sample
  const sample = {
    "consignment": {
      "events": [
        {"status":"LAHETYS LUOVUTETTU","dateTime":"2025-09-01T10:22:00","location":"Vantaa"},
        {"status":"Kuljetuksessa","dateTime":"2025-08-31T08:00:00","location":"Helsinki"}
      ]
    }
  };
  const latest = kk_pickLatestEventFromUnknownJson_(sample);
  Logger.log('kk_pickLatestEventFromUnknownJson_: %s', JSON.stringify(latest));

  // HTML fallback: sis. 'toimitettu' ja aikaleimoja
  const html = '<div>Toimitettu 17.09.2025 13:00 Vantaa</div>';
  const res = (function fakeScrape(html){
    const reIso = /(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2})?)/g;
    const reFi  = /(\d{1,2}\.\d{1,2}\.\d{4}[ T]\d{1,2}:\d{2}(?::\d{2})?)/g;
    const times = []; let m;
    while ((m = reIso.exec(html)) !== null){ times.push(parseDateFlexible_(m[1])); }
    while ((m = reFi.exec(html))  !== null){ times.push(parseDateFlexible_(m[1])); }
    const bestTime = times.filter(Boolean).sort((a,b)=>b-a)[0] || '';
    const delivered = String(html).toLowerCase().includes('toimitettu');
    return { status: delivered ? 'DELIVERED' : (bestTime ? 'IN_TRANSIT' : 'NOT_FOUND'), time: bestTime && fmtDateTime_(bestTime) };
  })(html);
  Logger.log('HTML fallback heuristic: %s', JSON.stringify(res));
}