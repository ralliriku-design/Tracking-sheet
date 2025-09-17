// bulk.gs — bulk runner with success stats

const EXTRA_PRIORITY_CALLS = 0;
function menuRefreshCarrier_MH() { bulkStartForSheetFiltered_(ACTION_SHEET, 'matkahuolto'); }
function menuRefreshCarrier_POSTI() { bulkStartForSheetFiltered_(ACTION_SHEET, 'posti'); }
function menuRefreshCarrier_BRING() { bulkStartForSheetFiltered_(ACTION_SHEET, 'bring'); }
function menuRefreshCarrier_GLS() { bulkStartForSheetFiltered_(ACTION_SHEET, 'gls'); }
function menuRefreshCarrier_DHL() { bulkStartForSheetFiltered_(ACTION_SHEET, 'dhl'); }
function menuRefreshCarrier_ALL() { bulkStartForSheetFiltered_(ACTION_SHEET, null); }

function bulkStartForSheetFiltered_(sheetName, carrierFilter) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const job = { sheet: sheetName, row: 2, calls: 0, started: Date.now(), total: sh.getLastRow()-1, carrierFilter: carrierFilter, stats:{tried:0,success:0,errors:0,rate429:0} };
  PropertiesService.getScriptProperties().setProperty('BULKJOB|' + sheetName, JSON.stringify(job));
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('bulkTick').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert(`Bulk päivitys aloitettu taululle "${sheetName}"${carrierFilter ? ' (suodatus: '+carrierFilter+')' : ''}.`);
}
function bulkStartForActiveSheet() { const sh = SpreadsheetApp.getActiveSheet(); if (!sh) return; bulkStartForSheetFiltered_(sh.getName(), null); }
function bulkStart_Vaatii() { bulkStartForSheetFiltered_(ACTION_SHEET, null); }
function bulkStart_Packages() { bulkStartForSheetFiltered_(TARGET_SHEET, null); }
function bulkStart_Archive() { bulkStartForSheetFiltered_(ARCHIVE_SHEET, null); }

function bulkStop() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  const sp = PropertiesService.getScriptProperties();
  Object.keys(sp.getProperties()).forEach(k => { if (k.indexOf('BULKJOB|')===0) sp.deleteProperty(k); });
  SpreadsheetApp.getUi().alert('Bulk-ajo pysäytetty.');
}

function bulkTick() {
  const sp = PropertiesService.getScriptProperties();
  const props = sp.getProperties();
  const jobKeys = Object.keys(props).filter(k => k.indexOf('BULKJOB|') === 0);
  if (!jobKeys.length) {
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
    return;
  }
  const ss = SpreadsheetApp.getActive();
  function paused_(carrier){ const until = parseInt(sp.getProperty('PAUSE_UNTIL_' + String(carrier||'').toUpperCase())||'0',10); return until && Date.now() < until; }

  jobKeys.forEach(jobKey => {
    try {
      const job = JSON.parse(sp.getProperty(jobKey));
      if (!job || !job.sheet) { sp.deleteProperty(jobKey); return; }
      const sh = ss.getSheetByName(job.sheet);
      if (!sh) { sp.deleteProperty(jobKey); return; }

      const data = sh.getDataRange().getValues();
      if (!data || data.length < 2) { sp.deleteProperty(jobKey); return; }

      const hdr = data[0];
      const colIndex = headerIndexMap_(hdr);
      const carrierI = colIndexOf_(colIndex, CARRIER_CANDIDATES);
      const trackI   = colIndexOf_(colIndex, TRACKING_CANDIDATES);
      const rCarrierI = colIndex['RefreshCarrier'];
      const rTimeI    = colIndex['RefreshTime'];

      const maxCalls = BULK_MAX_API_CALLS_PER_RUN;
      let callsThisTick = 0;
      job.stats = job.stats || {tried:0,success:0,errors:0,rate429:0};

      for (let r = job.row || 2; r < data.length && callsThisTick < maxCalls; r++) {
        const row = data[r];
        if (!row.join('')) continue;
        const carrier = String((row[rCarrierI] || row[carrierI] || '')).toLowerCase();
        if (job.carrierFilter && carrier.indexOf(job.carrierFilter) === -1) continue;
        if (paused_(carrier)) continue;

        const code = firstCode_(row[trackI] || '');
        const lastTime = row[rTimeI] || '';
        const expired = !lastTime || (Date.now() - new Date(lastTime).getTime() > 6*3600*1000);
        if (!code || !carrier || !expired) continue;

        trkRateLimitWait_(carrier);
        const res = TRK_trackByCarrier_(carrier, code);
        job.stats.tried++;

        const out = [
          res.carrier || carrier,
          res.status || '',
          res.time || fmtDateTime_(new Date()),
          res.location || '',
          res.raw || ''
        ];
        const firstCol = hdr.indexOf('RefreshCarrier') !== -1 ? hdr.indexOf('RefreshCarrier') + 1 : hdr.length + 1;
        try { sh.getRange(r + 1, firstCol, 1, out.length).setValues([out]); } catch(err) { logError_('bulkTick setValues', err); }

        if (isSuccessResult_(res)) job.stats.success++;
        else if (String(res.status||'')==='RATE_LIMIT_429') job.stats.rate429++; else job.stats.errors++;

        job.calls = (job.calls || 0) + 1;
        job.row = r + 1;
        callsThisTick++;

        if (res.status === 'RATE_LIMIT_429') { autoTuneRateLimitOn429_(carrier, 200, 2000); break; }
      }

      // finish or persist
      if ((job.row && job.row >= data.length) || (job.calls && job.calls >= (job.total || data.length-1))) {
        statsAppend_('bulk', job.sheet, job.stats.tried, job.stats.success, job.stats.errors, job.stats.rate429);
        sp.deleteProperty(jobKey);
      } else {
        sp.setProperty(jobKey, JSON.stringify(job));
      }
    } catch (e) {
      logError_('bulkTick processing ' + jobKey, e);
    }
  });

  const remaining = Object.keys(PropertiesService.getScriptProperties().getProperties()).filter(k => k.indexOf('BULKJOB|')===0);
  if (!remaining.length) {
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  }
}
