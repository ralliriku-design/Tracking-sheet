/**
 * 04_refresh_and_bulk.gs — CLEAN FIXED VERSION
 * - Safe, brace-checked implementation of bulk refresh
 * - No duplicate trailing blocks; no stray semicolons after function headers
 * 
 * Depends on helpers defined elsewhere in your project:
 *   - headerIndexMap_(hdr)
 *   - firstCode_(cell)
 *   - fmtDateTime_(date)
 *   - trkRateLimitWait_(carrier)
 *   - TRK_trackByCarrier_(carrier, code)
 *   - getCfgInt_(key, fallback)
 */

// ---- Config (with fallbacks to Script Properties in code) ----
var EXTRA_PRIORITY_CALLS = 0;

// ---- Bulk menu helpers ----
function bulkStartForSheetFiltered_(sheetName, carrierFilter) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { SpreadsheetApp.getUi().alert('Taulua ei löytynyt: ' + sheetName); return; }
  const job = {
    sheet: sheetName,
    row: 2,
    calls: 0,
    started: Date.now(),
    total: Math.max(0, sh.getLastRow() - 1),
    carrierFilter: carrierFilter ? String(carrierFilter).toLowerCase() : null
  };
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty('BULKJOB|' + sheetName, JSON.stringify(job));

  // ensure only one bulkTick trigger exists
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('bulkTick').timeBased().everyMinutes(1).create();

  SpreadsheetApp.getUi().alert('Bulk-päivitys aloitettu taululle "'+sheetName+'"' + (carrierFilter ? (' (suodatus: '+carrierFilter+')') : '') + '.');
}

function bulkStartForActiveSheet() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh) return;
  bulkStartForSheetFiltered_(sh.getName(), null);
}
function bulkStart_Vaatii()  { bulkStartForSheetFiltered_('Vaatii_toimenpiteitä', null); }
function bulkStart_Packages(){ bulkStartForSheetFiltered_('Packages', null); }
function bulkStart_Archive() { bulkStartForSheetFiltered_('Packages_Archive', null); }

function bulkStop() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  const sp = PropertiesService.getScriptProperties();
  Object.keys(sp.getProperties()).forEach(k => { if (k.indexOf('BULKJOB|')===0) sp.deleteProperty(k); });
  SpreadsheetApp.getUi().alert('Bulk-ajo pysäytetty.');
}

// ---- Bulk worker ----
function bulkTick() {
  const sp = PropertiesService.getScriptProperties();
  const props = sp.getProperties();
  const jobKeys = Object.keys(props).filter(k => k.indexOf('BULKJOB|') === 0);

  // nothing queued → remove trigger
  if (!jobKeys.length) {
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
    return;
  }

  const ss = SpreadsheetApp.getActive();

  jobKeys.forEach(jobKey => {
    try {
      const job = JSON.parse(sp.getProperty(jobKey));
      if (!job || !job.sheet) { sp.deleteProperty(jobKey); return; }

      // carrier pause (rate limit backoff) - optional
      if (job.carrierFilter) {
        const pauseKey = 'PAUSE_UNTIL_' + String(job.carrierFilter || '').toUpperCase();
        const until = Number(sp.getProperty(pauseKey) || '0');
        if (until && Date.now() < until) return; // still paused
      }

      const sh = ss.getSheetByName(job.sheet);
      if (!sh) { sp.deleteProperty(jobKey); return; }

      const data = sh.getDataRange().getValues();
      if (!data || data.length < 2) { sp.deleteProperty(jobKey); return; }

      const hdr = data[0];
      const colIndex = headerIndexMap_(hdr);

      const output = [];
      const maxCallsFromProp = getCfgInt_ && getCfgInt_('BULK_MAX_API_CALLS_PER_RUN', 300) || 300;
      const maxCallsThisRun = Math.min(job.total || (data.length - 1), (job.calls || 0) + (EXTRA_PRIORITY_CALLS || 0) + maxCallsFromProp);

      const refreshCutoffMs = 6 * 3600 * 1000; // 6h TTL

      for (let r = job.row || 2; r < data.length && (job.calls || 0) < maxCallsThisRun; r++) {
        const row = data[r];
        if (!row || !row.join('')) continue;

        // carrier from either RefreshCarrier or Carrier/LogisticsProvider/Shipper
        const rawCarrier = row[colIndex['RefreshCarrier']] || row[colIndex['Carrier']] || row[colIndex['LogisticsProvider']] || row[colIndex['Shipper']] || '';
        const carrier = String(rawCarrier || '').toLowerCase();
        if (job.carrierFilter && carrier.indexOf(job.carrierFilter) === -1) continue;

        // tracking code (supports "Package Number" and "Tracking Number")
        const code = firstCode_( row[colIndex['TrackingNumber']] || row[colIndex['Package Number']] || '' );

        const lastTime = row[colIndex['RefreshTime']] || '';
        const expired = !lastTime || (Date.now() - (new Date(lastTime).getTime())) > refreshCutoffMs;
        if (!code || !carrier || !expired) continue;

        // rate-limit gate
        trkRateLimitWait_(carrier);

        const res = TRK_trackByCarrier_(carrier, code);
        output.push({
          rowIndex: r,
          data: [
            res.carrier || carrier || '',
            res.status  || '',
            res.time    || fmtDateTime_(new Date()),
            res.location|| '',
            res.raw     || ''
          ]
        });

        job.calls = (job.calls || 0) + 1;
        job.row = r + 1;

        if (res.status === 'RATE_LIMIT_429') {
          const backoffMin = getCfgInt_ && getCfgInt_('BULK_BACKOFF_MINUTES_BASE', 30) || 30;
          const pauseKey = 'PAUSE_UNTIL_' + String(carrier).toUpperCase();
          sp.setProperty(pauseKey, String(Date.now() + backoffMin * 60000));
          break;
        }
      }

      // write results
      if (output.length) {
        const firstCol = hdr.indexOf('RefreshCarrier') !== -1 ? (hdr.indexOf('RefreshCarrier') + 1) : (hdr.length + 1);
        output.forEach(e => {
          try {
            sh.getRange(e.rowIndex + 1, firstCol, 1, e.data.length).setValues([e.data]);
          } catch(err) {
            if (typeof logError_ === 'function') logError_('bulkTick setValues', err);
          }
        });
      }

      // update or clear job
      if ((job.row && job.row >= data.length) || ((job.calls || 0) >= (job.total || (data.length - 1)))) {
        sp.deleteProperty(jobKey);
      } else {
        sp.setProperty(jobKey, JSON.stringify(job));
      }
    } catch (e) {
      if (typeof logError_ === 'function') logError_('bulkTick processing ' + jobKey, e);
    }
  });

  // if no jobs remain, remove trigger
  const remaining = Object.keys(PropertiesService.getScriptProperties().getProperties()).filter(k => k.indexOf('BULKJOB|') === 0);
  if (!remaining.length) {
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  }
}
