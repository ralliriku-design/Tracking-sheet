/******************************************************
 * Bulk tracking with throttling, cache, retry-after
 ******************************************************/

const BULK_MAX_API_CALLS_PER_RUN = 20;
const BULK_TIME_LIMIT_MS = 20 * 1000;
const BULK_COOLDOWN_NO_CODE_H = 1;
const BULK_BACKOFF_MINUTES_BASE = 5;
const BULK_BACKOFF_MINUTES_MAX  = 60;

function bulkStartForActiveSheet(){ bulkStart_(SpreadsheetApp.getActiveSheet().getName()); }
function bulkStart_Vaatii(){ bulkStart_(ACTION_SHEET); }
function bulkStart_Packages(){ bulkStart_(TARGET_SHEET); }
function bulkStart_Archive(){ bulkStart_(ARCHIVE_SHEET); }

function bulkStop(){
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'bulkWorker_') ScriptApp.deleteTrigger(t); });
  withLock_('BULK', () => {
    const sp = PropertiesService.getScriptProperties();
    Object.keys(sp.getProperties()).filter(k => k.startsWith('BULKJOB|')).forEach(k => sp.deleteProperty(k));
  });
  SpreadsheetApp.getActive().toast('Bulk-ajo pysäytetty');
}

function bulkStart_(sheetName){
  const sp = PropertiesService.getScriptProperties();
  const key = 'BULKJOB|' + sheetName;
  withLock_('BULK', () => {
    const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) throw new Error(`${sheetName}: ei rivejä`);
    const total = Math.max(0, sh.getLastRow() - 1);
    sp.setProperty(key, JSON.stringify({
      sheet: sheetName, row: 2, calls: 0, started: Date.now(), total, updated: Date.now(), done: 0, left: total
    }));
  });
  bulkWorker_();
  ScriptApp.newTrigger('bulkWorker_').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getActive().toast(`Bulk-ajo aloitettu: ${sheetName}`);
}

function bulkWorker_(){
  withLock_('BULK', () => {
    const ss = SpreadsheetApp.getActive();
    const sp = PropertiesService.getScriptProperties();
    const jobs = Object.entries(sp.getProperties()).filter(([k]) => k.startsWith('BULKJOB|'));
    if (!jobs.length) return;

    const callsByCarrier = { posti:0, gls:0, dhl:0, matkahuolto:0, bring:0, other:0 };
    let didCalls = 0;

    for (const [key, val] of jobs){
      let job = {};
      try { job = JSON.parse(val || '{}'); } catch(e){}
      const sh = ss.getSheetByName(job.sheet);
      if (!sh || sh.getLastRow() < 2){ sp.deleteProperty(key); continue; }

      const tStart = Date.now();
      const data = sh.getDataRange().getDisplayValues();
      const hdr  = data[0].map(String);
      const idxRef = ensureRefreshCols_(sh, hdr);
      const finalHdr = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);
      const carrierI   = pickAnyIndex_(finalHdr, CARRIER_CANDIDATES);
      const codeI      = pickAnyIndex_(finalHdr, TRACKING_CODE_CANDIDATES);
      if (carrierI < 0 || codeI < 0){ sp.deleteProperty(key); SpreadsheetApp.getActive().toast(`${job.sheet}: ei Carrier/Tracking -sarakkeita`); continue; }

      let r = Math.max(2, job.row || 2);
      let wrote = 0;
      const baseBackoff = Number(spGet_('BULK_BACKOFF_MINUTES_BASE') || BULK_BACKOFF_MINUTES_BASE);

      while (r <= data.length && (Date.now() - tStart) < BULK_TIME_LIMIT_MS){
        const allowedCalls = BULK_MAX_API_CALLS_PER_RUN + Math.min(EXTRA_PRIORITY_CALLS, (callsByCarrier.posti + callsByCarrier.gls));
        if (didCalls >= allowedCalls) break;

        const row = (() => { const rr = (data[r-1] || []).slice(0, finalHdr.length); while (rr.length < finalHdr.length) rr.push(''); return rr; })();
        const nextAt = parseDateFlexible_(row[idxRef.nextAt]);
        if (nextAt && nextAt instanceof Date && nextAt.getTime() > Date.now()){ r++; continue; }

        const carrier = String(row[carrierI]||'').trim();
        const codeRaw = firstCode_(row[codeI]);
        if (!carrier || !codeRaw){
          row[idxRef.status]   = 'SKIP_NO_CODE';
          row[idxRef.at]       = fmtDateTime_(new Date());
          row[idxRef.nextAt]   = fmtDateTime_(new Date(Date.now() + BULK_COOLDOWN_NO_CODE_H*3600*1000));
          data[r-1] = row; wrote++; r++; continue;
        }
        if (!isCodeLikely_(codeRaw)){
          row[idxRef.status] = 'SKIP_INVALID_CODE';
          row[idxRef.at]     = fmtDateTime_(new Date());
          data[r-1] = row; wrote++; r++; continue;
        }

        const cc = canonicalCarrier_(carrier);
        let res = trkGetCached_(carrier, codeRaw);
        if (!res){
          didCalls++;
          if (callsByCarrier.hasOwnProperty(cc)) callsByCarrier[cc]++; else callsByCarrier.other++;
          // Throttle per carrier
          const tag = (cc==='posti'?'POSTI':cc==='gls'?'GLS':cc==='dhl'?'DHL':cc==='matkahuolto'?'MH':cc==='bring'?'BRING':'OTHER');
          trkThrottle_(tag);
          res = TRK_trackByCarrier_(carrier, codeRaw);
          job.calls = (job.calls || 0) + 1;
          trkSetCached_(carrier, codeRaw, res);
        }

        row[idxRef.carrier]  = res.carrier || carrier;
        row[idxRef.status]   = res.status  || row[idxRef.status] || '';
        row[idxRef.time]     = res.time    || row[idxRef.time] || '';
        row[idxRef.location] = res.location|| row[idxRef.location] || '';
        row[idxRef.raw]      = res.raw     || '';
        row[idxRef.at]       = fmtDateTime_(new Date());

        if (res.status === 'RATE_LIMIT_429' && typeof res.retryAfter === 'number'){
          row[idxRef.attempts] = (parseInt(row[idxRef.attempts]||'0', 10) || 0) + 1;
          row[idxRef.nextAt]   = fmtDateTime_(trkComputeNextAtFromRetryAfter_(res.retryAfter));
          data[r-1] = row; wrote++; r++; continue;
        }

        const isDeliveredNow = isDeliveredByText_(res.status || '');
        if (isDeliveredNow && res.time){
          if (typeof idxRef.delivConfirmed === 'number' && !row[idxRef.delivConfirmed]) row[idxRef.delivConfirmed] = res.time;
          if (typeof idxRef.delivSource === 'number') row[idxRef.delivSource] = 'tracking';
        }

        if (!isDeliveredNow && !res.found){
          const attempts = (parseInt(row[idxRef.attempts]||'0', 10) || 0) + 1;
          row[idxRef.attempts] = attempts;
          const mins = Math.min(BULK_BACKOFF_MINUTES_MAX, baseBackoff * Math.pow(2, attempts - 1));
          row[idxRef.nextAt] = fmtDateTime_(new Date(Date.now() + mins*60*1000));
        } else {
          row[idxRef.nextAt] = '';
        }

        data[r-1] = row;
        wrote++;
        r++;
      }

      if (wrote){
        sh.getRange(1,1,1,finalHdr.length).setValues([finalHdr]);
        sh.getRange(2,1,data.length-1,finalHdr.length).setValues(data.slice(1).map(x => { const rr = x.slice(0, finalHdr.length); while (rr.length < finalHdr.length) rr.push(''); return rr; }));
        sh.setFrozenRows(1);
      }

      if (r > data.length){
        sp.deleteProperty(key);
        SpreadsheetApp.getActive().toast(`Bulk-ajo valmis: ${job.sheet} (API-kutsuja: ${job.calls || 0})`);
      } else {
        job.row    = r;
        const total = job.total || Math.max(0, data.length - 1);
        const done  = Math.max(0, r - 2);
        job.total   = total;
        job.done    = done;
        job.left    = Math.max(0, total - done);
        job.updated = Date.now();
        sp.setProperty(key, JSON.stringify(job));
      }

      if ((Date.now() - tStart) >= BULK_TIME_LIMIT_MS || didCalls >= BULK_MAX_API_CALLS_PER_RUN) break;
    }
  });
}