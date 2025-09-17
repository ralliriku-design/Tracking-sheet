/**
 * z_fallbacks_safe.gs
 * Provides missing helper functions if they are not already defined.
 * Safe to add to any project; uses typeof-guards to avoid redefinition.
 */

// --- Properties getter (fallback) ---
if (typeof TRK_props_ !== 'function') {
  function TRK_props_(k) {
    try { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
    catch(_) { return ''; }
  }
}

// --- Date formatter (fallback) ---
if (typeof fmtDateTime_ !== 'function') {
  function fmtDateTime_(d) {
    if (!d) return '';
    try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"); }
    catch(e){ try { return new Date(d).toISOString(); } catch(_) { return String(d); } }
  }
}

// --- Error logger (fallback) ---
if (typeof logError_ !== 'function') {
  function logError_(context, error){
    try{
      const ss = SpreadsheetApp.getActive();
      const sh = ss.getSheetByName('ErrorLog') || ss.insertSheet('ErrorLog');
      sh.appendRow([new Date(), String(context||''), String(error||'')]);
      const lr = sh.getLastRow();
      if (lr > 2000) sh.deleteRows(1, lr-2000);
    }catch(_){}
  }
}

// --- getCfgInt_ (fallback) ---
if (typeof getCfgInt_ !== 'function') {
  function getCfgInt_(key, fallback) {
    try {
      const v = PropertiesService.getScriptProperties().getProperty(key);
      return v ? Number(v) : fallback;
    } catch(_){
      return fallback;
    }
  }
}

// --- Rate limit wait (fallback) ---
if (typeof trkRateLimitWait_ !== 'function') {
  function trkRateLimitWait_(carrier) {
    // Minimal no-op limiter with light locking so parallel triggers won't collide.
    try {
      const lock = LockService.getScriptLock();
      lock.waitLock(5000);
      // basic min-ms per carrier
      const sp = PropertiesService.getScriptProperties();
      const keyLast = 'RATE_LAST_' + String(carrier||'').toUpperCase();
      const keyMinMs = 'RATE_MINMS_' + String(carrier||'').toUpperCase();
      const last = Number(sp.getProperty(keyLast) || '0');
      const minMs = Number(sp.getProperty(keyMinMs) || '200'); // default 200ms
      const now = Date.now();
      const elapsed = now - last;
      if (minMs && elapsed < minMs) Utilities.sleep(minMs - elapsed);
      sp.setProperty(keyLast, String(Date.now()));
      try { lock.releaseLock(); } catch(_){}
    } catch(e) {
      // swallow
    }
  }
}

// --- Auto-tune rate limit on 429 (fallback) ---
if (typeof autoTuneRateLimitOn429_ !== 'function') {
  function autoTuneRateLimitOn429_(carrier, addMs, capMs) {
    try {
      const sp = PropertiesService.getScriptProperties();
      const key = 'RATE_MINMS_' + String(carrier||'').toUpperCase();
      const cur = Number(sp.getProperty(key) || '200');
      const upd = Math.min(cur + (addMs||100), (capMs|| (cur + (addMs||100))));
      sp.setProperty(key, String(upd));
    } catch(_) {}
  }
}
