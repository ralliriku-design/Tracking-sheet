/** Tracking_Enhanced_AllInOne.gs
 * Drop-in script for Google Apps Script (V8).
 * - Enhanced tracking with retry/throttle/cache/normalization
 * - Carrier credentials helpers
 * - Script Properties import/export helpers
 * - Sheet writer
 *
 * Version: 2025-09-17
 */

/* =========================
   Public Entry Points
   ========================= */

/**
 * Track shipment(s) with enhanced orchestration.
 * ids: string | string[] | {carrier:string,id:string}[]
 * opts:
 *  - carrier?: default carrier when ids are plain strings
 *  - useCacheSeconds?: number (default 600)
 *  - perCarrierMinMs?: { [carrier]: number } (overrides RATE_MINMS props)
 *  - maxRetries?: number (default 3)
 *  - timeoutMs?: number (default 25000)
 *  - debug?: boolean
 */
function TRK_trackByCarrierEnhanced(ids, opts) {
  opts = opts || {};
  var list = normalizeInput_(ids, opts);
  var results = [];

  for (var i = 0; i < list.length; i++) {
    var req = list[i];
    try {
      var res = trackSingle_(req.carrier, req.id, opts);
      results.push(res);
    } catch (e) {
      logWarn_('trackSingle_ failed: ' + req.carrier + ' ' + req.id + ' ' + (e && e.message));
      results.push({
        id: req.id,
        carrier: req.carrier,
        delivered: false,
        status: 'ERROR',
        lastEvent: null,
        events: [],
        error: stringifySafe_(e)
      });
    }
  }
  return results;
}

/**
 * Append normalized tracking results to a sheet.
 * rows: array of {carrier, id} or simple IDs if opts.carrier is set
 */
function TRK_writeToSheetEnhanced(rows, sheetName, opts) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName || 'Tracking_Results') || ss.insertSheet(sheetName || 'Tracking_Results');

  if (sh.getLastRow() === 0) {
    sh.appendRow(['Carrier', 'ID', 'Delivered', 'Status', 'LastEventTime', 'LastEvent', 'Location', 'Events(JSON)']);
  }

  var res = TRK_trackByCarrierEnhanced(rows, opts || {});
  var out = [];
  for (var i = 0; i < res.length; i++) {
    var r = res[i];
    var lastTime = r.lastEvent && r.lastEvent.time ? r.lastEvent.time : '';
    var lastDesc = r.lastEvent && r.lastEvent.description ? r.lastEvent.description : '';
    var lastLoc = r.lastEvent && r.lastEvent.location ? r.lastEvent.location : '';
    out.push([r.carrier, r.id, r.delivered, r.status, lastTime, lastDesc, lastLoc, JSON.stringify(r.events)]);
  }
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  return res;
}

/* =========================
   Quick Test Helpers
   ========================= */

function TRK_Test_Single() {
  // Replace with a real tracking number and carrier you use
  return TRK_trackByCarrierEnhanced({ carrier: 'POSTI', id: 'JJFI123456789000000' }, { debug: true });
}

function TRK_Test_Batch() {
  return TRK_trackByCarrierEnhanced([
    { carrier: 'POSTI', id: 'JJFI123' },
    { carrier: 'GLS', id: '00ABC' },
    { carrier: 'DHL', id: '0034...' },
    { carrier: 'BRING', id: 'PNO...' }
  ], { useCacheSeconds: 600, debug: true });
}

/* =========================
   Credential Helpers (Script Properties)
   ========================= */

/** Generic set property */
function Props_set(key, value) {
  PropertiesService.getScriptProperties().setProperty(String(key), value == null ? '' : String(value));
}

/** Generic get property */
function Props_get(key, def) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(String(key));
    return (v == null || v === '') ? def : v;
  } catch (e) { return def; }
}

/** Export all script properties to a CSV string: key,value */
function Props_exportCsv() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var rows = ['key,value'];
  Object.keys(props).sort().forEach(function (k) {
    rows.push(csvEscape_(k) + ',' + csvEscape_(props[k]));
  });
  return rows.join('\n');
}

/** Import script properties from a CSV string "key,value" */
function Props_importCsv(csvText) {
  var lines = String(csvText || '').split(/\r?\n/);
  if (!lines.length) return 0;
  var count = 0;
  for (var i = 0; i < lines.length; i++) {
    if (i === 0 && /^key\s*,\s*value$/i.test(lines[i].trim())) continue;
    var line = lines[i].trim();
    if (!line) continue;
    var parts = parseCsvLine_(line);
    if (!parts || parts.length < 2) continue;
    var k = parts[0], v = parts[1];
    if (k) { Props_set(k, v); count++; }
  }
  return count;
}

/** Clear common credential keys safely (does not clear non-credential props) */
function Creds_clearAll() {
  var keys = [
    'POSTI_TRK_BASIC', 'POSTI_BASIC', 'POSTI_TRK_USER', 'POSTI_TRK_PASS',
    'GLS_FI_API_KEY', 'GLS_BASIC',
    'DHL_API_KEY',
    'BRING_UID', 'BRING_KEY',
    'MH_BASIC',
    'KAUKO_BASIC', 'KAUKO_TOKEN'
  ];
  var sp = PropertiesService.getScriptProperties();
  for (var i = 0; i < keys.length; i++) sp.deleteProperty(keys[i]);
}

/* Carrier-specific setters */
function setPostiBasicBase64(b64UserColonPass) { Props_set('POSTI_TRK_BASIC', b64UserColonPass); }
function setPostiBasicUserPass(user, pass) { var basic = Utilities.base64Encode(String(user) + ':' + String(pass)); Props_set('POSTI_TRK_BASIC', basic); }
function setPostiLegacyUserPass(user, pass) { Props_set('POSTI_TRK_USER', user); Props_set('POSTI_TRK_PASS', pass); }
function setDhlApiKey(key) { Props_set('DHL_API_KEY', key); }
function setBring(uid, key, clientUrl) { Props_set('BRING_UID', uid); Props_set('BRING_KEY', key); if (clientUrl) Props_set('BRING_CLIENT_URL', clientUrl); }
function setGlsFi(apiKey, senderIdsCsv, trackUrl) { Props_set('GLS_FI_API_KEY', apiKey); if (senderIdsCsv) Props_set('GLS_FI_SENDER_IDS', senderIdsCsv); if (trackUrl) Props_set('GLS_FI_TRACK_URL', trackUrl); }
function setGlsBasic(base64UserColonPass) { Props_set('GLS_BASIC', base64UserColonPass); }
function setMatkahuoltoBasic(base64UserColonPass) { Props_set('MH_BASIC', base64UserColonPass); }
function setKaukokiitoBasic(base64UserColonPass) { Props_set('KAUKO_BASIC', base64UserColonPass); }
function setKaukokiitoToken(token) { Props_set('KAUKO_TOKEN', token); }

/* Rate limit and pause controls */
function Rate_setMinMs(carrier, minMs) { var c = canonicalCarrier_(carrier); Props_set('RATE_MINMS_' + c, String(Math.max(1, Number(minMs || 1)))); }
function Rate_getMinMs(carrier) { var c = canonicalCarrier_(carrier); var def = defaultMinMs_[c] || 500; return Number(Props_get('RATE_MINMS_' + c, def)); }
function Rate_pauseUntil(carrier, epochMs) { var c = canonicalCarrier_(carrier); Props_set('PAUSE_UNTIL_' + c, String(epochMs)); }
function Rate_resume(carrier) { var c = canonicalCarrier_(carrier); PropertiesService.getScriptProperties().deleteProperty('PAUSE_UNTIL_' + c); }

/* =========================
   Core Engine
   ========================= */

function normalizeInput_(ids, opts) {
  var out = [];
  if (ids == null) return out;
  if (Array.isArray(ids)) {
    for (var i = 0; i < ids.length; i++) {
      var v = ids[i];
      if (v && typeof v === 'object') {
        out.push({ carrier: canonicalCarrier_(v.carrier || opts.carrier), id: String(v.id) });
      } else {
        out.push({ carrier: canonicalCarrier_(opts.carrier), id: String(v) });
      }
    }
  } else if (typeof ids === 'object') {
    out.push({ carrier: canonicalCarrier_(ids.carrier || opts.carrier), id: String(ids.id) });
  } else {
    out.push({ carrier: canonicalCarrier_(opts.carrier), id: String(ids) });
  }
  out = out.filter(function (x) { return x.carrier && x.id; });
  return out;
}

function canonicalCarrier_(s) {
  if (!s) return '';
  var k = String(s).trim().toUpperCase();
  var map = {
    'POSTI': 'POSTI',
    'POSTI API (FI)': 'POSTI',
    'POSTI OY (FI)': 'POSTI',
    'GLS': 'GLS',
    'GLS FINLAND (FI)': 'GLS',
    'DHL': 'DHL',
    'DHL FINLAND (FI)': 'DHL',
    'BRING': 'BRING',
    'MATKAHUOLTO': 'MATKAHUOLTO',
    'MH': 'MATKAHUOLTO',
    'KAUKOKIITO': 'KAUKOKIITO',
    'KAUKO': 'KAUKOKIITO'
  };
  return map[k] || k;
}

function parseDateFlexible_(str) {
  if (!str && str !== 0) return null;
  var s = String(str).trim();
  if (!s) return null;

  var iso = Date.parse(s);
  if (!isNaN(iso)) return new Date(iso);

  var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2})[:\.](\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    var d = Number(m[1]), mo = Number(m[2]) - 1, y = Number(m[3]);
    var hh = Number(m[4] || 0), mm = Number(m[5] || 0), ss = Number(m[6] || 0);
    return new Date(y, mo, d, hh, mm, ss);
  }

  m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]));

  var t = Date.parse(s.replace(/\./g, '/'));
  if (!isNaN(t)) return new Date(t);

  return null;
}

function toIsoOrNull_(d) { return d instanceof Date && !isNaN(d) ? d.toISOString() : null; }

function trackSingle_(carrier, id, opts) {
  if (!carrier) throw new Error('carrier is required');
  if (!id) throw new Error('id is required');

  var cacheSecs = Number(opts.useCacheSeconds || 600);
  var cache = CacheService.getScriptCache();
  var key = 'trk:' + carrier + ':' + id;
  if (cacheSecs > 0) {
    var cached = cache.get(key);
    if (cached) return JSON.parse(cached);
  }

  // Check pause
  var pauseUntil = Number(Props_get('PAUSE_UNTIL_' + carrier, 0));
  if (pauseUntil && Date.now() < pauseUntil) {
    throw new Error('PAUSED for ' + carrier + ' until ' + new Date(pauseUntil).toISOString());
  }

  // Throttle
  throttle_(carrier, opts);

  // Fetch with retry/backoff
  var maxRetries = Number(opts.maxRetries || 3);
  var resp = fetchWithRetry_(function () {
    return carrierHandlers_[carrier].track(id, opts);
  }, maxRetries, opts);

  var normalized = carrierHandlers_[carrier].normalize(resp, { id: id, parseDate: parseDateFlexible_ });

  if (cacheSecs > 0) cache.put(key, JSON.stringify(normalized), Math.min(cacheSecs, 21600));
  return normalized;
}

var defaultMinMs_ = {
  'POSTI': 500,
  'GLS': 800,
  'DHL': 800,
  'BRING': 800,
  'MATKAHUOLTO': 500,
  'KAUKOKIITO': 800
};

function throttle_(carrier, opts) {
  var minMs = (opts.perCarrierMinMs && Number(opts.perCarrierMinMs[carrier])) || Rate_getMinMs(carrier);
  minMs = Math.max(1, Number(minMs || 500));
  var propKey = 'RATE_LAST_' + carrier;
  var sp = PropertiesService.getScriptProperties();
  var last = Number(sp.getProperty(propKey) || 0);
  var now = Date.now();
  var wait = last + minMs - now;
  if (wait > 0 && wait < 60000) Utilities.sleep(wait);
  sp.setProperty(propKey, String(Date.now()));
}

function fetchWithRetry_(fn, maxRetries, opts) {
  var delay = 750;
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (e) {
      if (!isRetryableError_(e) || attempt === maxRetries) throw e;
      if (opts && opts.debug) logInfo_('Retry #' + (attempt + 1) + ' after: ' + (e && e.message));
      Utilities.sleep(delay);
      delay = Math.min(delay * 2, 8000);
    }
  }
}

function isRetryableError_(e) {
  var msg = (e && e.message) || '';
  var code = (e && e.code) || (e && e.responseCode) || null;
  if (code && String(code).match(/^(429|5\d\d)$/)) return true;
  if (/timeout|timed out|quota|rate|temporar|socket|network/i.test(msg)) return true;
  return false;
}

function httpJson_(url, params) {
  params = params || {};
  params.muteHttpExceptions = true;
  params.followRedirects = true;
  if (!params.method) params.method = 'get';
  if (!params.headers) params.headers = {};
  if (params.timeoutMs == null) params.timeoutMs = 25000;

  var res = UrlFetchApp.fetch(url, params);
  var code = res.getResponseCode();
  var body = res.getContentText() || '';
  if (code < 200 || code >= 300) {
    var err = new Error('HTTP ' + code + ' ' + url + ' ' + body.slice(0, 400));
    err.code = code;
    throw err;
  }
  if (!body) return null;
  try { return JSON.parse(body); } catch (e) { return { raw: body }; }
}

function applyTemplateUrl_(tmpl, data) {
  return String(tmpl || '').replace(/\{\{\s*(\w+)\s*\}\}/g, function (_, k) {
    return data && data[k] != null ? encodeURIComponent(String(data[k])) : '';
  });
}

/* =========================
   Carrier Handlers
   ========================= */

var carrierHandlers_ = {
  'POSTI': {
    track: function (id, opts) {
      var trackUrl = Props_get('POSTI_TRACK_URL', '') || Props_get('POSTI_TRK_URL', '');
      var headers = {};
      var basic = Props_get('POSTI_TRK_BASIC', '') || Props_get('POSTI_BASIC', '');
      if (basic) headers.Authorization = 'Basic ' + basic;

      var url;
      if (trackUrl) {
        url = applyTemplateUrl_(trackUrl, { code: id });
      } else {
        url = 'https://api2.posti.fi/tracking/shipments?trackingNumbers=' + encodeURIComponent(id);
        var apiKey = Props_get('POSTI_API_KEY', '');
        if (apiKey && !headers.Authorization) headers['x-api-key'] = apiKey;
      }
      return httpJson_(url, { headers: headers, timeoutMs: opts.timeoutMs || 25000 });
    },
    normalize: function (payload, ctx) {
      var events = [], delivered = false, status = 'UNKNOWN', lastEvent = null;
      try {
        var shipments = (payload && payload.shipments) || (payload && payload.TrackingInfo && payload.TrackingInfo.Shipments) || [];
        var s = shipments[0] || payload;
        var evs = (s && (s.events || s.history || s.TrackingEvents)) || [];
        for (var i = 0; i < evs.length; i++) {
          var e = evs[i];
          var dt = ctx.parseDate(e.timestamp || e.eventTime || e.dateTime || e.time || e.DateTime);
          var desc = e.description || e.eventDescription || e.status || e.state || e.EventDescription || '';
          var loc = e.location || e.city || e.country || e.Location || '';
          events.push({ time: toIsoOrNull_(dt), description: desc, location: loc, code: e.code || e.statusCode || e.EventCode || '' });
        }
        if (events.length) { lastEvent = events[events.length - 1]; status = lastEvent.description || 'OK'; }
        delivered = Boolean((s && (s.delivered || s.isDelivered)) || /delivered|toimitettu/i.test(status));
      } catch (e) {}
      return { id: ctx.id, carrier: 'POSTI', delivered: delivered, status: status, lastEvent: lastEvent, events: events, raw: payload };
    }
  },

  'GLS': {
    track: function (id, opts) {
      var headers = {};
      var url = '';
      var fiApiKey = Props_get('GLS_FI_API_KEY', '');
      var fiTrackUrl = Props_get('GLS_FI_TRACK_URL', '');
      var senderIds = Props_get('GLS_FI_SENDER_IDS', Props_get('GLS_FI_SENDER_ID', ''));
      var sender = (String(senderIds || '').split(',')[0] || '').trim();
      var basic = Props_get('GLS_BASIC', '');

      if (fiTrackUrl) {
        url = applyTemplateUrl_(fiTrackUrl, { code: id, sender: sender });
        if (fiApiKey) headers['x-api-key'] = fiApiKey;
      } else if (basic) {
        url = 'https://api.gls-group.eu/public/v1/tracking/' + encodeURIComponent(id);
        headers.Authorization = 'Basic ' + basic;
      } else {
        var glsTrack = Props_get('GLS_TRACK_URL', '');
        url = glsTrack ? applyTemplateUrl_(glsTrack, { code: id }) : 'https://api.gls-group.net/track-and-trace-v1/tracking/simple/references/' + encodeURIComponent(id);
        if (fiApiKey) headers['x-api-key'] = fiApiKey;
      }

      return httpJson_(url, { headers: headers, timeoutMs: opts.timeoutMs || 25000 });
    },
    normalize: function (payload, ctx) {
      var events = [], delivered = false, status = 'UNKNOWN', lastEvent = null;
      try {
        var evs = (payload && (payload.events || payload.history || payload.result || payload.Results)) || [];
        if (payload && payload.results && payload.results[0] && payload.results[0].events) {
          evs = payload.results[0].events;
        }
        for (var i = 0; i < evs.length; i++) {
          var e = evs[i];
          var dt = ctx.parseDate(e.dateTime || e.time || e.timestamp || e.Date);
          var desc = e.description || e.status || e.EventReason || '';
          var loc = e.location || e.city || e.SiteName || '';
          events.push({ time: toIsoOrNull_(dt), description: desc, location: loc, code: e.code || e.EventCode || '' });
        }
        if (events.length) { lastEvent = events[events.length - 1]; status = lastEvent.description || 'OK'; }
        delivered = /delivered|signed/i.test(stringifySafe_(lastEvent));
      } catch (e) {}
      return { id: ctx.id, carrier: 'GLS', delivered: delivered, status: status, lastEvent: lastEvent, events: events, raw: payload };
    }
  },

  'DHL': {
    track: function (id, opts) {
      var key = Props_get('DHL_API_KEY', '');
      if (!key) throw new Error('DHL_API_KEY missing');
      var tmpl = Props_get('DHL_TRACK_URL', 'https://api-eu.dhl.com/track/shipments?trackingNumber={{code}}');
      var url = applyTemplateUrl_(tmpl, { code: id });
      return httpJson_(url, { headers: { 'DHL-API-Key': key }, timeoutMs: opts.timeoutMs || 25000 });
    },
    normalize: function (payload, ctx) {
      var events = [], delivered = false, status = 'UNKNOWN', lastEvent = null;
      try {
        var s = payload && payload.shipments && payload.shipments[0];
        var evs = (s && s.events) || [];
        for (var i = 0; i < evs.length; i++) {
          var e = evs[i];
          var dt = ctx.parseDate(e.timestamp || e.eventTime);
          var desc = e.description || (e.status && e.status.status) || e.statusCode || '';
          var loc = (e.location && e.location.address && e.location.address.addressLocality) || e.location || '';
          events.push({ time: toIsoOrNull_(dt), description: desc, location: loc, code: e.statusCode || '' });
        }
        if (events.length) { lastEvent = events[events.length - 1]; status = lastEvent.description || 'OK'; }
        delivered = Boolean(s && s.status && /delivered/i.test(s.status.status || s.status.statusCode || ''));
      } catch (e) {}
      return { id: ctx.id, carrier: 'DHL', delivered: delivered, status: status, lastEvent: lastEvent, events: events, raw: payload };
    }
  },

  'BRING': {
    track: function (id, opts) {
      var uid = Props_get('BRING_UID', ''), key = Props_get('BRING_KEY', '');
      var clientUrl = Props_get('BRING_CLIENT_URL', 'apps-script');
      if (!uid || !key) throw new Error('BRING_UID/BRING_KEY missing');
      var tmpl = Props_get('BRING_TRACK_URL', 'https://api.bring.com/tracking/api/v2/tracking.json?query={{code}}');
      var url = applyTemplateUrl_(tmpl, { code: id });
      var headers = { 'X-MyBring-API-Uid': uid, 'X-MyBring-API-Key': key, 'X-Bring-Client-URL': clientUrl };
      return httpJson_(url, { headers: headers, timeoutMs: opts.timeoutMs || 25000 });
    },
    normalize: function (payload, ctx) {
      var events = [], delivered = false, status = 'UNKNOWN', lastEvent = null;
      try {
        var cons = payload && payload.consignmentSet && payload.consignmentSet[0];
        var evs = cons && cons.eventSet || [];
        for (var i = 0; i < evs.length; i++) {
          var e = evs[i];
          var dt = ctx.parseDate(e.dateIso || e.displayDate);
          var desc = e.description || e.status || '';
          var loc = e.city || e.postalCode || e.country || '';
          events.push({ time: toIsoOrNull_(dt), description: desc, location: loc, code: e.status || '' });
        }
        if (events.length) { lastEvent = events[events.length - 1]; status = lastEvent.description || 'OK'; }
        delivered = /delivered|utlevert|toimitettu/i.test(stringifySafe_(lastEvent));
      } catch (e) {}
      return { id: ctx.id, carrier: 'BRING', delivered: delivered, status: status, lastEvent: lastEvent, events: events, raw: payload };
    }
  },

  'MATKAHUOLTO': {
    track: function (id, opts) {
      var basic = Props_get('MH_BASIC', '');
      if (!basic) throw new Error('MH_BASIC missing');
      var tmpl = Props_get('MH_TRACK_URL', 'https://extservices.matkahuolto.fi/mpaketti/public/tracking?ids={{code}}');
      var url = applyTemplateUrl_(tmpl, { code: id });
      return httpJson_(url, { headers: { Authorization: 'Basic ' + basic }, timeoutMs: opts.timeoutMs || 25000 });
    },
    normalize: function (payload, ctx) {
      var events = [], delivered = false, status = 'UNKNOWN', lastEvent = null;
      try {
        var evs = payload && (payload.events || payload.history || payload.results) || [];
        for (var i = 0; i < evs.length; i++) {
          var e = evs[i];
          var dt = ctx.parseDate(e.time || e.timestamp || e.date);
          var desc = e.description || e.status || '';
          var loc = e.location || e.city || '';
          events.push({ time: toIsoOrNull_(dt), description: desc, location: loc, code: e.code || '' });
        }
        if (events.length) { lastEvent = events[events.length - 1]; status = lastEvent.description || 'OK'; }
        delivered = /delivered|toimitettu/i.test(stringifySafe_(lastEvent));
      } catch (e) {}
      return { id: ctx.id, carrier: 'MATKAHUOLTO', delivered: delivered, status: status, lastEvent: lastEvent, events: events, raw: payload };
    }
  },

  'KAUKOKIITO': {
    track: function (id, opts) {
      var basic = Props_get('KAUKO_BASIC', '');
      var token = Props_get('KAUKO_TOKEN', '');
      var headers = {};
      if (token) headers.Authorization = 'Bearer ' + token;
      else if (basic) headers.Authorization = 'Basic ' + basic;
      else throw new Error('KAUKOKIITO credentials missing (KAUKO_BASIC or KAUKO_TOKEN)');

      var tmpl = Props_get('KAUKO_TRACK_URL', '');
      var url = tmpl ? applyTemplateUrl_(tmpl, { code: id }) : ('https://api.kaukokiito.fi/tracking?consignment=' + encodeURIComponent(id));
      return httpJson_(url, { headers: headers, timeoutMs: opts.timeoutMs || 25000 });
    },
    normalize: function (payload, ctx) {
      var events = [], delivered = false, status = 'UNKNOWN', lastEvent = null;
      try {
        var evs = payload && (payload.events || payload.history || payload.shipmentEvents) || [];
        for (var i = 0; i < evs.length; i++) {
          var e = evs[i];
          var dt = ctx.parseDate(e.timestamp || e.time || e.date);
          var desc = e.description || e.status || '';
          var loc = e.location || e.city || '';
          events.push({ time: toIsoOrNull_(dt), description: desc, location: loc, code: e.code || '' });
        }
        if (events.length) { lastEvent = events[events.length - 1]; status = lastEvent.description || 'OK'; }
        delivered = /delivered|toimitettu/i.test(stringifySafe_(lastEvent));
      } catch (e) {}
      return { id: ctx.id, carrier: 'KAUKOKIITO', delivered: delivered, status: status, lastEvent: lastEvent, events: events, raw: payload };
    }
  }
};

/* =========================
   Misc Utilities
   ========================= */

function csvEscape_(val) {
  var s = String(val == null ? '' : val);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
function parseCsvLine_(line) {
  var out = [], i = 0, inQ = false, cur = '';
  while (i < line.length) {
    var ch = line[i];
    if (inQ) {
      if (ch === '"') {
        if (line[i + 1] === '"') { cur += '"'; i += 2; continue; }
        inQ = false; i++; continue;
      }
      cur += ch; i++; continue;
    } else {
      if (ch === '"') { inQ = true; i++; continue; }
      if (ch === ',') { out.push(cur); cur = ''; i++; continue; }
      cur += ch; i++; continue;
    }
  }
  out.push(cur);
  return out;
}

function logInfo_(msg) { try { console.log('[INFO] ' + msg); } catch (e) {} }
function logWarn_(msg) { try { console.warn('[WARN] ' + msg); } catch (e) {} }
function stringifySafe_(obj) { try { return JSON.stringify(obj); } catch (e) { return String(obj); } }

/* =========================
   OPTIONAL: Sheet Range Helper
   ========================= */
function TRK_trackSheetRange(sheetName, idCol, carrierCol, startRow, numRows, opts) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  var data = sh.getRange(startRow, Math.min(idCol, carrierCol), numRows, Math.abs(carrierCol - idCol) + 1).getValues();

  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var id = row[Math.abs(idCol - Math.min(idCol, carrierCol))];
    var carrier = row[Math.abs(carrierCol - Math.min(idCol, carrierCol))];
    if (id) rows.push({ carrier: carrier, id: id });
  }
  return TRK_writeToSheetEnhanced(rows, (sheetName + '_Results'), opts || {});
}
/** Adhoc_Tracker.gs
 * Ad hoc -tuonti ja seurantojen päivitys:
 * 1) Tuo seurantalista annetusta Google Sheetistä
 * 2) Deduplikoi jo olemassa oleviin (Packages_Archive, Packages, Adhoc_Results)
 * 3) Tunnista carrier -> hae status -> poimi created & delivered päivät
 * 4) Kirjoita Adhoc_Results- ja Adhoc_KPI -tauluihin
 *
 * Edellyttää: Tracking_Enhanced_AllInOne.gs (TRK_trackByCarrierEnhanced ym.)
 */

// Fallbackit, jos Props_get/Props_set ei ole globaaleina:
function __propsGetFallback__(k, def) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(String(k));
    return (v == null || v === '') ? def : v;
  } catch (e) { return def; }
}
function __propsSetFallback__(k, v) {
  PropertiesService.getScriptProperties().setProperty(String(k), v == null ? '' : String(v));
}
var _Props_get = (typeof Props_get === 'function') ? Props_get : __propsGetFallback__;
var _Props_set = (typeof Props_set === 'function') ? Props_set : __propsSetFallback__;

// Pääkohteet
var DEFAULT_ARCHIVE_SHEET = _Props_get('ARCHIVE_SHEET', 'Packages_Archive');
var DEFAULT_PACKAGES_SHEET = _Props_get('TARGET_SHEET', 'Packages');
var DEFAULT_RESULTS_SHEET = 'Adhoc_Results';
var DEFAULT_KPI_SHEET = 'Adhoc_KPI';

// Käynnistys suoraan URL:lla (esim. antamasi linkki)
function ADHOC_RunFromUrl() {
  var url = Browser.inputBox('Lähde-Spreadsheet URL', 'Liitä Google Sheets -linkki', Browser.Buttons.OK_CANCEL);
  if (url === 'cancel' || !url) return;
  var id = extractSpreadsheetIdFromUrl_(url);
  if (!id) throw new Error('Ei kyetty tunnistamaan Spreadsheet ID:tä URL:sta.');

  var opts = {
    // jos haluat lukea myös nShift-kansiosta/Drive: anna folderId
    nshiftFolderId: '', // esim. '1AbcDEF...'
    trackingHeaderHints: ['tracking','seuranta','package number','paketin numero','paketinro','package','pakkausnumero'],
    countryHeaderHints: ['country','maa','destination country','kohdemaa']
  };
  return ADHOC_ProcessSheet(id, opts);
}

// Ydin: tuo lähde, deduplikoi, hae seurannat, kirjoita ulos
function ADHOC_ProcessSheet(sourceSpreadsheetId, opts) {
  opts = opts || {};
  var trackingColHints = (opts.trackingHeaderHints || []).map(String);
  var countryColHints = (opts.countryHeaderHints || []).map(String);
  var ssSrc = SpreadsheetApp.openById(sourceSpreadsheetId);
  var shSrc = ssSrc.getSheets()[0]; // ensimmäinen sivu
  var data = shSrc.getDataRange().getValues();
  if (!data.length) throw new Error('Lähdetaulu on tyhjä');

  var header = data[0].map(function (h) { return String(h || '').trim(); });
  var trackingIdx = findColumnByHints_(header, trackingColHints, 'Package Number');
  if (trackingIdx < 0) throw new Error('Seurantakoodin saraketta ei löydy. Nimeä esim. "Tracking" tai "Package Number".');
  var countryIdx = findColumnByHints_(header, countryColHints, '');

  var srcRows = [];
  for (var i = 1; i < data.length; i++) {
    var trk = String(data[i][trackingIdx] || '').trim();
    if (!trk) continue;
    srcRows.push({
      tracking: trk,
      country: countryIdx >= 0 ? String(data[i][countryIdx] || '').trim() : ''
    });
  }
  if (!srcRows.length) throw new Error('Lähteessä ei ole yhtään seurantakoodia.');

  // Kerää jo olemassa olevat seurannat aktiivisesta työvihkosta
  var ss = SpreadsheetApp.getActive();
  var existing = new Set();
  addExistingFromSheet_(ss, DEFAULT_ARCHIVE_SHEET, existing);
  addExistingFromSheet_(ss, DEFAULT_PACKAGES_SHEET, existing);
  addExistingFromSheet_(ss, DEFAULT_RESULTS_SHEET, existing);

  // (Valinnainen) kerää nShift-raporteista Drive-kansiosta
  if (opts.nshiftFolderId) addExistingFromDriveFolder_(opts.nshiftFolderId, existing);

  // Suodata käsiteltävät (ei vielä olemassa)
  var todo = srcRows.filter(function (r) { return !existing.has(r.tracking); });
  if (!todo.length) {
    SpreadsheetApp.getActive().toast('Ei uusia seurantoja käsiteltäväksi.');
    return { processed: 0, written: 0 };
  }

  // Käsittele yksi kerrallaan: tunnista carrier -> hae status
  var results = [];
  for (var j = 0; j < todo.length; j++) {
    var row = todo[j];
    var id = row.tracking;
    var carrier = detectCarrierFromTracking_(id);
    var res = null;

    try {
      if (!carrier) {
        // jos ei tunnistunut, kokeile AUTO_SCAN_ORDER -listaa
        res = tryScanAcrossCarriers_(id, { debug: false, useCacheSeconds: 600 });
      } else {
        res = TRK_trackByCarrierEnhanced({ carrier: carrier, id: id }, { debug: false, useCacheSeconds: 600 })[0];
      }
    } catch (e) {
      res = {
        id: id, carrier: carrier || '', delivered: false, status: 'ERROR',
        lastEvent: null, events: [], error: String(e && e.message || e)
      };
    }

    // Poimi created & delivered päivät sekä laske SLA
    var createdDt = pickCreatedDate_(res.events || []);
    var deliveredDt = pickDeliveredDate_(res.events || []);
    var daysToDeliver = (createdDt && deliveredDt) ? daysBetween_(createdDt, deliveredDt) : '';
    
    // Laske SLA käyttäen maakohtaisia rajoja (SLA_RAJAT)
    var country = row.country || guessCountryFromEvents_(res.events || []);
    var slaResult = SLA_computeRuleBased_(res.events || [], country);

    results.push({
      id: id,
      carrier: res.carrier || carrier || '',
      delivered: !!res.delivered,
      status: res.status || '',
      createdISO: createdDt ? createdDt.toISOString() : '',
      deliveredISO: deliveredDt ? deliveredDt.toISOString() : '',
      daysToDeliver: daysToDeliver,
      country: country,
      slaStatus: slaResult.status,
      slaTransportDays: slaResult.transportDays !== null ? slaResult.transportDays : '',
      slaLimitDays: slaResult.slaLimitDays !== null ? slaResult.slaLimitDays : '',
      lastEventTime: res.lastEvent && res.lastEvent.time ? res.lastEvent.time : '',
      lastEventDesc: res.lastEvent && res.lastEvent.description ? res.lastEvent.description : '',
      eventsJson: JSON.stringify(res.events || [])
    });
  }

  // Kirjoita tulokset
  var shOut = ss.getSheetByName(DEFAULT_RESULTS_SHEET) || ss.insertSheet(DEFAULT_RESULTS_SHEET);
  ensureAdhocResultsHeader_(shOut);
  var outRows = results.map(function (r) {
    return [
      r.id, r.carrier, r.delivered, r.status,
      r.createdISO, r.deliveredISO, r.daysToDeliver, r.country,
      r.slaStatus, r.slaTransportDays, r.slaLimitDays,
      yearWeek_(r.deliveredISO || r.createdISO || r.lastEventTime || ''),
      r.lastEventTime, r.lastEventDesc, r.eventsJson
    ];
  });
  if (outRows.length) {
    shOut.getRange(shOut.getLastRow() + 1, 1, outRows.length, outRows[0].length).setValues(outRows);
  }

  // Rakenna viikkotasoinen KPI (keskimääräinen viive maittain)
  ADHOC_BuildWeeklyKPI();
  SpreadsheetApp.getActive().toast('Ad hoc -päivitys valmis: ' + outRows.length + ' riviä.');
  return { processed: todo.length, written: outRows.length };
}

/* ------- KPI / raportointi ------- */

function ADHOC_BuildWeeklyKPI() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DEFAULT_RESULTS_SHEET);
  if (!sh) return;

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  var hdr = data[0];
  var idxCountry = hdr.indexOf('Country');
  var idxWeek = hdr.indexOf('Week');
  var idxDays = hdr.indexOf('DaysToDeliver');

  var groups = {}; // key = country|week
  for (var i = 1; i < data.length; i++) {
    var cty = String(data[i][idxCountry] || '').trim() || 'UNKNOWN';
    var wk = String(data[i][idxWeek] || '').trim() || 'UNKNOWN';
    var days = Number(data[i][idxDays] || '');
    if (!wk || isNaN(days)) continue;
    var key = cty + '|' + wk;
    if (!groups[key]) groups[key] = { sum: 0, cnt: 0 };
    groups[key].sum += days;
    groups[key].cnt += 1;
  }

  var out = [['Country','Week','AvgDays','Count']];
  Object.keys(groups).sort().forEach(function (k) {
    var parts = k.split('|');
    var g = groups[k];
    out.push([parts[0], parts[1], g.cnt ? (g.sum / g.cnt) : '', g.cnt]);
  });

  var kpi = ss.getSheetByName(DEFAULT_KPI_SHEET) || ss.insertSheet(DEFAULT_KPI_SHEET);
  kpi.clearContents();
  kpi.getRange(1,1,out.length, out[0].length).setValues(out);
}

/* ------- Apufunktiot ------- */

function extractSpreadsheetIdFromUrl_(url) {
  var m = String(url || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : '';
}

function findColumnByHints_(header, hints, fallbackExact) {
  var idx = -1;
  var lower = header.map(function (h) { return h.toLowerCase(); });
  for (var i = 0; i < lower.length && idx < 0; i++) {
    var h = lower[i];
    for (var j = 0; j < hints.length; j++) {
      var hint = String(hints[j] || '').toLowerCase();
      if (!hint) continue;
      if (h === hint || h.indexOf(hint) >= 0) { idx = i; break; }
    }
  }
  if (idx < 0 && fallbackExact) {
    idx = header.indexOf(fallbackExact);
  }
  return idx;
}

function addExistingFromSheet_(ss, sheetName, setOut) {
  if (!sheetName) return;
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  // yritä löytää seurantakolumni headerin perusteella
  var hdr = data[0].map(function (h){ return String(h || '').toLowerCase(); });
  var colIdx = hdr.indexOf('package number');
  if (colIdx < 0) colIdx = hdr.indexOf('tracking');
  if (colIdx < 0) colIdx = hdr.indexOf('seuranta');

  // fallback: jos tulostaulussa, ensimmäinen sarake on ID
  if (colIdx < 0 && sheetName === DEFAULT_RESULTS_SHEET) colIdx = 0;

  if (colIdx < 0) return;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][colIdx] || '').trim();
    if (id) setOut.add(id);
  }
}

function addExistingFromDriveFolder_(folderId, setOut) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var name = (f.getName() || '').toLowerCase();
    try {
      if (name.endsWith('.csv')) {
        var csv = Utilities.parseCsv(f.getBlob().getDataAsString('UTF-8'));
        if (csv && csv.length) {
          var hdr = csv[0].map(function (h){ return String(h || '').toLowerCase(); });
          var idx = hdr.indexOf('package number');
          if (idx < 0) idx = hdr.indexOf('tracking');
          if (idx < 0) idx = hdr.indexOf('seuranta');
          if (idx >= 0) {
            for (var i = 1; i < csv.length; i++) {
              var id = String((csv[i] && csv[i][idx]) || '').trim();
              if (id) setOut.add(id);
            }
          }
        }
      }
      // .xlsx ja muut: ohitetaan (tai tee konversio jos tarvitaan)
    } catch (e) {
      // hiljainen ohitus
    }
  }
}

function detectCarrierFromTracking_(id) {
  var s = String(id || '').trim().toUpperCase();
  if (!s) return '';

  // Yleisheuriikat (täydennä tarvittaessa omien koodien pohjalta)
  var rules = [
    { re: /^JJFI[0-9A-Z]+$/, carrier: 'POSTI' },            // Posti JJFI...
    { re: /^JKFI[0-9A-Z]+$/, carrier: 'POSTI' },            // Posti JKFI...
    { re: /^[A-Z]{2}[0-9]{9}[A-Z]{2}$/, carrier: 'POSTI' }, // UPU / CN..FI
    { re: /^003\d{8,}$/, carrier: 'GLS' },                  // GLS viite
    { re: /^00\d{8,}$/, carrier: 'GLS' },
    { re: /^3S[A-Z0-9]+$/, carrier: 'DHL' },                // DHL 3S...
    { re: /^\d{10}$/, carrier: 'DHL' },                     // DHL usein 10-num
    { re: /^PNO[A-Z0-9]+$/, carrier: 'BRING' },             // Bring PNO...
    { re: /^MH[A-Z0-9]+$/, carrier: 'MATKAHUOLTO' },        // MH...
    { re: /^\d{12,}$/, carrier: 'KAUKOKIITO' }              // Kaukokiito (pitkä numerokoodi)
  ];
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].re.test(s)) return rules[i].carrier;
  }
  return ''; // tuntematon -> skannataan
}

function tryScanAcrossCarriers_(id, opts) {
  // Lue AUTO_SCAN_ORDER, esim: "posti,gls,matkahuolto,dhl,bring"
  var orderCsv = _Props_get('AUTO_SCAN_ORDER', 'posti,gls,matkahuolto,dhl,bring');
  var order = String(orderCsv || '').split(',').map(function (x){ return String(x || '').trim().toUpperCase(); }).filter(Boolean);

  for (var i = 0; i < order.length; i++) {
    var c = order[i];
    try {
      var res = TRK_trackByCarrierEnhanced({ carrier: c, id: id }, opts || {});
      if (res && res[0] && (res[0].events || []).length) return res[0];
      // jos ei tapahtumia, kokeile seuraavaa
    } catch (e) {
      // jatka seuraavaan
    }
  }
  // kaikki epäonnistuivat
  return { id: id, carrier: '', delivered: false, status: 'NOT_FOUND', events: [] };
}

/**
 * Country-specific SLA limits (in days).
 * Defines the maximum allowed delivery time for each country.
 * Used by SLA_computeRuleBased_ to determine if delivery meets SLA.
 * 
 * Key: Country code (ISO 3166-1 alpha-2)
 * Value: Maximum delivery days
 * 
 * Examples:
 * - Finland (FI): 2 days
 * - Sweden (SE): 3 days
 * - Norway (NO): 3 days
 * - Denmark (DK): 3 days
 * - Estonia (EE): 2 days
 * - Default: 5 days for unknown countries
 * 
 * Note: Both 'UK' and 'GB' are supported for United Kingdom (GB is official ISO code)
 */
var SLA_RAJAT = {
  'FI': 2,    // Finland - domestic, fast delivery
  'SE': 3,    // Sweden - Nordic neighbor
  'NO': 3,    // Norway - Nordic neighbor
  'DK': 3,    // Denmark - Nordic neighbor
  'EE': 2,    // Estonia - Baltic, close
  'LV': 3,    // Latvia - Baltic
  'LT': 3,    // Lithuania - Baltic
  'DE': 4,    // Germany - Central Europe
  'PL': 4,    // Poland - Eastern Europe
  'NL': 4,    // Netherlands - Western Europe
  'BE': 4,    // Belgium - Western Europe
  'FR': 5,    // France - Western Europe
  'ES': 5,    // Spain - Southern Europe
  'IT': 5,    // Italy - Southern Europe
  'GB': 4,    // Great Britain (official ISO code)
  'UK': 4     // United Kingdom (common alternative code)
};

var ACCEPTED_PATTERNS_ = [
  /accepted/i, /received/i, /handed\s*over/i, /lodged/i, /picked\s*up/i,
  /vastaanotettu/i, /luovutettu/i, /postitettu/i, /rekisteröity/i
];
var DELIVERED_PATTERNS_ = [
  /delivered/i, /signed/i, /delivered\s+to\s+recipient/i,
  /toimitettu/i, /luovutettu\s+vastaanottajalle/i, /utlevert/i
];

/**
 * Pick created/departure date from tracking events.
 * Identifies when package journey started (accepted, picked up, handed over).
 * 
 * Priority logic:
 * 1. Look for "accepted"/"received"/"handed over" events (package entered system)
 * 2. Use earliest such event as start date
 * 3. Fallback to earliest event if no accepted event found
 * 
 * This function is used by:
 * - SLA_computeRuleBased_ for transport time start date
 * - ADHOC_ProcessSheet for delivery time calculation
 * - Transport time = pickDeliveredDate_ - pickCreatedDate_
 * 
 * @param {Array} events - Array of tracking events
 * @return {Date|null} Created/departure date or null
 */
function pickCreatedDate_(events) {
  var best = null;
  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    var t = parseIsoDate_(e && e.time);
    var dsc = (e && e.description) ? String(e.description) : '';
    if (!t) continue;
    if (ACCEPTED_PATTERNS_.some(function (re){ return re.test(dsc); })) {
      if (!best || t < best) best = t;
    }
  }
  // fallback: aikaisin tapahtuma
  if (!best) {
    for (var j = 0; j < events.length; j++) {
      var tt = parseIsoDate_(events[j] && events[j].time);
      if (tt && (!best || tt < best)) best = tt;
    }
  }
  return best;
}

/**
 * Pick delivered/closing date from tracking events.
 * Identifies when package was actually delivered based on event descriptions
 * and location information. Uses tracking location to verify true delivery.
 * 
 * Priority logic:
 * 1. Look for "delivered" events in tracking history with location info
 * 2. Use earliest delivered event (first delivery attempt)
 * 3. Verify delivery by checking location field is populated
 * 
 * This function is used by:
 * - SLA_computeRuleBased_ for SLA calculation
 * - ADHOC_ProcessSheet for delivery time analysis
 * - Transport time calculations (departure → delivery)
 * 
 * @param {Array} events - Array of tracking events with time, description, location
 * @return {Date|null} Delivered date or null if not delivered
 */
function pickDeliveredDate_(events) {
  var best = null;
  var bestWithLocation = null;
  
  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    var t = parseIsoDate_(e && e.time);
    var dsc = (e && e.description) ? String(e.description) : '';
    var loc = (e && e.location) ? String(e.location) : '';
    if (!t) continue;
    
    // Check if event indicates delivery (using description patterns)
    if (DELIVERED_PATTERNS_.some(function (re){ return re.test(dsc); })) {
      // Track earliest delivered event overall
      if (!best || t < best) {
        best = t;
      }
      
      // Prefer events with location information (indicates actual delivery point)
      // Location helps verify that package truly reached destination
      if (loc && (!bestWithLocation || t < bestWithLocation)) {
        bestWithLocation = t;
      }
    }
  }
  
  // Return event with location if found, otherwise return any delivered event
  return bestWithLocation || best;
}

/**
 * Compute SLA status based on country-specific rules.
 * Calculates transport time from departure (created) date to delivery date,
 * then compares against country-specific SLA limit from SLA_RAJAT.
 * 
 * Uses:
 * - pickCreatedDate_(events) to get departure/start date
 * - pickDeliveredDate_(events) to get arrival/delivered date (with location verification)
 * - guessCountryFromEvents_(events) to determine destination country
 * - daysBetween_(d1, d2) to calculate transport time
 * - SLA_RAJAT[country] to get country-specific limit
 * 
 * @param {Array} events - Array of tracking events
 * @param {string} countryHint - Optional country code hint (if known from other sources)
 * @return {Object} SLA result with status, days, limit, country
 * 
 * Example result:
 * {
 *   status: 'OK' | 'LATE' | 'PENDING' | 'UNKNOWN',
 *   transportDays: 2,
 *   slaLimitDays: 2,
 *   country: 'FI',
 *   createdDate: Date object,
 *   deliveredDate: Date object
 * }
 * 
 * Test cases:
 * - FI domestic: created Mon 10:00, delivered Wed 12:00 → 2 days, SLA limit 2 → OK
 * - SE delivery: created Mon, delivered Fri → 4 days, SLA limit 3 → LATE
 * - NO delivery: created Mon, delivered Wed → 2 days, SLA limit 3 → OK
 * - Unknown country: uses default limit 5 days
 */
function SLA_computeRuleBased_(events, countryHint) {
  var result = {
    status: 'UNKNOWN',
    transportDays: null,
    slaLimitDays: null,
    country: '',
    createdDate: null,
    deliveredDate: null
  };
  
  // Get created and delivered dates from events
  var createdDate = pickCreatedDate_(events);
  var deliveredDate = pickDeliveredDate_(events);
  
  result.createdDate = createdDate;
  result.deliveredDate = deliveredDate;
  
  // Determine country (use hint if provided, otherwise guess from events)
  var country = countryHint ? String(countryHint).trim().toUpperCase() : '';
  if (!country || country.length !== 2) {
    country = guessCountryFromEvents_(events);
  }
  result.country = country;
  
  // Get country-specific SLA limit (default to 5 days if country unknown)
  var slaLimit = SLA_RAJAT[country] || 5;
  result.slaLimitDays = slaLimit;
  
  // If not delivered yet, status is PENDING
  if (!deliveredDate) {
    result.status = 'PENDING';
    return result;
  }
  
  // If no created date, cannot calculate SLA
  if (!createdDate) {
    result.status = 'UNKNOWN';
    return result;
  }
  
  // Calculate transport time using daysBetween_
  var transportDays = daysBetween_(createdDate, deliveredDate);
  result.transportDays = transportDays;
  
  // Compare against SLA limit
  if (transportDays <= slaLimit) {
    result.status = 'OK';
  } else {
    result.status = 'LATE';
  }
  
  return result;
}

/**
 * Guess country from tracking events.
 * Examines location field in events to extract country code.
 * Used for country-specific SLA calculation.
 * @param {Array} events - Array of tracking events
 * @return {string} Country code (e.g., 'FI', 'SE') or empty string
 */
function guessCountryFromEvents_(events) {
  // Yritä poimia viimeisistä tapahtumista maa (jos mainitaan)
  for (var i = events.length - 1; i >= 0; i--) {
    var loc = String((events[i] && events[i].location) || '').trim();
    if (!loc) continue;
    // Etsi 2‑kirjaiminen maakoodi tai maa‑nimi (hyvin kevyt)
    var m = loc.match(/\b([A-Z]{2})\b/);
    if (m) return m[1];
    // heikko fallback: jos loppuosa näyttää maatiedolta
    if (loc.split(',').length >= 2) {
      var last = loc.split(',').pop().trim();
      if (last && last.length >= 2 && last.length <= 30) return last;
    }
  }
  return '';
}

function ensureAdhocResultsHeader_(sh) {
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'ID','Carrier','Delivered','Status',
      'CreatedISO','DeliveredISO','DaysToDeliver','Country',
      'SLA_Status','SLA_TransportDays','SLA_LimitDays',
      'Week','LastEventTime','LastEventDesc','Events(JSON)'
    ]);
  }
}

function parseIsoDate_(s) { if (!s) return null; var t = Date.parse(String(s)); return isNaN(t) ? null : new Date(t); }
/**
 * Calculate days between two dates (transport time calculation).
 * Used for SLA and lead time calculations.
 * @param {Date} d1 - Start date (departure/created)
 * @param {Date} d2 - End date (delivered/closing)
 * @return {number} Days between dates (rounded)
 */
function daysBetween_(d1, d2) { var ms = d2.getTime() - d1.getTime(); return Math.round(ms / 86400000); }

function yearWeek_(isoStr) {
  if (!isoStr) return '';
  var d = parseIsoDate_(isoStr); if (!d) return '';
  var tmp = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  var dayNum = tmp.getUTCDay() || 7;
  tmp.setUTCDate(tmp.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
  var week = Math.ceil((((tmp - yearStart) / 86400000) + 1) / 7);
  var y = tmp.getUTCFullYear();
  return y + '-W' + ('0' + week).slice(-2);
}
