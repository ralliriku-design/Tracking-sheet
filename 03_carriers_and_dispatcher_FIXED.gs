/**
 * 03_carriers_and_dispatcher.gs — CLEAN FIXED VERSION
 * Fixes: balanced braces, no stray tokens, safe fallbacks.
 * 
 * Depends on helpers that should exist elsewhere:
 *  - TRK_props_(key)
 *  - trkRateLimitWait_(carrier)
 *  - autoTuneRateLimitOn429_(carrier, addMs, capMs)
 *  - getCfgInt_(key, fallback)           (optional, used in bulk elsewhere)
 *  - fmtDateTime_(date)                  (optional fallback inside)
 *  - logError_(context, error)           (optional)
 */

/* ---------- Small utils (local, safe) ---------- */

function TRK_fmt_(d) {
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"); }
  catch(e){ try { return new Date(d).toISOString(); } catch(_) { return String(d||''); } }
}

function TRK_pickLatestEvent_(events) {
  if (!Array.isArray(events) || !events.length) return null;
  const ts = function(e){
    const cand = (e && (e.eventDateTime || e.eventTime || e.timestamp || e.dateTime || e.dateIso || e.date || e.time)) || '';
    const d = new Date(cand);
    return isNaN(d) ? 0 : d.getTime();
  };
  var arr = events.slice();
  arr.sort(function(a,b){ return ts(a) - ts(b); });
  return arr[arr.length-1];
}

function TRK_safeFetch_(url, opt) {
  try { return UrlFetchApp.fetch(url, opt); }
  catch(e) {
    return {
      getResponseCode: function(){ return 0; },
      getContentText: function(){ return String(e); },
      getAllHeaders: function(){ return {}; }
    };
  }
}

function TRK_canon_(raw){
  try {
    if (typeof canonicalCarrier_ === 'function') return canonicalCarrier_(raw);
  } catch(_) {}
  var s = String(raw||'').toLowerCase();
  if (/posti/.test(s)) return 'posti';
  if (/gls/.test(s)) return 'gls';
  if (/dhl/.test(s)) return 'dhl';
  if (/bring/.test(s)) return 'bring';
  if (/matkahuolto|\bmh\b/.test(s)) return 'matkahuolto';
  if (/kaukokiito/.test(s)) return 'kaukokiito';
  return 'other';
}

/* ---------- POSTI ---------- */

function TRK_trackPosti(code) {
  var carrierName = 'Posti';
  trkRateLimitWait_(carrierName);

  var tokenUrl = TRK_props_('POSTI_TOKEN_URL');
  var basic    = TRK_props_('POSTI_BASIC');
  var trackUrl = (TRK_props_('POSTI_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));

  if (tokenUrl && basic && trackUrl) {
    var tok = TRK_safeFetch_(tokenUrl, {
      method:'post',
      headers:{ 'Authorization':'Basic ' + basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload:{ grant_type:'client_credentials' },
      muteHttpExceptions:true
    });
    var access = '';
    try { access = JSON.parse(tok.getContentText()).access_token || ''; } catch(_){}
    if (access) {
      var res = TRK_safeFetch_(trackUrl, {
        method:'get',
        headers:{ 'Authorization':'Bearer ' + access, 'Accept':'application/json' },
        muteHttpExceptions:true
      });
      var http = res.getResponseCode ? res.getResponseCode() : 0;
      var body = ''; try { body = res.getContentText() || ''; } catch(_){}
      if (http === 429) {
        autoTuneRateLimitOn429_('POSTI', 100, 1500);
        return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
      }
      if (http >= 200 && http < 300) {
        try {
          var j = JSON.parse(body);
          var pickText = function(v){ return (v && typeof v === 'object') ? (v.fi || v.en || v.sv || v.name || v.description || '') : (v || ''); };
          var status='', time='', location='';
          if (Array.isArray(j && j.parcelShipments) && j.parcelShipments.length) {
            var p = j.parcelShipments[0];
            var last = TRK_pickLatestEvent_(p.events || []) || p.latestEvent || {};
            status = pickText(last.eventDescription) || pickText(last.eventShortName) || last.status || '';
            time   = last.eventDateTime || last.eventTime || last.timestamp || last.dateTime || last.dateIso || last.date || '';
            location = (last.location && (pickText(last.location.displayName) || pickText(last.location.name))) || last.location || last.city || '';
          }
          if (!status && ((j && j.shipments && j.shipments.length) || (j && j.items && j.items.length))) {
            var ship = (j.shipments && j.shipments[0]) || (j.items && j.items[0]) || null;
            var last2 = (ship && ship.events && ship.events.length) ? (TRK_pickLatestEvent_(ship.events) || {}) : {};
            status = pickText(last2.eventDescription) || last2.description || last2.status || '';
            time   = last2.timestamp || last2.dateTime || last2.dateIso || last2.date || '';
            location = (last2.location && (pickText(last2.location.displayName) || pickText(last2.location.name))) || last2.location || last2.city || '';
          }
          if (!status && Array.isArray(j && j.freightShipments) && j.freightShipments.length) {
            var f = j.freightShipments[0];
            var last3 = (Array.isArray(f.events) && f.events.length) ? (TRK_pickLatestEvent_(f.events) || {}) : (f.latestEvent || f.latestStatus || {});
            status = pickText(last3.eventDescription) || last3.statusText || last3.status || (f.latestStatus && pickText(f.latestStatus.description)) || '';
            time   = last3.timestamp || last3.dateTime || last3.dateIso || last3.date || (f.latestStatus && f.latestStatus.timestamp) || '';
            location = (last3.location && (pickText(last3.location.displayName) || pickText(last3.location.name))) || last3.location || last3.city || '';
          }
          if (status) return { carrier: carrierName, found:true, status: status, time: time, location: location, raw: body.slice(0,2000) };
        } catch(_){}
      }
    }
  }

  // Fallback to legacy Basic endpoint (Atlas)
  var fbUrl = (TRK_props_('POSTI_TRK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  var fbBasic = TRK_props_('POSTI_TRK_BASIC');
  var fbUser  = TRK_props_('POSTI_TRK_USER');
  var fbPass  = TRK_props_('POSTI_TRK_PASS');
  if (fbUrl && (fbBasic || (fbUser && fbPass))) {
    if (!fbBasic) fbBasic = Utilities.base64Encode(fbUser + ':' + fbPass);
    var res2 = TRK_safeFetch_(fbUrl, {
      method:'get',
      headers:{ 'Authorization':'Basic ' + fbBasic, 'Accept':'application/json' },
      muteHttpExceptions:true
    });
    var http2 = res2.getResponseCode ? res2.getResponseCode() : 0;
    var body2 = ''; try { body2 = res2.getContentText() || ''; } catch(_){}
    if (http2 === 429) {
      autoTuneRateLimitOn429_('POSTI', 100, 1500);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body2.slice(0,1000), retryAfter:15 };
    }
    if (http2 >= 400) return { carrier: carrierName, status:'HTTP_'+http2, raw: body2.slice(0,1000) };
    try {
      var j2 = JSON.parse(body2);
      var ev = j2.events || j2.trackingEvents || (j2.parcelShipments && j2.parcelShipments[0] && j2.parcelShipments[0].events) || [];
      ev = Array.isArray(ev) ? ev : [];
      ev.sort(function(a,b){
        var ta = Date.parse((a && (a.eventDateTime || a.timestamp || a.dateTime || a.date)) || '') || 0;
        var tb = Date.parse((b && (b.eventDateTime || b.timestamp || b.dateTime || b.date)) || '') || 0;
        return ta - tb;
      });
      var last = ev[ev.length-1] || {};
      var status2 = last.eventDescription || last.description || last.status || '';
      var time2   = last.eventDateTime || last.timestamp || last.dateTime || last.date || '';
      var location2 = (last.location && (last.location.name || last.location.displayName)) || last.location || last.city || '';
      if (status2) return { carrier: carrierName, found:true, status: status2, time: time2, location: location2, raw: body2.slice(0,2000) };
    } catch(_){}
    return { carrier: carrierName, status:'NO_DATA', raw: body2.slice(0,2000) };
  }

  return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
}

/* ---------- GLS ---------- */

function TRK_trackGLS(code) {
  var carrierName = 'GLS';
  trkRateLimitWait_(carrierName);

  // FI Customer API v2 (x-api-key)
  var base   = TRK_props_('GLS_FI_TRACK_URL');
  var key    = TRK_props_('GLS_FI_API_KEY');
  var senderSingle = TRK_props_('GLS_FI_SENDER_ID');
  var senderList   = TRK_props_('GLS_FI_SENDER_IDS');
  var sender = senderSingle || (senderList ? String(senderList).split(',')[0].trim() : '');
  var method = (TRK_props_('GLS_FI_METHOD') || 'POST').toUpperCase();

  if (base && key) {
    var url = base;
    var headers = { 'x-api-key': key, 'accept':'application/json' };
    var opt = { muteHttpExceptions:true, headers: headers };

    if (method === 'POST') {
      opt.method = 'post';
      opt.contentType = 'application/json';
      opt.payload = JSON.stringify(sender ? { senderId: sender, parcelNo: code } : { parcelNo: code });
    } else {
      url = url + (url.indexOf('?')>=0 ? '&' : '?') + (sender ? ('senderId='+encodeURIComponent(sender)+'&') : '') + 'parcelNo=' + encodeURIComponent(code);
      opt.method = 'get';
    }

    var res = TRK_safeFetch_(url, opt);
    var http = res.getResponseCode ? res.getResponseCode() : 0;
    var body = ''; try { body = res.getContentText() || ''; } catch(_){}
    if (http === 429) {
      autoTuneRateLimitOn429_('GLS', 200, 2000);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
    }
    if (http >= 200 && http < 300) {
      try {
        var j = JSON.parse(body);
        var p = (j && (j.parcels && j.parcels[0])) || j.parcel || j;
        if (p) {
          var ev = Array.isArray(p.events) ? p.events : [];
          var last = ev.length ? (TRK_pickLatestEvent_(ev) || {}) : {};
          var status = last.description || p.status || last.code || '';
          var time   = p.statusDateTime || last.eventDateTime || last.timestamp || last.dateTime || '';
          var location = [last.city || '', last.postalCode || '', last.country || ''].filter(function(x){return !!x;}).join(' ');
          if (status) return { carrier: carrierName, found:true, status: status, time: time, location: location, raw: body.slice(0,2000) };
        }
        // general fallback in FI response
        var ev2 = (j && (j.events || j.trackingEvents || (j.parcel && j.parcel.events) || (j.parcelShipments && j.parcelShipments[0] && j.parcelShipments[0].events) || (j.parcels && j.parcels[0] && j.parcels[0].events) || (j.items && j.items[0] && j.items[0].events) || j.eventList || j.tracking)) || [];
        ev2 = Array.isArray(ev2) ? ev2 : [];
        var last2 = TRK_pickLatestEvent_(ev2) || {};
        var status2 = last2.eventDescription || last2.description || (last2.eventShortName && (last2.eventShortName.fi || last2.eventShortName.en)) || last2.name || last2.status || last2.code || '';
        var time2   = last2.eventTime || last2.dateTime || last2.timestamp || last2.dateIso || last2.date || last2.time || '';
        var location2 = last2.locationName || last2.location || last2.depot || last2.city || '';
        if (status2) return { carrier: carrierName, found:true, status: status2, time: time2, location: location2, raw: body.slice(0,2000) };
      } catch(_){}
    }
    // else fall through to global OAuth
  }

  // Global OAuth + track-and-trace
  var tokenUrl = TRK_props_('GLS_TOKEN_URL');
  var basic    = TRK_props_('GLS_BASIC');
  var trackUrl = (TRK_props_('GLS_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));

  if (tokenUrl && basic && trackUrl) {
    var tok = TRK_safeFetch_(tokenUrl, {
      method:'post',
      headers:{ 'Authorization':'Basic ' + basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload:{ grant_type:'client_credentials' },
      muteHttpExceptions:true
    });
    var access = ''; try { access = JSON.parse(tok.getContentText()).access_token || ''; } catch(_){}
    if (!access) return { carrier: carrierName, status:'TOKEN_FAIL', raw: (tok && tok.getContentText && tok.getContentText() || '').slice(0,500) };

    var res2 = TRK_safeFetch_(trackUrl, {
      method:'get',
      headers:{ 'Authorization':'Bearer ' + access, 'Accept':'application/json' },
      muteHttpExceptions:true
    });
    var http2 = res2.getResponseCode ? res2.getResponseCode() : 0;
    var body2 = ''; try { body2 = res2.getContentText() || ''; } catch(_){}
    if (http2 === 429) {
      autoTuneRateLimitOn429_('GLS', 200, 2000);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body2.slice(0,1000), retryAfter:15 };
    }
    if (http2 >= 400) return { carrier: carrierName, status:'HTTP_'+http2, raw: body2.slice(0,1000) };
    try {
      var j2 = JSON.parse(body2);
      var it = (j2 && (j2.parcels && j2.parcels[0])) || (j2 && j2.items && j2.items[0]) || j2.shipment || (j2.consignment && j2.consignment[0]) || j2;
      var ev3 = (it && (it.events || it.milestones || it.statusHistory || (it.trackAndTraceInfo && it.trackAndTraceInfo.events) || it.trackings)) || [];
      ev3 = Array.isArray(ev3) ? ev3 : [];
      var last3 = TRK_pickLatestEvent_(ev3) || {};
      var status3 = last3.description || last3.eventDescription || last3.name || last3.status || last3.code || '';
      var time3   = last3.dateTime || last3.timestamp || last3.dateIso || last3.date || last3.time || '';
      var location3 = (last3.location && (last3.location.name || last3.location.displayName)) || last3.location || last3.depot || last3.city || '';
      return { carrier: carrierName, found: !!status3, status: status3, time: time3, location: location3, raw: body2.slice(0,2000) };
    } catch(_){}
  }

  return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
}

/* ---------- DHL ---------- */

function TRK_trackDHL(code) {
  var carrierName = 'DHL';
  trkRateLimitWait_(carrierName);

  var url = (TRK_props_('DHL_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  var key = TRK_props_('DHL_API_KEY');
  if (!url || !key) return { carrier: carrierName, status:'MISSING_CREDENTIALS' };

  var res = TRK_safeFetch_(url, { method:'get', headers:{ 'DHL-API-Key': key, 'Accept':'application/json' }, muteHttpExceptions:true });
  var http = res.getResponseCode ? res.getResponseCode() : 0;
  var body = res.getContentText ? (res.getContentText() || '') : '';
  if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter: 15 };
  if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };

  try {
    var j = JSON.parse(body);
    if (Array.isArray(j && j.parcels) && j.parcels.length) {
      var p = j.parcels[0];
      var ev = Array.isArray(p.events) ? p.events : [];
      var last = ev.length ? (TRK_pickLatestEvent_(ev) || {}) : {};
      var status = last.description || p.status || last.code || '';
      var time   = p.statusDateTime || last.eventDateTime || last.timestamp || last.dateTime || '';
      var location = [last.city || '', last.postalCode || '', last.country || ''].filter(function(x){return !!x;}).join(', ');
      if (status) return { carrier: carrierName, found:true, status: status, time: time, location: location, raw: body.slice(0,2000) };
    }
    if (Array.isArray(j && j.shipments) && j.shipments.length) {
      var s = j.shipments[0];
      var ev2 = Array.isArray(s.events) ? s.events : [];
      if (!ev2.length && s.status) ev2 = [s.status];
      var last2 = ev2.length ? (TRK_pickLatestEvent_(ev2) || {}) : (s.status || {});
      var status2 = last2.description || last2.status || (s.status && s.status.status) || '';
      var time2   = last2.timestamp || last2.dateTime || last2.date || (s.status && s.status.timestamp) || '';
      var addr = last2.location && last2.location.address;
      var location2 = '';
      if (addr) location2 = [addr.addressLocality, (addr.countryCode || addr.addressCountryCode)].filter(function(x){return !!x;}).join(', ');
      return { carrier: carrierName, found: !!status2, status: status2, time: time2, location: location2, raw: body.slice(0,2000) };
    }
  } catch(_){}

  return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
}

/* ---------- Bring ---------- */

function TRK_trackBring(code) {
  var carrierName = 'Bring';
  trkRateLimitWait_(carrierName);

  var url = (TRK_props_('BRING_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  var uid = TRK_props_('BRING_UID');
  var key = TRK_props_('BRING_KEY');
  var cli = TRK_props_('BRING_CLIENT_URL');
  if (!url) return { carrier: carrierName, status:'MISSING_CREDENTIALS' };

  var res = TRK_safeFetch_(url, { method:'get',
    headers:{ 'X-MyBring-API-Uid': uid, 'X-MyBring-API-Key': key, 'X-Bring-Client-URL': cli, 'Accept':'application/json' },
    muteHttpExceptions:true
  });
  var http = res.getResponseCode ? res.getResponseCode() : 0;
  var body = res.getContentText ? (res.getContentText() || '') : '';
  if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
  if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };

  try {
    var j = JSON.parse(body);
    var c = (j && (j.consignments && j.consignments[0])) || (j && j.items && j.items[0]) || j;
    var ev = Array.isArray(c && c.events) ? c.events : [];
    var last = TRK_pickLatestEvent_(ev) || {};
    var status = last.description || last.eventDescription || last.status || '';
    var time   = last.dateIso || last.timestamp || last.dateTime || last.date || '';
    var location = last.location || last.city || '';
    return { carrier: carrierName, found: !!status, status: status, time: time, location: location, raw: body.slice(0,2000) };
  } catch(_){
    return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
  }
}

/* ---------- Matkahuolto ---------- */

function TRK_trackMH(code) {
  trkRateLimitWait_('Matkahuolto');

  var url = (TRK_props_('MH_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  var basic = TRK_props_('MH_BASIC');
  if (!url || !basic) return { carrier:'Matkahuolto', status:'MISSING_CREDENTIALS' };

  var res = TRK_safeFetch_(url, { method:'get', headers:{ 'Authorization':'Basic '+basic, 'Accept':'application/json' }, muteHttpExceptions:true });
  var body = res.getContentText ? (res.getContentText() || '') : '';

  var status='', time='', location='';
  try {
    var j = JSON.parse(body);
    var first = (j && j[0]) || (j && j.items && j.items[0]) || j.consignment || j.shipment || j;
    var ev = (first && (first.events || first.history)) || [];
    ev = Array.isArray(ev) ? ev : [];
    var last = TRK_pickLatestEvent_(ev) || {};
    status = last.status || last.description || last.eventDescription || last.name || '';
    time   = last.time || last.timestamp || last.dateTime || last.dateIso || last.date || '';
    location = last.location || last.place || last.depot || last.city || '';
  } catch(_){}
  return { carrier:'Matkahuolto', found: !!status, status: status, time: time, location: location, raw: body.slice(0,2000) };
}

/* ---------- Dispatcher ---------- */

function TRK_trackByCarrier_(carrier, code) {
  var c = String(TRK_canon_(carrier));
  if (c === 'kaukokiito') return { carrier: 'Kaukokiito', status: 'UNSUPPORTED_CARRIER' };
  if (c === 'gls')        return TRK_trackGLS(code);
  if (c === 'posti')      return TRK_trackPosti(code);
  if (c === 'dhl')        return TRK_trackDHL(code);
  if (c === 'bring')      return TRK_trackBring(code);
  if (c === 'matkahuolto')return TRK_trackMH(code);
  return { carrier: carrier, status:'UNKNOWN_CARRIER' };
}
