/******************************************************
 * Tracking adapters + safe fetch + token caching
 ******************************************************/

function TRK_trackByCarrier_(carrier, code){
  const c = String(carrier||'').toLowerCase();
  try {
    if (c.includes('matkahuolto') || c.includes('mh')) return TRK_trackMatkahuolto(code);
    if (c.includes('posti') || c.includes('itella')) return TRK_trackPosti(code);
    if (c.includes('gls')) return TRK_trackGLS(code);
    if (c.includes('dhl')) return TRK_trackDHL(code);
    if (c.includes('bring')) return TRK_trackBring(code);
    return { carrier, status: 'UNKNOWN_CARRIER' };
  } catch(e){
    const msg = String(e && e.message || e);
    if (/MISSING_CREDENTIALS/.test(msg)) return { carrier, status: 'MISSING_CREDENTIALS' };
    if (/NOT_FOUND|NO_DATA/.test(msg)) return { carrier, status: 'NOT_FOUND' };
    if (/RATE_LIMIT_429/.test(msg)) return { carrier, status: 'RATE_LIMIT_429' };
    return { carrier, status: 'ERROR', error: msg };
  }
}

function TRK_safeFetch_(url, params){
  try {
    const resp = UrlFetchApp.fetch(url, params);
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    const headers = (resp.getAllHeaders ? resp.getAllHeaders() : resp.getHeaders()) || {};
    const ra = headers['Retry-After'] || headers['retry-after'];
    const retryAfter = ra ? parseInt(ra, 10) : null;

    if (code === 429) return { ok:false, code, status:'RATE_LIMIT_429', text, headers, retryAfter };
    if (code >= 200 && code < 300) return { ok:true, code, text, headers };
    if (code === 404) return { ok:false, code, status:'NOT_FOUND', text, headers };
    return { ok:false, code, status:'HTTP_'+code, text, headers };
  } catch(e){
    return { ok:false, code:0, status:'NETWORK_ERROR', text:String(e && e.message || e) };
  }
}

function trkComputeNextAtFromRetryAfter_(retryAfterSeconds){
  const s = Math.max(1, parseInt(retryAfterSeconds || 0, 10));
  return new Date(Date.now() + s*1000);
}

/** Token cache helper (Script Properties) */
function getCachedToken_(prefix, fetcher){
  const now = Math.floor(Date.now()/1000);
  const token = spGet_(`${prefix}_TOKEN`);
  const exp   = Number(spGet_(`${prefix}_EXPIRES`) || '0');
  if (token && exp && exp - now > 60) return token;
  const t = fetcher();
  if (!t || !t.access_token) return '';
  const expiry = t.expires_in ? now + Number(t.expires_in) : now + 3600;
  spSet_(`${prefix}_TOKEN`, t.access_token);
  spSet_(`${prefix}_EXPIRES`, String(expiry));
  return t.access_token;
}

/******** Matkahuolto ********/
function TRK_trackMatkahuolto(code){
  const carrierName = 'Matkahuolto';
  const url = (spGet_('MH_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const basic = spGet_('MH_BASIC');
  if (!url || !basic) return { carrier: carrierName, status: 'MISSING_CREDENTIALS' };

  trkThrottle_('MH');
  const res = TRK_safeFetch_(url, {
    method:'get',
    headers:{ 'Authorization':'Basic '+basic, 'Accept':'application/json' },
    muteHttpExceptions:true
  });

  if (!res.ok){
    if (res.code === 429){
      autoTuneRateLimitOn429_('MH', 200, 2000);
      logHttpError_(carrierName, code, 'TRK_trackMatkahuolto', url, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw:String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
    }
    if (res.status === 'NOT_FOUND') return { carrier: carrierName, status:'NOT_FOUND' };
    logHttpError_(carrierName, code, 'TRK_trackMatkahuolto', url, res.code, res.status, res.text, null);
    return { carrier: carrierName, status: res.status, raw: String(res.text||'').slice(0,1000) };
  }

  try {
    const data = JSON.parse(res.text);
    const cons = data.consignment && data.consignment[0];
    const events = (cons && cons.event) || [];
    if (!events.length) return { carrier: carrierName, found:true, status:'IN_TRANSIT', time:'', location:'', raw:'' };
    const latest = events.slice().sort((a,b) => new Date(b.eventTime) - new Date(a.eventTime))[0];
    return {
      carrier: carrierName, found: true,
      status: latest.status || latest.description || '',
      time: latest.eventTime ? fmtDateTime_(latest.eventTime) : '',
      location: latest.location || '',
      raw: ''
    };
  } catch(e){
    return { carrier: carrierName, status:'PARSING_ERROR', raw: String(res.text||'').slice(0,1000) };
  }
}

/******** Posti (OAuth) ********/
function TRK_trackPosti(code){
  const carrierName = 'Posti';
  const tokenUrl = spGet_('POSTI_TOKEN_URL');
  const basic    = spGet_('POSTI_BASIC');
  const url      = (spGet_('POSTI_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  if (!url) return { carrier: carrierName, status: 'MISSING_CREDENTIALS' };

  const token = (tokenUrl && basic) ? getCachedToken_('POSTI', () => {
    const tRes = TRK_safeFetch_(tokenUrl, {
      method: 'post',
      headers: { 'Authorization':'Basic '+basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload: 'grant_type=client_credentials&scope=shipment.read',
      muteHttpExceptions:true
    });
    if (!tRes.ok) return {};
    try { return JSON.parse(tRes.text) } catch(e){ return {}; }
  }) : '';

  trkThrottle_('POSTI');
  const headers = token ? { 'Authorization':'Bearer '+token, 'Accept':'application/json' } : { 'Accept':'application/json' };
  const res = TRK_safeFetch_(url, { method:'get', headers, muteHttpExceptions:true });

  if (!res.ok){
    if (res.code === 429){
      autoTuneRateLimitOn429_('POSTI', 200, 2000);
      logHttpError_(carrierName, code, 'TRK_trackPosti', url, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw:String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
    }
    if (res.status === 'NOT_FOUND') return { carrier: carrierName, status:'NOT_FOUND' };
    logHttpError_(carrierName, code, 'TRK_trackPosti', url, res.code, res.status, res.text, null);
    return { carrier: carrierName, status: res.status, raw: String(res.text||'').slice(0,1000) };
  }

  try {
    const data = JSON.parse(res.text);
    const ship = data.shipments && data.shipments[0];
    const evts = (ship && ship.events) || [];
    if (!evts.length) return { carrier: carrierName, found:true, status:'IN_TRANSIT', time:'', location:'', raw:'' };
    const latest = evts.slice().sort((a,b)=> new Date(b.timestamp) - new Date(a.timestamp))[0];
    return {
      carrier: carrierName, found: true,
      status: latest.description || latest.eventType || '',
      time: latest.timestamp ? fmtDateTime_(latest.timestamp) : '',
      location: latest.locationCode || latest.location || '',
      raw: ''
    };
  } catch(e){
    return { carrier: carrierName, status:'PARSING_ERROR', raw: String(res.text||'').slice(0,1000) };
  }
}

/******** GLS (FI x-api-key, fallback OAuth) ********/
function TRK_trackGLS(code){
  const carrierName = 'GLS';
  const fiUrl    = (spGet_('GLS_FI_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const apiKey   = spGet_('GLS_FI_API_KEY') || '';
  const senderId = spGet_('GLS_FI_SENDER_ID') || '';
  const method   = spGet_('GLS_FI_METHOD') || 'GET';

  if (fiUrl){
    trkThrottle_('GLS');
    const headers = { 'x-api-key': apiKey, 'Accept':'application/json' };
    if (senderId) headers['x-ib-sender-id'] = senderId;
    const res = TRK_safeFetch_(fiUrl, { method, headers, muteHttpExceptions:true });
    if (!res.ok){
      if (res.code === 429){
        autoTuneRateLimitOn429_('GLS', 200, 2000);
        logHttpError_(carrierName, code, 'TRK_trackGLS_v2', fiUrl, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
        return { carrier: carrierName, status:'RATE_LIMIT_429', raw: String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
      }
      if (res.status === 'NOT_FOUND') return { carrier: carrierName, status:'NOT_FOUND' };
      logHttpError_(carrierName, code, 'TRK_trackGLS_v2', fiUrl, res.code, res.status, res.text, null);
      return { carrier: carrierName, status: res.status, raw: String(res.text||'').slice(0,1000) };
    }
    try {
      const data = JSON.parse(res.text);
      const pkg = data.statuses && data.statuses[0];
      const hist = (pkg && pkg.history) || [];
      if (!hist.length) return { carrier: carrierName, found:true, status:'IN_TRANSIT', time:'', location:'', raw:'' };
      const latest = hist.slice().sort((a,b)=> new Date(b.dateTime) - new Date(a.dateTime))[0];
      return {
        carrier: carrierName, found: true,
        status: latest.statusText || latest.status || '',
        time: latest.dateTime ? fmtDateTime_(latest.dateTime) : '',
        location: latest.location || '',
        raw: ''
      };
    } catch(e){
      return { carrier: carrierName, status:'PARSING_ERROR', raw: String(res.text||'').slice(0,1000) };
    }
  }

  // OAuth fallback (global)
  const tokenUrl = spGet_('GLS_TOKEN_URL'), basic = spGet_('GLS_BASIC');
  const trackUrl = (spGet_('GLS_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  if (!tokenUrl || !basic || !trackUrl) return { carrier: carrierName, status: 'MISSING_CREDENTIALS' };

  const token = getCachedToken_('GLS', () => {
    const tRes = TRK_safeFetch_(tokenUrl, {
      method:'post',
      headers:{ 'Authorization':'Basic '+basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload:'grant_type=client_credentials',
      muteHttpExceptions:true
    });
    if (!tRes.ok) return {};
    try { return JSON.parse(tRes.text) } catch(e){ return {}; }
  });

  trkThrottle_('GLS');
  const res = TRK_safeFetch_(trackUrl, { method:'get', headers: { 'Authorization':'Bearer '+token, 'Accept':'application/json' }, muteHttpExceptions:true });

  if (!res.ok){
    if (res.code === 429){
      autoTuneRateLimitOn429_('GLS', 200, 2000);
      logHttpError_(carrierName, code, 'TRK_trackGLS_oauth', trackUrl, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw:String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
    }
    if (res.status === 'NOT_FOUND') return { carrier: carrierName, status:'NOT_FOUND' };
    logHttpError_(carrierName, code, 'TRK_trackGLS_oauth', trackUrl, res.code, res.status, res.text, null);
    return { carrier: carrierName, status: res.status, raw: String(res.text||'').slice(0,1000) };
  }

  try {
    const data = JSON.parse(res.text);
    const shipment = data.track && data.track.shipment && data.track.shipment[0];
    const events = (shipment && shipment.events) || [];
    if (!events.length) return { carrier: carrierName, found:true, status:'IN_TRANSIT', time:'', location:'', raw:'' };
    const latest = events.slice().sort((a,b)=> new Date(b.date) - new Date(a.date))[0];
    return {
      carrier: carrierName, found: true,
      status: latest.statusDescription || latest.status || latest.eventName || '',
      time: latest.date ? fmtDateTime_(latest.date) : '',
      location: latest.location || '',
      raw: ''
    };
  } catch(e){
    return { carrier: carrierName, status:'PARSING_ERROR', raw: String(res.text||'').slice(0,1000) };
  }
}

/******** DHL ********/
function TRK_trackDHL(code){
  const carrierName = 'DHL';
  const url = (spGet_('DHL_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const key = spGet_('DHL_API_KEY');
  if (!url || !key) return { carrier: carrierName, status: 'MISSING_CREDENTIALS' };

  trkThrottle_('DHL');
  const res = TRK_safeFetch_(url, {
    method:'get',
    headers:{ 'DHL-API-Key': key, 'Accept':'application/json' },
    muteHttpExceptions:true
  });

  if (!res.ok){
    if (res.code === 429){
      autoTuneRateLimitOn429_('DHL', 200, 2000);
      logHttpError_(carrierName, code, 'TRK_trackDHL', url, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw:String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
    }
    if (res.status === 'NOT_FOUND') return { carrier: carrierName, status:'NOT_FOUND' };
    logHttpError_(carrierName, code, 'TRK_trackDHL', url, res.code, res.status, res.text, null);
    return { carrier: carrierName, status: res.status, raw: String(res.text||'').slice(0,1000) };
  }

  try {
    const data = JSON.parse(res.text);
    const ship = data.shipments && data.shipments[0];
    if (!ship || !ship.status) return { carrier: carrierName, status:'NOT_FOUND' };
    const ev = (ship.events && ship.events[0]) || {};
    return {
      carrier: carrierName, found: true,
      status: ship.status.status || ship.status.statusCode || '',
      time: ev.timestamp ? fmtDateTime_(ev.timestamp) : '',
      location: ev.location && (ev.location.address && ev.location.address.addressLocality || ev.location.name) || '',
      raw: ''
    };
  } catch(e){
    return { carrier: carrierName, status:'PARSING_ERROR', raw: String(res.text||'').slice(0,1000) };
  }
}

/******** Bring ********/
function TRK_trackBring(code){
  const carrierName = 'Bring';
  const url = (spGet_('BRING_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const uid = spGet_('BRING_UID'), key = spGet_('BRING_KEY');
  const clientUrl = spGet_('BRING_CLIENT_URL') || 'https://example.com';
  if (!url || !uid || !key) return { carrier: carrierName, status: 'MISSING_CREDENTIALS' };

  trkThrottle_('BRING');
  const res = TRK_safeFetch_(url, {
    method:'get',
    headers:{
      'X-MyBring-API-Uid': uid,
      'X-MyBring-API-Key': key,
      'X-Bring-Client-URL': clientUrl,
      'Accept':'application/json',
      'api-version':'2'
    },
    muteHttpExceptions:true
  });

  if (!res.ok){
    if (res.code === 429){
      autoTuneRateLimitOn429_('BRING', 200, 2000);
      logHttpError_(carrierName, code, 'TRK_trackBring', url, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw:String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
    }
    if (res.status === 'NOT_FOUND') return { carrier: carrierName, status:'NOT_FOUND' };
    logHttpError_(carrierName, code, 'TRK_trackBring', url, res.code, res.status, res.text, null);
    return { carrier: carrierName, status: res.status, raw: String(res.text||'').slice(0,1000) };
  }

  try {
    const data = JSON.parse(res.text);
    const set = data.consignmentSet && data.consignmentSet[0];
    const pkg = set && set.packageSet && set.packageSet[0];
    const events = (pkg && pkg.eventSet) || [];
    if (!events.length) return { carrier: carrierName, found:true, status:'IN_TRANSIT', time:'', location:'', raw:'' };
    const latest = events.slice().sort((a,b)=> new Date(b.dateIso) - new Date(a.dateIso))[0];
    return {
      carrier: carrierName, found: true,
      status: latest.description || latest.status || '',
      time: latest.dateIso ? fmtDateTime_(latest.dateIso) : '',
      location: latest.postalCode || latest.countryCode || latest.city || '',
      raw: ''
    };
  } catch(e){
    return { carrier: carrierName, status:'PARSING_ERROR', raw: String(res.text||'').slice(0,1000) };
  }
}