
// 03_carriers_and_dispatcher.gs — canonical mapping + API calls
function canonicalCarrier_(s){
  const c = String(s||'').toLowerCase().replace(/\s+/g,' ').replace(/[()]/g,'').trim();
  if (/bring/.test(c)) return 'bring';
  if (/matkahuolto|matka ?huolto|mh\b/.test(c)) return 'matkahuolto';
  if (/\bdhl\b|dhl parcel europe|dhl finland|dhl express/.test(c)) return 'dhl';
  if (/gls|gls finland|general logistics/.test(c)) return 'gls';
  if (/posti api|posti oy|posti fi|posten fi|posti group/.test(c)) return 'posti';
  if (/kaukokiito/.test(c)) return 'kaukokiito';
  return 'other';
}

function TRK_safeFetch_(url, opt) {
  try { return UrlFetchApp.fetch(url, opt); }
  catch(e) {
    return { getResponseCode: () => 0, getContentText: () => String(e), getAllHeaders: () => ({}) };
  }
}

function pickLatestEvent_(events) {
  if (!Array.isArray(events) || !events.length) return null;
  const ts = e => {
    const cand = e.eventDateTime || e.eventTime || e.timestamp || e.dateTime || e.dateIso || e.date || e.time || '';
    const d = new Date(cand);
    return isNaN(d) ? 0 : d.getTime();
  };
  return events.slice().sort((a,b) => ts(a)-ts(b))[events.length-1];
}

// ===== Posti
function TRK_trackPosti(code) {
  const carrierName = 'Posti';
  const tokenUrl = TRK_props_('POSTI_TOKEN_URL');
  const basic = TRK_props_('POSTI_BASIC');
  const trackUrl = (TRK_props_('POSTI_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  if (tokenUrl && basic && trackUrl) {
    const tok = TRK_safeFetch_(tokenUrl, {
      method:'post',
      headers:{ 'Authorization':'Basic ' + basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload:{ grant_type:'client_credentials' },
      muteHttpExceptions:true
    });
    let access=''; try { access = JSON.parse(tok.getContentText()).access_token || ''; } catch(e) {}
    if (access) {
      const res = TRK_safeFetch_(trackUrl, {
        method:'get',
        headers:{ 'Authorization':'Bearer ' + access, 'Accept':'application/json' },
        muteHttpExceptions:true
      });
      const http = res.getResponseCode ? res.getResponseCode() : 0;
      let body=''; try { body = res.getContentText() || ''; } catch(e) {}
      if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
      if (http >= 200 && http < 300) {
        try {
          const j = JSON.parse(body);
          const pickText = v => (v && typeof v === 'object') ? (v.fi || v.en || v.sv || v.name || v.description || '') : (v || '');
          let status='', time='', location='';
          if (Array.isArray(j?.parcelShipments) && j.parcelShipments.length) {
            const p = j.parcelShipments[0];
            const last = pickLatestEvent_(p.events || []) || p.latestEvent || {};
            status = pickText(last.eventDescription) || pickText(last.eventShortName) || last.status || '';
            time = last.eventDateTime || last.eventTime || last.timestamp || last.dateTime || last.dateIso || last.date || '';
            location = (last.location && (pickText(last.location.displayName) || pickText(last.location.name))) || last.location || last.city || '';
          }
          if (!status && (j?.shipments?.length || j?.items?.length)) {
            const ship = j.shipments?.[0] || j.items?.[0];
            const last = (ship?.events && ship.events.length) ? (pickLatestEvent_(ship.events) || {}) : {};
            status = pickText(last.eventDescription) || last.description || last.status || '';
            time = last.timestamp || last.dateTime || last.dateIso || last.date || '';
            location = (last.location && (pickText(last.location.displayName) || pickText(last.location.name))) || last.location || last.city || '';
          }
          if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
        } catch(e) {}
      }
    }
  }
  // fallback Basic
  const fbUrl = (TRK_props_('POSTI_TRK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  let fbBasic = TRK_props_('POSTI_TRK_BASIC');
  const fbUser = TRK_props_('POSTI_TRK_USER');
  const fbPass = TRK_props_('POSTI_TRK_PASS');
  if (fbUrl && (fbBasic || (fbUser && fbPass))) {
    if (!fbBasic) fbBasic = Utilities.base64Encode(fbUser + ':' + fbPass);
    const res = TRK_safeFetch_(fbUrl, {
      method:'get',
      headers:{ 'Authorization':'Basic ' + fbBasic, 'Accept':'application/json' },
      muteHttpExceptions:true
    });
    const http = res.getResponseCode ? res.getResponseCode() : 0;
    let body=''; try { body = res.getContentText() || ''; } catch(e) {}
    if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
    if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
    try {
      const j = JSON.parse(body);
      let ev = j.events || j.trackingEvents || j.parcelShipments?.[0]?.events || [];
      ev = Array.isArray(ev) ? ev : [];
      const parseTs = e => Date.parse(e?.eventDateTime || e?.timestamp || e?.dateTime || e?.date || '');
      ev.sort((a,b) => (parseTs(a) || 0) - (parseTs(b) || 0));
      const last = ev[ev.length-1] || {};
      const status = last.eventDescription || last.description || last.status || '';
      const time = last.eventDateTime || last.timestamp || last.dateTime || last.date || '';
      const location = last.location?.name || last.location || last.city || '';
      if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
    } catch(e) {}
    return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
  }
  return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
}

// ===== GLS
function TRK_trackGLS(code) {
  const carrierName = 'GLS';
  const base = TRK_props_('GLS_FI_TRACK_URL');
  const key = TRK_props_('GLS_FI_API_KEY');
  const senderSingle = TRK_props_('GLS_FI_SENDER_ID');
  const senderList = TRK_props_('GLS_FI_SENDER_IDS');
  const sender = senderSingle || (senderList ? senderList.split(',')[0].trim() : '');
  const method = (TRK_props_('GLS_FI_METHOD') || 'POST').toUpperCase();
  if (base && key) {
    let url = base;
    const headers = { 'x-api-key': key, 'accept':'application/json' };
    let opt = { muteHttpExceptions:true, headers: headers };
    if (method === 'POST') {
      opt.method = 'post';
      opt.contentType = 'application/json';
      opt.payload = JSON.stringify(sender ? { senderId: sender, parcelNo: code } : { parcelNo: code });
    } else {
      url = url + (url.includes('?') ? '&' : '?') + (sender ? ('senderId='+encodeURIComponent(sender)+'&') : '') + 'parcelNo=' + encodeURIComponent(code);
      opt.method = 'get';
    }
    const res = TRK_safeFetch_(url, opt);
    const http = res.getResponseCode ? res.getResponseCode() : 0;
    let body=''; try { body = res.getContentText() || ''; } catch(e) {}
    if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
    if (http >= 200 && http < 300) {
      try {
        const j = JSON.parse(body);
        const p = j?.parcels?.[0] || j?.parcel || j;
        if (p) {
          const ev = Array.isArray(p.events) ? p.events : [];
          const last = ev.length ? (pickLatestEvent_(ev) || {}) : {};
          const status = last.description || p.status || last.code || '';
          const time = p.statusDateTime || last.eventDateTime || last.timestamp || last.dateTime || '';
          const location = [last.city || '', last.postalCode || '', last.country || ''].filter(Boolean).join(' ');
          if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
        }
        // general fallback
        let ev = j.events || j.trackingEvents || j.parcel?.events || j.parcelShipments?.[0]?.events || j.parcels?.[0]?.events || j.items?.[0]?.events || j.eventList || j.tracking || [];
        ev = Array.isArray(ev) ? ev : [];
        const last = pickLatestEvent_(ev) || {};
        const status = last.eventDescription || last.description || last.eventShortName?.fi || last.eventShortName?.en || last.name || last.status || last.code || '';
        const time = last.eventTime || last.dateTime || last.timestamp || last.dateIso || last.date || last.time || '';
        const location = last.locationName || last.location || last.depot || last.city || '';
        if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
      } catch(e) {}
    }
  }
  // OAuth fallback
  const tokenUrl = TRK_props_('GLS_TOKEN_URL');
  const basic = TRK_props_('GLS_BASIC');
  const trackUrl = (TRK_props_('GLS_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  if (tokenUrl && basic && trackUrl) {
    const tok = TRK_safeFetch_(tokenUrl, {
      method:'post',
      headers:{ 'Authorization':'Basic ' + basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload:{ grant_type:'client_credentials' },
      muteHttpExceptions:true
    });
    let access=''; try { access = JSON.parse(tok.getContentText()).access_token || ''; } catch(e) {}
    if (!access) return { carrier: carrierName, status:'TOKEN_FAIL', raw: tok.getContentText().slice(0,500) };
    const res = TRK_safeFetch_(trackUrl, {
      method:'get',
      headers:{ 'Authorization':'Bearer ' + access, 'Accept':'application/json' },
      muteHttpExceptions:true
    });
    const http = res.getResponseCode ? res.getResponseCode() : 0;
    let body=''; try { body = res.getContentText() || ''; } catch(e) {}
    if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
    if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
    try {
      const j = JSON.parse(body);
      const it = j?.parcels?.[0] || j?.items?.[0] || j?.shipment || j?.consignment?.[0] || j;
      let ev = it?.events || it?.milestones || it?.statusHistory || it?.trackAndTraceInfo?.events || it?.trackings || [];
      ev = Array.isArray(ev) ? ev : [];
      const last = pickLatestEvent_(ev) || {};
      const status = last.description || last.eventDescription || last.name || last.status || last.code || '';
      const time = last.dateTime || last.timestamp || last.dateIso || last.date || last.time || '';
      const location = (last.location && (last.location.name || last.location.displayName)) || last.location || last.depot || last.city || '';
      return { carrier: carrierName, found:!!status, status, time, location, raw: body.slice(0,2000) };
    } catch(e) {}
  }
  return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
}

// ===== DHL
function TRK_trackDHL(code) {
  const carrierName = 'DHL';
  const url = (TRK_props_('DHL_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const key = TRK_props_('DHL_API_KEY');
  if (!url || !key) return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
  const res = TRK_safeFetch_(url, { method:'get', headers:{ 'DHL-API-Key': key, 'Accept':'application/json' }, muteHttpExceptions:true });
  const http = res.getResponseCode ? res.getResponseCode() : 0;
  const body = res.getContentText ? (res.getContentText() || '') : '';
  if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter: 15 };
  if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
  try {
    const j = JSON.parse(body);
    if (Array.isArray(j?.parcels) && j.parcels.length) {
      const p = j.parcels[0];
      const ev = Array.isArray(p.events) ? p.events : [];
      const last = ev.length ? (pickLatestEvent_(ev) || {}) : {};
      const status = last.description || p.status || last.code || '';
      const time = p.statusDateTime || last.eventDateTime || last.timestamp || last.dateTime || '';
      const location = [last.city || '', last.postalCode || '', last.country || ''].filter(Boolean).join(', ');
      if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
    }
    if (Array.isArray(j?.shipments) && j.shipments.length) {
      const s = j.shipments[0];
      let ev = Array.isArray(s.events) ? s.events : [];
      if (!ev.length && s.status) ev = [s.status];
      const last = ev.length ? (pickLatestEvent_(ev) || {}) : s.status || {};
      const status = last.description || last.status || s.status?.status || '';
      const time = last.timestamp || last.dateTime || last.date || s.status?.timestamp || '';
      const addr = last.location?.address;
      let location='';
      if (addr) location = [addr.addressLocality, (addr.countryCode || addr.addressCountryCode)].filter(Boolean).join(', ');
      return { carrier: carrierName, found: !!status, status, time, location, raw: body.slice(0,2000) };
    }
  } catch(e) {}
  return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
}

// ===== Bring
function TRK_trackBring(code) {
  const carrierName = 'Bring';
  const url = (TRK_props_('BRING_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const uid = TRK_props_('BRING_UID'), key = TRK_props_('BRING_KEY'), cli = TRK_props_('BRING_CLIENT_URL');
  if (!url) return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
  const res = TRK_safeFetch_(url, { method:'get', headers:{ 'X-MyBring-API-Uid': uid, 'X-MyBring-API-Key': key, 'X-Bring-Client-URL': cli, 'Accept':'application/json' }, muteHttpExceptions:true });
  const http = res.getResponseCode ? res.getResponseCode() : 0;
  const body = res.getContentText ? (res.getContentText() || '') : '';
  if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
  if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
  try {
    const j = JSON.parse(body);
    const c = j?.consignments?.[0] || j?.items?.[0] || j;
    const ev = Array.isArray(c?.events) ? c.events : [];
    const last = pickLatestEvent_(ev) || {};
    const status = last.description || last.eventDescription || last.status || '';
    const time = last.dateIso || last.timestamp || last.dateTime || last.date || '';
    const location = last.location || last.city || '';
    return { carrier: carrierName, found: !!status, status, time, location, raw: body.slice(0,2000) };
  } catch(e) {
    return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
  }
}

// ===== Matkahuolto
function TRK_trackMH(code) {
  const carrierName = 'Matkahuolto';
  const url = (TRK_props_('MH_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const basic = TRK_props_('MH_BASIC');
  if (!url || !basic) return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
  const res = TRK_safeFetch_(url, { method:'get', headers:{ 'Authorization':'Basic '+basic, 'Accept':'application/json' }, muteHttpExceptions:true });
  const body = res.getContentText() || '';
  let status='', time='', location='';
  try {
    const j = JSON.parse(body);
    const first = j?.[0] || j?.items?.[0] || j?.consignment || j?.shipment || j;
    let ev = first?.events || first?.history || [];
    ev = Array.isArray(ev) ? ev : [];
    const last = pickLatestEvent_(ev) || {};
    status = last.status || last.description || last.eventDescription || last.name || '';
    time = last.time || last.timestamp || last.dateTime || last.dateIso || last.date || '';
    location = last.location || last.place || last.depot || last.city || '';
  } catch(e) {}
  return { carrier: carrierName, found: !!status, status, time, location, raw: body.slice(0,2000) };
}

function TRK_trackByCarrier_(carrier, code) {
  const c = canonicalCarrier_(carrier);
  if (c === 'gls') return TRK_trackGLS(code);
  if (c === 'posti') return TRK_trackPosti(code);
  if (c === 'dhl') return TRK_trackDHL(code);
  if (c === 'bring') return TRK_trackBring(code);
  if (c === 'matkahuolto') return TRK_trackMH(code);
  if (c === 'kaukokiito') return { carrier: 'Kaukokiito', status: 'UNSUPPORTED_CARRIER' };
  return { carrier: carrier, status:'UNKNOWN_CARRIER' };
}
