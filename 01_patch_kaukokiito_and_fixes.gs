/**
 * PATCH 01 — Kaukokiito + korjaukset (Ralli Logistics)
 * ----------------------------------------------------
 * Tämä tiedosto on DROP-IN -päivitys aiempaan "Shipment Tracking Toolkit" -projektiisi.
 * Liitä tämä samaan Apps Script -projektiin uutena .gs-tiedostona (esim. "patch_kaukokiito.gs").
 *
 * SISÄLTÖ:
 *  - canonicalCarrier_(): lisätty Kaukokiito
 *  - parseDateFlexible_(): korjattu dd.MM.yyyy HH:mm -tulkinta
 *  - TRK_trackKaukokiito(): uusi seurantafunktio (API tai HTML fallback)
 *  - TRK_trackByCarrier_(): reititys päivitetty tunnistamaan Kaukokiito
 *  - readMissingProps_(): lisätty KAUKO_* avaimet tarkistuksiin
 *  - seedKnownAccountsAndKeys(): lisätty KAUKO_TRACK_URL oletus
 *  - Valikkoon: menuRefreshCarrier_KAUKO()
 *
 * HUOM:
 *  - Tämä tiedosto KORVAA projektissasi samannimiset funktiot (canonicalCarrier_, parseDateFlexible_,
 *    TRK_trackByCarrier_, readMissingProps_, seedKnownAccountsAndKeys). Poista vanhat tai anna tämän
 *    olla viimeisenä tiedostolistassa, jolloin uudempi määritelmä varjostaa vanhan.
 *  - DELIVERED_KEYWORDS ja muut helperit tulevat aiemmasta paketistasi — älä duplikoi niitä.
 */

/********************* CARRIER-NORMALISOINTI **********************/
function canonicalCarrier_(s){
  const c = String(s||'').toLowerCase();
  if (/kaukokiito|kki/.test(c)) return 'kaukokiito';
  if (/posti/.test(c)) return 'posti';
  if (/gls/.test(c)) return 'gls';
  if (/dhl/.test(c)) return 'dhl';
  if (/bring/.test(c)) return 'bring';
  if (/matkahuolto|mh/.test(c)) return 'matkahuolto';
  return 'other';
}

/********************* PÄIVÄYKSEN JOUSTAVA TULKINTA **************/
function parseDateFlexible_(val){
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) return val;
  let s = String(val);
  // ISO tai RFC
  let d = new Date(s);
  if (!isNaN(d)) return d;
  // dd.MM.yyyy [HH:mm[:ss]]
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m){
    const dd = m[1].padStart(2,'0');
    const MM = m[2].padStart(2,'0');
    const yyyy = m[3];
    const hh = (m[4]||'00').padStart(2,'0');
    const mi = (m[5]||'00').padStart(2,'0');
    const ss = (m[6]||'00').padStart(2,'0');
    d = new Date(`${yyyy}-${MM}-${dd}T${hh}:${mi}:${ss}`);
    if (!isNaN(d)) return d;
  }
  // yyyy-MM-dd HH:mm:ss
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (m2){
    d = new Date(`${m2[1]}-${m2[2]}-${m2[3]}T${m2[4]}:${m2[5]}:${m2[6]||'00'}`);
    if (!isNaN(d)) return d;
  }
  return null;
}

/********************* KAUKOKIITO TRACKING ************************
 * Tuki kahdelle moodille:
 *  A) API (JSON): aseta Script Properties:
 *     - KAUKO_TRACK_URL   esim. https://api.kaukokiito.example/track?code={{code}}
 *     - KAUKO_API_KEY     (vaihtoehtoisesti KAUKO_BASIC = Base64("user:pass"))
 *  B) HTML fallback (kevyt): aseta
 *     - KAUKO_SCRAPE_URL  esim. https://www.kaukokiito.fi/seuranta/{{code}}
 *
 * Palauttaa {carrier, found, status, time, location, raw, [retryAfter]}
 ******************************************************************/
function TRK_trackKaukokiito(code){
  const carrierName = 'Kaukokiito';
  const sp = PropertiesService.getScriptProperties();
  const url = (sp.getProperty('KAUKO_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const apiKey = sp.getProperty('KAUKO_API_KEY') || '';
  const basic  = sp.getProperty('KAUKO_BASIC') || '';
  const scrapeUrl = (sp.getProperty('KAUKO_SCRAPE_URL') || '').replace('{{code}}', encodeURIComponent(code));

  if (!url && !scrapeUrl){
    return { carrier: carrierName, status: 'MISSING_CREDENTIALS' };
  }

  if (url){
    // JSON API -kutsu
    const headers = { 'Accept':'application/json' };
    if (apiKey) headers['x-api-key'] = apiKey;
    if (basic) headers['Authorization'] = 'Basic ' + basic;
    const res = TRK_safeFetch_(url, { method:'get', headers, muteHttpExceptions:true });
    if (!res.ok){
      if (res.code === 429){
        autoTuneRateLimitOn429_('KAUKO', 200, 2000);
        logHttpError_(carrierName, code, 'TRK_trackKaukokiito', url, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
        return { carrier: carrierName, status:'RATE_LIMIT_429', raw:String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
      }
      if (res.code === 404) return { carrier: carrierName, status:'NOT_FOUND' };
      logHttpError_(carrierName, code, 'TRK_trackKaukokiito', url, res.code, res.status, res.text, null);
      return { carrier: carrierName, status: res.status || 'HTTP_'+res.code, raw: String(res.text||'').slice(0,1000) };
    }
    try{
      const data = JSON.parse(res.text);
      const latest = kk_pickLatestEventFromUnknownJson_(data);
      if (!latest) return { carrier: carrierName, found:true, status:'IN_TRANSIT', time:'', location:'', raw:'' };
      return {
        carrier: carrierName,
        found: true,
        status: latest.status || latest.description || latest.text || '',
        time: latest.time ? fmtDateTime_(latest.time) : '',
        location: latest.location || latest.city || latest.postalCode || '',
        raw: ''
      };
    } catch(e){
      // JSON parse epäonnistui → yritä HTML fallbackia jos asetettu
    }
  }

  if (scrapeUrl){
    // HTML fallback: etsi "delivered/toimitettu/luovutettu" tai viimeinen tapahtuma
    const res = TRK_safeFetch_(scrapeUrl, { method:'get', headers:{ 'Accept':'text/html' }, muteHttpExceptions:true });
    if (!res.ok){
      if (res.code === 429){
        autoTuneRateLimitOn429_('KAUKO', 200, 2000);
        logHttpError_(carrierName, code, 'TRK_trackKaukokiito_scrape', scrapeUrl, res.code, 'RATE_LIMIT_429', res.text, res.retryAfter);
        return { carrier: carrierName, status:'RATE_LIMIT_429', raw:String(res.text||'').slice(0,1000), retryAfter: res.retryAfter };
      }
      if (res.code === 404) return { carrier: carrierName, status:'NOT_FOUND' };
      logHttpError_(carrierName, code, 'TRK_trackKaukokiito_scrape', scrapeUrl, res.code, res.status, res.text, null);
      return { carrier: carrierName, status: res.status || 'HTTP_'+res.code, raw: String(res.text||'').slice(0,1000) };
    }
    const html = String(res.text||'');
    // Poimi aikaleima (yyyy-mm-dd hh:mm tai dd.mm.yyyy hh:mm)
    const reIso = /(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2})?)/g;
    const reFi  = /(\d{1,2}\.\d{1,2}\.\d{4}[ T]\d{1,2}:\d{2}(?::\d{2})?)/g;
    const times = [];
    let m;
    while ((m = reIso.exec(html)) !== null){ times.push(parseDateFlexible_(m[1])); }
    while ((m = reFi.exec(html))  !== null){ times.push(parseDateFlexible_(m[1])); }
    const bestTime = times.filter(Boolean).sort((a,b)=>b-a)[0] || '';
    const lowered = html.toLowerCase();
    const delivered = (typeof DELIVERED_KEYWORDS !== 'undefined') && DELIVERED_KEYWORDS.some(w => lowered.includes(String(w).toLowerCase()));
    const status = delivered ? 'DELIVERED' : (bestTime ? 'IN_TRANSIT' : 'NOT_FOUND');
    return {
      carrier: carrierName,
      found: status !== 'NOT_FOUND',
      status: status,
      time: bestTime ? fmtDateTime_(bestTime) : '',
      location: '',
      raw: ''
    };
  }

  return { carrier: carrierName, status: 'NOT_FOUND' };
}

// Heuristiikka: etsi JSONista "event"-taulut ja poimi uusin aikaleiman perusteella
function kk_pickLatestEventFromUnknownJson_(obj){
  let best = null;
  function visit(node){
    if (!node) return;
    if (Array.isArray(node)){
      node.forEach(visit);
    } else if (typeof node === 'object'){
      // jos näyttää eventiltä
      const keys = Object.keys(node).map(k=>k.toLowerCase());
      const hasStatus = keys.some(k=>/status|description|eventname|text/.test(k));
      const timeKey = keys.find(k=>/time|timestamp|date/.test(k));
      if (hasStatus && timeKey){
        const t = parseDateFlexible_(node[timeKey]);
        if (t && (!best || t > best.time)){
          best = {
            status: String(node.status || node.description || node.eventName || node.text || ''),
            time: t,
            location: String(node.location || node.city || node.postalCode || node.locationCode || '')
          };
        }
      }
      for (var k in node){ if (node.hasOwnProperty(k)) visit(node[k]); }
    }
  }
  visit(obj);
  return best;
}

/********************* REITITYS: lisää Kaukokiito *****************/
function TRK_trackByCarrier_(carrier, code){
  const c = String(carrier||'').toLowerCase();
  try {
    if (c.includes('kaukokiito') || c.includes('kki')){
      return TRK_trackKaukokiito(code);
    } else if (c.includes('matkahuolto') || c.includes('mh')){
      return TRK_trackMatkahuolto(code);
    } else if (c.includes('posti') || c.includes('posti.fi') || c.includes('itella')){
      return TRK_trackPosti(code);
    } else if (c.includes('gls')){
      return TRK_trackGLS(code);
    } else if (c.includes('dhl')){
      return TRK_trackDHL(code);
    } else if (c.includes('bring')){
      return TRK_trackBring(code);
    } else {
      return { carrier: carrier, status: 'UNKNOWN_CARRIER' };
    }
  } catch(e){
    const msg = String(e && e.message ? e.message : e);
    if (/MISSING_CREDENTIALS/.test(msg)) {
      return { carrier: carrier, status: 'MISSING_CREDENTIALS' };
    } else if (/NOT_FOUND/.test(msg) || /NO_DATA/.test(msg)) {
      return { carrier: carrier, status: 'NOT_FOUND' };
    } else if (/RATE_LIMIT_429/.test(msg)){
      return { carrier: carrier, status: 'RATE_LIMIT_429' };
    } else {
      return { carrier: carrier, status: 'ERROR', error: msg };
    }
  }
}

/********************* PROPS: lisää KAUKO-avaimet *****************/
function readMissingProps_(){
  const sp = PropertiesService.getScriptProperties();
  const have = sp.getProperties();
  const mustKeys = [
    // Matkahuolto
    'MH_TRACK_URL','MH_BASIC',
    // Posti
    'POSTI_TOKEN_URL','POSTI_BASIC','POSTI_TRACK_URL',
    // (tuki myös vanhoille nimille jos käytössä)
    // GLS (joko FI tai OAuth fallback)
    'GLS_TOKEN_URL','GLS_BASIC','GLS_TRACK_URL',
    'GLS_FI_TRACK_URL','GLS_FI_API_KEY',
    // DHL
    'DHL_API_KEY','DHL_TRACK_URL',
    // Bring
    'BRING_TRACK_URL','BRING_UID','BRING_KEY','BRING_CLIENT_URL',
    // Kaukokiito (vähintään toinen: API tai SCRAPE)
    'KAUKO_TRACK_URL',  // jos käytät virallista APIa
    // throttlaus-konfig
    'BULK_BACKOFF_MINUTES_BASE'
  ];
  const optionalKeys = [
    'POSTI_TRK_URL','POSTI_TRK_BASIC','POSTI_TRK_USER','POSTI_TRK_PASS',
    'GLS_FI_SENDER_ID','GLS_FI_SENDER_IDS','GLS_FI_METHOD',
    'RATE_MINMS_DHL','RATE_MINMS_GLS','RATE_MINMS_POSTI',
    'KAUKO_API_KEY','KAUKO_BASIC','KAUKO_SCRAPE_URL'
  ];
  return {
    missingMust: mustKeys.filter(k => !have[k]),
    missingOpt:  optionalKeys.filter(k => !have[k])
  };
}

function seedKnownAccountsAndKeys(){
  const defaults = {
    MH_TRACK_URL:     'https://extservices.matkahuolto.fi/mpaketti/public/tracking?ids={{code}}',
    POSTI_TOKEN_URL:  'https://oauth2.posti.com/oauth/token',
    POSTI_TRACK_URL:  'https://api.posti.fi/tracking/7/shipments/trackingnumbers/{{code}}',
    GLS_TOKEN_URL:    'https://api.gls-group.net/oauth2/v2/token',
    GLS_TRACK_URL:    'https://api.gls-group.net/track-and-trace-v1/tracking/simple/references/{{code}}',
    DHL_TRACK_URL:    'https://api-eu.dhl.com/track/shipments?trackingNumber={{code}}',
    BRING_TRACK_URL:  'https://api.bring.com/tracking/api/v2/tracking.json?q={{code}}',
    KAUKO_TRACK_URL:  'https://api.kaukokiito.example/track?code={{code}}', // VAIHDA OIKEAAN
    BULK_BACKOFF_MINUTES_BASE: '5'
  };
  const sp = PropertiesService.getScriptProperties();
  Object.entries(defaults).forEach(function([k,v]){
    if (!sp.getProperty(k)) sp.setProperty(k, String(v));
  });
  SpreadsheetApp.getActive().toast('Oletusavaimet asetettu (sis. KAUKO_TRACK_URL placeholderin)');
}

/********************* VALIKKO: Kaukokiito ************************/
function menuRefreshCarrier_KAUKO(){
  refreshStatuses_Filtered(getActiveSheetName_(), ['kaukokiito','kki'], false);
}