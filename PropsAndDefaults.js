/******************************************************
 * Credentials (Script Properties) + defaults
 ******************************************************/

function credGetProps() {
  const sp = PropertiesService.getScriptProperties();
  const keys = [
    // Matkahuolto
    'MH_TRACK_URL','MH_BASIC',
    // Posti (OAuth)
    'POSTI_TOKEN_URL','POSTI_BASIC','POSTI_TRACK_URL',
    // GLS FI + OAuth fallback
    'GLS_FI_TRACK_URL','GLS_FI_API_KEY','GLS_FI_SENDER_ID','GLS_FI_METHOD',
    'GLS_TOKEN_URL','GLS_BASIC','GLS_TRACK_URL',
    // DHL
    'DHL_TRACK_URL','DHL_API_KEY',
    // Bring
    'BRING_TRACK_URL','BRING_UID','BRING_KEY','BRING_CLIENT_URL',
    // rate/backoff
    'BULK_BACKOFF_MINUTES_BASE',
    // optional per-carrier throttle ms
    'RATE_MINMS_POSTI','RATE_MINMS_GLS','RATE_MINMS_DHL','RATE_MINMS_MH','RATE_MINMS_BRING'
  ];
  const obj = {}; keys.forEach(k => obj[k] = sp.getProperty(k) || '');
  return obj;
}

function credSaveProps(data) {
  if (!data || typeof data !== 'object') return { ok:false, msg:'No data' };
  const sp = PropertiesService.getScriptProperties();
  Object.keys(data).forEach(k => sp.setProperty(k, String(data[k] ?? '')));
  return { ok:true };
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
    BULK_BACKOFF_MINUTES_BASE: '5'
  };
  const sp = PropertiesService.getScriptProperties();
  Object.entries(defaults).forEach(([k, v]) => { if (!sp.getProperty(k)) sp.setProperty(k, String(v)); });
  SpreadsheetApp.getActive().toast('Defaults set in Script Properties');
}

function readMissingProps_(){
  const have = PropertiesService.getScriptProperties().getProperties();
  const must = [
    'MH_TRACK_URL','MH_BASIC',
    'POSTI_TOKEN_URL','POSTI_BASIC','POSTI_TRACK_URL',
    // GLS FI or OAuth (one path must be complete for GLS usage)
    'DHL_API_KEY','DHL_TRACK_URL',
    'BRING_TRACK_URL','BRING_UID','BRING_KEY','BRING_CLIENT_URL',
    'BULK_BACKOFF_MINUTES_BASE'
  ];
  const opt = [
    'GLS_FI_TRACK_URL','GLS_FI_API_KEY','GLS_FI_SENDER_ID','GLS_FI_METHOD',
    'GLS_TOKEN_URL','GLS_BASIC','GLS_TRACK_URL',
    'RATE_MINMS_POSTI','RATE_MINMS_GLS','RATE_MINMS_DHL','RATE_MINMS_MH','RATE_MINMS_BRING'
  ];
  return { missingMust: must.filter(k => !have[k]), missingOpt: opt.filter(k => !have[k]) };
}

function showMissingProperties(){
  const res = readMissingProps_();
  const ui = SpreadsheetApp.getUi();
  const must = res.missingMust || [], opt = res.missingOpt || [];
  const msg =
    (must.length ? 'PUUTTUU (pakolliset):\n• ' + must.join('\n• ') + '\n\n' : 'Pakolliset OK\n\n') +
    (opt.length ? 'Puuttuu (valinnaiset):\n• ' + opt.join('\n• ') : 'Valinnaiset OK');
  ui.alert('Integraatioavaimet (Script Properties)', msg, ui.ButtonSet.OK);
}