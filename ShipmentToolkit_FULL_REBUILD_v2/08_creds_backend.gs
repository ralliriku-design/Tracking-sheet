
// 08_creds_backend.gs
function credGetProps(){
  const sp = PropertiesService.getScriptProperties();
  const keys = [
    'CARRIER_DEFAULT_COL','TRACKING_HEADER_HINT','TRACKING_DEFAULT_COL','REFRESH_MIN_AGE_MINUTES','AUTO_SCAN_ORDER','DEDUP_JOIN_CARRIER',
    'POSTI_TOKEN_URL','POSTI_BASIC','POSTI_TRACK_URL','POSTI_TRK_URL','POSTI_TRK_BASIC','POSTI_TRK_USER','POSTI_TRK_PASS',
    'GLS_FI_TRACK_URL','GLS_FI_API_KEY','GLS_FI_SENDER_ID','GLS_FI_SENDER_IDS','GLS_FI_METHOD',
    'GLS_TOKEN_URL','GLS_BASIC','GLS_TRACK_URL',
    'DHL_TRACK_URL','DHL_API_KEY',
    'BRING_TRACK_URL','BRING_UID','BRING_KEY','BRING_CLIENT_URL',
    'MH_TRACK_URL','MH_BASIC'
  ];
  const props = {}; keys.forEach(k => props[k] = sp.getProperty(k) || '');
  return props;
}
function credSaveProps(data){
  const sp = PropertiesService.getScriptProperties();
  try {
    Object.keys(data||{}).forEach(k => {
      if (data[k] !== '') sp.setProperty(k, String(data[k]));
      else sp.deleteProperty(k);
    });
    return { ok:true };
  } catch(e) {
    return { ok:false, msg:String(e) };
  }
}
function showCredentialsHub(){
  const html = HtmlService.createHtmlOutputFromFile('07_sidebar_creds').setTitle('Integraatioavaimet').setWidth(760).setHeight(820);
  SpreadsheetApp.getUi().showSidebar(html);
}
