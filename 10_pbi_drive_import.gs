/** 99_defaults_overwrite.gs — Oletusasetusten lataus (ylikirjoittaa) **/
     

function SeedDefaultsOverwrite(){
  const P = PropertiesService.getScriptProperties();
  const SET = (k,v) => { P.setProperty(k, String(v)); return true; };
  const changes = [];
  function put(k,v){
    const before = P.getProperty(k);
    SET(k,v);
    const after = P.getProperty(k);
    changes.push({k,before,after});
  }

  /** --------- POSTI (legacy fallback) --------- **/
  put('POSTI_TOKEN_URL', 'https://oauth2.posti.com/oauth/token');
  put('POSTI_TRACK_URL', 'https://api.posti.fi/tracking/7/shipments/trackingnumbers/{{code}}');
  put('POSTI_BASIC', ''); // OAuth ei käytössä
  put('POSTI_TRK_URL', 'https://atlas.posti.fi/track-shipment-json?ShipmentId={{code}}');
  const postiUser = 'ma_09931637_1P';
  const postiPass = 'ZatuCE8isPeStlSijl4u';
  put('POSTI_TRK_BASIC', Utilities.base64Encode(postiUser + ':' + postiPass));

  /** --------- GLS (FI API) --------- **/
  put('GLS_TOKEN_URL', 'https://api.gls-group.net/oauth2/v2/token');
  put('GLS_BASIC', ''); // OAuth ei käytössä
  put('GLS_TRACK_URL', 'https://api.gls-group.net/track-and-trace-v1/tracking/simple/references/{{code}}');
  put('GLS_FI_API_KEY','vAUg0en8vKC6wufKwGWxIw8hL23IfuzL4u14Ioce9PkFkAqQgnnYTvfFyjnONWzR');
  put('GLS_FI_SENDER_ID','006112,007413,007380,007529,007488,007580,007951,007952');
  put('GLS_FI_METHOD','POST');

  /** --------- DHL --------- **/
  put('DHL_TRACK_URL','https://api-eu.dhl.com/track/shipments?trackingNumber={{code}}');
  put('DHL_API_KEY','MC4grRiZ7dDsokW2ltPG1Qfj17d0enZA');

  /** --------- BRING --------- **/
  put('BRING_TRACK_URL','https://api.bring.com/tracking/api/v2/tracking.json?q={{code}}');
  put('BRING_UID','riku.ralli@ip-agency.fi');
  put('BRING_KEY','5db5bcb9-d8ca-471b-89f2-db625f965228');
  put('BRING_CLIENT_URL','https://omaappisi.fi');

  /** --------- MATKAHUOLTO --------- **/
  put('MH_TRACK_URL','https://extservices.matkahuolto.fi/mpaketti/public/tracking?ids={{code}}');
  put('MH_BASIC','OTQwMzI3Njp5MHRmWkd3ZGpp');

  /** --------- Rate-limit ja bulk --------- **/
  put('RATE_MINMS_POSTI', '500');
  put('RATE_MINMS_GLS', '800');
  put('RATE_MINMS_DHL', '800');
  put('RATE_MINMS_BRING', '800');
  put('RATE_MINMS_MATKAHUOLTO', '500');
  put('BULK_MAX_API_CALLS_PER_RUN', '300');
  put('BULK_BACKOFF_MINUTES_BASE','30');

  /** --------- Gmail & Drive --------- **/
  put('GMAIL_QUERY_PACKAGES','label:"package report" OR subject:"Packages Report" newer_than:90d');
  put('PBI_FOLDER_ID','1G6OyD9vNKq2DTT2YlNfGc48Mt3QPA_Un');
  put('DRIVE_IMPORT_FOLDER_ID','1yAkYYR6hetV3XATEJqg7qvy5NAJrFgKh');

  /** --------- Taulujen nimet --------- **/
  put('ACTION_SHEET','Vaatii_toimenpiteitä');
  put('TARGET_SHEET','Packages');
  put('ARCHIVE_SHEET','Packages_Archive');
  put('KK_PASTE_SHEET','KK_Paste');

  /** --------- Päivitys ja mittarit --------- **/
  put('REFRESH_COOLDOWN_MS', String(6*3600*1000));
  put('DISPATCH_METRICS','true');
  put('SLA_OPEN_DAYS','5');
  put('SLA_PD_DAYS','5');

  /** --------- Näytä yhteenveto --------- **/
  showDefaultsSummary_(changes);
}

function showDefaultsSummary_(changes){
  function esc(s){return String(s||'').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
  function secretKey(k){ return /_BASIC$|_KEY$|_PASS$|_TOKEN$|SECRET|PASSWORD|AUTH/i.test(k); }
  function mask(v){ if(!v) return ''; const s=String(v); return s.slice(0,6)+'...'+s.slice(-4)+' (len:'+s.length+')'; }

  const rows = changes.map(r => {
    const b = secretKey(r.k) ? mask(r.before||'') : (r.before||'');
    const a = secretKey(r.k) ? mask(r.after||'')  : (r.after||'');
    return `<tr>
      <td><code>${esc(r.k)}</code></td>
      <td>${esc(b)}</td>
      <td>${esc(a)}</td>
    </tr>`;
  }).join('');

  const html = HtmlService.createHtmlOutput(
    `<div style="font:14px system-ui;padding:12px">
      <h3>Oletusasetukset asennettu (ylikirjoitettu)</h3>
      <p>Alempana näet ennen → jälkeen (salaisuudet maskattu).</p>
      <table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:#f6f8fa">
          <th style="text-align:left;border:1px solid #e5e7eb;padding:6px 8px">KEY</th>
          <th style="text-align:left;border:1px solid #e5e7eb;padding:6px 8px">Ennen</th>
          <th style="text-align:left;border:1px solid #e5e7eb;padding:6px 8px">Jälkeen</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <p style="color:#666;font-size:12px;margin-top:8px">
        OAuth-käytössä muista täyttää <code>POSTI_BASIC</code> Base64(client_id:client_secret) → token haetaan automaattisesti.
        Legacy-fallback Postille (<code>POSTI_TRK_* </code>) on jo valmis.
      </p>
    </div>`).setWidth(800).setHeight(520).setTitle('Asetukset – yhteenveto');
  SpreadsheetApp.getUi().showSidebar(html);
}
