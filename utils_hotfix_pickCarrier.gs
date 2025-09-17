// ===== HOTFIX: define pickCarrierIndex_ (and helper) =====
function findHeaderIndexIncludes_(headers, substrs){
  const n = (headers || []).map(h => normalize_(h));
  for (let i=0;i<n.length;i++){
    for (const s of (substrs || [])){
      if (n[i].includes(normalize_(s))) return i;
    }
  }
  return -1;
}

const CARRIER_CANDIDATES = [
  'RefreshCarrier','Carrier','Carrier name','CarrierName','Delivery Carrier','Courier','Courier name',
  'LogisticsProvider','Logistics Provider','Shipper','Service provider','Forwarder','Forwarder name',
  'Transporter','Kuljetusliike','Kuljetusyhtiö','Kuljetus','Toimitustapa','Delivery method'
];

function pickCarrierIndex_(hdr){
  const m = headerIndexMap_(hdr);
  // 1) property hint(s)
  const hint = TRK_props_('CARRIER_HEADER_HINT');
  if (hint){
    const names = String(hint).split(',').map(s=>s.trim()).filter(Boolean);
    for (const h of names){
      const direct = colIndexOf_(m, [h]);
      if (direct >= 0) return direct;
      const inc = findHeaderIndexIncludes_(hdr, [h]);
      if (inc >= 0) return inc;
    }
  }
  // 2) standard candidates
  let idx = colIndexOf_(m, CARRIER_CANDIDATES);
  // 3) broader includes
  if (idx < 0) idx = findHeaderIndexIncludes_(hdr, ['carrier','courier','provider','forwarder','kuljetus','toimitustapa','shipper','delivery']);
  return idx;
}
// ===== END HOTFIX =====
