importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");

function normText(s){ return String(s||'').trim(); }
function normCode(s){ return String(s||'').trim(); }

self.onmessage = async function(e){
  const { buf } = e.data;

  const wb = XLSX.read(buf, { type:'array' });

  const wsRev   = wb.Sheets['Revenue'];
  const wsLinks = wb.Sheets['Links'];
  const wsDown  = wb.Sheets['DownLinks'];

  // =========================
  // STEP 1：months（快速）
  // =========================
  const headerRow = XLSX.utils.sheet_to_json(wsRev, { header:1 })[0] || [];
  const found = new Set();

  for (const h of headerRow) {
    if (!h) continue;
    const m = String(h).match(/(\d{4}).*?(\d{1,2})/);
    if (m) {
      found.add(m[1] + String(m[2]).padStart(2,'0'));
    }
  }

  const months = Array.from(found).sort((a,b)=>Number(b)-Number(a));

  // ⚡ 先回 UI
  self.postMessage({
    type: 'months_ready',
    months
  });

  // =========================
  // STEP 2：只回 raw data（不建 map）
  // =========================
  const revenueRows = XLSX.utils.sheet_to_json(wsRev, { defval:null });
  const linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });
  const downRows    = wsDown ? XLSX.utils.sheet_to_json(wsDown, { defval:null }) : [];

  // =========================
  // STEP 3：回傳（完全不做加工）
  // =========================
  self.postMessage({
    type: 'ready',
    payload: {
      revenueRows,
      linksRows,
      downRows,
      months
    }
  });
};
