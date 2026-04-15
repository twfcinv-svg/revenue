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
  // STEP 1：快速抓 months（先回 UI）
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

  // ⚡ 先讓 UI 出現（超重要）
  self.postMessage({
    type: 'months_ready',
    months
  });

  // =========================
  // STEP 2：慢處理（背景）
  // =========================
  const revenueRows = XLSX.utils.sheet_to_json(wsRev, { defval:null });
  const linksRows    = XLSX.utils.sheet_to_json(wsLinks, { defval:null });
  const downRows     = wsDown ? XLSX.utils.sheet_to_json(wsDown, { defval:null }) : [];

  const byCode = new Map();
  const byName = new Map();

  for (const r of revenueRows) {
    const code = normCode(
      r['個股'] || r['代號'] || r['股票代碼'] ||
      r['股票代號'] || r['公司代號'] || r['證券代號'] || ''
    );

    const name = normText(
      r['名稱'] || r['公司名稱'] || r['證券名稱'] || ''
    );

    if (code) byCode.set(code, r);
    if (name) byName.set(name, r);
  }

  // =========================
  // STEP 3：最後完整回傳
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
