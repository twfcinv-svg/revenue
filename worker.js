importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");

let byCode = new Map();
let byName = new Map();
let revenueRows = [];
let linksRows = [];
let downRows = [];
let COL_MAP = {};
let months = [];

function normText(s){ return String(s||'').trim(); }
function normCode(s){ return String(s||'').trim(); }

self.onmessage = async function(e){
  const { buf } = e.data;

  const wb = XLSX.read(buf, { type:'array' });

  const wsRev = wb.Sheets['Revenue'];
  const wsLinks = wb.Sheets['Links'];
  const wsDown = wb.Sheets['DownLinks'];

  const rowsHeaderFirst = XLSX.utils.sheet_to_json(wsRev, { header:1, blankrows:false });
  const headerRow = rowsHeaderFirst[0] || [];

  const found = new Set();

  for (const h of headerRow) {
    if (!h) continue;
    const m = String(h).match(/(\d{4}).*?(\d{1,2})/);
    if (m) {
      const ym = m[1] + String(m[2]).padStart(2,'0');
      found.add(ym);
    }
  }

  months = Array.from(found);

  revenueRows = XLSX.utils.sheet_to_json(wsRev, { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });
  downRows    = XLSX.utils.sheet_to_json(wsDown, { defval:null });

  for (const r of revenueRows) {
    const code = normCode(r['代號'] || r['股票代碼'] || '');
    const name = normText(r['名稱'] || '');
    if (code) byCode.set(code, r);
    if (name) byName.set(name, r);
  }

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
