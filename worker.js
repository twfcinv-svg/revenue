/* worker.js
 * 功能：
 * 1) 在 Worker 中解析 data.xlsx，避免主執行緒卡住
 * 2) 先回傳 months_ready（讓月份下拉先出現）
 * 3) 再回傳 ready（完整資料）
 * 4) 若 worker 內找不到 XLSX，回傳 error 讓 app.js 自動 fallback
 */

const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET = 'Links';
const DOWNLINKS_SHEET = 'DownLinks';

function z(s) {
  return String(s == null ? '' : s);
}

function normText(s) {
  return z(s)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[\u3000]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildColMapFromHeader(headerRow) {
  const found = new Set();
  const colMap = Object.create(null);

  for (const rawHeader of headerRow || []) {
    if (!rawHeader) continue;
    const h = normText(String(rawHeader));

    let m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*年[成增]長\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if (m) {
      const ym = m[1] + String(m[2]).padStart(2, '0');
      (colMap[ym] ??= {}).YoY = rawHeader;
      found.add(ym);
      continue;
    }

    m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*月[變增]動\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if (m) {
      const ym = m[1] + String(m[2]).padStart(2, '0');
      (colMap[ym] ??= {}).MoM = rawHeader;
      found.add(ym);
      continue;
    }
  }

  const months = Array.from(found).sort((a, b) => Number(b) - Number(a));
  return { months, colMap };
}

function ensureXLSX(candidates = []) {
  if (self.XLSX) return true;

  const list = Array.isArray(candidates) && candidates.length > 0
    ? candidates
    : [
        './xlsx.full.min.js',
        './xlsx.min.js',
        './libs/xlsx.full.min.js',
        './vendor/xlsx.full.min.js',
        'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
        'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
      ];

  for (const url of list) {
    try {
      importScripts(url);
      if (self.XLSX) return true;
    } catch (_) {
      // 繼續嘗試下一個路徑
    }
  }

  return !!self.XLSX;
}

self.onmessage = function (e) {
  try {
    const { buf, xlsxLibCandidates } = e.data || {};

    if (!buf) {
      self.postMessage({
        type: 'error',
        message: 'worker 未收到 data.xlsx 的 ArrayBuffer'
      });
      return;
    }

    const ok = ensureXLSX(xlsxLibCandidates);
    if (!ok || !self.XLSX) {
      self.postMessage({
        type: 'error',
        message: 'worker 內無法載入 XLSX 函式庫，請確認 xlsx.full.min.js 路徑'
      });
      return;
    }

    const wb = self.XLSX.read(buf, { type: 'array' });

    const wsRev = wb.Sheets[REVENUE_SHEET];
    const wsLinks = wb.Sheets[LINKS_SHEET];
    const wsDown = wb.Sheets[DOWNLINKS_SHEET];

    if (!wsRev || !wsLinks) {
      self.postMessage({
        type: 'error',
        message: '找不到必要工作表 Revenue 或 Links'
      });
      return;
    }

    // 先讀 header，快速回月份
    const rowsHeaderFirst = self.XLSX.utils.sheet_to_json(wsRev, {
      header: 1,
      blankrows: false
    });

    const headerRow =
      Array.isArray(rowsHeaderFirst) && rowsHeaderFirst.length > 0
        ? rowsHeaderFirst[0]
        : [];

    const { months, colMap } = buildColMapFromHeader(headerRow);

    self.postMessage({
      type: 'months_ready',
      months,
      colMap
    });

    // 再完整轉資料
    const revenueRows = self.XLSX.utils.sheet_to_json(wsRev, { defval: null });
    const linksRows = self.XLSX.utils.sheet_to_json(wsLinks, { defval: null });
    const downRows = wsDown
      ? self.XLSX.utils.sheet_to_json(wsDown, { defval: null })
      : [];

    self.postMessage({
      type: 'ready',
      payload: {
        months,
        colMap,
        revenueRows,
        linksRows,
        downRows
      }
    });
  } catch (err) {
    self.postMessage({
      type: 'error',
      message: err && err.message ? err.message : String(err)
    });
  }
};
