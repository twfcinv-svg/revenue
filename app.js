/* app.js — fixed worker/cache version
 * 修正重點：
 * 1) 移除 Date.now() 當預設版本號，避免每次進站都強制重抓 data.xlsx
 * 2) worker 失敗時自動 fallback 到主執行緒解析
 * 3) worker ready 後，完整重建 COL_MAP / linksByUp / linksByDown / downstreamHJ
 * 4) 月份先出現、資料後到，改善操作體驗
 * 5) 增加 IndexedDB 快取（同版本資料可直接秒開）
 * 6) 修正 renderTreemap 重複綁 resize listener 的問題
 */

const APP_DATA_VERSION = '20260416-1'; // ★ 你資料有更新時，手動改這裡版本號即可
const URL_VER = new URLSearchParams(location.search).get('v') || APP_DATA_VERSION;

const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
const WORKER_FILE = new URL(`./worker.js?v=${URL_VER}`, location.href).toString();

const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET = 'Links';
const DOWNLINKS_SHEET = 'DownLinks';

const CODE_FIELDS = ['個股', '代號', '股票代碼', '股票代號', '公司代號', '證券代號'];
const NAME_FIELDS = ['名稱', '公司名稱', '證券名稱'];

const COL_MAP = Object.create(null);

// ===== 可調參數 =====
const HEADER_H = 22;
const GROUP_KEEP_MAX = 7;
const GROUP_WEIGHT_MODE = 'RANK';
const RANK_WEIGHT_MIN = 1.3;
const RANK_WEIGHT_MAX = 1.8;

// ===== 上游類股篩選規則 =====
// true  = 上游只顯示平均值 > 0 的類股
// false = 上游只依平均值排序，允許負值類股進榜
const UPSTREAM_ONLY_POSITIVE = false;

const ENABLE_NODE_CLICK = true;
const MIN_RENDER_W = 75;
const MIN_RENDER_H = 20;
const MIN_RENDER_AREA = 400;

let revenueRows = [];
let linksRows = [];
let downRows = [];
let months = [];

let byCode = new Map();
let byName = new Map();

let worker = null;
let DATA_READY = false;
let MONTHS_READY = false;

let linksByUp = new Map();
let linksByDown = new Map();
let downstreamHJ = [];

const XLSX_WORKER_CANDIDATES = [
  './xlsx.full.min.js',
  './xlsx.min.js',
  './libs/xlsx.full.min.js',
  './vendor/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
];

// ===== IndexedDB Cache =====
const DB_NAME = 'revenue_cache';
const STORE = 'data';
const CACHE_KEY = `main:${URL_VER}`;

function z(s) {
  return String(s == null ? '' : s);
}
function toHalfWidth(str) {
  return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)
  );
}
function normText(s) {
  return z(s)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[\u3000]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
function normCode(s) {
  return toHalfWidth(z(s))
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, '')
    .trim();
}
function displayPct(v) {
  if (v == null || !isFinite(v)) return '—';
  const s = v.toFixed(1) + '%';
  return v > 0 ? '+' + s : s;
}
function colorFor(v, mode) {
  if (v == null || !isFinite(v)) return '#0f172a';
  const t = Math.min(1, Math.abs(v) / 80);
  const alpha = 0.25 + 0.35 * t;
  const good = mode === 'greenPositive';
  const pos = good ? '156,163,175' : '59,130,246';
  const neg = good ? '59,130,246' : '156,163,175';
  const rgb = v >= 0 ? pos : neg;
  return `rgba(${rgb},${alpha})`;
}
function safe(s) {
  return z(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));
}
function isUSCode(code) {
  return /\.US$/i.test(String(code || '').trim());
}

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);

    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE);
      }
    };

    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveCache(data) {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(data, CACHE_KEY);

    await new Promise((resolve, reject) => {
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });

    db.close();
  } catch (err) {
    console.warn('saveCache 失敗：', err);
  }
}

async function loadCache() {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).get(CACHE_KEY);

    const result = await new Promise(resolve => {
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    });

    db.close();
    return result || null;
  } catch (err) {
    console.warn('loadCache 失敗：', err);
    return null;
  }
}

// ===== UI =====
function initControls() {
  console.log('初始化 UI 控制項');
}

function setMonthSelectLoading(text = '載入中...') {
  const sel = document.querySelector('#monthSelect');
  if (!sel) return;
  sel.innerHTML = `<option value="">${safe(text)}</option>`;
}

function updateControls() {
  const sel = document.querySelector('#monthSelect');

  if (!sel) {
    console.error('❌ 找不到 #monthSelect');
    return;
  }

  if (!Array.isArray(months) || months.length === 0) {
    sel.innerHTML = `<option value="">載入中...</option>`;
    return;
  }

  const prev = sel.value;
  sel.innerHTML = '';

  for (const m of months) {
    const o = document.createElement('option');
    o.value = m;
    o.textContent = `${m.slice(0, 4)}年${m.slice(4, 6)}月`;
    sel.appendChild(o);
  }

  if (months.includes(prev)) {
    sel.value = prev;
  } else {
    sel.value = months[0];
  }

  console.log('✅ 月份下拉完成', months);
}

// ===== 資料處理 =====
function detectCodeKey(sample = {}) {
  return CODE_FIELDS.find(k => k in sample) || CODE_FIELDS[0];
}

function detectNameKey(sample = {}) {
  return NAME_FIELDS.find(k => k in sample) || NAME_FIELDS[0];
}

function buildColMapFromHeader(headerRow) {
  const found = new Set();
  const nextMap = Object.create(null);

  for (const rawHeader of headerRow || []) {
    if (!rawHeader) continue;
    const h = normText(String(rawHeader));

    let m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*年[成增]長\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if (m) {
      const ym = m[1] + String(m[2]).padStart(2, '0');
      (nextMap[ym] ??= {}).YoY = rawHeader;
      found.add(ym);
      continue;
    }

    m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*月[變增]動\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if (m) {
      const ym = m[1] + String(m[2]).padStart(2, '0');
      (nextMap[ym] ??= {}).MoM = rawHeader;
      found.add(ym);
      continue;
    }
  }

  const nextMonths = Array.from(found).sort((a, b) => Number(b) - Number(a));
  return { colMap: nextMap, months: nextMonths };
}

function applyColMap(nextMap) {
  for (const k of Object.keys(COL_MAP)) delete COL_MAP[k];
  Object.assign(COL_MAP, nextMap || {});
}

function rebuildMaps() {
  byCode.clear();
  byName.clear();

  const sample = revenueRows[0] || {};
  const codeKeyName = detectCodeKey(sample);
  const nameKeyName = detectNameKey(sample);

  for (const r of revenueRows) {
    const code = normCode(
      r[codeKeyName] ||
      r['個股'] ||
      r['代號'] ||
      r['股票代碼'] ||
      r['股票代號'] ||
      r['公司代號'] ||
      r['證券代號'] ||
      ''
    );

    const name = normText(
      r[nameKeyName] ||
      r['名稱'] ||
      r['公司名稱'] ||
      r['證券名稱'] ||
      ''
    );

    if (code) byCode.set(code, r);
    if (name) byName.set(name, r);
  }

  console.log('✔ rebuildMaps 完成', {
    byCodeSize: byCode.size,
    byNameSize: byName.size
  });
}

function rebuildRelationMaps() {
  linksByUp.clear();
  linksByDown.clear();

  for (const e of linksRows || []) {
    const A = normCode(e['上游代號']);
    const B = normCode(e['下游代號']);
    const C = normText(e['關係類型']);

    if (A && B && C) {
      if (!linksByUp.has(A)) linksByUp.set(A, []);
      linksByUp.get(A).push(e);

      if (!linksByDown.has(B)) linksByDown.set(B, []);
      linksByDown.get(B).push(e);
    }
  }

  downstreamHJ = [];
  for (const row of downRows || []) {
    const up = normCode(row['上游代號']);
    const down = normCode(row['下游代號']);
    const type = normText(row['關係類型']);

    if (up && down && type) {
      downstreamHJ.push({
        上游代號: up,
        下游代號: down,
        關係類型: type
      });
    }
  }

  console.log('✔ rebuildRelationMaps 完成', {
    linksRows: linksRows.length,
    downRows: downRows.length,
    linksByUpSize: linksByUp.size,
    linksByDownSize: linksByDown.size,
    downstreamHJSize: downstreamHJ.length
  });
}

function applyPayload(p, source = 'unknown') {
  revenueRows = Array.isArray(p?.revenueRows) ? p.revenueRows : [];
  linksRows = Array.isArray(p?.linksRows) ? p.linksRows : [];
  downRows = Array.isArray(p?.downRows) ? p.downRows : [];

  months = Array.isArray(p?.months) ? p.months : months;
  if (p?.colMap) applyColMap(p.colMap);

  if (months.length > 0) {
    MONTHS_READY = true;
    updateControls();
  }

  rebuildMaps();
  rebuildRelationMaps();

  DATA_READY = true;

  console.log(`✅ DATA_READY 完成（來源：${source}）`);
}

function getMetricValue(row, month, metric) {
  if (!row || !month || !metric) return null;

  const codeOfRow = normCode(
    row['個股'] ||
    row['代號'] ||
    row['股票代碼'] ||
    row['股票代號'] ||
    row['公司代號'] ||
    row['證券代號'] ||
    ''
  );

  if (isUSCode(codeOfRow)) return null;

  const col = (COL_MAP[month] || {})[metric];
  if (!col) return null;

  let v = row[col];
  if (v == null || v === '') return null;

  if (typeof v === 'string') v = v.replace('%', '').replace('％', '').trim();
  v = Number(v);

  return Number.isFinite(v) ? v : null;
}

// ===== Treemap 分群 =====
function getTreemapGroupName(svgId, edge, row) {
  if (svgId === 'upTreemap') {
    return normText(row?.['產業別'] || '未分類');
  }
  return normText(edge['關係類型'] || edge['type'] || '未分類');
}

function selectTreemapGroups(svgId, summaries) {
  if (svgId !== 'upTreemap') {
    return [...summaries]
      .sort((a, b) => b.list.length - a.list.length)
      .slice(0, GROUP_KEEP_MAX);
  }

  let arr = summaries.filter(g => Number.isFinite(g.avg));

  if (UPSTREAM_ONLY_POSITIVE) {
    arr = arr.filter(g => g.avg > 0);
  }

  return arr
    .sort((a, b) => b.avg - a.avg)
    .slice(0, GROUP_KEEP_MAX);
}

// ===== 載入流程 =====
async function loadWorkbookFallback() {
  if (typeof XLSX === 'undefined') {
    throw new Error('主執行緒找不到 XLSX，請確認 index.html 已先載入 xlsx.full.min.js');
  }

  const res = await fetch(XLSX_FILE, { cache: 'default' });
  if (!res.ok) throw new Error('讀取 data.xlsx 失敗 HTTP ' + res.status);

  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });

  const wsRev = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];
  const wsDown = wb.Sheets[DOWNLINKS_SHEET];

  if (!wsRev || !wsLinks) {
    throw new Error('找不到必要工作表 Revenue 或 Links');
  }

  const rowsHeaderFirst = XLSX.utils.sheet_to_json(wsRev, { header: 1, blankrows: false });
  const headerRow = Array.isArray(rowsHeaderFirst) && rowsHeaderFirst.length > 0 ? rowsHeaderFirst[0] : [];

  const { colMap, months: nextMonths } = buildColMapFromHeader(headerRow);

  MONTHS_READY = nextMonths.length > 0;
  months = nextMonths;
  applyColMap(colMap);
  updateControls();

  const payload = {
    months: nextMonths,
    colMap,
    revenueRows: XLSX.utils.sheet_to_json(wsRev, { defval: null }),
    linksRows: XLSX.utils.sheet_to_json(wsLinks, { defval: null }),
    downRows: wsDown ? XLSX.utils.sheet_to_json(wsDown, { defval: null }) : []
  };

  applyPayload(payload, 'fallback-main-thread');
  await saveCache({ ts: Date.now(), payload });
}

async function startWorkerLoad() {
  if (!('Worker' in window)) {
    console.warn('瀏覽器不支援 Worker，改用主執行緒解析');
    await loadWorkbookFallback();
    return;
  }

  try {
    worker = new Worker(WORKER_FILE);
  } catch (err) {
    console.warn('建立 Worker 失敗，改用主執行緒解析：', err);
    await loadWorkbookFallback();
    return;
  }

  let fellBack = false;

  const doFallbackOnce = async (reason) => {
    if (fellBack) return;
    fellBack = true;
    console.warn('worker 失敗，改用 fallback：', reason);

    try {
      if (worker) worker.terminate();
    } catch (_) {}

    await loadWorkbookFallback();
  };

  worker.onmessage = async (e) => {
    const msg = e.data || {};
    console.log('worker msg:', msg.type);

    if (msg.type === 'months_ready') {
      months = Array.isArray(msg.months) ? msg.months : [];
      applyColMap(msg.colMap || {});
      MONTHS_READY = months.length > 0;
      updateControls();
      return;
    }

    if (msg.type === 'ready') {
      const p = msg.payload || {};
      applyPayload(p, 'worker');
      await saveCache({ ts: Date.now(), payload: p });
      return;
    }

    if (msg.type === 'error') {
      console.error('worker error message:', msg.message);
      await doFallbackOnce(msg.message || 'unknown worker error');
    }
  };

  worker.onerror = async (err) => {
    console.error('worker.onerror:', err);
    await doFallbackOnce(err?.message || 'worker onerror');
  };

  worker.onmessageerror = async (err) => {
    console.error('worker.onmessageerror:', err);
    await doFallbackOnce('worker onmessageerror');
  };

  const res = await fetch(XLSX_FILE, { cache: 'default' });
  if (!res.ok) {
    await doFallbackOnce('fetch data.xlsx failed HTTP ' + res.status);
    return;
  }

  const buf = await res.arrayBuffer();
  worker.postMessage({ buf, xlsxLibCandidates: XLSX_WORKER_CANDIDATES }, [buf]);
}

async function initData() {
  setMonthSelectLoading('載入月份中...');

  // 先讀快取（若有快取，使用者幾乎可立即操作）
  const cached = await loadCache();
  if (cached?.payload) {
    console.log('✅ 使用 IndexedDB 快取資料');
    applyPayload(cached.payload, 'cache');
  }

  // 再背景刷新最新版本
  try {
    await startWorkerLoad();
  } catch (err) {
    console.error('initData 啟動失敗：', err);
    if (!DATA_READY) {
      await loadWorkbookFallback();
    }
  }
}

window.addEventListener('DOMContentLoaded', async () => {
  initControls();
  setMonthSelectLoading('載入月份中...');

  document.querySelector('#runBtn')?.addEventListener('click', handleRun);

  try {
    await initData();
  } catch (err) {
    console.error(err);
    alert('資料載入失敗，請重新整理頁面，或確認 data.xlsx / worker.js / xlsx.full.min.js 是否存在。');
  }
});

// ===== 查詢 =====
function handleRun() {
  if (!MONTHS_READY) {
    alert('月份仍在載入中，請稍候 1–3 秒');
    return;
  }

  if (!DATA_READY) {
    alert('資料仍在載入中，請稍候 1–3 秒');
    return;
  }

  const inputEl = document.querySelector('#stockInput');
  const raw = inputEl ? inputEl.value : '';
  const month = document.querySelector('#monthSelect')?.value || '';
  const metric = document.querySelector('#metricSelect')?.value || 'MoM';
  const colorMode = document.querySelector('#colorMode')?.value || 'redPositive';

  if (!raw || !raw.trim()) {
    alert('請輸入股票代號或公司名稱');
    return;
  }

  let codeKey = normCode(String(raw).toUpperCase());
  let rowSelf = byCode.get(codeKey);

  if (!rowSelf) {
    const nameQ = normText(raw);
    rowSelf =
      byName.get(nameQ) ||
      revenueRows.find(r =>
        normText(r['名稱'] || r['公司名稱'] || r['證券名稱'] || '').startsWith(nameQ)
      );

    if (rowSelf) {
      codeKey = normCode(
        rowSelf['個股'] ??
        rowSelf['代號'] ??
        rowSelf['股票代碼'] ??
        rowSelf['股票代號'] ??
        rowSelf['公司代號'] ??
        rowSelf['證券代號']
      );
    }
  }

  if (!rowSelf) {
    alert('找不到此代號/名稱');
    return;
  }

  try {
    const codeLabel =
      (rowSelf['個股'] ||
        rowSelf['代號'] ||
        rowSelf['股票代碼'] ||
        rowSelf['股票代號'] ||
        rowSelf['公司代號'] ||
        rowSelf['證券代號'] ||
        '').trim();

    const nameLabel =
      (rowSelf['名稱'] || rowSelf['公司名稱'] || rowSelf['證券名稱'] || '').trim();

    const extra = `${month.slice(0, 4)}/${month.slice(4, 6)} · ${metric}`;
    if (window.setResultChipLink) window.setResultChipLink(codeLabel, nameLabel, extra);
  } catch (_) {}

  const upstreamEdges = linksByDown.get(codeKey) || [];
  let downstreamEdges = downstreamHJ.filter(e => e['上游代號'] === codeKey);

  downstreamEdges = downstreamEdges.filter(e => !String(e['下游代號']).endsWith('.US'));

  console.log('查詢代號 =', codeKey);
  console.log('上游筆數 =', upstreamEdges.length, upstreamEdges);
  console.log('下游筆數 =', downstreamEdges.length, downstreamEdges);

  requestAnimationFrame(() => {
    renderResultChip(rowSelf, month, metric, colorMode);
    renderTreemap('upTreemap', 'upHint', upstreamEdges, '上游代號', month, metric, colorMode);
  });

  requestAnimationFrame(() => {
    renderTreemap('downTreemap', 'downHint', downstreamEdges, '下游代號', month, metric, colorMode);
  });
}

function renderResultChip(selfRow, month, metric, colorMode) {
  const host = document.querySelector('#resultChip');
  if (!host) return;

  const v = getMetricValue(selfRow, month, metric);
  const bg = colorFor(v, colorMode);

  const showCode =
    selfRow['個股'] ||
    selfRow['代號'] ||
    selfRow['股票代碼'] ||
    selfRow['股票代號'] ||
    selfRow['公司代號'] ||
    selfRow['證券代號'] ||
    '';

  const showName =
    selfRow['名稱'] || selfRow['公司名稱'] || selfRow['證券名稱'] || '';

  host.innerHTML = `
    <div class="result-card" style="background:${bg}">
      <div class="row1"><strong>${safe(showCode)}｜${safe(showName)}</strong><span>${month.slice(0,4)}/${month.slice(4,6)} / ${metric}</span></div>
      <div class="row2"><span>${safe(selfRow['產業別'] || '')}</span><span>${displayPct(v)}</span></div>
    </div>`;
}

// ========= 個股標籤適配 =========
const LabelFit = {
  paddingBase: 8,
  maxFont: 36,
  minFontSoft: 9,
  minFontHard: 8,
  lineHeight: 1.15,

  dynPadding(w, h) {
    const m = Math.min(w, h);
    return Math.max(2, Math.min(this.paddingBase, Math.floor(m * 0.08)));
  },

  centerText(el, w, h, p) {
    el.setAttribute('text-anchor', 'middle');
    el.setAttribute('dominant-baseline', 'middle');
    el.setAttribute('x', p + Math.max(0, (w - p * 2) / 2));
    el.setAttribute('y', p + Math.max(0, (h - p * 2) / 2));
  },

  ensureClip(gEl, w, h) {
    const inset = 2;
    const svg = gEl.ownerSVGElement;
    let defs = svg.querySelector('defs');

    if (!defs) {
      defs = svg.insertBefore(
        document.createElementNS('http://www.w3.org/2000/svg', 'defs'),
        svg.firstChild
      );
    }

    const id = gEl.dataset.clipId || ('clip-' + Math.random().toString(36).slice(2));
    gEl.dataset.clipId = id;

    let clip = svg.querySelector('#' + id);
    if (!clip) {
      clip = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
      clip.setAttribute('id', id);
      const r = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      clip.appendChild(r);
      defs.appendChild(clip);
    }

    const rect = clip.firstChild;
    rect.setAttribute('x', inset);
    rect.setAttribute('y', inset);
    rect.setAttribute('width', Math.max(0, w - inset * 2));
    rect.setAttribute('height', Math.max(0, h - inset * 2));

    gEl.querySelectorAll('text').forEach(t => t.setAttribute('clip-path', `url(#${id})`));
  },

  ellipsizeNameToWidth(textEl, maxW) {
    const t1 = textEl.querySelector('tspan');
    if (!t1) return;

    const full = t1.textContent || '';
    const m = full.match(/^(\S+)\s*(.*)$/);
    let code = '';
    let name = full;

    if (m) {
      code = m[1];
      name = m[2] || '';
    }

    t1.textContent = code + (name ? ' ' + name : '');

    while (t1.getComputedTextLength() > maxW && name.length > 0) {
      name = name.slice(0, -1);
      t1.textContent = code + (name ? ' ' + name + '…' : '');
    }
  },

  fitBlock(textEl, w, h) {
    const p = this.dynPadding(w, h);
    const targetW = Math.max(1, w - p * 2);
    const targetH = Math.max(1, h - p * 2);

    const code = textEl.dataset.code || '';
    const name = textEl.dataset.name || '';
    const pct = textEl.dataset.pct || '';

    const layouts = [
      () => [`${code}${name ? ' ' + name : ''}`, pct],
      () => [code, pct]
    ];

    const k = 0.12;
    const areaFont = Math.sqrt(targetW * targetH) * k;
    const logicalMax = Math.min(this.maxFont, Math.floor(targetH * 0.5));

    for (const L of layouts) {
      while (textEl.firstChild) textEl.removeChild(textEl.firstChild);

      L().forEach(s => {
        const t = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
        t.textContent = s;
        textEl.appendChild(t);
      });

      let f = Math.max(this.minFontHard, Math.min(logicalMax, Math.floor(areaFont)));
      textEl.setAttribute('font-size', f);
      this.centerText(textEl, w, h, p);
      this.ellipsizeNameToWidth(textEl, targetW);

      let guard = 0;
      while (guard++ < 60) {
        const bb = textEl.getBBox();
        const sW = targetW / Math.max(1, bb.width);
        const sH = targetH / Math.max(1, bb.height);
        const s = Math.min(sW, sH, 1);
        const next = Math.max(this.minFontHard, Math.floor(f * s));

        if (next < f) {
          f = next;
          textEl.setAttribute('font-size', f);
          this.centerText(textEl, w, h, p);
          continue;
        }
        break;
      }

      const tsp = textEl.querySelectorAll('tspan');
      const n = Math.max(1, tsp.length);
      const offsetEm = -((n - 1) * this.lineHeight / 2);

      tsp.forEach((t, i) => {
        t.setAttribute('x', textEl.getAttribute('x'));
        t.setAttribute('dy', i === 0 ? `${offsetEm}em` : `${this.lineHeight}em`);
      });

      const box = textEl.getBBox();
      if (box.width <= targetW + 0.1 && box.height <= targetH + 0.1) {
        textEl.removeAttribute('display');
        return true;
      }
    }

    textEl.setAttribute('display', 'none');
    return false;
  }
};

// ========= 群組標題 =========
const GroupTitleFit = {
  minFont: 5,
  lineHeight: 1.12,
  inset: 4,
  k: 0.12,

  ensureHeaderClip(svg, gEl, d, headerH) {
    const id = gEl.dataset.headerClipId || ('hclip-' + Math.random().toString(36).slice(2));
    gEl.dataset.headerClipId = id;
    let defs = svg.querySelector('defs');

    if (!defs) {
      defs = svg.insertBefore(
        document.createElementNS('http://www.w3.org/2000/svg', 'defs'),
        svg.firstChild
      );
    }

    let clip = svg.querySelector('#' + id);
    if (!clip) {
      clip = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
      clip.setAttribute('id', id);
      clip.setAttribute('clipPathUnits', 'userSpaceOnUse');
      const r = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      clip.appendChild(r);
      defs.appendChild(clip);
    }

    const r = clip.firstChild;
    const w = Math.max(0, d.x1 - d.x0);
    const h = Math.max(0, headerH);
    r.setAttribute('x', d.x0 + this.inset);
    r.setAttribute('y', d.y0 + this.inset);
    r.setAttribute('width', Math.max(0, w - this.inset * 2));
    r.setAttribute('height', Math.max(0, h - this.inset * 2));
    return `url(#${id})`;
  },

  mountOneLine(text, d) {
    while (text.firstChild) text.removeChild(text.firstChild);

    const tName = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tName.textContent = d.data.name || '';

    const tSep = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tSep.textContent = '  ';

    const tAvg = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tAvg.textContent = `平均：${displayPct(d.data.avg)}`;

    text.appendChild(tName);
    text.appendChild(tSep);
    text.appendChild(tAvg);
    text.dataset.mode = 'one';
  },

  mountTwoLines(text, d) {
    while (text.firstChild) text.removeChild(text.firstChild);

    const tName = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tName.textContent = d.data.name || '';

    const tAvg = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tAvg.textContent = `平均：${displayPct(d.data.avg)}`;

    text.appendChild(tName);
    text.appendChild(tAvg);
    text.dataset.mode = 'two';
  },

  ellipsizeName(text) {
    const tName = text.querySelector('tspan');
    if (!tName) return false;
    let nm = tName.textContent || '';
    if (nm.length === 0) return false;
    tName.textContent = nm.slice(0, -1) + '…';
    return true;
  },

  shortenAvg(text) {
    const tsp = text.querySelectorAll('tspan');
    if (tsp.length === 0) return;

    const last = tsp[tsp.length - 1];
    const m = String(last.textContent || '').match(/([+\-]?[0-9]+(?:\.[0-9])?%)/);
    if (m) last.textContent = m[1];
  },

  fit(text, d, headerH) {
    const wMaxFull = Math.max(0, d.x1 - d.x0) - this.inset * 2 - 2;
    const hMax = Math.max(0, headerH) - this.inset * 2 - 1;
    if (wMaxFull <= 0 || hMax <= 0) return;

    text.setAttribute('text-anchor', 'start');
    text.setAttribute('dominant-baseline', 'middle');
    text.setAttribute('x', d.x0 + this.inset + 4);
    text.setAttribute('y', d.y0 + headerH / 2);
    text.removeAttribute('lengthAdjust');
    text.removeAttribute('textLength');
    text.setAttribute('clip-path', this.ensureHeaderClip(text.ownerSVGElement, text.parentNode, d, headerH));

    this.mountOneLine(text, d);

    let f = Math.max(
      this.minFont,
      Math.floor(Math.min(Math.sqrt(Math.max(1, wMaxFull * hMax)) * this.k, hMax * 0.95))
    );

    let guard = 0;

    const loop = () => {
      if (++guard > 160) return;

      text.setAttribute('font-size', f);
      const mode = text.dataset.mode || 'one';
      const bb = text.getBBox();
      const sW = wMaxFull / Math.max(1, bb.width);
      const sH = hMax / Math.max(1, bb.height);
      const s = Math.min(sW, sH, 1);
      const next = Math.max(this.minFont, Math.floor(f * s));

      if (next < f) {
        f = next;
        return loop();
      }

      if (sW < 1 && f <= this.minFont) {
        if (mode === 'one') {
          if (!this.ellipsizeName(text)) {
            if (hMax >= this.minFont * 2 * this.lineHeight + 2) {
              this.mountTwoLines(text, d);
              return loop();
            }
          }
          return loop();
        } else {
          if (this.ellipsizeName(text)) return loop();
        }
      }
    };

    loop();

    let bb = text.getBBox();
    if (bb.width > wMaxFull + 0.1) {
      this.shortenAvg(text);
      text.setAttribute(
        'font-size',
        Math.max(this.minFont, parseInt(text.getAttribute('font-size') || this.minFont, 10) - 1)
      );
      bb = text.getBBox();
    }

    if (bb.width > wMaxFull + 0.1) {
      text.setAttribute('lengthAdjust', 'spacingAndGlyphs');
      text.setAttribute('textLength', Math.max(1, Math.floor(wMaxFull)));
    }
  }
};

// ===== Treemap =====
function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode) {
  const svg = d3.select('#' + svgId);
  if (svg.empty()) return;

  svg.selectAll('*').remove();

  const wrap = svg.node().parentElement;
  const W = Math.max(240, (wrap?.clientWidth || 600) - 16);
  const H = parseInt(getComputedStyle(svg.node()).height, 10) || 560;
  svg.attr('width', W).attr('height', H);

  const groups = new Map();

  for (const e of edges || []) {
    const keyRaw = normCode(e[codeField] || e['down'] || e['up']);
    if (!keyRaw || isUSCode(keyRaw)) continue;

    const r = byCode.get(keyRaw);
    const groupName = getTreemapGroupName(svgId, e, r);

    if (!groups.has(groupName)) groups.set(groupName, []);

    if (!r) {
      groups.get(groupName).push({
        code: keyRaw,
        name: '',
        raw: null,
        rel: groupName
      });
      continue;
    }

    const v = getMetricValue(r, month, metric);
    const codeVal =
      r['個股'] ??
      r['代號'] ??
      r['股票代碼'] ??
      r['股票代號'] ??
      r['公司代號'] ??
      r['證券代號'];

    const nameVal = r['名稱'] ?? r['公司名稱'] ?? r['證券名稱'];

    groups.get(groupName).push({
      code: codeVal,
      name: nameVal,
      raw: v,
      rel: groupName
    });
  }

  const hint = document.getElementById(hintId);
  if (groups.size === 0) {
    if (hint) hint.textContent = '此區在選定月份沒有可用數據';
    return;
  }

  const EPS = 0.01;
  const allSummaries = [];

  for (const [rel, list] of groups) {
    const avg = d3.mean(list, d => (Number.isFinite(d.raw) ? d.raw : null));
    const minLeafRaw = d3.min(list.map(d => (Number.isFinite(d.raw) ? d.raw : 0)));

    const baseValues = list.map(s => {
      const valNum = Number.isFinite(s.raw) ? s.raw : minLeafRaw;
      return { s, base: Math.max(EPS, valNum - minLeafRaw + EPS) };
    });

    const baseSum = d3.sum(baseValues, d => d.base) || EPS;

    allSummaries.push({
      rel,
      list,
      avg,
      baseValues,
      baseSum
    });
  }

  const groupSummaries = selectTreemapGroups(svgId, allSummaries);

  if (groupSummaries.length === 0) {
    if (hint) hint.textContent = '此區在選定月份沒有符合條件的類股';
    return;
  }

  if (hint) {
    if (svgId === 'upTreemap') {
      hint.textContent =
        '已顯示上游平均' +
        metric +
        '最佳的 ' +
        groupSummaries.length +
        ' 個類股：' +
        groupSummaries.map(g => `${g.rel}（${displayPct(g.avg)}）`).join('、');
    } else {
      hint.textContent = '';
    }
  }

  const groupWeights = new Map();

  if (GROUP_WEIGHT_MODE === 'AVG') {
    const minAvg = d3.min(groupSummaries.map(d => (Number.isFinite(d.avg) ? d.avg : 0)));
    for (const g of groupSummaries) {
      const a = Number.isFinite(g.avg) ? g.avg : minAvg;
      groupWeights.set(g.rel, Math.max(EPS, a - minAvg + EPS));
    }
  } else {
    const sorted = [...groupSummaries].sort(
      (a, b) =>
        (Number.isFinite(a.avg) ? a.avg : -Infinity) -
        (Number.isFinite(b.avg) ? b.avg : -Infinity)
    );
    const n = Math.max(1, sorted.length - 1);

    sorted.forEach((g, i) => {
      const t = i / n;
      const w = RANK_WEIGHT_MIN + t * (RANK_WEIGHT_MAX - RANK_WEIGHT_MIN);
      groupWeights.set(g.rel, w);
    });
  }

  let children = [];
  for (const g of groupSummaries) {
    const gw = groupWeights.get(g.rel) || 1;
    const scale = gw / (g.baseSum || EPS);

    const kids = g.baseValues.map(({ s, base }) => ({
      name: s.name || '',
      code: s.code,
      raw: s.raw,
      rel: s.rel || g.rel,
      value: base * scale
    }));

    children.push({
      name: g.rel,
      avg: g.avg,
      children: kids
    });
  }

  // 第一次 layout
  let root = d3
    .hierarchy({ children })
    .sum(d => d.value)
    .sort((a, b) => (b.value || 0) - (a.value || 0));

  d3.treemap().size([W, H]).paddingOuter(8).paddingInner(3).paddingTop(HEADER_H)(root);

  // 過濾太小的個股
  const filteredChildren = (root.children || [])
    .map(parent => {
      const keptLeaves = (parent.children || [])
        .filter(leaf => {
          const w = Math.max(0, leaf.x1 - leaf.x0);
          const h = Math.max(0, leaf.y1 - leaf.y0);
          const area = w * h;
          return w >= MIN_RENDER_W && h >= MIN_RENDER_H && area >= MIN_RENDER_AREA;
        })
        .map(leaf => leaf.data);

      return {
        name: parent.data.name,
        avg: parent.data.avg,
        children: keptLeaves
      };
    })
    .filter(g => g.children && g.children.length > 0);

  if (filteredChildren.length === 0) {
    if (hint) hint.textContent = '此區個股方塊過小，已自動省略';
    return;
  }

  // 第二次 layout
  root = d3
    .hierarchy({ children: filteredChildren })
    .sum(d => d.value)
    .sort((a, b) => (b.value || 0) - (a.value || 0));

  d3.treemap().size([W, H]).paddingOuter(8).paddingInner(3).paddingTop(HEADER_H)(root);

  const g = svg.append('g');

  const parents = g.selectAll('g.parent')
    .data(root.children || [])
    .enter()
    .append('g')
    .attr('class', 'parent');

  parents.append('rect')
    .attr('class', 'group-bg')
    .attr('x', d => d.x0)
    .attr('y', d => d.y0)
    .attr('width', d => Math.max(0, d.x1 - d.x0))
    .attr('height', d => Math.max(0, d.y1 - d.y0))
    .attr('fill', d => colorFor(d.data.avg, colorMode));

  parents.append('rect')
    .attr('class', 'group-border')
    .attr('x', d => d.x0)
    .attr('y', d => d.y0)
    .attr('width', d => Math.max(0, d.x1 - d.x0))
    .attr('height', d => Math.max(0, d.y1 - d.y0));

  const titles = parents.append('text')
    .attr('class', 'node-title')
    .attr('fill', '#fff')
    .style('paint-order', 'stroke')
    .style('stroke', 'rgba(0,0,0,0.35)')
    .style('stroke-width', '2px');

  titles.each(function (d) {
    GroupTitleFit.fit(this, d, HEADER_H);
  });

  const node = g.selectAll('g.node')
    .data(root.leaves())
    .enter()
    .append('g')
    .attr('class', 'node')
    .attr('transform', d => `translate(${d.x0},${d.y0})`);

  node.append('rect')
    .attr('class', 'node-rect')
    .attr('width', d => Math.max(0, d.x1 - d.x0))
    .attr('height', d => Math.max(0, d.y1 - d.y0))
    .attr('fill', d => colorFor(d.data.raw, colorMode));

  const labels = node.append('text')
    .attr('class', 'node-label')
    .attr('fill', '#fff')
    .style('paint-order', 'stroke')
    .style('stroke', 'rgba(0,0,0,0.35)')
    .style('stroke-width', '2px')
    .style('text-rendering', 'geometricPrecision');

  labels.each(function (d) {
    const code = `${d.data.code || ''}`.trim();
    const name = `${d.data.name || ''}`.trim();
    const pct = displayPct(d.data.raw);
    const rel = `${d.data.rel || ''}`.trim();

    this.dataset.code = code;
    this.dataset.name = name;
    this.dataset.pct = pct;

    const t1 = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    t1.textContent = `${code}${name ? ' ' + name : ''}`;

    const t2 = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    t2.textContent = pct;

    this.appendChild(t1);
    this.appendChild(t2);

    const title = document.createElementNS('http://www.w3.org/2000/svg', 'title');
    title.textContent = `${code} ${name}\n${rel}\n${month.slice(0, 4)}/${month.slice(4, 6)} ${metric}: ${pct}`;
    this.appendChild(title);
  });

  if (ENABLE_NODE_CLICK) {
    node
      .style('cursor', 'pointer')
      .on('click', function (event, d) {
        const code = `${d.data.code || ''}`.trim();
        if (!code) return;

        const input = document.querySelector('#stockInput');
        if (input) input.value = code;

        handleRun();
      });
  }

  requestAnimationFrame(() => {
    node.each(function (d) {
      const w = Math.max(0, d.x1 - d.x0);
      const h = Math.max(0, d.y1 - d.y0);
      const textEl = this.querySelector('text');
      if (!textEl) return;
      LabelFit.fitBlock(textEl, w, h);
      LabelFit.ensureClip(this, w, h);
    });

    parents.select('text').each(function (d) {
      GroupTitleFit.fit(this, d, HEADER_H);
    });
  });

  // 修正：避免每次 render 都無限增加 resize listener
  const resizeKey = `__treemapResize_${svgId}`;
  if (window[resizeKey]) {
    window.removeEventListener('resize', window[resizeKey]);
  }

  const onResize = () => {
    parents.select('text').each(function (d) {
      GroupTitleFit.fit(this, d, HEADER_H);
    });
  };

  window[resizeKey] = onResize;
  window.addEventListener('resize', onResize, { passive: true });
}
