/* app.js — v3.16 (已移除下載按鈕版) */

const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';
const DOWNLINKS_SHEET = 'DownLinks';

const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];
const COL_MAP = {};

// ===== 可調參數 =====
const HEADER_H = 22;
const GROUP_KEEP_MAX = 7;
const GROUP_WEIGHT_MODE = 'RANK';
const RANK_WEIGHT_MIN = 1.3;
const RANK_WEIGHT_MAX = 1.8;

const UPSTREAM_ONLY_POSITIVE = false;

const ENABLE_NODE_CLICK = true;
const MIN_RENDER_W = 75;
const MIN_RENDER_H = 20;
const MIN_RENDER_AREA = 400;

let revenueRows = [], linksRows = [], downRows = [], months = [];
let byCode = new Map();
let byName = new Map();
let linksByUp = new Map();
let linksByDown = new Map();
let downstreamHJ = [];

// ===== 基本工具 =====
function z(s){ return String(s == null ? '' : s); }
function toHalfWidth(str){ return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)); }
function normText(s){ return z(s).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/[\u3000]/g,' ').replace(/\s+/g,' ').trim(); }
function normCode(s){ return toHalfWidth(z(s)).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\s+/g,'').trim(); }

function displayPct(v){
  if(v == null || !isFinite(v)) return '—';
  const s = v.toFixed(1) + '%';
  return v > 0 ? ('+' + s) : s;
}

function colorFor(v, mode){
  if(v == null || !isFinite(v)) return '#0f172a';
  const t = Math.min(1, Math.abs(v)/80);
  const alpha = 0.25 + 0.35*t;
  const good = (mode === 'greenPositive');
  const pos = good ? '16,185,129' : '239,68,68';
  const neg = good ? '239,68,68' : '16,185,129';
  const rgb = (v >= 0) ? pos : neg;
  return `rgba(${rgb},${alpha})`;
}

function isUSCode(code){
  return /\.US$/i.test(String(code || '').trim());
}

// ===== 初始化 =====
window.addEventListener('DOMContentLoaded', async () => {
  try {
    await loadWorkbook();
    initControls();
    // ❌ 已移除 setupDownloadButton();
  } catch (e) {
    console.error(e);
    alert('載入失敗：' + e.message);
  }

  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
});

// ===== 載入 Excel =====
async function loadWorkbook(){
  const res = await fetch(XLSX_FILE, { cache:'no-store' });
  if (!res.ok) throw new Error('讀取 data.xlsx 失敗 HTTP ' + res.status);

  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf, { type:'array' });

  const wsRev = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];
  const wsDown = wb.Sheets[DOWNLINKS_SHEET];

  if (!wsRev || !wsLinks) throw new Error('找不到 Revenue 或 Links');

  revenueRows = XLSX.utils.sheet_to_json(wsRev, { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });
  downRows    = wsDown ? XLSX.utils.sheet_to_json(wsDown, { defval:null }) : [];

  byCode.clear();
  byName.clear();

  const sample = revenueRows[0] || {};
  const codeKeyName = CODE_FIELDS.find(k => k in sample) || CODE_FIELDS[0];
  const nameKeyName = NAME_FIELDS.find(k => k in sample) || NAME_FIELDS[0];

  for (const r of revenueRows) {
    const code = normCode(r[codeKeyName]);
    const name = normText(r[nameKeyName]);
    if (code) byCode.set(code, r);
    if (name) byName.set(name, r);
  }

  linksByUp.clear();
  linksByDown.clear();

  for (const e of linksRows) {
    const A = normCode(e['上游代號']);
    const B = normCode(e['下游代號']);

    if (A && B) {
      if (!linksByUp.has(A)) linksByUp.set(A, []);
      linksByUp.get(A).push(e);

      if (!linksByDown.has(B)) linksByDown.set(B, []);
      linksByDown.get(B).push(e);
    }
  }

  downstreamHJ = downRows.map(r => ({
    '上游代號': normCode(r['上游代號']),
    '下游代號': normCode(r['下游代號']),
    '關係類型': normText(r['關係類型'])
  }));
}

// ===== UI =====
function initControls(){
  const sel = document.querySelector('#monthSelect');
  sel.innerHTML = '';

  months = Object.keys(COL_MAP).sort((a,b)=>b.localeCompare(a));

  for (const m of months) {
    const o = document.createElement('option');
    o.value = m;
    o.textContent = `${m.slice(0,4)}年${m.slice(4,6)}月`;
    sel.appendChild(o);
  }

  if (!sel.value && months.length > 0) sel.value = months[0];
}

// ===== 查詢 =====
function handleRun(){
  const raw = document.querySelector('#stockInput').value;
  if (!raw) return alert('請輸入股票');

  let codeKey = normCode(raw);
  let rowSelf = byCode.get(codeKey);

  if (!rowSelf) {
    const nameQ = normText(raw);
    rowSelf = byName.get(nameQ);
  }

  if (!rowSelf) return alert('找不到');

  const upstreamEdges = linksByDown.get(codeKey) || [];
  const downstreamEdges = downstreamHJ.filter(e => e['上游代號'] === codeKey);

  renderTreemap('upTreemap', upstreamEdges);
  renderTreemap('downTreemap', downstreamEdges);
}

// ===== Treemap（簡化保留核心）=====
function renderTreemap(id, edges){
  const svg = d3.select('#' + id);
  svg.selectAll('*').remove();

  const wrap = svg.node().parentElement;
  const W = wrap.clientWidth;
  const H = 400;

  svg.attr('width', W).attr('height', H);

  if (!edges.length) return;

  const data = {
    children: edges.map(e => ({
      name: e['下游代號'] || e['上游代號'],
      value: 1
    }))
  };

  const root = d3.hierarchy(data).sum(d => d.value);

  d3.treemap().size([W, H])(root);

  const g = svg.append('g');

  g.selectAll('rect')
    .data(root.leaves())
    .enter()
    .append('rect')
    .attr('x', d => d.x0)
    .attr('y', d => d.y0)
    .attr('width', d => d.x1 - d.x0)
    .attr('height', d => d.y1 - d.y0)
    .attr('fill', '#3b82f6');
}
