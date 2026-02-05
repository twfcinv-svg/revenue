/* app.js — 欄位自動對齊＋查詢卡片變色＋下載連結白色（請覆蓋既有檔案） */

// 版本參數（避免快取）
const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
// 以相對路徑組合，確保在 GitHub Pages 專案站點 repo 子路徑下也可正確讀取
const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();

// 工作表名稱
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';

// 指標對照
const COL_SUFFIX = { YoY:'年成長', MoM:'月變動' };

// 欄位別名（不同資料來源的表頭容錯）
const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];

// 每個月份對應到實際欄位名（避免括號/全半形等差異）
// 形如：COL_MAP['202512'] = { YoY: '202512單月合併營收年成長（％）', MoM: '...' }
const COL_MAP = {};

let revenueRows = [], linksRows = [], months = [];
let byCode = new Map();
let byName = new Map();

// --------- 工具函式 ---------
function z(s){ return String(s==null?'':s); }
function toHalfWidth(str){ return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); }
function normText(s){ return z(s).replace(/[​-‍﻿]/g,'').replace(/[　]/g,' ').replace(/\s+/g,' ').trim(); }
function normCode(s){ return toHalfWidth(z(s)).replace(/[​-‍﻿]/g,'').replace(/\s+/g,'').trim(); }
function displayPct(v){ if(v==null||!isFinite(v)) return '—'; const s=v.toFixed(1)+'%'; return v>0?('+'+s):s; }
function colorFor(v, mode){ if(v==null||!isFinite(v)) return '#0f172a'; const t=Math.min(1,Math.abs(v)/80); const alpha=0.25+0.35*t; const good=(mode==='greenPositive'); const pos=good?'16,185,129':'239,68,68'; const neg=good?'239,68,68':'16,185,129'; const rgb=(v>=0)?pos:neg; return `rgba(${rgb},${alpha})`; }
function safe(s){ return z(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[c])); }

// --------- 進入點 ---------
window.addEventListener('DOMContentLoaded', async()=>{
  // 下載連結白色 & 帶版本
  const a=document.getElementById('dlData'); if(a){ a.href='data.xlsx?v='+URL_VER; a.style.color='#fff'; }
  try{ await loadWorkbook(); initControls(); }catch(e){ console.error(e); alert('載入失敗：'+e.message); }
  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
});

// --------- 載入與索引 ---------
async function loadWorkbook(){
  const res = await fetch(XLSX_FILE, { cache:'no-store' });
  if(!res.ok) throw new Error('讀取 data.xlsx 失敗 HTTP '+res.status);
  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf, { type:'array' });

  const wsRev   = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];
  if(!wsRev || !wsLinks) throw new Error('找不到必要工作表 Revenue 或 Links');

  revenueRows = XLSX.utils.sheet_to_json(wsRev,   { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });

  byCode.clear(); byName.clear();

  // 1) 依實際欄位名建立代號/名稱索引
  const sample = revenueRows[0] || {};
  const codeKeyName = CODE_FIELDS.find(k => k in sample) || '個股';
  const nameKeyName = NAME_FIELDS.find(k => k in sample) || '名稱';

  for(const r of revenueRows){
    const code = normCode(r[codeKeyName]);
    const name = normText(r[nameKeyName]);
    if(code) byCode.set(code, r);
    if(name) byName.set(name, r);
  }

  // 2) 掃描表頭 → 建立月份清單與欄位對照（保留原始欄名）
  const found = new Set();
  for(const rawHeader of Object.keys(sample)){
    const h = normText(rawHeader);

    // 年成長（YoY）
    let m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*年[成增]長\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if(m){
      const ym = m[1] + String(m[2]).padStart(2,'0');
      COL_MAP[ym] = COL_MAP[ym] || {};
      COL_MAP[ym].YoY = rawHeader; // 用原始表頭名稱取值
      found.add(ym);
      continue;
    }
    // 月變動（MoM）
    m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*月[變增]動\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if(m){
      const ym = m[1] + String(m[2]).padStart(2,'0');
      COL_MAP[ym] = COL_MAP[ym] || {};
      COL_MAP[ym].MoM = rawHeader;
      found.add(ym);
      continue;
    }
  }
  months = Array.from(found).sort((a,b)=>b.localeCompare(a));
}

function initControls(){
  const sel=document.querySelector('#monthSelect');
  sel.innerHTML='';
  for(const m of months){
    const o=document.createElement('option');
    o.value=m;
    o.textContent=`${m.slice(0,4)}年${m.slice(4,6)}月`;
    sel.appendChild(o);
  }
  if(!sel.value && months.length>0) sel.value=months[0];
}

// 依據 COL_MAP 讀取實際欄位，容錯移除百分號
function getMetricValue(row, month, metric){
  if(!row || !month || !metric) return null;
  const col = (COL_MAP[month] || {})[metric];
  if(!col) return null;
  let v = row[col];
  if(v==null || v==='') return null;
  if(typeof v === 'string') v = v.replace('%','').replace('％','').trim();
  v = Number(v);
  return Number.isFinite(v) ? v : null;
}

// --------- 查詢 ---------
function handleRun(){
  const raw     = document.querySelector('#stockInput').value;
  const month   = (document.querySelector('#monthSelect')?.value)||'';
  const metric  = (document.querySelector('#metricSelect')?.value)||'MoM';
  const colorMode=(document.querySelector('#colorMode')?.value)||'redPositive';

  if(!raw || !raw.trim()){ alert('請輸入股票代號或公司名稱'); return; }

  let codeKey = normCode(raw);
  let rowSelf = byCode.get(codeKey);

  if(!rowSelf){
    // 名稱完整或前綴比對
    const nameQ = normText(raw);
    rowSelf = byName.get(nameQ) || revenueRows.find(r => normText(r['名稱']||r['公司名稱']||r['證券名稱']||'').startsWith(nameQ));
    if(rowSelf){
      codeKey = normCode(
        rowSelf['個股'] ?? rowSelf['代號'] ?? rowSelf['股票代碼'] ??
        rowSelf['股票代號'] ?? rowSelf['公司代號'] ?? rowSelf['證券代號']
      );
    }
  }

  if(!rowSelf){ alert('找不到此代號/名稱'); return; }

  const upstreamEdges   = linksRows.filter(r => normCode(r['下游代號']) === codeKey);
  const downstreamEdges = linksRows.filter(r => normCode(r['上游代號']) === codeKey);

  renderResultChip(rowSelf, month, metric, colorMode);
  renderTreemap('upTreemap','upHint',   upstreamEdges,  '上游代號', month, metric, colorMode);
  renderTreemap('downTreemap','downHint',downstreamEdges,'下游代號', month, metric, colorMode);
}

// --------- 呈現 ---------
function renderResultChip(selfRow, month, metric, colorMode){
  const host=document.querySelector('#resultChip');
  const v=getMetricValue(selfRow,month,metric);
  const bg=colorFor(v, colorMode);
  const showCode = selfRow['個股'] || selfRow['代號'] || selfRow['股票代碼'] || selfRow['股票代號'] || selfRow['公司代號'] || selfRow['證券代號'] || '';
  const showName = selfRow['名稱'] || selfRow['公司名稱'] || selfRow['證券名稱'] || '';
  host.innerHTML=`
    <div class="result-card" style="background:${bg}">
      <div class="row1"><strong>${safe(showCode)}｜${safe(showName)}</strong><span>${month.slice(0,4)}/${month.slice(4,6)} / ${metric}</span></div>
      <div class="row2"><span>${safe(selfRow['產業別']||'')}</span><span>${displayPct(v)}</span></div>
    </div>`;
}

function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode){
  const svg=d3.select('#'+svgId); svg.selectAll('*').remove();
  const wrap=svg.node().parentElement;
  const W=wrap.clientWidth-16;
  const H=parseInt(getComputedStyle(svg.node()).height)||560;
  svg.attr('width',W).attr('height',H);

  const groups=new Map();
  for(const e of edges){
    const rel=normText(e['關係類型']||'未分類');
    const key=normCode(e[codeField]);
    const r=byCode.get(key);
    if(!r) continue;
    const v=getMetricValue(r,month,metric);
    if(v==null) continue;
    if(!groups.has(rel)) groups.set(rel,[]);
    const codeVal = r['個股'] ?? r['代號'] ?? r['股票代碼'] ?? r['股票代號'] ?? r['公司代號'] ?? r['證券代號'];
    const nameVal = r['名稱'] ?? r['公司名稱'] ?? r['證券名稱'];
    groups.get(rel).push({ code:codeVal, name:nameVal, value:v });
  }

  const hint=document.getElementById(hintId);
  if(groups.size===0){ hint.textContent='此區在選定月份沒有可用數據'; return; } else { hint.textContent=''; }

  const children=[];
  for(const [rel,list] of groups){
    const avg=d3.mean(list,d=>d.value);
    const kids=list.map(s=>({ name: s.name||'', code:s.code, value:Math.max(0.01,Math.abs(s.value)), raw:s.value }));
    children.push({ name:rel, avg, children:kids });
  }

  const root=d3.hierarchy({ children }).sum(d=>d.value).sort((a,b)=>(b.value||0)-(a.value||0));
  d3.treemap().size([W,H]).paddingOuter(8).paddingInner(3).paddingTop(22)(root);

  const g=svg.append('g');

  // 群組底色（先畫）
  const parents=g.selectAll('g.parent').data(root.children||[]).enter().append('g').attr('class','parent');
  parents.append('rect').attr('class','group-bg')
    .attr('x',d=>d.x0).attr('y',d=>d.y0)
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=> colorFor(d.data.avg, colorMode));
  parents.append('rect').attr('class','group-border')
    .attr('x',d=>d.x0).attr('y',d=>d.y0)
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0));
  parents.append('text').attr('class','node-title')
    .attr('x', d=>d.x0+6).attr('y', d=>d.y0+16)
    .text(d=>`${d.data.name}  平均：${displayPct(d.data.avg)}`);

  // 葉節點（個股）— 單行：代號 中文名 數值
  const node=g.selectAll('g.node').data(root.leaves()).enter().append('g').attr('class','node').attr('transform',d=>`translate(${d.x0},${d.y0})`);
  node.append('rect').attr('class','node-rect')
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=> colorFor(d.data.raw, colorMode));
  node.append('text').attr('class','node-line').attr('x',6).attr('y',16)
    .text(d=>`${(d.data.code||'')} ${safe(d.data.name||'')} ${displayPct(d.data.raw)}`)
    .each(function(d){ const w=d.x1-d.x0; if(this.getBBox().width>w-8){ d3.select(this).attr('opacity',0.9).attr('font-size',10); }});
}
