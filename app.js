
/* app.js — 上下游 Treemap（皆依關係類型分群） */
const XLSX_FILE = 'data.xlsx';
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';
const COL_SUFFIX = { YoY:'單月合併營收年成長(%)', MoM:'單月合併營收月變動(%)' };

let revenueRows=[], linksRows=[], byCode=new Map(), months=[];

window.addEventListener('DOMContentLoaded', async () => {
  setStatus('載入資料中…');
  try{ await loadWorkbook(XLSX_FILE); initControls(); setStatus('資料就緒'); }
  catch(e){ console.error(e); setStatus('載入失敗：'+e.message); }
  document.querySelector('#runBtn').addEventListener('click', handleRun);
  window.addEventListener('resize', throttle(()=> { // 視窗改變時重畫
    const code = document.querySelector('#stockInput').value.trim();
    if(code && byCode.has(code)) handleRun();
  }, 300));
});

function setStatus(s){ document.querySelector('#status').textContent = s || ''; }

async function loadWorkbook(url){
  const res = await fetch(url); if(!res.ok) throw new Error('HTTP '+res.status);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type:'array' });
  const wsRev = wb.Sheets[REVENUE_SHEET]; const wsLinks=wb.Sheets[LINKS_SHEET];
  if(!wsRev || !wsLinks) throw new Error('找不到必要工作表 Revenue 或 Links');
  revenueRows = XLSX.utils.sheet_to_json(wsRev, { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });
  byCode.clear();
  for(const r of revenueRows){ const c=String(r['個股']).trim(); if(c) byCode.set(c, r); }
  const headers = Object.keys(revenueRows[0]||{});
  const y=headers.filter(h=>h.endsWith(COL_SUFFIX.YoY)).map(h=>h.slice(0,6));
  const m=headers.filter(h=>h.endsWith(COL_SUFFIX.MoM)).map(h=>h.slice(0,6));
  months = Array.from(new Set([...y,...m])).sort((a,b)=>b.localeCompare(a));
}

function initControls(){
  const sel=document.querySelector('#monthSelect'); sel.innerHTML='';
  for(const mm of months){ const o=document.createElement('option'); o.value=mm; o.textContent=`${mm.slice(0,4)}年${mm.slice(4,6)}月`; sel.appendChild(o);} 
  if(months.length===0) setStatus('警告：尚未偵測到月份欄位');
}

function handleRun(){
  const code = document.querySelector('#stockInput').value.trim();
  const month= document.querySelector('#monthSelect').value;
  const metric=document.querySelector('#metricSelect').value;
  const colorMode=document.querySelector('#colorMode').value;
  if(!code) return setStatus('請輸入股票代號');
  if(!byCode.has(code)) return setStatus(`找不到代號 ${code} 的公司於 Revenue 表`);
  setStatus('');

  // 查詢結果 chip（按鈕旁）
  renderResultChip(code, month, metric);

  // 上游：Links 中 下游代號 == code → 使用 上游代號
  const upstreamEdges = linksRows.filter(r=>String(r['下游代號']).trim()===code);
  renderTreemap('upTreemap','upHint', upstreamEdges, '上游代號', month, metric, colorMode);

  // 下游：Links 中 上游代號 == code → 使用 下游代號
  const downstreamEdges = linksRows.filter(r=>String(r['上游代號']).trim()===code);
  renderTreemap('downTreemap','downHint', downstreamEdges, '下游代號', month, metric, colorMode);
}

function renderResultChip(code, month, metric){
  const host = document.querySelector('#resultChip');
  const row = byCode.get(code);
  const v = getMetricValue(row, month, metric);
  const cls = v==null?'':(v>=0?'good':'bad');
  host.innerHTML = `
    <div class="result-card">
      <div class="row1"><strong>${code}｜${safe(row['名稱'])}</strong><span>${month.slice(0,4)}/${month.slice(4,6)} / ${metric}</span></div>
      <div class="row2"><span>${safe(row['產業別']||'')}</span><span class="${cls}">${displayPct(v)}</span></div>
    </div>`;
}

/**
 * 通用 Treemap 渲染（上下游共用）
 * @param {string} svgId  - SVG 元素 id
 * @param {string} hintId - 提示元素 id
 * @param {Array<Object>} edges - Links 篩選後列
 * @param {string} codeField - '上游代號' 或 '下游代號'（要取個股的欄位）
 */
function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode){
  const svg = d3.select('#'+svgId); svg.selectAll('*').remove();
  const wrap = svg.node().parentElement; // .treemap-wrap
  const W = wrap.clientWidth - 16; const H = parseInt(getComputedStyle(svg.node()).height);
  svg.attr('width', W).attr('height', H);

  const groups = new Map(); // rel -> [{code,name,value}]
  for(const e of edges){
    const rel = String(e['關係類型']||'未分類').trim();
    const c = String(e[codeField]).trim();
    const r = byCode.get(c); if(!r) continue;
    const v = getMetricValue(r, month, metric); if(v==null) continue;
    if(!groups.has(rel)) groups.set(rel, []);
    groups.get(rel).push({ code:c, name:r['名稱'], value:v });
  }
  const hint = document.getElementById(hintId);
  if(groups.size===0){ hint.textContent='此區在選定月份沒有可用數據'; return; } else { hint.textContent=''; }

  // 建立階層資料：父層=關係類型，子層=個股
  const children = [];
  for(const [rel, list] of groups){
    const avg = d3.mean(list, d=>d.value);
    const kids = list.map(s=>({ name:`${s.code} ${s.name||''}`.trim(), code:s.code, value:Math.max(0.01, Math.abs(s.value)), raw:s.value }));
    children.push({ name: rel, avg, children: kids });
  }
  const root = d3.hierarchy({ name:'root', children }).sum(d=>d.value).sort((a,b)=>(b.value||0)-(a.value||0));
  d3.treemap().size([W,H]).paddingOuter(8).paddingInner(3)(root);

  const g = svg.append('g');

  // 群組框與標題
  const parents = g.selectAll('g.parent').data(root.children||[]).enter().append('g').attr('class','parent');
  parents.append('rect').attr('class','group-border')
    .attr('x', d=>d.x0).attr('y', d=>d.y0)
    .attr('width', d=>Math.max(0,d.x1-d.x0)).attr('height', d=>Math.max(0,d.y1-d.y0));
  parents.append('text').attr('class','node-title')
    .attr('x', d=>d.x0+6).attr('y', d=>d.y0+16)
    .text(d=>`${d.data.name}  平均：${displayPct(d.data.avg)}`);

  // 子節點（個股）
  const node = g.selectAll('g.node').data(root.leaves()).enter().append('g').attr('class','node')
    .attr('transform', d=>`translate(${d.x0},${d.y0})`);
  node.append('rect').attr('class','node-rect')
    .attr('width', d=>Math.max(0,d.x1-d.x0)).attr('height', d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=>colorFor(d.data.raw, colorMode));
  node.append('text').attr('class','node-sub').attr('x',6).attr('y',16)
    .text(d=>`${(d.data.code||'').slice(0,6)} ${displayPct(d.data.raw)}`)
    .each(function(d){ const w=d.x1-d.x0; if(this.getBBox().width>w-8){ d3.select(this).attr('opacity',0.85).attr('font-size',10); }});
  node.append('text').attr('class','node-sub').attr('x',6).attr('y',30)
    .text(d=>`${safe((d.data.name||'').slice(0,8))}`)
    .each(function(d){ const w=d.x1-d.x0; if(this.getBBox().width>w-8){ d3.select(this).attr('opacity',0.8).attr('font-size',10); }});
}

// ------- utils -------
function getMetricValue(row, month, metric){ if(!row) return null; const col=`${month}${COL_SUFFIX[metric]}`; let v=row[col]; if(v==null||v==='') return null; v=Number(v); return Number.isFinite(v)?v:null; }
function displayPct(v){ if(v==null||!isFinite(v)) return '—'; const s=v.toFixed(1)+'%'; return v>0?('+'+s):s; }
function safe(s){ return String(s??'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function colorFor(v, mode){ if(v==null||!isFinite(v)) return '#0f172a'; const t=Math.min(1, Math.abs(v)/80); const alpha=0.18+0.42*t; const good=(mode==='greenPositive'); const pos=good? '16,185,129':'239,68,68'; const neg=good? '239,68,68':'16,185,129'; const rgb=(v>=0)?pos:neg; return `rgba(${rgb},${alpha})`; }
function throttle(fn, wait){ let t=0; return (...args)=>{ const now=Date.now(); if(now-t>wait){ t=now; fn(...args); } }; }
