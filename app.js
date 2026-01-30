
/* app.js — 左側 Treemap（上游分群）＋ 右側清單（下游不分群） */
const XLSX_FILE = 'data.xlsx';
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';
const COL_SUFFIX = { YoY:'單月合併營收年成長(%)', MoM:'單月合併營收月變動(%)' };

let revenueRows = []; let linksRows = []; let byCode = new Map(); let months = [];

window.addEventListener('DOMContentLoaded', async () => {
  setStatus('載入資料中…');
  try{
    await loadWorkbook(XLSX_FILE);
    initControls(); setStatus('資料就緒');
  }catch(e){ console.error(e); setStatus('載入失敗：'+e.message); }
  document.querySelector('#runBtn').addEventListener('click', handleRun);
  window.addEventListener('resize', () => {
    // 若已有查詢，重畫 treemap
    const code = document.querySelector('#stockInput').value.trim();
    if(code && byCode.has(code)) handleRun();
  });
});

function setStatus(s){ document.querySelector('#status').textContent = s || ''; }

async function loadWorkbook(url){
  const res = await fetch(url); if(!res.ok) throw new Error('HTTP '+res.status);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type:'array' });
  const wsRev = wb.Sheets[REVENUE_SHEET]; const wsLinks = wb.Sheets[LINKS_SHEET];
  if(!wsRev || !wsLinks) throw new Error('找不到必要工作表 Revenue 或 Links');
  revenueRows = XLSX.utils.sheet_to_json(wsRev, { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });
  byCode.clear();
  for(const r of revenueRows){ const c = String(r['個股']).trim(); if(c) byCode.set(c, r); }
  const headers = Object.keys(revenueRows[0]||{});
  const y = headers.filter(h=>h.endsWith(COL_SUFFIX.YoY)).map(h=>h.slice(0,6));
  const m = headers.filter(h=>h.endsWith(COL_SUFFIX.MoM)).map(h=>h.slice(0,6));
  months = Array.from(new Set([...y,...m])).sort((a,b)=>b.localeCompare(a));
}

function initControls(){
  const sel = document.querySelector('#monthSelect'); sel.innerHTML='';
  for(const mm of months){ const o=document.createElement('option'); o.value=mm; o.textContent=`${mm.slice(0,4)}年${mm.slice(4,6)}月`; sel.appendChild(o);} 
  if(months.length===0) setStatus('警告：尚未偵測到月份欄位');
}

function handleRun(){
  const code = document.querySelector('#stockInput').value.trim();
  const month = document.querySelector('#monthSelect').value;
  const metric = document.querySelector('#metricSelect').value;
  const colorMode = document.querySelector('#colorMode').value; // redPositive | greenPositive
  if(!code) return setStatus('請輸入股票代號');
  if(!byCode.has(code)) return setStatus(`找不到代號 ${code} 的公司於 Revenue 表`);
  setStatus('');

  renderFocus(code, month, metric);
  // 左側 Treemap = 上游（Links: 下游代號 == code）依關係類型分群
  const upstreamEdges = linksRows.filter(r => String(r['下游代號']).trim() === code);
  renderUpstreamTreemap(upstreamEdges, month, metric, colorMode);
  // 右側清單 = 下游（Links: 上游代號 == code）不分群
  const downstreamEdges = linksRows.filter(r => String(r['上游代號']).trim() === code);
  renderDownstreamList(downstreamEdges, month, metric, colorMode);
}

function renderFocus(code, month, metric){
  const wrap = document.querySelector('#focusCard'); const row = byCode.get(code);
  const v = getMetricValue(row, month, metric); const cls = v==null?'':(v>=0?'good':'bad');
  wrap.innerHTML = `<div class="title" style="display:flex;justify-content:space-between;gap:8px">
      <strong>${code}｜${safe(row['名稱'])}</strong>
      <span class="legend chip">${month.slice(0,4)}/${month.slice(4,6)} / ${metric}</span>
    </div>
    <div class="meta">${safe(row['產業別']||'')}</div>
    <div class="meta" style="margin-top:4px">本月：<span class="${cls}">${displayPct(v)}</span></div>`;
}

// --------- 左側 Treemap ---------
function renderUpstreamTreemap(edges, month, metric, colorMode){
  const svg = d3.select('#treemap');
  const wrap = document.getElementById('treemapWrap');
  const W = wrap.clientWidth - 16, H = document.getElementById('treemap').clientHeight; // height from CSS
  svg.attr('width', W).attr('height', H);
  svg.selectAll('*').remove();

  const groups = new Map(); // rel -> [{code,name,value}]
  for(const e of edges){
    const rel = String(e['關係類型']||'未分類').trim();
    const c = String(e['上游代號']).trim();
    const r = byCode.get(c); if(!r) continue;
    const v = getMetricValue(r, month, metric); if(v==null) continue;
    if(!groups.has(rel)) groups.set(rel, []);
    groups.get(rel).push({ code:c, name:r['名稱'], value:v });
  }
  if(groups.size===0){ d3.select('#treemapHint').text('此月份在上游沒有可用數據'); return; } else { d3.select('#treemapHint').text(''); }

  // 構建 d3-hierarchy
  const children = [];
  for(const [rel, list] of groups){
    const relAvg = d3.mean(list, d=>d.value);
    children.push({ name: rel, avg: relAvg, children: list.map(s=>({ name: `${s.code} ${s.name||''}`.trim(), code:s.code, value: Math.max(0.01, Math.abs(s.value)), raw: s.value })) });
  }
  const root = d3.hierarchy({ name:'root', children }).sum(d=>d.value).sort((a,b)=> (b.value||0)-(a.value||0));
  d3.treemap().size([W,H]).paddingOuter(8).paddingInner(3)(root);

  const g = svg.append('g');

  const node = g.selectAll('g.node').data(root.leaves()).enter().append('g').attr('class','node')
    .attr('transform', d=>`translate(${d.x0},${d.y0})`);

  // 葉子（個股）矩形
  node.append('rect')
      .attr('width', d=>Math.max(0, d.x1-d.x0))
      .attr('height', d=>Math.max(0, d.y1-d.y0))
      .attr('fill', d=>colorFor(d.data.raw, colorMode))
      .attr('rx', 6).attr('ry', 6);

  // 群組框（在 parent 層畫一次）
  const parents = g.selectAll('g.parent').data(root.children||[]).enter().append('g').attr('class','parent');
  parents.append('rect')
    .attr('x', d=>d.x0).attr('y', d=>d.y0)
    .attr('width', d=>Math.max(0, d.x1-d.x0)).attr('height', d=>Math.max(0, d.y1-d.y0))
    .attr('fill','none').attr('stroke','#334155').attr('stroke-width',1.2).attr('rx',10).attr('ry',10);

  // 群組標題
  parents.append('text')
    .attr('x', d=>d.x0+6).attr('y', d=>d.y0+16)
    .attr('class','node-title')
    .text(d=>`${d.data.name}  平均：${displayPct(d.data.avg)}`);

  // 葉子標籤（代號 + 值）
  node.append('text')
      .attr('x', 6).attr('y', 16)
      .attr('class','node-sub')
      .text(d=> `${(d.data.code||'').slice(0,6)}  ${displayPct(d.data.raw)}`)
      .each(function(d){ // 超出寬度就隱藏
        const w = d.x1-d.x0; if(this.getBBox().width > w-8){ d3.select(this).attr('opacity',0.85).attr('font-size', 10); }
      });
}

// --------- 右側清單（不分群） ---------
function renderDownstreamList(edges, month, metric, colorMode){
  const list = document.getElementById('downstreamList'); list.innerHTML='';
  if(edges.length===0){ list.innerHTML='<div class="item"><div class="meta">沒有連結資料</div></div>'; return; }
  const codes = edges.map(e => String(e['下游代號']).trim());
  for(const c of codes){
    const r = byCode.get(c);
    if(!r){ list.insertAdjacentHTML('beforeend', `<div class="item"><div class="meta">${c}（在 Revenue 中找不到）</div></div>`); continue; }
    const v = getMetricValue(r, month, metric);
    const cls = v==null?'':(v>=0?'good':'bad');
    const width = pctWidth(v);
    const fill = barColor(v, colorMode);
    list.insertAdjacentHTML('beforeend', `
      <div class="item">
        <div class="title"><strong>${c}｜${safe(r['名稱'])}</strong><span class="val ${cls}">${displayPct(v)}</span></div>
        <div class="meta">${safe(r['產業別']||'')}</div>
        <div class="mini-bar"><div class="mini-fill" style="width:${width}%; background:${fill}"></div></div>
      </div>`);
  }
}

// ---------- Utilities ----------
function getMetricValue(row, month, metric){ if(!row) return null; const col = `${month}${COL_SUFFIX[metric]}`; let v=row[col]; if(v==null||v==='') return null; v=Number(v); return Number.isFinite(v)?v:null; }
function displayPct(v){ if(v==null||!isFinite(v)) return '—'; const s=v.toFixed(1)+'%'; return v>0?('+'+s):s; }
function pctWidth(v){ if(v==null||!isFinite(v)) return 0; const cap=80; return Math.max(0, Math.min(100, Math.round(Math.abs(v)/cap*100))); }
function safe(s){ return String(s??'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

function colorFor(v, mode){
  // 以透明度與色系表達強度；紅=好/綠=弱 或 反之（可切換）
  if(v==null||!isFinite(v)) return '#0f172a';
  const t = Math.min(1, Math.abs(v)/80); // 強度 0~1
  const alpha = 0.18 + 0.42*t; // 透明度 0.18~0.6
  const good = (mode==='greenPositive'); // 綠=好？
  const pos = good? '16,185,129' : '239,68,68';  // 綠 or 紅
  const neg = good? '239,68,68' : '16,185,129';  // 紅 or 綠
  const rgb = (v>=0)? pos : neg;
  return `rgba(${rgb},${alpha})`;
}
function barColor(v, mode){
  if(v==null||!isFinite(v)) return 'var(--neutral)';
  const good = (mode==='greenPositive');
  return (v>=0) ? (good? 'var(--good2)' : 'var(--good)')
                : (good? 'var(--bad2)'  : 'var(--bad)');
}
