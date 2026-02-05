/* app.js — v3.2 標籤仿 stock-heatmap1229 的字級：面積驅動 + 嚴格 fit + clipPath */

const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';
const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];
const COL_MAP = {};

let revenueRows = [], linksRows = [], months = [];
let byCode = new Map();
let byName = new Map();
let linksByUp = new Map();
let linksByDown = new Map();

function z(s){ return String(s==null?'':s); }
function toHalfWidth(str){ return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); }
function normText(s){ return z(s).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/[\u3000]/g,' ').replace(/\s+/g,' ').trim(); }
function normCode(s){ return toHalfWidth(z(s)).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\s+/g,'').trim(); }
function displayPct(v){ if(v==null||!isFinite(v)) return '—'; const s=v.toFixed(1)+'%'; return v>0?('+'+s):s; }
function colorFor(v, mode){ if(v==null||!isFinite(v)) return '#0f172a'; const t=Math.min(1,Math.abs(v)/80); const alpha=0.25+0.35*t; const good=(mode==='greenPositive'); const pos=good?'16,185,129':'239,68,68'; const neg=good?'239,68,68':'16,185,129'; const rgb=(v>=0)?pos:neg; return `rgba(${rgb},${alpha})`; }
function safe(s){ return z(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[c])); }

window.addEventListener('DOMContentLoaded', async()=>{
  try{ await loadWorkbook(); initControls(); setupDownloadButton(); }
  catch(e){ console.error(e); alert('載入失敗：'+e.message); }
  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
});

function setupDownloadButton(){
  const old = document.getElementById('dlData'); if (old) old.style.display = 'none';
  const a = document.createElement('a');
  a.href = 'data.xlsx?v='+URL_VER; a.textContent = '下載 data.xlsx';
  a.setAttribute('download',''); a.setAttribute('rel','noopener');
  Object.assign(a.style, { position:'fixed', top:'10px', right:'12px', zIndex:1000, background:'#fff', color:'#0f172a', padding:'6px 10px', borderRadius:'6px', textDecoration:'none', boxShadow:'0 1px 2px rgba(0,0,0,.25)', border:'1px solid rgba(15,23,42,.15)', fontSize:'13px', lineHeight:'1.2', fontWeight:'600' });
  document.body.appendChild(a);
}

async function loadWorkbook(){
  const res = await fetch(XLSX_FILE, { cache:'no-store' });
  if(!res.ok) throw new Error('讀取 data.xlsx 失敗 HTTP '+res.status);
  const buf = await res.arrayBuffer(); const wb  = XLSX.read(buf, { type:'array' });
  const wsRev = wb.Sheets[REVENUE_SHEET]; const wsLinks = wb.Sheets[LINKS_SHEET];
  if(!wsRev || !wsLinks) throw new Error('找不到必要工作表 Revenue 或 Links');

  const rowsHeaderFirst = XLSX.utils.sheet_to_json(wsRev, { header:1, blankrows:false });
  const headerRow = Array.isArray(rowsHeaderFirst) && rowsHeaderFirst.length>0 ? rowsHeaderFirst[0] : [];
  const found = new Set();
  for(const rawHeader of headerRow){
    if (!rawHeader) continue; const h = normText(String(rawHeader));
    let m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*年[成增]長\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if(m){ const ym=m[1]+String(m[2]).padStart(2,'0'); (COL_MAP[ym]??=({})).YoY = rawHeader; found.add(ym); continue; }
    m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*月[變增]動\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if(m){ const ym=m[1]+String(m[2]).padStart(2,'0'); (COL_MAP[ym]??=({})).MoM = rawHeader; found.add(ym); continue; }
  }
  months = Array.from(found).sort((a,b)=>b.localeCompare(a));

  revenueRows = XLSX.utils.sheet_to_json(wsRev,   { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });

  byCode.clear(); byName.clear();
  const sample = revenueRows[0] || {};
  const codeKeyName = CODE_FIELDS.find(k => k in sample) || CODE_FIELDS[0];
  const nameKeyName = NAME_FIELDS.find(k => k in sample) || NAME_FIELDS[0];
  for(const r of revenueRows){
    const code = normCode(r[codeKeyName]); const name = normText(r[nameKeyName]);
    if(code) byCode.set(code, r); if(name) byName.set(name, r);
  }

  linksByUp.clear(); linksByDown.clear();
  for(const e of linksRows){
    const up   = normCode(e['上游代號']); const down = normCode(e['下游代號']);
    if (up)   { if(!linksByUp.has(up))   linksByUp.set(up, []);   linksByUp.get(up).push(e); }
    if (down) { if(!linksByDown.has(down)) linksByDown.set(down, []); linksByDown.get(down).push(e); }
  }
}

function initControls(){
  const sel=document.querySelector('#monthSelect'); sel.innerHTML='';
  for(const m of months){ const o=document.createElement('option'); o.value=m; o.textContent=`${m.slice(0,4)}年${m.slice(4,6)}月`; sel.appendChild(o); }
  if(!sel.value && months.length>0) sel.value=months[0];
}

function getMetricValue(row, month, metric){
  if(!row || !month || !metric) return null; const col = (COL_MAP[month] || {})[metric];
  if(!col) return null; let v = row[col]; if(v==null || v==='') return null;
  if(typeof v === 'string') v = v.replace('%','').replace('％','').trim(); v = Number(v);
  return Number.isFinite(v) ? v : null;
}

function handleRun(){
  const raw = document.querySelector('#stockInput').value;
  const month = (document.querySelector('#monthSelect')?.value)||'';
  const metric = (document.querySelector('#metricSelect')?.value)||'MoM';
  const colorMode=(document.querySelector('#colorMode')?.value)||'redPositive';
  if(!raw || !raw.trim()){ alert('請輸入股票代號或公司名稱'); return; }

  let codeKey = normCode(raw); let rowSelf = byCode.get(codeKey);
  if(!rowSelf){
    const nameQ = normText(raw);
    rowSelf = byName.get(nameQ) || revenueRows.find(r => normText(r['名稱']||r['公司名稱']||r['證券名稱']||'').startsWith(nameQ));
    if(rowSelf){ codeKey = normCode(rowSelf['個股'] ?? rowSelf['代號'] ?? rowSelf['股票代碼'] ?? rowSelf['股票代號'] ?? rowSelf['公司代號'] ?? rowSelf['證券代號']); }
  }
  if(!rowSelf){ alert('找不到此代號/名稱'); return; }

  try{
    const codeLabel = (rowSelf['個股'] || rowSelf['代號'] || rowSelf['股票代碼'] || rowSelf['股票代號'] || rowSelf['公司代號'] || rowSelf['證券代號'] || '').trim();
    const nameLabel = (rowSelf['名稱'] || rowSelf['公司名稱'] || rowSelf['證券名稱'] || '').trim();
    const extra = `${month.slice(0,4)}/${month.slice(4,6)} · ${metric}`;
    if (window.setResultChipLink) window.setResultChipLink(codeLabel, nameLabel, extra);
  }catch(_){ }

  const upstreamEdges   = linksByDown.get(codeKey) || [];
  const downstreamEdges = linksByUp.get(codeKey)   || [];

  requestAnimationFrame(()=>{ renderResultChip(rowSelf, month, metric, colorMode);
    renderTreemap('upTreemap','upHint',   upstreamEdges,  '上游代號', month, metric, colorMode); });
  requestAnimationFrame(()=>{ renderTreemap('downTreemap','downHint',downstreamEdges,'下游代號', month, metric, colorMode); });
}

function renderResultChip(selfRow, month, metric, colorMode){
  const host=document.querySelector('#resultChip');
  const v=getMetricValue(selfRow,month,metric); const bg=colorFor(v, colorMode);
  const showCode = selfRow['個股'] || selfRow['代號'] || selfRow['股票代碼'] || selfRow['股票代號'] || selfRow['公司代號'] || selfRow['證券代號'] || '';
  const showName = selfRow['名稱'] || selfRow['公司名稱'] || selfRow['證券名稱'] || '';
  host.innerHTML=`
    <div class="result-card" style="background:${bg}">
      <div class="row1"><strong>${safe(showCode)}｜${safe(showName)}</strong><span>${month.slice(0,4)}/${month.slice(4,6)} / ${metric}</span></div>
      <div class="row2"><span>${safe(selfRow['產業別']||'')}</span><span>${displayPct(v)}</span></div>
    </div>`;
}

// ====== 字級演算法（面積驅動 + 嚴格 fit + clip） ======
const LabelFit = {
  padding: 8,
  maxFont: 36,     // 全域最大上限（保護超大格）
  minFont: 9,
  lineHeight: 1.15,

  centerText(el, w, h) {
    el.setAttribute('text-anchor', 'middle');
    el.setAttribute('dominant-baseline', 'middle');
    el.setAttribute('x', this.padding + Math.max(0, (w - this.padding*2) / 2));
    el.setAttribute('y', this.padding + Math.max(0, (h - this.padding*2) / 2));
  },
  ensureClip(gEl, w, h){
    const svg = gEl.ownerSVGElement; let defs = svg.querySelector('defs');
    if (!defs) defs = svg.insertBefore(document.createElementNS('http://www.w3.org/2000/svg','defs'), svg.firstChild);
    const id = gEl.dataset.clipId || ('clip-' + Math.random().toString(36).slice(2)); gEl.dataset.clipId = id;
    let clip = svg.querySelector('#'+id);
    if (!clip){ clip = document.createElementNS('http://www.w3.org/2000/svg','clipPath'); clip.setAttribute('id', id);
      const r = document.createElementNS('http://www.w3.org/2000/svg','rect'); clip.appendChild(r); defs.appendChild(clip); }
    const rect = clip.firstChild; rect.setAttribute('x', 0); rect.setAttribute('y', 0);
    rect.setAttribute('width', Math.max(0, w)); rect.setAttribute('height', Math.max(0, h));
    gEl.querySelectorAll('text').forEach(t => t.setAttribute('clip-path', `url(#${id})`));
  },
  fitBlock(textEl, w, h){
    const targetW = Math.max(1, w - this.padding*2);
    const targetH = Math.max(1, h - this.padding*2);

    // 面積驅動的字級：與 sqrt(area) 成正比，兩行建議 k ≈ 0.11~0.14
    const k = 0.12;
    const area = Math.max(1, targetW * targetH);
    const areaFont = Math.sqrt(area) * k;

    // 上限依 cell 高度限制 45%，避免極端暴衝
    const logicalMax = Math.min(this.maxFont, Math.floor(targetH * 0.45));
    const logicalMin = this.minFont;
    const baseFont = Math.max(logicalMin, Math.min(logicalMax, Math.floor(areaFont)));

    textEl.setAttribute('font-size', baseFont);
    this.centerText(textEl, w, h);

    // 嚴格 fit：防止超寬/超高（例如名字很長）
    const bbox = textEl.getBBox();
    const scale = Math.min(targetW / Math.max(1,bbox.width), targetH / Math.max(1,bbox.height));
    const finalSize = Math.max(logicalMin, Math.min(logicalMax, Math.floor(baseFont * Math.min(1, scale))));

    if (finalSize < this.minFont) {
      textEl.setAttribute('display','none');
    } else {
      textEl.removeAttribute('display'); textEl.setAttribute('font-size', finalSize);
      this.centerText(textEl, w, h);
    }
  }
};

function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode){
  const svg=d3.select('#'+svgId); svg.selectAll('*').remove();
  const wrap=svg.node().parentElement; const W=wrap.clientWidth-16; const H=parseInt(getComputedStyle(svg.node()).height)||560;
  svg.attr('width',W).attr('height',H);

  const groups=new Map();
  for(const e of edges){
    const rel=normText(e['關係類型']||'未分類'); const key=normCode(e[codeField]); const r=byCode.get(key);
    if(!r) continue; const v=getMetricValue(r,month,metric); if(v==null) continue;
    if(!groups.has(rel)) groups.set(rel,[]);
    const codeVal = r['個股'] ?? r['代號'] ?? r['股票代碼'] ?? r['股票代號'] ?? r['公司代號'] ?? r['證券代號'];
    const nameVal = r['名稱'] ?? r['公司名稱'] ?? r['證券名稱'];
    groups.get(rel).push({ code:codeVal, name:nameVal, value:Math.max(0.01,Math.abs(v)), raw:v });
  }

  const hint=document.getElementById(hintId);
  if(groups.size===0){ hint.textContent='此區在選定月份沒有可用數據'; return; } else { hint.textContent=''; }

  const children=[]; for(const [rel,list] of groups){ const avg=d3.mean(list,d=>d.raw);
    const kids=list.map(s=>({ name: s.name||'', code:s.code, value:Math.max(0.01,Math.abs(s.raw)), raw:s.raw }));
    children.push({ name:rel, avg, children:kids }); }

  const root=d3.hierarchy({ children }).sum(d=>d.value).sort((a,b)=>(b.value||0)-(a.value||0));
  d3.treemap().size([W,H]).paddingOuter(8).paddingInner(3).paddingTop(22)(root);

  const g=svg.append('g');

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

  const node=g.selectAll('g.node').data(root.leaves()).enter().append('g').attr('class','node').attr('transform',d=>`translate(${d.x0},${d.y0})`);
  node.append('rect').attr('class','node-rect')
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=> colorFor(d.data.raw, colorMode));

  const labels = node.append('text')
    .attr('class','node-label')
    .attr('fill','#fff')
    .style('paint-order','stroke')
    .style('stroke','rgba(0,0,0,0.35)')
    .style('stroke-width','2px');

  labels.each(function(d){
    const name = `${d.data.code||''} ${d.data.name||''}`.trim();
    const vstr = displayPct(d.data.raw);
    const t1 = document.createElementNS('http://www.w3.org/2000/svg','tspan'); t1.setAttribute('x','0'); t1.setAttribute('dy','0'); t1.textContent = name;
    const t2 = document.createElementNS('http://www.w3.org/2000/svg','tspan'); t2.setAttribute('x','0'); t2.setAttribute('dy', `${LabelFit.lineHeight}em`); t2.textContent = vstr;
    this.appendChild(t1); this.appendChild(t2);
    const title = document.createElementNS('http://www.w3.org/2000/svg','title'); title.textContent = `${name} ${vstr}`; this.appendChild(title);
  });

  requestAnimationFrame(()=>{
    node.each(function(d){
      const w = Math.max(0, d.x1 - d.x0); const h = Math.max(0, d.y1 - d.y0);
      const textEl = this.querySelector('text'); if (!textEl) return;
      LabelFit.fitBlock(textEl, w, h); LabelFit.ensureClip(this, w, h);
    });
  });
}
