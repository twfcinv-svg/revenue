/* app.js — v2 性能修正版：
 1) 預先索引 Links（上/下游）
 2) 避免每個葉節點呼叫 getBBox（大量節點會讓瀏覽器卡死）
 3) 更嚴謹的 normText（顯式 Unicode 範圍）
*/

const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';
const COL_SUFFIX = { YoY:'年成長', MoM:'月變動' };
const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];
const COL_MAP = {};

let revenueRows = [], linksRows = [], months = [];
let byCode = new Map();
let byName = new Map();
let linksByUp = new Map();   // key: 上游代號（被當作對方的上游）
let linksByDown = new Map(); // key: 下游代號

function z(s){ return String(s==null?'':s); }
function toHalfWidth(str){ return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); }
// 嚴謹版本：移除零寬空白/連字/不換行 BOM 等隱形字元
function normText(s){ return z(s).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/[\u3000]/g,' ').replace(/\s+/g,' ').trim(); }
function normCode(s){ return toHalfWidth(z(s)).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\s+/g,'').trim(); }
function displayPct(v){ if(v==null||!isFinite(v)) return '—'; const s=v.toFixed(1)+'%'; return v>0?('+'+s):s; }
function colorFor(v, mode){ if(v==null||!isFinite(v)) return '#0f172a'; const t=Math.min(1,Math.abs(v)/80); const alpha=0.25+0.35*t; const good=(mode==='greenPositive'); const pos=good?'16,185,129':'239,68,68'; const neg=good?'239,68,68':'16,185,129'; const rgb=(v>=0)?pos:neg; return `rgba(${rgb},${alpha})`; }
function safe(s){ return z(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[c])); }

window.addEventListener('DOMContentLoaded', async()=>{
  const a=document.getElementById('dlData'); if(a){ a.href='data.xlsx?v='+URL_VER; a.style.color='#fff'; }
  try{ await loadWorkbook(); initControls(); }catch(e){ console.error(e); alert('載入失敗：'+e.message); }
  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
});

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

  byCode.clear(); byName.clear(); linksByUp.clear(); linksByDown.clear();

  const sample = revenueRows[0] || {};
  const codeKeyName = CODE_FIELDS.find(k => k in sample) || '個股';
  const nameKeyName = NAME_FIELDS.find(k => k in sample) || '名稱';

  for(const r of revenueRows){
    const code = normCode(r[codeKeyName]);
    const name = normText(r[nameKeyName]);
    if(code) byCode.set(code, r);
    if(name) byName.set(name, r);
  }

  // 建立月份對照
  const found = new Set();
  for(const rawHeader of Object.keys(sample)){
    const h = normText(rawHeader);
    let m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*年[成增]長\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if(m){ const ym=m[1]+String(m[2]).padStart(2,'0'); (COL_MAP[ym]??=( {} )).YoY = rawHeader; found.add(ym); continue; }
    m = h.match(/^(\d{4})[\/年-]?\s*(\d{1,2})\s*單月合併營收\s*月[變增]動\s*[\(（]?\s*(?:%|％)\s*[\)）]?$/);
    if(m){ const ym=m[1]+String(m[2]).padStart(2,'0'); (COL_MAP[ym]??=( {} )).MoM = rawHeader; found.add(ym); continue; }
  }
  months = Array.from(found).sort((a,b)=>b.localeCompare(a));

  // 預先索引 Links：避免每次查詢都全表掃描
  for(const e of linksRows){
    const up   = normCode(e['上游代號']);
    const down = normCode(e['下游代號']);
    if (up)   { if(!linksByUp.has(up))   linksByUp.set(up, []);   linksByUp.get(up).push(e); }
    if (down) { if(!linksByDown.has(down)) linksByDown.set(down, []); linksByDown.get(down).push(e); }
  }
}

function initControls(){
  const sel=document.querySelector('#monthSelect');
  sel.innerHTML='';
  for(const m of months){
    const o=document.createElement('option'); o.value=m; o.textContent=`${m.slice(0,4)}年${m.slice(4,6)}月`; sel.appendChild(o);
  }
  if(!sel.value && months.length>0) sel.value=months[0];
}

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

function handleRun(){
  const raw     = document.querySelector('#stockInput').value;
  const month   = (document.querySelector('#monthSelect')?.value)||'';
  const metric  = (document.querySelector('#metricSelect')?.value)||'MoM';
  const colorMode=(document.querySelector('#colorMode')?.value)||'redPositive';

  if(!raw || !raw.trim()){ alert('請輸入股票代號或公司名稱'); return; }

  let codeKey = normCode(raw);
  let rowSelf = byCode.get(codeKey);

  if(!rowSelf){
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

  // 使用預索引，避免在大量資料上兩次 filter
  const upstreamEdges   = linksByDown.get(codeKey) || []; // 下游=自己 => 找上游群
  const downstreamEdges = linksByUp.get(codeKey)   || []; // 上游=自己 => 找下游群

  // 讓渲染分幀，避免主執行緒長時間阻塞
  requestAnimationFrame(()=>{
    renderResultChip(rowSelf, month, metric, colorMode);
    renderTreemap('upTreemap','upHint',   upstreamEdges,  '上游代號', month, metric, colorMode);
  });
  requestAnimationFrame(()=>{
    renderTreemap('downTreemap','downHint',downstreamEdges,'下游代號', month, metric, colorMode);
  });
}

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

  // 群組底色
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

  // 葉節點（個股）— 為避免卡頓：取消逐個 getBBox 量測，只用固定字級，長度過長靠 CSS/視覺容忍
  const node=g.selectAll('g.node').data(root.leaves()).enter().append('g').attr('class','node').attr('transform',d=>`translate(${d.x0},${d.y0})`);
  node.append('rect').attr('class','node-rect')
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=> colorFor(d.data.raw, colorMode));
  node.append('text').attr('class','node-line').attr('x',6).attr('y',16)
    .attr('font-size', 11)
    .text(d=>`${(d.data.code||'')} ${safe(d.data.name||'')} ${displayPct(d.data.raw)}`);
}
