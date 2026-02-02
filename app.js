
/* app.js — 不快取 + 修正顏色 + 移除藍字狀態列 */
const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
const XLSX_FILE = new URL(`data.xlsx?v=${URL_VER}`, location.href).toString();
const REVENUE_SHEET='Revenue';
const LINKS_SHEET='Links';
const COL_SUFFIX={ YoY:'年成長', MoM:'月變動' };

let revenueRows=[], linksRows=[], months=[]; let byCode=new Map();
function z(s){ return String(s==null?'':s); }
function toHalfWidth(str){ return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); }
function normText(s){ return z(s).replace(/[​-‍﻿]/g,'').replace(/[　]/g,' ').replace(/\s+/g,' ').trim(); }
function normCode(s){ return toHalfWidth(z(s)).replace(/[​-‍﻿]/g,'').replace(/\s+/g,'').trim(); }
function displayPct(v){ if(v==null||!isFinite(v)) return '—'; const s=v.toFixed(1)+'%'; return v>0?('+'+s):s; }
function colorFor(v, mode){ if(v==null||!isFinite(v)) return '#0f172a'; const t=Math.min(1,Math.abs(v)/80); const alpha=0.18+0.42*t; const good=(mode==='greenPositive'); const pos=good?'16,185,129':'239,68,68'; const neg=good?'239,68,68':'16,185,129'; const rgb=(v>=0)?pos:neg; return `rgba(${rgb},${alpha})`; }
function safe(s){ return z(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

window.addEventListener('DOMContentLoaded', async()=>{
  // 設定「下載目前版本 data.xlsx」的連結，點了也不會拿到舊檔
  const a=document.getElementById('dlData'); if(a){ a.href='data.xlsx?v='+URL_VER; }
  try{ await loadWorkbook(); initControls(); }catch(e){ console.error(e); alert('載入失敗：'+e.message); }
  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
});

async function loadWorkbook(){
  const res=await fetch(XLSX_FILE,{cache:'no-store'});
  if(!res.ok) throw new Error('讀取 data.xlsx 失敗 HTTP '+res.status);
  const buf=await res.arrayBuffer();
  const wb=XLSX.read(buf,{type:'array'});
  const wsRev=wb.Sheets[REVENUE_SHEET]; const wsLinks=wb.Sheets[LINKS_SHEET];
  if(!wsRev||!wsLinks) throw new Error('找不到必要工作表 Revenue 或 Links');
  revenueRows=XLSX.utils.sheet_to_json(wsRev,{defval:null});
  linksRows  =XLSX.utils.sheet_to_json(wsLinks,{defval:null});
  byCode.clear();
  for(const r of revenueRows){ const key=normCode(r['個股']??r['代號']); if(key) byCode.set(key,r); }
  const headers=Object.keys(revenueRows[0]||{}).map(normText);
  const found=new Set();
  for(const h of headers){
    const m1=h.match(/^(\d{6})\s*單月合併營收\s*年成長\s*[\(（]?\s*(?:%|％)\s*[\)）]?\s*$/); if(m1) found.add(m1[1]);
    const m2=h.match(/^(\d{6})\s*單月合併營收\s*月變動\s*[\(（]?\s*(?:%|％)\s*[\)）]?\s*$/); if(m2) found.add(m2[1]);
  }
  months=Array.from(found).sort((a,b)=>b.localeCompare(a));
}

function initControls(){ const sel=document.querySelector('#monthSelect'); sel.innerHTML=''; for(const m of months){ const o=document.createElement('option'); o.value=m; o.textContent=`${m.slice(0,4)}年${m.slice(4,6)}月`; sel.appendChild(o);} if(!sel.value&&months.length>0) sel.value=months[0]; }

function getMetricValue(row, month, metric){ if(!row) return null; const part=COL_SUFFIX[metric]; const patterns=[`${month}單月合併營收${part}(%)`, `${month}單月合併營收${part}(％)`, `${month}單月合併營收${part}（%）`, `${month}單月合併營收${part}（％）`]; let v=null; for(const p of patterns){ if(row[p]!=null&&row[p]!=='' ){ v=row[p]; break; } } if(v==null) return null; v=Number(v); return Number.isFinite(v)?v:null; }

function handleRun(){
  const raw=document.querySelector('#stockInput').value; let codeKey=normCode(raw); const month=(document.querySelector('#monthSelect')?.value)||''; const metric=(document.querySelector('#metricSelect')?.value)||'MoM'; const colorMode=(document.querySelector('#colorMode')?.value)||'redPositive';
  if(!codeKey){ alert('請輸入股票代號或公司名稱'); return; }
  if(!byCode.has(codeKey)){ const hit=revenueRows.find(r=>z(r['名稱']).includes(z(raw).trim())); if(hit){ codeKey=normCode(hit['個股']??hit['代號']); } }
  if(!byCode.has(codeKey)){ alert('找不到此代號/名稱'); return; }
  const upstreamEdges=linksRows.filter(r=>normCode(r['下游代號'])===codeKey);
  const downstreamEdges=linksRows.filter(r=>normCode(r['上游代號'])===codeKey);
  const rowSelf=byCode.get(codeKey);
  renderResultChip(rowSelf, month, metric);
  renderTreemap('upTreemap','upHint', upstreamEdges,'上游代號',month,metric,colorMode);
  renderTreemap('downTreemap','downHint',downstreamEdges,'下游代號',month,metric,colorMode);
}

function renderResultChip(selfRow, month, metric){ const host=document.querySelector('#resultChip'); const v=getMetricValue(selfRow,month,metric); const cls=v==null?'':(v>=0?'good':'bad'); host.innerHTML=`<div class="result-card"><div class="row1"><strong>${safe(selfRow['個股'])}｜${safe(selfRow['名稱'])}</strong><span>${month.slice(0,4)}/${month.slice(4,6)} / ${metric}</span></div><div class="row2"><span>${safe(selfRow['產業別']||'')}</span><span class="${cls}">${displayPct(v)}</span></div></div>`; }

function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode){
  const svg=d3.select('#'+svgId); svg.selectAll('*').remove(); const wrap=svg.node().parentElement; const W=wrap.clientWidth-16; const H=parseInt(getComputedStyle(svg.node()).height)||560; svg.attr('width',W).attr('height',H);
  const groups=new Map();
  for(const e of edges){ const rel=normText(e['關係類型']||'未分類'); const key=normCode(e[codeField]); const r=byCode.get(key); if(!r) continue; const v=getMetricValue(r,month,metric); if(v==null) continue; if(!groups.has(rel)) groups.set(rel,[]); groups.get(rel).push({ code:r['個股'], name:r['名稱'], value:v }); }
  const hint=document.getElementById(hintId); if(groups.size===0){ hint.textContent='此區在選定月份沒有可用數據'; return; } else { hint.textContent=''; }
  const children=[]; for(const [rel,list] of groups){ const avg=d3.mean(list,d=>d.value); const kids=list.map(s=>({ name:`${s.code} ${s.name||''}`.trim(), code:s.code, value:Math.max(0.01,Math.abs(s.value)), raw:s.value })); children.push({ name:rel, avg, children:kids }); }
  const root=d3.hierarchy({ children }).sum(d=>d.value).sort((a,b)=>(b.value||0)-(a.value||0)); d3.treemap().size([W,H]).paddingOuter(8).paddingInner(3)(root);
  const g=svg.append('g');
  const parents=g.selectAll('g.parent').data(root.children||[]).enter().append('g').attr('class','parent');
  parents.append('rect').attr('class','group-border').attr('x',d=>d.x0).attr('y',d=>d.y0).attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0));
  parents.append('text').attr('class','node-title').attr('x',d=>d.x0+6).attr('y',d=>d.y0+16).text(d=>`${d.data.name}  平均：${displayPct(d.data.avg)}`);
  const node=g.selectAll('g.node').data(root.leaves()).enter().append('g').attr('class','node').attr('transform',d=>`translate(${d.x0},${d.y0})`);
  node.append('rect').attr('class','node-rect').attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0)).attr('fill', d=>colorFor(d.data.raw, colorMode));
  node.append('text').attr('class','node-sub').attr('x',6).attr('y',16).text(d=>`${(d.data.code||'').toString().slice(0,6)} ${displayPct(d.data.raw)}`).each(function(d){ const w=d.x1-d.x0; if(this.getBBox().width>w-8){ d3.select(this).attr('opacity',0.85).attr('font-size',10);} });
  node.append('text').attr('class','node-sub').attr('x',6).attr('y',30).text(d=>`${safe((d.data.name||'').slice(0,8))}`).each(function(d){ const w=d.x1-d.x0; if(this.getBBox().width>w-8){ d3.select(this).attr('opacity',0.8).attr('font-size',10);} });
}
