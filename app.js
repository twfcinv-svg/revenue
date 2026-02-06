/* app.js — v3.9
 * 修正：
 *  1) 類股標題確實「跟著格子縮小」：更強的 GroupTitleFit（單行→兩行自動切換、逐字省略、連續縮放）；
 *  2) 標題永不外溢：header clipPath（內縮 3px）+ 重新量測循環；
 *  3) 平均值永不消失：無論單行或兩行，皆保證呈現；
 *  4) 視窗縮放/容器變動時自動重算標題字級；
 *  5) 其他行為沿用 v3.8（群組面積=平均數值，個股標籤由豐到簡、最小 4px、clip 內縮等）。
 */

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

  const wsRev = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];
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
    const code = normCode(r[codeKeyName]);
    const name = normText(r[nameKeyName]);
    if(code) byCode.set(code, r);
    if(name) byName.set(name, r);
  }

  linksByUp.clear(); linksByDown.clear();
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
    const o=document.createElement('option');
    o.value=m; o.textContent=`${m.slice(0,4)}年${m.slice(4,6)}月`;
    sel.appendChild(o);
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

  try{
    const codeLabel = (rowSelf['個股'] || rowSelf['代號'] || rowSelf['股票代碼'] || rowSelf['股票代號'] || rowSelf['公司代號'] || rowSelf['證券代號'] || '').trim();
    const nameLabel = (rowSelf['名稱'] || rowSelf['公司名稱'] || rowSelf['證券名稱'] || '').trim();
    const extra = `${month.slice(0,4)}/${month.slice(4,6)} · ${metric}`;
    if (window.setResultChipLink) window.setResultChipLink(codeLabel, nameLabel, extra);
  }catch(_){ }

  const upstreamEdges   = linksByDown.get(codeKey) || [];
  const downstreamEdges = linksByUp.get(codeKey)   || [];

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

// ========= 葉節點（延用 v3.8/v3.7） =========
const LabelFit = {
  paddingBase: 8,
  maxFont: 36,
  minFontSoft: 9,
  minFontHard: 4,
  lineHeight: 1.15,
  dynPadding(w,h){ const m=Math.min(w,h); return Math.max(2, Math.min(this.paddingBase, Math.floor(m*0.08))); },
  centerText(el,w,h,p){ el.setAttribute('text-anchor','middle'); el.setAttribute('dominant-baseline','middle'); el.setAttribute('x', p + Math.max(0,(w-p*2)/2)); el.setAttribute('y', p + Math.max(0,(h-p*2)/2)); },
  ensureClip(gEl,w,h){ const inset=2; const svg=gEl.ownerSVGElement; let defs=svg.querySelector('defs'); if(!defs) defs=svg.insertBefore(document.createElementNS('http://www.w3.org/2000/svg','defs'), svg.firstChild); const id=gEl.dataset.clipId||('clip-'+Math.random().toString(36).slice(2)); gEl.dataset.clipId=id; let clip=svg.querySelector('#'+id); if(!clip){ clip=document.createElementNS('http://www.w3.org/2000/svg','clipPath'); clip.setAttribute('id',id); const r=document.createElementNS('http://www.w3.org/2000/svg','rect'); clip.appendChild(r); defs.appendChild(clip);} const rect=clip.firstChild; rect.setAttribute('x',inset); rect.setAttribute('y',inset); rect.setAttribute('width',Math.max(0,w-inset*2)); rect.setAttribute('height',Math.max(0,h-inset*2)); gEl.querySelectorAll('text').forEach(t=>t.setAttribute('clip-path',`url(#${id})`)); },
  ellipsizeNameToWidth(textEl,maxW){ const t1=textEl.querySelector('tspan'); if(!t1) return; const full=t1.textContent||''; const m=full.match(/^(\d{4})\s*(.*)$/); let code='', name=full; if(m){ code=m[1]; name=m[2]||''; } t1.textContent=code+(name?(' '+name):''); while(t1.getComputedTextLength()>maxW && name.length>0){ name=name.slice(0,-1); t1.textContent=code+(name?(' '+name+'…'):''); } },
  fitBlock(textEl,w,h){ const p=this.dynPadding(w,h); const targetW=Math.max(1,w-p*2), targetH=Math.max(1,h-p*2); const code=textEl.dataset.code||''; const name=textEl.dataset.name||''; const pct=textEl.dataset.pct||''; const layouts=[ ()=>[`${code}${name?(' '+name):''}`, pct], ()=>[code, pct], ()=>[pct] ]; const k=0.12; const areaFont=Math.sqrt(targetW*targetH)*k; const logicalMax=Math.min(this.maxFont, Math.floor(targetH*0.5)); for(const L of layouts){ while(textEl.firstChild) textEl.removeChild(textEl.firstChild); L().forEach(s=>{ const t=document.createElementNS('http://www.w3.org/2000/svg','tspan'); t.textContent=s; textEl.appendChild(t); }); let f=Math.max(this.minFontHard, Math.min(logicalMax, Math.floor(areaFont))); textEl.setAttribute('font-size',f); this.centerText(textEl,w,h,p); this.ellipsizeNameToWidth(textEl, targetW); let guard=0; while(guard++<60){ const bb=textEl.getBBox(); const sW=targetW/Math.max(1,bb.width), sH=targetH/Math.max(1,bb.height); const s=Math.min(sW,sH,1); const next=Math.max(this.minFontHard, Math.floor(f*s)); if(next<f){ f=next; textEl.setAttribute('font-size',f); this.centerText(textEl,w,h,p); continue; } if(sW<1 && f<=this.minFontHard){ this.ellipsizeNameToWidth(textEl, targetW); } break; } const tsp=textEl.querySelectorAll('tspan'); const n=Math.max(1,tsp.length); const offsetEm=-((n-1)*this.lineHeight/2); tsp.forEach((t,i)=>{ t.setAttribute('x', textEl.getAttribute('x')); t.setAttribute('dy', i===0?`${offsetEm}em`:`${this.lineHeight}em`); }); const box=textEl.getBBox(); if(box.width<=targetW+0.1 && box.height<=targetH+0.1){ textEl.removeAttribute('display'); return; } } textEl.setAttribute('display','none'); }
};

// ========= 類股標題自動縮放（更強版） =========
const GroupTitleFit = {
  minFont: 5,
  lineHeight: 1.12,
  inset: 3,
  k: 0.12,        // 面積驅動初始字級係數
  // 只給 header 區域 clip，避免溢出
  ensureHeaderClip(svg, gEl, d, headerH){
    const id = gEl.dataset.headerClipId || ('hclip-'+Math.random().toString(36).slice(2));
    gEl.dataset.headerClipId = id;
    let defs = svg.querySelector('defs');
    if(!defs) defs = svg.insertBefore(document.createElementNS('http://www.w3.org/2000/svg','defs'), svg.firstChild);
    let clip = svg.querySelector('#'+id);
    if(!clip){
      clip = document.createElementNS('http://www.w3.org/2000/svg','clipPath');
      clip.setAttribute('id', id);
      const r = document.createElementNS('http://www.w3.org/2000/svg','rect');
      clip.appendChild(r);
      defs.appendChild(clip);
    }
    const r = clip.firstChild;
    const w = Math.max(0, d.x1-d.x0), h = Math.max(0, headerH);
    r.setAttribute('x', d.x0 + this.inset);
    r.setAttribute('y', d.y0 + this.inset);
    r.setAttribute('width',  Math.max(0, w - this.inset*2));
    r.setAttribute('height', Math.max(0, h - this.inset*2));
    return `url(#${id})`;
  },

  // 單行布局：『名稱  平均：xx%』
  mountOneLine(text, d, headerH){
    while(text.firstChild) text.removeChild(text.firstChild);
    const tName = document.createElementNS('http://www.w3.org/2000/svg','tspan'); tName.textContent = d.data.name || ''; text.appendChild(tName);
    const tSep  = document.createElementNS('http://www.w3.org/2000/svg','tspan'); tSep.textContent = '  '; text.appendChild(tSep);
    const tAvg  = document.createElementNS('http://www.w3.org/2000/svg','tspan'); tAvg.textContent = `平均：${displayPct(d.data.avg)}`; text.appendChild(tAvg);
    text.dataset.mode = 'one';
  },

  // 兩行布局：第一行 名稱（可省略）/ 第二行 平均（必保留）
  mountTwoLines(text, d, headerH){
    while(text.firstChild) text.removeChild(text.firstChild);
    const tName = document.createElementNS('http://www.w3.org/2000/svg','tspan'); tName.textContent = d.data.name || ''; text.appendChild(tName);
    const tAvg  = document.createElementNS('http://www.w3.org/2000/svg','tspan'); tAvg.textContent = `平均：${displayPct(d.data.avg)}`; text.appendChild(tAvg);
    text.dataset.mode = 'two';
  },

  // 名稱逐字省略（單行或兩行的第一個 tspan）
  ellipsizeName(text, maxW){
    const tName = text.querySelector('tspan');
    if(!tName) return false;
    let nm = tName.textContent || '';
    if (nm.length===0) return false;
    tName.textContent = nm.slice(0, -1) + '…';
    return true;
  },

  fit(text, d, headerH){
    const svg = text.ownerSVGElement;
    const wMax = Math.max(0, d.x1-d.x0) - this.inset*2;
    const hMax = Math.max(0, headerH)  - this.inset*2;
    if (wMax<=0 || hMax<=0) return;

    // 位置與 clip
    text.setAttribute('text-anchor','start');
    text.setAttribute('dominant-baseline','middle');
    text.setAttribute('x', d.x0 + this.inset + 4);
    text.setAttribute('y', d.y0 + headerH/2);
    text.setAttribute('clip-path', this.ensureHeaderClip(svg, text.parentNode, d, headerH));

    // 先用單行嘗試
    this.mountOneLine(text, d, headerH);

    // 初始字級：面積驅動 + 高度限制
    const f0 = Math.floor(Math.min(Math.sqrt(Math.max(1,wMax*hMax))*this.k, hMax*0.95));
    let f   = Math.max(this.minFont, f0);

    let guard = 0;
    const fitLoop = ()=>{
      if (++guard>120) return; // 保護
      text.setAttribute('font-size', f);
      // 單行對齊（兩行時會在後面改 dy）
      const mode = text.dataset.mode || 'one';
      if (mode==='one'){
        // 直接量測
        const bb = text.getBBox();
        const sW = wMax / Math.max(1, bb.width);
        const sH = hMax / Math.max(1, bb.height);
        const scale = Math.min(sW, sH, 1);
        const next  = Math.max(this.minFont, Math.floor(f * scale));
        if (next < f){ f = next; return fitLoop(); }
        if (sW < 1 && f <= this.minFont){
          // 仍超寬 → 省略名稱；若名稱已空仍超寬且高度足夠 → 切兩行
          if (!this.ellipsizeName(text, wMax)){
            // 切換兩行（若高度足夠）
            if (hMax >= this.minFont*2*this.lineHeight + 2){
              this.mountTwoLines(text, d, headerH);
              return fitLoop();
            }
          }
          return fitLoop();
        }
        // 成功：單行 fits
        return;
      } else {
        // 兩行：第一行名稱、第二行平均
        const tsp = text.querySelectorAll('tspan');
        if (tsp.length<2) return;
        // 調整行距：第一行上提，第二行往下
        tsp[0].setAttribute('dy', `${-this.lineHeight/2}em`);
        tsp[1].setAttribute('dy', `${this.lineHeight}em`);
        tsp.forEach(t=>t.setAttribute('x', text.getAttribute('x')));

        const bb = text.getBBox();
        const sW = wMax / Math.max(1, bb.width);
        const sH = hMax / Math.max(1, bb.height);
        const scale = Math.min(sW, sH, 1);
        const next  = Math.max(this.minFont, Math.floor(f * scale));
        if (next < f){ f = next; return fitLoop(); }
        if (sW < 1 && f <= this.minFont){
          // 再省略名稱直到 fits
          if (this.ellipsizeName(text, wMax)) return fitLoop();
        }
        return; // 成功
      }
    };

    fitLoop();
  }
};

function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode){
  const svg=d3.select('#'+svgId); svg.selectAll('*').remove();
  const wrap=svg.node().parentElement; const W=wrap.clientWidth-16; const H=parseInt(getComputedStyle(svg.node()).height)||560;
  svg.attr('width',W).attr('height',H);

  // ====== 分群 ======
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
    groups.get(rel).push({ code:codeVal, name:nameVal, raw:v });
  }

  // 只留前 8 類（依個股數量）
  const entries = Array.from(groups.entries()).sort((a,b)=> b[1].length - a[1].length).slice(0,8);
  const kept = new Map(entries);

  const hint=document.getElementById(hintId);
  if(kept.size===0){ hint.textContent='此區在選定月份沒有可用數據'; return; } else { hint.textContent=''; }

  // ====== 葉節點 base 值（簽名平移）與 類股平均權重（簽名平移） ======
  const EPS = 0.01;
  const groupSummaries = [];
  for (const [rel, list] of kept){
    const avg = d3.mean(list, d=>d.raw);
    const minLeafRaw = d3.min(list, d=>d.raw);
    const baseValues = list.map(s => ({ s, base: Math.max(EPS, (s.raw - minLeafRaw + EPS)) }));
    const baseSum = d3.sum(baseValues, d=>d.base) || EPS;
    groupSummaries.push({ rel, list, avg, baseValues, baseSum });
  }
  const minAvg = d3.min(groupSummaries, d=>d.avg);

  const children=[];
  for (const g of groupSummaries){
    const groupWeight = Math.max(EPS, (g.avg - minAvg + EPS));
    const scale = groupWeight / (g.baseSum || EPS);
    const kids = g.baseValues.map(({s, base})=>({ name:s.name||'', code:s.code, raw:s.raw, value: base*scale }));
    children.push({ name:g.rel, avg:g.avg, children:kids });
  }

  const HEADER_H = 22;
  const root=d3.hierarchy({ children }).sum(d=>d.value).sort((a,b)=>(b.value||0)-(a.value||0));
  d3.treemap().size([W,H]).paddingOuter(8).paddingInner(3).paddingTop(HEADER_H)(root);

  const g=svg.append('g');

  // —— 類股群組底與框 ——
  const parents=g.selectAll('g.parent').data(root.children||[]).enter().append('g').attr('class','parent');
  parents.append('rect').attr('class','group-bg')
    .attr('x',d=>d.x0).attr('y',d=>d.y0)
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=> colorFor(d.data.avg, colorMode));
  parents.append('rect').attr('class','group-border')
    .attr('x',d=>d.x0).attr('y',d=>d.y0)
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0));

  // —— 類股標題 ——
  const titles = parents.append('text')
    .attr('class','node-title')
    .attr('fill','#fff')
    .style('paint-order','stroke')
    .style('stroke','rgba(0,0,0,0.35)')
    .style('stroke-width','2px');

  titles.each(function(d){ GroupTitleFit.fit(this, d, HEADER_H); });

  // —— 葉節點（個股） ——
  const node=g.selectAll('g.node').data(root.leaves()).enter().append('g').attr('class','node').attr('transform',d=>`translate(${d.x0},${d.y0})`);
  node.append('rect').attr('class','node-rect')
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=> colorFor(d.data.raw, colorMode));

  const labels = node.append('text')
    .attr('class','node-label')
    .attr('fill','#fff')
    .style('paint-order','stroke')
    .style('stroke','rgba(0,0,0,0.35)')
    .style('stroke-width','2px')
    .style('text-rendering','geometricPrecision');

  labels.each(function(d){
    const code = `${d.data.code||''}`.trim();
    const name = `${d.data.name||''}`.trim();
    const pct  = displayPct(d.data.raw);
    this.dataset.code = code; this.dataset.name = name; this.dataset.pct = pct;
    const t1 = document.createElementNS('http://www.w3.org/2000/svg','tspan'); t1.textContent = `${code}${name?(' '+name):''}`;
    const t2 = document.createElementNS('http://www.w3.org/2000/svg','tspan'); t2.textContent = pct;
    this.appendChild(t1); this.appendChild(t2);
    const title = document.createElementNS('http://www.w3.org/2000/svg','title'); title.textContent = `${code} ${name} ${pct}`; this.appendChild(title);
  });

  requestAnimationFrame(()=>{
    // 個股文字 fit + 裁切
    node.each(function(d){
      const w = Math.max(0, d.x1 - d.x0); const h = Math.max(0, d.y1 - d.y0);
      const textEl = this.querySelector('text'); if (!textEl) return;
      LabelFit.fitBlock(textEl, w, h); LabelFit.ensureClip(this, w, h);
    });
    // 標題在 draw 後再多跑一次 fit（避免邊界條件）
    parents.select('text').each(function(d){ GroupTitleFit.fit(this, d, HEADER_H); });
  });

  // 視窗改變時重算標題（避免左下角等位置在容器變動時被擋）
  const onResize = ()=>{
    parents.select('text').each(function(d){ GroupTitleFit.fit(this, d, HEADER_H); });
  };
  window.addEventListener('resize', onResize, { passive:true });
}
