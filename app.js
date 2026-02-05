/* app.js — v3.5
 * 這版回應：
 *  1) 有空間卻不顯示中文＋代號：先嘗試完整(代號+名稱 / 百分比)；若太長，會「逐字省略」名稱(…)，盡量保留代號與名稱。
 *  2) 數值或資訊被切到：字級允許持續縮小到硬下限(5px)；同時反覆量測直到完全塞入，不再裁切。
 *  3) 維持 v3.4：群組標題靠左垂直中線、clipPath 內縮 1px、防溢出；上/下游僅保留前 8 類(依個股數量)。
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

// ========= 字級與裁切工具 =========
const LabelFit = {
  paddingBase: 8,
  maxFont: 36,      // 全域上限（保護特大格）
  minFontSoft: 9,   // 盡量維持的可讀下限
  minFontHard: 5,   // 真的塞不下時的硬下限（允許更小字以避免裁切）
  lineHeight: 1.15,

  dynPadding(w,h){
    // 小格子自動縮小 padding，釋出更多空間
    const m = Math.min(w,h);
    return Math.max(2, Math.min(this.paddingBase, Math.floor(m * 0.12)));
  },

  centerText(el, w, h, pad) {
    el.setAttribute('text-anchor', 'middle');
    el.setAttribute('dominant-baseline', 'middle');
    el.setAttribute('x', pad + Math.max(0, (w - pad*2) / 2));
    el.setAttribute('y', pad + Math.max(0, (h - pad*2) / 2));
  },

  ensureClip(gEl, w, h){
    const inset = 1; // 內縮 1px，避免邊緣吃字
    const svg = gEl.ownerSVGElement; let defs = svg.querySelector('defs');
    if (!defs) defs = svg.insertBefore(document.createElementNS('http://www.w3.org/2000/svg','defs'), svg.firstChild);
    const id = gEl.dataset.clipId || ('clip-' + Math.random().toString(36).slice(2)); gEl.dataset.clipId = id;
    let clip = svg.querySelector('#'+id);
    if (!clip){ clip = document.createElementNS('http://www.w3.org/2000/svg','clipPath'); clip.setAttribute('id', id);
      const r = document.createElementNS('http://www.w3.org/2000/svg','rect'); clip.appendChild(r); defs.appendChild(clip); }
    const rect = clip.firstChild; rect.setAttribute('x', inset); rect.setAttribute('y', inset);
    rect.setAttribute('width', Math.max(0, w - inset*2)); rect.setAttribute('height', Math.max(0, h - inset*2));
    gEl.querySelectorAll('text').forEach(t => t.setAttribute('clip-path', `url(#${id})`));
  },

  // 將 tspan[0] 的名稱做「逐字省略」，直到寬度塞入
  ellipsizeNameToWidth(textEl, maxWidth){
    const tspans = Array.from(textEl.querySelectorAll('tspan'));
    if (!tspans.length) return;
    const t1 = tspans[0];
    const full = t1.textContent;
    let s = full;
    if (!s) return;
    // 只省略「中文名稱」部分；盡量保留代號
    const m = s.match(/^(\d{4})\s*(.*)$/);
    let code = '', name = s;
    if (m){ code = m[1]; name = m[2] || ''; }
    // 先把 t1 設成原樣（每次試一個字）
    t1.textContent = code + (name ? (' ' + name) : '');
    while (t1.getComputedTextLength() > maxWidth && name.length > 0){
      name = name.slice(0, -1);
      t1.textContent = code + (name ? (' ' + name + '…') : '');
    }
  },

  fitBlock(textEl, w, h){
    const pad = this.dynPadding(w,h);
    const targetW = Math.max(1, w - pad*2);
    const targetH = Math.max(1, h - pad*2);

    // 取原始資訊（在建立 text 時已寫入 dataset）
    const code = textEl.dataset.code || '';
    const name = textEl.dataset.name || '';
    const pct  = textEl.dataset.pct  || '';

    // 嘗試不同版型：full -> code+pct -> pct-only
    const tryLayouts = [
      () => [ `${code}${name?(' '+name):''}`, pct ],
      () => [ code, pct ],
      () => [ pct ]
    ];

    const k = 0.12; // 面積驅動係數
    const areaFont = Math.sqrt(targetW * targetH) * k;
    const logicalMax = Math.min(this.maxFont, Math.floor(targetH * 0.5));

    // 嘗試三種版型；每種版型內部採「先大後小、反覆量測直到塞入」
    for (let layout of tryLayouts){
      // 填入候選版型
      while (textEl.firstChild) textEl.removeChild(textEl.firstChild);
      const lines = layout();
      lines.forEach((s,i)=>{
        const t = document.createElementNS('http://www.w3.org/2000/svg','tspan');
        t.textContent = s; textEl.appendChild(t);
      });

      // 初始字級：取 areaFont，夾在 [minHard, logicalMax]
      let font = Math.max(this.minFontHard, Math.min(logicalMax, Math.floor(areaFont)));
      textEl.setAttribute('font-size', font);
      this.centerText(textEl, w, h, pad);

      // 先對第一行做逐字省略（避免長公司名撐寬）
      this.ellipsizeNameToWidth(textEl, targetW);

      // 量測 → 迭代縮小直到完全塞入（或到 minFontHard 為止）
      let safety = 0;
      while (safety++ < 50){
        const bbox = textEl.getBBox();
        const scaleW = targetW / Math.max(1,bbox.width);
        const scaleH = targetH / Math.max(1,bbox.height);
        const scale = Math.min(scaleW, scaleH, 1);
        const next = Math.floor(font * scale);
        if (next >= this.minFontHard && next < font){
          font = next; textEl.setAttribute('font-size', font); this.centerText(textEl, w, h, pad); continue;
        }
        // 若還是超寬（scaleW<1）而字已到硬下限，再次針對第一行做省略
        if (scaleW < 1 && font <= this.minFontHard){
          this.ellipsizeNameToWidth(textEl, targetW);
        }
        break;
      }

      // 調整多行垂直置中
      const tspans = Array.from(textEl.querySelectorAll('tspan'));
      const n = Math.max(1, tspans.length);
      const offsetEm = -((n - 1) * this.lineHeight / 2);
      tspans.forEach((tsp,i)=>{
        tsp.setAttribute('x', textEl.getAttribute('x'));
        tsp.setAttribute('dy', i===0 ? `${offsetEm}em` : `${this.lineHeight}em`);
      });

      // 再次檢查：若現在 bbox 完全落在 target 內，就採用此版型
      const box = textEl.getBBox();
      if (box.width <= targetW + 0.1 && box.height <= targetH + 0.1){
        // 若字級 >= minFontSoft，或這是最後一個版型，就用它
        if (font >= this.minFontSoft || layout === tryLayouts[tryLayouts.length-1]){
          textEl.removeAttribute('display');
          return;
        }
        // 否則嘗試更緊湊的下一種版型
      }
    }

    // 三種版型都無法塞入 → 隱藏
    textEl.setAttribute('display','none');
  }
};

function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode){
  const svg=d3.select('#'+svgId); svg.selectAll('*').remove();
  const wrap=svg.node().parentElement; const W=wrap.clientWidth-16; const H=parseInt(getComputedStyle(svg.node()).height)||560;
  svg.attr('width',W).attr('height',H);

  // ====== 分群 ======
  const groups=new Map();
  for(const e of edges){
    const rel=normText(e['關係類型']||'未分類'); const key=normCode(e[codeField]); const r=byCode.get(key);
    if(!r) continue; const v=getMetricValue(r,month,metric); if(v==null) continue;
    if(!groups.has(rel)) groups.set(rel,[]);
    const codeVal = r['個股'] ?? r['代號'] ?? r['股票代碼'] ?? r['股票代號'] ?? r['公司代號'] ?? r['證券代號'];
    const nameVal = r['名稱'] ?? r['公司名稱'] ?? r['證券名稱'];
    groups.get(rel).push({ code:codeVal, name:nameVal, value:Math.max(0.01,Math.abs(v)), raw:v });
  }

  // ====== 僅保留前 8 大類（依個股數量多寡） ======
  const entries = Array.from(groups.entries());
  entries.sort((a,b)=> b[1].length - a[1].length);
  const kept = new Map(entries.slice(0,8));

  const hint=document.getElementById(hintId);
  if(kept.size===0){ hint.textContent='此區在選定月份沒有可用數據'; return; } else { hint.textContent=''; }

  const children=[]; for(const [rel,list] of kept){ const avg=d3.mean(list,d=>d.raw);
    const kids=list.map(s=>({ name: s.name||'', code:s.code, value:Math.max(0.01,Math.abs(s.raw)), raw:s.raw }));
    children.push({ name:rel, avg, children:kids }); }

  const HEADER_H = 22; // 與 treemap().paddingTop(HEADER_H) 一致
  const root=d3.hierarchy({ children }).sum(d=>d.value).sort((a,b)=>(b.value||0)-(a.value||0));
  d3.treemap().size([W,H]).paddingOuter(8).paddingInner(3).paddingTop(HEADER_H)(root);

  const g=svg.append('g');

  // —— 類股群組 ——
  const parents=g.selectAll('g.parent').data(root.children||[]).enter().append('g').attr('class','parent');
  parents.append('rect').attr('class','group-bg')
    .attr('x',d=>d.x0).attr('y',d=>d.y0)
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0))
    .attr('fill', d=> colorFor(d.data.avg, colorMode));
  parents.append('rect').attr('class','group-border')
    .attr('x',d=>d.x0).attr('y',d=>d.y0)
    .attr('width',d=>Math.max(0,d.x1-d.x0)).attr('height',d=>Math.max(0,d.y1-d.y0));

  // 類股標題：靠左 + 在 HEADER_H 區塊的「垂直置中」；字級依群組面積而變
  const titles = parents.append('text')
    .attr('class','node-title')
    .attr('text-anchor','start')
    .attr('dominant-baseline','middle')
    .style('paint-order','stroke')
    .style('stroke','rgba(0,0,0,0.35)')
    .style('stroke-width','2px')
    .attr('fill','#fff');

  titles.each(function(d){
    const w = Math.max(0, d.x1 - d.x0), h = Math.max(0, d.y1 - d.y0);
    const area = w * h; const kGroup = 0.085; // 可微調：0.08~0.10
    const fs = Math.max(11, Math.min(22, Math.floor(Math.sqrt(area) * kGroup)));
    const el = d3.select(this)
      .attr('x', d.x0 + 6)
      .attr('y', d.y0 + HEADER_H/2)
      .attr('font-size', fs);

    this.textContent = '';
    const t1 = document.createElementNS('http://www.w3.org/2000/svg','tspan'); t1.textContent = d.data.name; this.appendChild(t1);
    if (fs >= 13) {
      const t2 = document.createElementNS('http://www.w3.org/2000/svg','tspan'); t2.textContent = `  平均：${displayPct(d.data.avg)}`; t2.setAttribute('dx','6'); this.appendChild(t2);
    }
  });

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

  // 批次縮放與裁切
  requestAnimationFrame(()=>{
    node.each(function(d){
      const w = Math.max(0, d.x1 - d.x0); const h = Math.max(0, d.y1 - d.y0);
      const textEl = this.querySelector('text'); if (!textEl) return;
      LabelFit.fitBlock(textEl, w, h); LabelFit.ensureClip(this, w, h);
    });
  });
}
