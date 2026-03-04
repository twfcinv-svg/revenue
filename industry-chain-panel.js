
/* industry-chain-panel.js (foldable v5)
 * 強化：左側清單「同頂同底」絕對對齊右側 chart-wrap（含邊框）
 * 作法：
 *   1) 以 #icp-fold-wrap 為坐標系，計算 chart-wrap 的 top/bottom → 對應成 scroll.marginTop 與 wrap 的 targetHeight
 *   2) 折疊：scroll.maxHeight = targetHeight - marginTop；展開：scroll.maxHeight = scroll.scrollHeight
 *   3) 「＋/－」無圓圈粗體，垂直中心對齊 chart-wrap 的底緣
 */
(function(){
  const $ = (s)=> document.querySelector(s);
  const encode = (s)=> encodeURIComponent(s);

  function norm(s){ return String(s||'').trim(); }
  function detectCol(headers, patterns){ for(const p of patterns){ for(const h of headers){ if(p.test(h)) return h; } } return null; }
  function groupBy(arr, keyFn){ const m=new Map(); for(const x of arr){ const k=keyFn(x)||''; if(!m.has(k)) m.set(k,[]); m.get(k).push(x);} return m; }

  const cache = { loaded:false };
  async function loadAll(){
    if(cache.loaded) return cache;
    const res = await fetch('data.xlsx');
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array' });

    const revName = wb.SheetNames.find(n=>/revenue|營收/i.test(n)) || wb.SheetNames[0];
    const rev = XLSX.utils.sheet_to_json(wb.Sheets[revName], { defval:null });
    if(!rev.length) throw new Error('Revenue 工作表為空');
    const rh = Object.keys(rev[0]);
    const colCode = detectCol(rh, [/個股|代號|股票代號|code|symbol/i]);
    const colName = detectCol(rh, [/名稱|公司|name/i]);
    const colInd  = detectCol(rh, [/產業別|產業|industry/i]);

    const revRows = rev.map(r=>({ code:norm(r[colCode]), name:norm(r[colName]), industry:norm(r[colInd]) }))
                       .filter(r=>r.code || r.name);

    const linkName = wb.SheetNames.find(n=>/link|關聯|關係|供應鏈/i.test(n));
    let linksRows = [];
    if(linkName){
      const links = XLSX.utils.sheet_to_json(wb.Sheets[linkName], { defval:null });
      if(links.length){
        const lh = Object.keys(links[0]);
        const colUp = detectCol(lh, [/上游代號|上游|up(stream)?_?code/i]);
        const colDn = detectCol(lh, [/下游代號|下游|down(stream)?_?code/i]);
        const colTyp= detectCol(lh, [/關係類型|類型|type/i]);
        linksRows = links.map(r=>({ up:norm(r[colUp]), down:norm(r[colDn]), type:norm(r[colTyp]) }))
                         .filter(r=>r.up||r.down);
      }
    }

    const byCode = new Map(); for(const r of revRows){ if(r.code) byCode.set(r.code,r); }
    const upstreamOf = new Map(), downstreamOf = new Map();
    for(const l of linksRows){ if(l.up && l.down){
      if(!downstreamOf.has(l.up)) downstreamOf.set(l.up, []);
      downstreamOf.get(l.up).push(l.down);
      if(!upstreamOf.has(l.down)) upstreamOf.set(l.down, []);
      upstreamOf.get(l.down).push(l.up);
    }}

    Object.assign(cache,{loaded:true,revRows,byCode,upstreamOf,downstreamOf});
    return cache;
  }

  function stockLiHtml(code, byCode){ const r=byCode.get(code)||{code,name:''}; return `<li class="icp-stock" data-code="${r.code}"><span class="code">${r.code}</span> ${r.name||''}</li>`; }
  function renderGroupList(el, codes, byCode){
    if(!el) return;
    const items = Array.from(new Set(codes)).map(c=>byCode.get(c)||{code:c,name:'',industry:''});
    const m = groupBy(items, it=>it.industry);
    const parts=[]; for(const [ind, arr] of m.entries()){
      const list = arr.sort((a,b)=>a.code.localeCompare(b.code)).map(r=>stockLiHtml(r.code,byCode)).join('');
      parts.push(`<div class="icp-card"><div class="icp-card-title">${ind||'未分類'}</div><ul>${list}</ul></div>`);
    }
    el.innerHTML = parts.join('');
  }

  function findStock(byCode, rows, kw){ const k=norm(kw); return k? (byCode.get(k) || rows.find(r=>r.name===k) || null) : null; }

  function setFoldToChartHeight(){
    const wrap = document.getElementById('icp-fold-wrap');
    const scroll = document.getElementById('icp-scroll');
    const btn = document.getElementById('icp-expander');
    const fade = document.getElementById('icp-fade');
    const chart = document.querySelector('#combo-section .chart-wrap');
    if(!wrap || !scroll || !btn || !chart) return;

    const wrapRect  = wrap.getBoundingClientRect();
    const chartRect = chart.getBoundingClientRect();
    const isExpanded = btn.getAttribute('aria-expanded') === 'true';

    // 目標可視高度 = chart 底 - chart 頂（含邊框）
    const targetTop    = Math.round(chartRect.top    - wrapRect.top);
    const targetBottom = Math.round(chartRect.bottom - wrapRect.top);
    const targetHeight = Math.max(0, targetBottom - targetTop);

    // 同頂：把 scroll 往下推 targetTop；同底：折疊時限制高度為 targetHeight
    scroll.style.marginTop = (targetTop < 0 ? 0 : targetTop) + 'px';
    if(!isExpanded){
      scroll.style.maxHeight = targetHeight + 'px';
      if(fade) fade.style.display = 'block';
    }else{
      scroll.style.maxHeight = scroll.scrollHeight + 'px';
      if(fade) fade.style.display = 'none';
    }

    // 「＋ / －」加號：對齊 chart 底的垂直中心
    const top = Math.round(targetBottom - (btn.offsetHeight/2));
    btn.style.top = (top < 0 ? 0 : top) + 'px';
  }

  function toggleFold(){
    const btn = document.getElementById('icp-expander');
    const expanded = btn.getAttribute('aria-expanded') === 'true';
    btn.setAttribute('aria-expanded', String(!expanded));
    btn.textContent = expanded ? '＋' : '－';
    setFoldToChartHeight();
  }

  function installObservers(){
    const chart = document.querySelector('#combo-section .chart-wrap');
    if(chart){ const ro = new ResizeObserver(()=> setFoldToChartHeight()); ro.observe(chart); }
    window.addEventListener('resize', setFoldToChartHeight);
  }

  async function updatePanel(){
    const { revRows, byCode, upstreamOf, downstreamOf } = await loadAll();
    const me = findStock(byCode, revRows, document.getElementById('stockInput')?.value || '');
    const upWrap = document.getElementById('icp-up-wrap');
    const downWrap = document.getElementById('icp-down-wrap');

    if(!me){ if(upWrap) upWrap.innerHTML=''; if(downWrap) downWrap.innerHTML=''; setFoldToChartHeight(); return; }

    renderGroupList(upWrap, upstreamOf.get(me.code) || [], byCode);
    renderGroupList(downWrap, downstreamOf.get(me.code) || [], byCode);

    // 點個股開外網
    function bindClick(root){ if(!root) return; root.addEventListener('click', (e)=>{
      const li = e.target.closest('.icp-stock'); if(!li) return; const code = li.getAttribute('data-code'); if(!code) return;
      const url = 'https://www.fbs.com.tw/MKT/Index?name=' + encode('Ｊ線圖') + '&stock=' + encode(code);
      window.open(url, '_blank');
    }); }
    bindClick(upWrap); bindClick(downWrap);

    // 完成後進行齊頂齊底計算
    setTimeout(setFoldToChartHeight, 0);
  }

  document.addEventListener('DOMContentLoaded', async ()=>{
    try{ await loadAll(); }catch(e){ console.error(e); }
    const btn = document.getElementById('icp-expander'); if(btn){ btn.addEventListener('click', toggleFold); }
    const run = document.getElementById('runBtn'); if(run){ run.addEventListener('click', ()=> setTimeout(updatePanel, 0)); }
    setTimeout(updatePanel, 0);
    installObservers();
  });
})();
