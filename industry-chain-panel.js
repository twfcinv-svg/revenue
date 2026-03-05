
/* industry-chain-panel.js (v8)
 * 1) 對齊改善（最小變動）：
 *    - Top anchor：#combo-section .section-head 底線
 *    - Bottom anchor：max(#combo-section .chart-wrap 底線, .legend-row 底線)
 *    - 左側 #icp-scroll 依 anchor 設定 margin-top / max-height；＋/－ 置於 bottom anchor 中線
 * 2) 產業群組來源改為 Links.關係類型（而非 Revenue.產業別）
 *    - 上游：所有 Links 中 down==當前個股 的記錄 → group=關係類型
 *    - 下游：所有 Links 中 up==當前個股   的記錄 → group=關係類型
 *    - 若關係類型缺失 → "未分類"
 */
(function(){
  const $ = (s)=> document.querySelector(s);
  const $$ = (s)=> Array.from(document.querySelectorAll(s));
  const encode = (s)=> encodeURIComponent(s);

  // ------------------ Helpers ------------------
  function norm(s){ return String(s==null? '': s).trim(); }
  function detectCol(headers, patterns){
    for(const p of patterns){ for(const h of headers){ if(p.test(h)) return h; } }
    return null;
  }
  function groupBy(arr, keyFn){ const m=new Map(); for(const x of arr){ const k=keyFn(x)||''; if(!m.has(k)) m.set(k,[]); m.get(k).push(x); } return m; }

  // ------------------ Load Excel once ------------------
  const cache = { loaded:false };
  async function loadAll(){
    if(cache.loaded) return cache;
    const res = await fetch('data.xlsx');
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array' });

    // Revenue（用於查公司名稱；不再用其「產業別」分群）
    const revName = wb.SheetNames.find(n=>/revenue|營收/i.test(n)) || wb.SheetNames[0];
    const rev = XLSX.utils.sheet_to_json(wb.Sheets[revName], { defval:null });
    if(!rev.length) throw new Error('Revenue 工作表為空');
    const rh = Object.keys(rev[0]);
    const colCode = detectCol(rh, [/個股|代號|股票代號|code|symbol/i]);
    const colName = detectCol(rh, [/名稱|公司|name/i]);

    const revRows = rev.map(r=>({ code:norm(r[colCode]), name:norm(r[colName]) }))
                       .filter(r=>r.code || r.name);

    // Links（供應鏈關係 + 關係類型）
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

    // 名稱查詢表
    const byCode = new Map(); for(const r of revRows){ if(r.code) byCode.set(r.code, r); }

    // 上/下游快速映射（僅代碼，用於相容舊函式）
    const upstreamOf = new Map(), downstreamOf = new Map();
    for(const l of linksRows){ if(l.up && l.down){
      if(!downstreamOf.has(l.up)) downstreamOf.set(l.up, []);
      downstreamOf.get(l.up).push(l.down);
      if(!upstreamOf.has(l.down)) upstreamOf.set(l.down, []);
      upstreamOf.get(l.down).push(l.up);
    }}

    Object.assign(cache,{ loaded:true, revRows, linksRows, byCode, upstreamOf, downstreamOf });
    return cache;
  }

  // ------------------ Rendering ------------------
  function stockLiHtml(code, byCode){
    const r = byCode.get(code) || { code, name:'' };
    return `<li class="icp-stock" data-code="${r.code}"><span class="code">${r.code}</span> ${r.name||''}</li>`;
  }

  function renderGroupListByPairs(el, pairs, byCode){
    // pairs: [{code, group}]
    if(!el) return;
    const normPairs = pairs.map(p=>({ code:norm(p.code), group:norm(p.group)||'未分類' }))
                           .filter(p=>p.code);
    const g = groupBy(normPairs, p=>p.group);
    const html = [];
    for(const [grp, arr] of g.entries()){
      // 去重 & 排序
      const uniqCodes = Array.from(new Set(arr.map(a=>a.code))).sort((a,b)=>a.localeCompare(b));
      const list = uniqCodes.map(code=>stockLiHtml(code, byCode)).join('');
      html.push(`<div class="icp-card"><div class="icp-card-title">${grp}</div><ul>${list}</ul></div>`);
    }
    el.innerHTML = html.join('');
  }

  function findStock(byCode, rows, kw){ const k=norm(kw); return k? (byCode.get(k) || rows.find(r=>r.name===k) || null) : null; }

  // ----- 對齊：以右側 section 的 head & chart/legend 為錨點 -----
  function setFoldToAnchors(){
    const wrap = $('#icp-fold-wrap');
    const scroll = $('#icp-scroll');
    const btn = $('#icp-expander');
    const fade = $('#icp-fade');
    const secHead = $('#combo-section .section-head');
    const chart = $('#combo-section .chart-wrap');
    const legend = $('#combo-section .legend-row');
    if(!wrap || !scroll || !btn || !secHead || !chart) return;

    const wrapRect  = wrap.getBoundingClientRect();
    const headRect  = secHead.getBoundingClientRect();
    const chartRect = chart.getBoundingClientRect();
    const legendRect= legend ? legend.getBoundingClientRect() : null;

    const topAnchor    = Math.round(headRect.bottom - wrapRect.top);
    const bottomAnchor = Math.round(Math.max(chartRect.bottom, legendRect? legendRect.bottom : -Infinity) - wrapRect.top);
    const targetHeight = Math.max(0, bottomAnchor - topAnchor);

    const isExpanded = btn.getAttribute('aria-expanded') === 'true';
    scroll.style.marginTop = (topAnchor < 0 ? 0 : topAnchor) + 'px';
    if(!isExpanded){
      scroll.style.maxHeight = targetHeight + 'px';
      if(fade) fade.style.display = 'block';
    }else{
      scroll.style.maxHeight = scroll.scrollHeight + 'px';
      if(fade) fade.style.display = 'none';
    }

    const btnTop = Math.round(bottomAnchor - (btn.offsetHeight/2));
    btn.style.top = (btnTop < 0 ? 0 : btnTop) + 'px';
  }

  function toggleFold(){
    const btn = document.getElementById('icp-expander');
    const expanded = btn.getAttribute('aria-expanded') === 'true';
    btn.setAttribute('aria-expanded', String(!expanded));
    btn.textContent = expanded ? '＋' : '－';
    setFoldToAnchors();
  }

  function installObservers(){
    const head = $('#combo-section .section-head');
    const chart = $('#combo-section .chart-wrap');
    const legend = $('#combo-section .legend-row');
    if(head){ const ro1 = new ResizeObserver(()=> setFoldToAnchors()); ro1.observe(head); }
    if(chart){ const ro2 = new ResizeObserver(()=> setFoldToAnchors()); ro2.observe(chart); }
    if(legend){ const ro3 = new ResizeObserver(()=> setFoldToAnchors()); ro3.observe(legend); }
    window.addEventListener('resize', setFoldToAnchors);
  }

  async function updatePanel(){
    const { revRows, byCode, upstreamOf, downstreamOf, linksRows } = await loadAll();
    const me = findStock(byCode, revRows, document.getElementById('stockInput')?.value || '');
    const upWrap = document.getElementById('icp-up-wrap');
    const downWrap = document.getElementById('icp-down-wrap');

    if(!me){ if(upWrap) upWrap.innerHTML=''; if(downWrap) downWrap.innerHTML=''; setFoldToAnchors(); return; }

    // 以 Links.關係類型 分群
    const upPairs   = linksRows.filter(l=> l.down === me.code).map(l=> ({ code:l.up,   group:l.type || '未分類' }));
    const downPairs = linksRows.filter(l=> l.up   === me.code).map(l=> ({ code:l.down, group:l.type || '未分類' }));

    renderGroupListByPairs(upWrap, upPairs, byCode);
    renderGroupListByPairs(downWrap, downPairs, byCode);

    // 事件委派：點個股開外部頁
    function bindClick(root){ if(!root) return; root.addEventListener('click', (e)=>{
      const li = e.target.closest('.icp-stock'); if(!li) return; const code = li.getAttribute('data-code'); if(!code) return;
      const url = 'https://www.fbs.com.tw/MKT/Index?name=' + encode('Ｊ線圖') + '&stock=' + encode(code);
      window.open(url, '_blank');
    }); }
    bindClick(upWrap); bindClick(downWrap);

    setTimeout(setFoldToAnchors, 0);
  }

  document.addEventListener('DOMContentLoaded', async ()=>{
    try{ await loadAll(); }catch(e){ console.error(e); }
    const btn = document.getElementById('icp-expander'); if(btn){ btn.addEventListener('click', toggleFold); }
    const run = document.getElementById('runBtn'); if(run){ run.addEventListener('click', ()=> setTimeout(updatePanel, 0)); }
    setTimeout(updatePanel, 0);
    installObservers();
  });
})();
