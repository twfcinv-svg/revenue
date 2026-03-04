
/* industry-chain-panel.js (foldable v3)
 * 新增：
 *  - 左/右清單支援「摺疊/展開」，預設顯示高度與右側「營收走勢」底部齊平。
 *  - 點擊「＋ / －」按鈕展開/收合；視窗縮放與圖表重繪時自動重算高度與按鈕位置。
 *  - 既有功能（讀 data.xlsx、渲染上/下游、群組卡 hover 亮邊、個股 hover 浮起、點擊開外網）維持。
 */
(function(){
  const $ = (s)=> document.querySelector(s);
  const $$ = (s)=> Array.from(document.querySelectorAll(s));
  const encode = (s)=> encodeURIComponent(s);

  // -------- Helpers --------
  function norm(s){ return String(s||'').trim(); }
  function detectCol(headers, patterns){
    for(const p of patterns){ for(const h of headers){ if(p.test(h)) return h; } }
    return null;
  }
  function groupBy(arr, keyFn){
    const map = new Map();
    for(const x of arr){ const k = keyFn(x)||''; if(!map.has(k)) map.set(k,[]); map.get(k).push(x); }
    return map;
  }

  // -------- Load Excel once --------
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

    const byCode = new Map(); for(const r of revRows){ if(r.code) byCode.set(r.code, r); }
    const upstreamOf = new Map(), downstreamOf = new Map();
    for(const l of linksRows){ if(l.up && l.down){
      if(!downstreamOf.has(l.up)) downstreamOf.set(l.up, []); downstreamOf.get(l.up).push(l.down);
      if(!upstreamOf.has(l.down)) upstreamOf.set(l.down, []); upstreamOf.get(l.down).push(l.up);
    }}

    Object.assign(cache, { loaded:true, revRows, byCode, upstreamOf, downstreamOf });
    return cache;
  }

  // -------- Renderers --------
  function stockLiHtml(code, byCode){
    const r = byCode.get(code) || { code, name:'' };
    return `<li class="icp-stock" data-code="${r.code}"><span class="code">${r.code}</span> ${r.name||''}</li>`;
  }

  function renderGroupList(el, codes, byCode){
    if(!el) return;
    const items = Array.from(new Set(codes)).map(c => byCode.get(c) || { code:c, name:'', industry:'' });
    const map = groupBy(items, it => it.industry);
    const chunks = [];
    for(const [ind, arr] of map.entries()){
      const list = arr.sort((a,b)=>a.code.localeCompare(b.code))
                      .map(r=>stockLiHtml(r.code, byCode)).join('');
      chunks.push(`<div class="icp-card"><div class="icp-card-title">${ind||'未分類'}</div><ul>${list}</ul></div>`);
    }
    el.innerHTML = chunks.join('');
  }

  function findStock(byCode, rows, keyword){
    const k = norm(keyword); if(!k) return null;
    return byCode.get(k) || rows.find(r=>r.name===k) || null;
  }

  // -------- Fold / Expand --------
  function setFoldToChartHeight(){
    const wrap = $('#icp-fold-wrap');
    const scroll = $('#icp-scroll');
    const btn = $('#icp-expander');
    const fade = $('#icp-fade');
    const chart = document.querySelector('#combo-section .chart-wrap');
    if(!wrap || !scroll || !btn || !chart) return;

    const chartRect = chart.getBoundingClientRect();
    const wrapRect = wrap.getBoundingClientRect();
    const isExpanded = btn.getAttribute('aria-expanded') === 'true';

    // 設定摺疊高度 = 圖表容器高度
    if(!isExpanded){
      const h = Math.max(0, Math.round(chartRect.height));
      scroll.style.maxHeight = h + 'px';
      if(fade) fade.style.display = 'block';
    }else{
      scroll.style.maxHeight = scroll.scrollHeight + 'px';
      if(fade) fade.style.display = 'none';
    }

    // 讓「＋」與摺疊線與圖表底對齊（同一水平線）
    const top = Math.round(chartRect.bottom - wrapRect.top - (btn.offsetHeight/2));
    btn.style.top = (top < 0 ? 0 : top) + 'px';
  }

  function toggleFold(){
    const btn = $('#icp-expander');
    const expanded = btn.getAttribute('aria-expanded') === 'true';
    btn.setAttribute('aria-expanded', String(!expanded));
    btn.textContent = expanded ? '＋' : '－';
    setFoldToChartHeight();
  }

  function installObservers(){
    const chart = document.querySelector('#combo-section .chart-wrap');
    if(chart){
      const ro = new ResizeObserver(()=> setFoldToChartHeight());
      ro.observe(chart);
    }
    window.addEventListener('resize', setFoldToChartHeight);
  }

  async function updatePanel(){
    const { revRows, byCode, upstreamOf, downstreamOf } = await loadAll();
    const kw = $('#stockInput') ? $('#stockInput').value.trim() : '';
    const me = findStock(byCode, revRows, kw);
    const upWrap = $('#icp-up-wrap');
    const downWrap = $('#icp-down-wrap');

    if(!me){ if(upWrap) upWrap.innerHTML=''; if(downWrap) downWrap.innerHTML=''; setFoldToChartHeight(); return; }

    renderGroupList(upWrap, upstreamOf.get(me.code) || [], byCode);
    renderGroupList(downWrap, downstreamOf.get(me.code) || [], byCode);

    // 事件委派：點個股開外部頁
    function bindClick(root){ if(!root) return; root.addEventListener('click', (e)=>{
      const li = e.target.closest('.icp-stock'); if(!li) return;
      const code = li.getAttribute('data-code'); if(!code) return;
      const url = 'https://www.fbs.com.tw/MKT/Index?name=' + encode('Ｊ線圖') + '&stock=' + encode(code);
      window.open(url, '_blank');
    }); }
    bindClick(upWrap); bindClick(downWrap);

    // 渲染完成後依圖表高度設定摺疊線與按鈕位置
    setTimeout(setFoldToChartHeight, 0);
  }

  document.addEventListener('DOMContentLoaded', async ()=>{
    try{ await loadAll(); }catch(e){ console.error(e); }

    // 建立「＋ / －」按鈕事件
    const btn = document.getElementById('icp-expander');
    if(btn){ btn.addEventListener('click', toggleFold); }

    // 綁定查詢與初始渲染
    const run = document.getElementById('runBtn');
    if(run){ run.addEventListener('click', ()=> setTimeout(updatePanel, 0)); }
    setTimeout(updatePanel, 0);

    // 監看圖表尺寸（RWD / 重繪）
    installObservers();
  });
})();
