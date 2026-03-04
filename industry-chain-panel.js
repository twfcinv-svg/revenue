
/* industry-chain-panel.js (v6, row-aligned)
 * 變更：將左「個股產業鏈」與右「營收走勢」放入相同 row（兩欄 panel）。
 * 折疊高度與加號位置以右側 panel-body 為基準，保證同頂同底。
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

  function anchorRects(){
    const leftBody  = document.querySelector('#icp-panel .panel-body');
    const rightBody = document.querySelector('#chart-panel .panel-body');
    if(!leftBody || !rightBody) return null;
    return { leftBody, rightBody, leftRect:leftBody.getBoundingClientRect(), rightRect:rightBody.getBoundingClientRect() };
  }

  function setFoldToAnchor(){
    const ctx = anchorRects();
    const scroll = document.getElementById('icp-scroll');
    const btn = document.getElementById('icp-expander');
    const fade = document.getElementById('icp-fade');
    if(!ctx || !scroll || !btn) return;

    const { leftRect, rightRect } = ctx;
    const isExpanded = btn.getAttribute('aria-expanded') === 'true';

    // 對齊右側 panel-body 的 top/bottom
    const targetTop    = Math.round(rightRect.top    - leftRect.top);
    const targetBottom = Math.round(rightRect.bottom - leftRect.top);
    const targetHeight = Math.max(0, targetBottom - targetTop);

    scroll.style.marginTop = (targetTop < 0 ? 0 : targetTop) + 'px';
    if(!isExpanded){
      scroll.style.maxHeight = targetHeight + 'px';
      if(fade) fade.style.display = 'block';
    }else{
      scroll.style.maxHeight = scroll.scrollHeight + 'px';
      if(fade) fade.style.display = 'none';
    }

    // 加號定位在右側 panel-body 底緣中線
    const top = Math.round(targetBottom - (btn.offsetHeight/2));
    btn.style.top = (top < 0 ? 0 : top) + 'px';
  }

  function toggleFold(){
    const btn = document.getElementById('icp-expander');
    const expanded = btn.getAttribute('aria-expanded') === 'true';
    btn.setAttribute('aria-expanded', String(!expanded));
    btn.textContent = expanded ? '＋' : '－';
    setFoldToAnchor();
  }

  function installObservers(){
    const rightBody = document.querySelector('#chart-panel .panel-body');
    if(rightBody){ const ro = new ResizeObserver(()=> setFoldToAnchor()); ro.observe(rightBody); }
    window.addEventListener('resize', setFoldToAnchor);
  }

  async function updatePanel(){
    const { revRows, byCode, upstreamOf, downstreamOf } = await loadAll();
    const me = findStock(byCode, revRows, document.getElementById('stockInput')?.value || '');
    const upWrap = document.getElementById('icp-up-wrap');
    const downWrap = document.getElementById('icp-down-wrap');

    if(!me){ if(upWrap) upWrap.innerHTML=''; if(downWrap) downWrap.innerHTML=''; setFoldToAnchor(); return; }

    renderGroupList(upWrap, upstreamOf.get(me.code) || [], byCode);
    renderGroupList(downWrap, downstreamOf.get(me.code) || [], byCode);

    // 點個股開外網
    function bindClick(root){ if(!root) return; root.addEventListener('click', (e)=>{
      const li = e.target.closest('.icp-stock'); if(!li) return; const code = li.getAttribute('data-code'); if(!code) return;
      const url = 'https://www.fbs.com.tw/MKT/Index?name=' + encode('Ｊ線圖') + '&stock=' + encode(code);
      window.open(url, '_blank');
    }); }
    bindClick(upWrap); bindClick(downWrap);

    setTimeout(setFoldToAnchor, 0);
  }

  document.addEventListener('DOMContentLoaded', async ()=>{
    try{ await loadAll(); }catch(e){ console.error(e); }
    const btn = document.getElementById('icp-expander'); if(btn){ btn.addEventListener('click', toggleFold); }
    const run = document.getElementById('runBtn'); if(run){ run.addEventListener('click', ()=> setTimeout(updatePanel, 0)); }
    setTimeout(updatePanel, 0);
    installObservers();
  });
})();
