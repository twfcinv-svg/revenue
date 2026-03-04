
/* industry-chain-panel.js  個股產業鏈（上游｜同業｜下游）資訊面板
 * 依據 data.xlsx：
 *   - Revenue 工作表（至少需要：代碼、名稱、產業別）
 *   - Links   工作表（至少需要：上游代號、下游代號、關係類型）
 *
 * 功能：
 *   - 當使用者按下頁面上的「查詢」按鈕（#runBtn）或首次載入後輸入代號/名稱時，
 *     顯示該檔個股的：
 *       左側：上游產業相關個股（依業別分群）
 *       中間：同產業（Peers）與該股產業名稱
 *       右側：下游產業相關個股（依業別分群）
 */
(function(){
  const $ = (s)=> document.querySelector(s);

  // -------- Helpers --------
  function norm(s){ return String(s||'').trim(); }
  function detectCol(headers, patterns){
    const hset = new Set(headers);
    for(const p of patterns){
      for(const h of headers){ if(p.test(h)) return h; }
    }
    return null;
  }

  function groupBy(arr, keyFn){
    const map = new Map();
    for(const x of arr){
      const k = keyFn(x) || '';
      if(!map.has(k)) map.set(k, []);
      map.get(k).push(x);
    }
    return map;
  }

  // -------- Loaders --------
  const cache = { loaded:false };

  async function loadAll(){
    if(cache.loaded) return cache;
    const res = await fetch('data.xlsx');
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array' });

    // Revenue（找名稱類似的表）
    const revName = wb.SheetNames.find(n=>/revenue|營收/i.test(n)) || wb.SheetNames[0];
    const rev = XLSX.utils.sheet_to_json(wb.Sheets[revName], { defval:null });
    if(!rev.length) throw new Error('Revenue 工作表為空');

    const rh = Object.keys(rev[0]);
    const colCode = detectCol(rh, [/個股|代號|股票代號|code|symbol/i]);
    const colName = detectCol(rh, [/名稱|公司|name/i]);
    const colInd  = detectCol(rh, [/產業別|產業|industry/i]);

    const revRows = rev.map(r=>({
      code: norm(r[colCode]),
      name: norm(r[colName]),
      industry: norm(r[colInd])
    })).filter(r=>r.code || r.name);

    // Links（上游/下游關係）
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

    // 建表
    const byCode = new Map();
    for(const r of revRows){ if(r.code) byCode.set(r.code, r); }

    const upstreamOf = new Map();   // key: code => [up codes]
    const downstreamOf = new Map(); // key: code => [down codes]

    for(const l of linksRows){
      if(l.up && l.down){
        if(!downstreamOf.has(l.up)) downstreamOf.set(l.up, []);
        downstreamOf.get(l.up).push(l.down);
        if(!upstreamOf.has(l.down)) upstreamOf.set(l.down, []);
        upstreamOf.get(l.down).push(l.up);
      }
    }

    Object.assign(cache, { loaded:true, revRows, byCode, upstreamOf, downstreamOf });
    return cache;
  }

  // -------- Rendering --------
  function li(code, byCode){
    const r = byCode.get(code) || { code, name:'' };
    return `<li><span class="code">${r.code || ''}</span> ${r.name || ''}</li>`;
  }

  function renderList(el, arr, byCode){
    if(!el) return;
    const uniq = Array.from(new Set(arr)).slice(0, 50);
    el.innerHTML = uniq.map(c => li(c, byCode)).join('');
  }

  function renderGroupList(el, codes, byCode){
    if(!el) return;
    const items = Array.from(new Set(codes)).map(c => byCode.get(c) || { code:c, name:'', industry:'' });
    const map = groupBy(items, it => it.industry);
    const chunks = [];
    for(const [ind, arr] of map.entries()){
      const list = arr.sort((a,b)=>a.code.localeCompare(b.code)).slice(0, 20)
        .map(r=>`<li><span class="code">${r.code}</span> ${r.name}</li>`).join('');
      chunks.push(`<div class="icp-card"><div class="icp-card-title">${ind||'未分類'}</div><ul>${list}</ul></div>`);
    }
    el.innerHTML = chunks.join('');
  }

  function findStock(byCode, rows, keyword){
    const k = norm(keyword);
    if(!k) return null;
    return byCode.get(k) || rows.find(r=>r.name===k) || null;
  }

  async function updatePanel(){
    const { revRows, byCode, upstreamOf, downstreamOf } = await loadAll();
    const kw = $('#stockInput') ? $('#stockInput').value.trim() : '';
    const me = findStock(byCode, revRows, kw);
    const centerTitle = $('#icp-center-title');
    const centerPeers = $('#icp-center-peers');
    const upWrap = $('#icp-up-wrap');
    const downWrap = $('#icp-down-wrap');

    if(!me){
      if(centerTitle) centerTitle.textContent = '（請輸入個股並按查詢）';
      if(centerPeers) centerPeers.innerHTML = '';
      if(upWrap) upWrap.innerHTML = '';
      if(downWrap) downWrap.innerHTML = '';
      return;
    }

    if(centerTitle) centerTitle.textContent = (me.industry || '產業') + '（Peers）';

    // Peers = 同產業（排除自己）
    const peers = revRows.filter(r => r.industry===me.industry && r.code!==me.code).slice(0, 30).map(r=>r.code);
    renderList(centerPeers, peers, byCode);

    // 上游 = 下游指向我（down = me.code）
    const ups = upstreamOf.get(me.code) || [];
    renderGroupList(upWrap, ups, byCode);

    // 下游 = 我指向下游（up = me.code）
    const downs = downstreamOf.get(me.code) || [];
    renderGroupList(downWrap, downs, byCode);
  }

  document.addEventListener('DOMContentLoaded', async ()=>{
    try{ await loadAll(); }catch(e){ console.error(e); }
    const btn = $('#runBtn');
    if(btn) btn.addEventListener('click', ()=> setTimeout(updatePanel, 0));
    // 初次也可嘗試更新（若已有輸入值）
    setTimeout(updatePanel, 0);
    window.addEventListener('resize', ()=>{});
  });
})();
