/* industry-chain-panel.js (v9.4)
 * 變更：
 *  - 下游客戶資料來源改讀 Links 工作表 **H~J 欄**（上游代號 / 下游客戶 / 關係類型）。
 *  - 上游供應鏈維持讀取 **A~C 欄**。
 *  - 採用 header:1 讀取整張 Sheet，以欄位位置分段，不受重複表頭影響。
 *  - 其他維持：.US 不外連、點擊只綁一次、對齊強化（dPR/字型載入/CSS nudge）。
 */
(function(){
  const $  = (s)=> document.querySelector(s);
  const $$ = (s)=> Array.from(document.querySelectorAll(s));
  const encode = (s)=> encodeURIComponent(s);
  const RAF = (fn)=> requestAnimationFrame(()=> requestAnimationFrame(fn));
  const isUS = (code)=> /\.US$/i.test(String(code||'').trim());
  const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();

  function norm(s){ return String(s==null?'':s).trim(); }
  function groupBy(arr, keyFn){ const m=new Map(); for(const x of arr){ const k=keyFn(x)||''; if(!m.has(k)) m.set(k,[]); m.get(k).push(x);} return m; }

  // ---------- dPR 四捨五入 & CSS Nudge ----------
  const pxRound =(x)=> Math.round(x * (window.devicePixelRatio||1)) / (window.devicePixelRatio||1);
  const cssNudge=(name)=>{ const v=getComputedStyle(document.documentElement).getPropertyValue(name).trim(); const n=parseInt(v,10); return Number.isFinite(n)? n:0; };

  // ---------- 讀取 Excel ----------
  const cache = { loaded:false };
  async function loadAll(){
    if(cache.loaded) return cache;
    const res = await fetch('data.xlsx?v='+URL_VER, { cache:'no-store' });
    if(!res.ok) throw new Error('讀取 data.xlsx 失敗 HTTP '+res.status);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array' });

    // Revenue（建立代號→名稱對照）
    const revName = wb.SheetNames.find(n=>/revenue|營收/i.test(n)) || wb.SheetNames[0];
    const revRowsObj = XLSX.utils.sheet_to_json(wb.Sheets[revName], { defval:null });
    if(!revRowsObj.length) throw new Error('Revenue 工作表為空');
    const rHeaders = Object.keys(revRowsObj[0]);
    const codeHeader = rHeaders.find(h=>/(個股|代號|股票代號|code|symbol)/i.test(String(h))) || rHeaders[0];
    const nameHeader = rHeaders.find(h=>/(名稱|公司|name)/i.test(String(h))) || rHeaders[1];
    const revRows = revRowsObj.map(r=>({ code:norm(r[codeHeader]), name:norm(r[nameHeader]) })).filter(r=>r.code||r.name);

    // Links：改用 header:1 讀整表，分 A~C 與 H~J 兩段
    const linkName = wb.SheetNames.find(n=>/link|關聯|關係|供應鏈/i.test(n));
    let linksUpstreamAC = [], linksDownstreamHJ = [];
    if(linkName){
      const rowsA1 = XLSX.utils.sheet_to_json(wb.Sheets[linkName], { header:1, defval:null, blankrows:false });
      // 取最長行數，從第2列起掃描
      for(let i=1;i<rowsA1.length;i++){
        const row = rowsA1[i] || [];
        const A = norm(row[0]); // 上游代號 (A)
        const B = norm(row[1]); // 下游代號 (B)
        const C = norm(row[2]); // 關係類型 (C)
        const H = norm(row[7]); // 上游代號 (H)
        const I = norm(row[8]); // 下游客戶 (I)
        const J = norm(row[9]); // 關係類型 (J)
        if(A || B || C){ linksUpstreamAC.push({ up:A, down:B, type:C }); }
        if(H || I || J){ linksDownstreamHJ.push({ up:H, down:I, type:J }); }
      }
      // 清掉空行
      linksUpstreamAC = linksUpstreamAC.filter(x=>x.up||x.down);
      linksDownstreamHJ = linksDownstreamHJ.filter(x=>x.up||x.down);
    }

    const byCode = new Map(); for(const r of revRows){ if(r.code) byCode.set(r.code,r); }

    Object.assign(cache,{
      loaded:true,
      revRows,
      byCode,
      linksUpstreamAC,
      linksDownstreamHJ
    });
    return cache;
  }

  // ---------- Rendering ----------
  function stockLiHtml(code, byCode){ const r=byCode.get(code)||{code,name:''}; return `<li class="icp-stock" data-code="${r.code}"><span class="code">${r.code}</span> ${r.name||''}</li>`; }

  function renderGroupListByPairs(el, pairs, byCode){
    if(!el) return;
    const normPairs = pairs.map(p=>({ code:norm(p.code), group:norm(p.group)||'未分類' })).filter(p=>p.code);
    const g = groupBy(normPairs, p=>p.group);
    const html=[]; for(const [grp, arr] of g.entries()){
      const uniqCodes = Array.from(new Set(arr.map(a=>a.code))).sort((a,b)=>a.localeCompare(b));
      const list = uniqCodes.map(code=>stockLiHtml(code, byCode)).join('');
      html.push(`<div class="icp-card"><div class="icp-card-title">${grp}</div><ul>${list}</ul></div>`);
    }
    el.innerHTML = html.join('');
  }

  function findStock(byCode, rows, kw){ const k=norm(kw); return k? (byCode.get(k) || rows.find(r=>r.name===k) || null) : null; }

  // ---------- 對齊（含 dPR / CSS nudge / 字型就緒） ----------
  function setFoldToAnchors(){
    const wrap   = $('#icp-fold-wrap');
    const scroll = $('#icp-scroll');
    const btn    = $('#icp-expander');
    const fade   = $('#icp-fade');

    const head   = $('#combo-section .section-head');
    const chartW = $('#combo-section .chart-wrap');
    const legend = $('#combo-section .legend-row');
    if(!wrap||!scroll||!btn||!head||!chartW) return;

    const wrapRect   = wrap.getBoundingClientRect();
    const headRect   = head.getBoundingClientRect();
    const chartRect  = chartW.getBoundingClientRect();
    const legendRect = legend ? legend.getBoundingClientRect() : null;

    const headCS  = getComputedStyle(head);
    const chartCS = getComputedStyle(chartW);

    const headBottom   = headRect.top + (parseFloat(headCS.paddingTop)||0) + (parseFloat(headCS.height)||headRect.height);
    const chartBottom  = chartRect.bottom + (parseFloat(chartCS.marginBottom)||0);
    const legendBottom = legendRect ? legendRect.bottom : -Infinity;

    const topAnchor    = pxRound(headBottom - wrapRect.top) + cssNudge('--icp-body-shift');
    const bottomAnchor = pxRound(Math.max(chartBottom, legendBottom) - wrapRect.top) + cssNudge('--icp-body-height');

    const firstBox = scroll.querySelector('.icp-box');
    let innerTopOffset = 0; if(firstBox){ const fb = firstBox.getBoundingClientRect(); const sc = scroll.getBoundingClientRect(); innerTopOffset = pxRound(fb.top - sc.top); }

    const isExpanded = btn.getAttribute('aria-expanded') === 'true';
    const marginTop = Math.max(-32, topAnchor - innerTopOffset);
    scroll.style.marginTop = marginTop + 'px';

    const targetHeight = Math.max(0, bottomAnchor - (marginTop + innerTopOffset));
    if(!isExpanded){ scroll.style.maxHeight = targetHeight + 'px'; if(fade) fade.style.display='block'; }
    else{ scroll.style.maxHeight = scroll.scrollHeight + 'px'; if(fade) fade.style.display='none'; }

    const btnTop = pxRound(bottomAnchor - (btn.offsetHeight/2));
    btn.style.top = (btnTop<0?0:btnTop) + 'px';
  }

  function toggleFold(){ const btn=$('#icp-expander'); const expanded=btn.getAttribute('aria-expanded')==='true'; btn.setAttribute('aria-expanded', String(!expanded)); btn.textContent = expanded ? '＋' : '－'; RAF(setFoldToAnchors); }

  function installObservers(){
    const head=$('#combo-section .section-head'); const chart=$('#combo-section .chart-wrap'); const legend=$('#combo-section .legend-row');
    if(head){ const ro1=new ResizeObserver(()=> RAF(setFoldToAnchors)); ro1.observe(head); }
    if(chart){ const ro2=new ResizeObserver(()=> RAF(setFoldToAnchors)); ro2.observe(chart); }
    if(legend){ const ro3=new ResizeObserver(()=> RAF(setFoldToAnchors)); ro3.observe(legend); }
    window.addEventListener('resize', ()=> RAF(setFoldToAnchors));
  }

  async function firstAlign(){ try{ if(document.fonts && document.fonts.ready) await document.fonts.ready; }catch(e){} RAF(setFoldToAnchors); }

  // ---------- 點擊委派（只綁一次；.US 不外連） ----------
  function bindClickOnce(root){
    if(!root) return;
    if(root.dataset.bound === '1') return; // 防重綁
    const handler = (e)=>{
      const li = e.target.closest('.icp-stock');
      if(!li) return;
      const code = li.getAttribute('data-code');
      if(!code || isUS(code)) return;
      const url='https://www.fbs.com.tw/MKT/Index?name='+encode('Ｊ線圖')+'&stock='+encode(code);
      window.open(url, '_blank', 'noopener');
    };
    root.addEventListener('click', handler);
    root.dataset.bound = '1';
  }

  // ---------- 主流程 ----------
  async function updatePanel(){
    const { revRows, byCode, linksUpstreamAC, linksDownstreamHJ } = await loadAll();
    const me = findStock(byCode, revRows, document.getElementById('stockInput')?.value || '');
    const upWrap = document.getElementById('icp-up-wrap');
    const downWrap = document.getElementById('icp-down-wrap');

    if(!me){ if(upWrap) upWrap.innerHTML=''; if(downWrap) downWrap.innerHTML=''; RAF(setFoldToAnchors); return; }

    // 以上游供應鏈（A~C）：找所有 down === 自己 → 取 up 分群
    const upPairs   = linksUpstreamAC.filter(l=> l.down === me.code).map(l=>({ code:l.up,   group:l.type || '未分類' }));

    // 以下游客戶（H~J）：找所有 up === 自己 → 取 I 欄（下游客戶）分群
    const downPairs = linksDownstreamHJ.filter(l=> l.up   === me.code).map(l=>({ code:l.down, group:l.type || '未分類' }));

    renderGroupListByPairs(upWrap, upPairs, byCode);
    renderGroupListByPairs(downWrap, downPairs, byCode);

    bindClickOnce(upWrap);
    bindClickOnce(downWrap);

    RAF(setFoldToAnchors);
  }

  document.addEventListener('DOMContentLoaded', async ()=>{
    try{ await loadAll(); }catch(e){ console.error(e); }
    const btn=$('#icp-expander'); if(btn) btn.addEventListener('click', toggleFold);
    const run=$('#runBtn'); if(run) run.addEventListener('click', ()=> RAF(updatePanel));
    installObservers();
    firstAlign();
    RAF(updatePanel);
  });
})();
