/* combo-chart.js (v2.1)  修復：查無個股導致圖表為空；強化載入與比對
 *  - 以 data.xlsx?v=VER + cache:'no-store' 破壞快取
 *  - 更寬鬆的個股比對：代號/名稱精確、code+空白+name、startsWith() 全面支援
 *  - 維持只顯示三個刻度（最新／去年同月／前年同月）
 */
(function(){
  const $ = (sel) => document.querySelector(sel);
  const $$closest = (el, sel) => (el && el.closest) ? el.closest(sel) : null;

  // Tooltip 容器
  const svgNode = document.getElementById('comboChart');
  const chartWrap = svgNode ? ($$closest(svgNode, '.chart-wrap') || document.body) : document.body;
  if(chartWrap && getComputedStyle(chartWrap).position === 'static'){
    chartWrap.style.position = 'relative';
  }
  const tooltip = document.createElement('div');
  tooltip.className = 'combo-tooltip';
  Object.assign(tooltip.style, {
    position: 'absolute', pointerEvents: 'none', display: 'none', padding: '8px 10px',
    fontSize: '12px', lineHeight: '1.4', color: '#fff', background: 'rgba(0,0,0,0.82)',
    border: '1px solid rgba(255,255,255,0.18)', borderRadius: '8px', boxShadow: '0 4px 14px rgba(0,0,0,.35)',
    zIndex: 1000, whiteSpace: 'nowrap'
  });
  chartWrap.appendChild(tooltip);

  const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
  function norm(s){ return String(s || '').trim(); }
  function fmtYM(ym){ return ym.slice(0,4)+'-'+ym.slice(4,6); }
  function fmtPct(v){ return (v==null)?'': d3.format('+.0f')(v)+'%'; }
  function fmtMoney(val){ if(val==null) return ''; const si=d3.format('.2s'); return si(val*1000).replace('G','B'); }
  function shiftYM(ym, delta){ let y = +ym.slice(0,4), m = +ym.slice(4,6); let total = y*12 + (m-1) + delta; let y2 = Math.floor(total/12); let m2 = total%12 + 1; return String(y2).padStart(4,'0') + String(m2).padStart(2,'0'); }

  const state = { loaded:false, rows:[], columns:null, months:null, indexByCode:new Map(), indexByName:new Map() };

  function detectColumns(headers){
    const col = { code:null, name:null, industry:null, amount:{}, mom:{}, yoy:{} };
    const reMonth = /^(20\d{2})[\/年-]?\s*(0[1-9]|1[0-2])/; // 支援 YYYYMM / YYYY-MM / YYYY/MM
    headers.forEach(h => {
      const H = norm(h); if(!H) return;
      if(!col.code && /(個股|代號|股票代號|Code|Symbol)/i.test(H)) col.code = h;
      if(!col.name && /(名稱|公司|Name)/i.test(H)) col.name = h;
      if(!col.industry && /(產業|產業別|Industry)/i.test(H)) col.industry = h;
      const m = H.match(reMonth);
      if(m){
        const ym = m[1]+m[2];
        if(/(單月)?(合併)?營收(?!.*(年|月))|營收$|金額|千元|仟/i.test(H)) col.amount[ym] = h; // 寬鬆偵測金額欄
        else if(/(月變動|月增率|MoM)/i.test(H)) col.mom[ym] = h;
        else if(/(年成長|年增率|YoY)/i.test(H)) col.yoy[ym] = h;
      }
    });
    return col;
  }

  function parseNumber(v){
    if(v===null || v===undefined || v==='') return null;
    if(typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/[% ,，]/g,'');
    const n = parseFloat(s);
    return isFinite(n) ? n : null;
  }

  async function loadRevenue(){
    if(state.loaded) return state;
    const res = await fetch('data.xlsx?v='+URL_VER, { cache:'no-store' });
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array' });
    const sheetName = wb.SheetNames.find(n=>/revenue|營收/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval:null });
    if(!rows.length) throw new Error('Revenue 工作表為空');

    const headers = Object.keys(rows[0]);
    const col = detectColumns(headers);
    const months_amount = Object.keys(col.amount).sort();
    const months_mom    = Object.keys(col.mom).sort();
    const months_yoy    = Object.keys(col.yoy).sort();
    const months_all    = Array.from(new Set([...months_amount, ...months_mom, ...months_yoy])).sort();

    const tidy = rows.map(r=>{
      const code = norm(r[col.code]);
      const name = norm(r[col.name]);
      const ind  = norm(r[col.industry]);
      const series = months_all.map(ym=>({
        ym,
        amount: parseNumber(r[col.amount[ym]]),
        mom:    parseNumber(r[col.mom[ym]]),
        yoy:    parseNumber(r[col.yoy[ym]])
      })).filter(d=> d.amount!==null || d.mom!==null || d.yoy!==null );
      return { code, name, industry:ind, series };
    }).filter(d=> d.code || d.name);

    const byCode=new Map(), byName=new Map();
    for(const r of tidy){ if(r.code) byCode.set(r.code, r); if(r.name) byName.set(r.name, r); }

    state.loaded = true;
    state.rows = tidy;
    state.columns = col;
    state.months = { amount:months_amount, mom:months_mom, yoy:months_yoy, all:months_all };
    state.indexByCode = byCode; state.indexByName = byName;
    return state;
  }

  function parseKeyword(k){
    const raw = norm(k);
    if(!raw) return { raw:'', code:'', name:'' };
    const m = raw.match(/^([A-Za-z0-9.]+)\s+(.*)$/); // 支援 "2330 台積電" 或 "TSM.US TSMC"
    if(m){ return { raw, code:norm(m[1]), name:norm(m[2]) }; }
    return { raw, code:raw, name:raw };
  }

  function pickStartsWith(rows, q){
    if(!q) return null; const Q = q.toLowerCase();
    return rows.find(r => (r.code||'').toLowerCase().startsWith(Q) || (r.name||'').toLowerCase().startsWith(Q)) || null;
  }

  function findStock(keyword){
    const {raw, code, name} = parseKeyword(keyword);
    return state.indexByCode.get(code)
        || state.indexByName.get(raw)
        || state.indexByName.get(name)
        || pickStartsWith(state.rows, raw)
        || pickStartsWith(state.rows, name)
        || pickStartsWith(state.rows, code)
        || null;
  }

  function placeTooltipNearMouse(evt){
    const rect = chartWrap.getBoundingClientRect();
    const mouseX = evt.clientX - rect.left;
    const mouseY = evt.clientY - rect.top;
    const offsetX = 14, offsetY = 12;
    tooltip.style.display = 'block';
    tooltip.style.visibility = 'hidden';
    const tw = tooltip.offsetWidth || 160;
    const th = tooltip.offsetHeight || 80;
    let left = mouseX + offsetX;
    let top  = mouseY - th - offsetY;
    if(left + tw > rect.width - 6){ left = mouseX - tw - 8; }
    if(top < 6){ top = mouseY + 12; }
    if(left < 6) left = 6;
    if(top  > rect.height - th - 6) top = rect.height - th - 6;
    tooltip.style.left = left + 'px';
    tooltip.style.top  = top  + 'px';
    tooltip.style.visibility = 'visible';
  }

  function renderCombo(stock){
    const svg = d3.select('#comboChart');
    const node = svg.node(); if(!node) return;
    const W = node.clientWidth || node.parentNode.clientWidth || 960;
    const H = node.clientHeight || 380;
    svg.attr('viewBox', `0 0 ${W} ${H}`);

    const margin = { top:10, right:60, bottom:36, left:56 };
    const w = W - margin.left - margin.right;
    const h = H - margin.top - margin.bottom;

    const root = svg.selectAll('g.root').data([null]).join('g')
      .attr('class','root')
      .attr('transform',`translate(${margin.left},${margin.top})`);

    const data = (stock && stock.series) ? stock.series.filter(d=>d.amount!==null || d.mom!==null || d.yoy!==null) : [];
    if(!data.length){
      root.selectAll('*').remove();
      const hint = $('#comboHint'); if(hint) hint.style.display='none';
      tooltip.style.display='none';
      return;
    }

    const months = data.map(d=>d.ym);
    const x = d3.scaleBand().domain(months).range([0,w]).padding(0.15);

    const maxAmt = d3.max(data, d=>d.amount||0) || 0;
    const yR = d3.scaleLinear().domain([0, Math.max(1, maxAmt)*1.1]).nice().range([h,0]);

    const maxPct = d3.max(data, d=>Math.max(Math.abs(d.mom||0), Math.abs(d.yoy||0))) || 10;
    const yL = d3.scaleLinear().domain([-maxPct*1.2, maxPct*1.2]).nice().range([h,0]);

    // 只顯示三個座標
    const latest = months[months.length-1];
    const ticksWanted = [ latest, shiftYM(latest, -12), shiftYM(latest, -24) ]
                          .filter(v => months.indexOf(v) !== -1);

    const xAxis = (sel)=> sel.call(
      d3.axisBottom(x).tickValues(ticksWanted).tickFormat(ym=>fmtYM(ym)).tickSizeOuter(0)
    );
    const yAxisL = (sel)=> sel.call(d3.axisLeft(yL).ticks(6).tickFormat(d=>d+'%'));
    const yAxisR = (sel)=> sel.call(d3.axisRight(yR).ticks(6).tickFormat(fmtMoney));

    root.selectAll('.x.axis').data([null]).join('g')
      .attr('class','x axis')
      .attr('transform',`translate(0,${h})`).call(xAxis);
    root.selectAll('.y.axis-left').data([null]).join('g')
      .attr('class','y axis-left').call(yAxisL);
    root.selectAll('.y.axis-right').data([null]).join('g')
      .attr('class','y axis-right')
      .attr('transform',`translate(${w},0)`).call(yAxisR);

    root.selectAll('.zero-line').data([0]).join('line')
      .attr('class','zero-line')
      .attr('x1',0).attr('x2',w)
      .attr('y1',yL(0)).attr('y2',yL(0));

    // 柱
    const bars = root.selectAll('rect.bar').data(data, d=>d.ym);
    bars.enter().append('rect')
      .attr('class','bar bar-fixed')
      .attr('x', d=>x(d.ym))
      .attr('width', x.bandwidth())
      .attr('y', h)
      .attr('height', 0)
      .merge(bars)
      .transition().duration(450)
      .attr('x', d=>x(d.ym))
      .attr('width', x.bandwidth())
      .attr('y', d=>yR(d.amount||0))
      .attr('height', d=>h - yR(d.amount||0));
    bars.exit().remove();

    // Tooltip 事件
    root.selectAll('rect.bar')
      .on('mousemove', (evt, d)=>{
        const ym = fmtYM(d.ym);
        tooltip.innerHTML = (
          '<div><b>'+ norm(stock.code) +' '+ norm(stock.name) +'</b>｜'+ ym +'</div>'+
          '<div>合併營收：<b>'+ fmtMoney(d.amount || 0) +'</b></div>'+
          '<div>月增率 (MoM)：<b>'+ fmtPct(d.mom) +'</b></div>'+
          '<div>年增率 (YoY)：<b>'+ fmtPct(d.yoy) +'</b></div>'
        );
        placeTooltipNearMouse(evt);
      })
      .on('mouseleave', ()=>{ tooltip.style.display='none'; });

    // 折線
    const lineL = d3.line().defined(v=>v!=null)
      .x((_,i)=> x(months[i]) + x.bandwidth()/2)
      .y(v=>yL(v))
      .curve(d3.curveMonotoneX);

    root.selectAll('path.line-mom').data([data.map(d=>d.mom)])
      .join('path').attr('class','line-mom').attr('d', lineL);
    root.selectAll('path.line-yoy').data([data.map(d=>d.yoy)])
      .join('path').attr('class','line-yoy').attr('d', lineL);

    const hint = $('#comboHint'); if(hint) hint.style.display='none';
  }

  async function boot(){
    try{ await loadRevenue(); }catch(err){ console.error(err); return; }
    // 點「查詢」渲染
    const btn = $('#runBtn');
    if(btn){ btn.addEventListener('click', ()=>{
      const kw = $('#stockInput') ? $('#stockInput').value.trim() : '';
      const stock = findStock(kw);
      renderCombo(stock);
    }); }
    // 視窗大小變更也重繪
    window.addEventListener('resize', ()=>{
      const kw = $('#stockInput') ? $('#stockInput').value.trim() : '';
      const stock = findStock(kw);
      renderCombo(stock);
    });
    // 若載入時輸入框已有值，也先畫一次（避免空白畫面）
    const initKw = $('#stockInput') ? $('#stockInput').value.trim() : '';
    if(initKw){ const stock = findStock(initKw); renderCombo(stock); }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
