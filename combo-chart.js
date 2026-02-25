
/* combo-chart.js  個股合併營收（長條）＋ 月增率 / 年增率（雙折線） */
(function(){
  const $ = (sel) => document.querySelector(sel);
  const tooltip = document.createElement('div');
  tooltip.className = 'tooltip';
  tooltip.style.display = 'none';
  document.body.appendChild(tooltip);

  const state = { loaded: false, rows: [], columns: { code: null, name: null, industry: null }, months: { amount: [], mom: [], yoy: [] }, indexByCode: new Map(), indexByName: new Map() };

  function norm(s){ return String(s || '').trim(); }
  function detectColumns(headers){
    const col = { code: null, name: null, industry: null, amount: {}, mom: {}, yoy: {} };
    const reMonth = /^(20\d{2})(0[1-9]|1[0-2])/;
    headers.forEach(h => {
      const H = norm(h); if(!H) return;
      if(!col.code && /(個股|代號|股票代號|Code|Symbol)/i.test(H)) col.code = h;
      if(!col.name && /(名稱|公司|Name)/i.test(H)) col.name = h;
      if(!col.industry && /(產業|產業別|Industry)/i.test(H)) col.industry = h;
      const m = H.match(reMonth);
      if(m){
        const ym = m[1]+m[2];
        if(/(合併)?營收.*(仟|千|千元|金額|營收$)/i.test(H)) col.amount[ym] = h;
        else if(/(月變動|月增率|MoM)/i.test(H)) col.mom[ym] = h;
        else if(/(年成長|年增率|YoY)/i.test(H)) col.yoy[ym] = h;
      }
    });
    return col;
  }
  function parseNumber(v){ if(v==null||v==='') return null; if(typeof v==='number') return isFinite(v)?v:null; const s=String(v).replace(/[% ,，]/g,''); const n=parseFloat(s); return isFinite(n)?n:null; }

  async function loadRevenue(){
    if(state.loaded) return state;
    const res = await fetch('data.xlsx');
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetName = wb.SheetNames.find(n => /revenue/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
    if(!rows.length) throw new Error('Revenue 工作表為空');

    const headers = Object.keys(rows[0]);
    const col = detectColumns(headers);
    const months_amount = Object.keys(col.amount).sort();
    const months_mom = Object.keys(col.mom).sort();
    const months_yoy = Object.keys(col.yoy).sort();
    const months_all = Array.from(new Set([...months_amount, ...months_mom, ...months_yoy])).sort();

    const tidy = rows.map(r => {
      const code = norm(r[col.code]);
      const name = norm(r[col.name]);
      const ind = norm(r[col.industry]);
      const series = months_all.map(ym => ({ ym, amount: parseNumber(r[col.amount[ym]]), mom: parseNumber(r[col.mom[ym]]), yoy: parseNumber(r[col.yoy[ym]]) }))
        .filter(d => d.amount!==null || d.mom!==null || d.yoy!==null);
      return { code, name, industry: ind, series };
    }).filter(d => d.code || d.name);

    const byCode=new Map(), byName=new Map();
    for(const r of tidy){ if(r.code) byCode.set(r.code,r); if(r.name) byName.set(r.name,r); }
    state.loaded=true; state.rows=tidy; state.columns=col; state.months={ amount:months_amount, mom:months_mom, yoy:months_yoy, all:months_all }; state.indexByCode=byCode; state.indexByName=byName; return state;
  }

  function findStock(input){ const key=norm(input); if(!key) return null; return state.indexByCode.get(key) || state.indexByName.get(key) || null; }
  function fmtYM(ym){ return ym.slice(0,4)+'-'+ym.slice(4,6); }
  function fmtPct(v){ return (v==null)?'': d3.format('+.0f')(v)+'%'; }
  function fmtMoney(val){ if(val==null) return ''; const si=d3.format('.2s'); return si(val*1000).replace('G','B'); }

  function renderCombo(stock){
    const svg = d3.select('#comboChart'); const node = svg.node(); if(!node) return;
    const W = node.clientWidth || node.parentNode.clientWidth || 960; const H = node.clientHeight || 380; svg.attr('viewBox',`0 0 ${W} ${H}`);
    const margin={top:10,right:60,bottom:36,left:56}; const w=W-margin.left-margin.right; const h=H-margin.top-margin.bottom;
    const g=svg.selectAll('g.root').data([null]); const root=g.enter().append('g').attr('class','root').merge(g).attr('transform',`translate(${margin.left},${margin.top})`);

    const data=stock.series.filter(d=>d.amount!==null||d.mom!==null||d.yoy!==null); if(!data.length){ d3.select('#comboHint').style('display','none'); root.selectAll('*').remove(); return; }
    const months=data.map(d=>d.ym);

    const x=d3.scaleBand().domain(months).range([0,w]).padding(0.15);
    const maxAmt=d3.max(data,d=>d.amount||0)||0; const yR=d3.scaleLinear().domain([0, d3.max([1,maxAmt])*1.1]).nice().range([h,0]);
    const maxPct=d3.max(data,d=>Math.max(Math.abs(d.mom||0), Math.abs(d.yoy||0)))||0; const yL=d3.scaleLinear().domain([-maxPct*1.2,maxPct*1.2]).nice().range([h,0]);

    const xAxis=sel=>sel.call(d3.axisBottom(x).tickFormat(ym=>fmtYM(ym)).tickSizeOuter(0));
    const yAxisL=sel=>sel.call(d3.axisLeft(yL).ticks(6).tickFormat(d=>d+'%'));
    const yAxisR=sel=>sel.call(d3.axisRight(yR).ticks(6).tickFormat(fmtMoney));

    root.selectAll('.x.axis').data([null]).join('g').attr('class','x axis').attr('transform',`translate(0,${h})`).call(xAxis);
    root.selectAll('.y.axis-left').data([null]).join('g').attr('class','y axis-left').call(yAxisL);
    root.selectAll('.y.axis-right').data([null]).join('g').attr('class','y axis-right').attr('transform',`translate(${w},0)`).call(yAxisR);

    root.selectAll('.zero-line').data([0]).join('line').attr('class','zero-line').attr('x1',0).attr('x2',w).attr('y1',yL(0)).attr('y2',yL(0));

    // 柱固定紅色
    const bars=root.selectAll('rect.bar').data(data, d=>d.ym);
    bars.enter().append('rect').attr('class','bar bar-fixed')
      .attr('x',d=>x(d.ym)).attr('width',x.bandwidth()).attr('y',h).attr('height',0)
      .merge(bars)
      .transition().duration(450)
      .attr('x',d=>x(d.ym)).attr('width',x.bandwidth()).attr('y',d=>yR(d.amount||0)).attr('height',d=>h-yR(d.amount||0));
    bars.exit().remove();

    const lineL=d3.line().defined(d=>d!=null).x((_,i)=>x(months[i])+x.bandwidth()/2).y(v=>yL(v)).curve(d3.curveMonotoneX);
    root.selectAll('path.line-mom').data([data.map(d=>d.mom)]).join('path').attr('class','line-mom').attr('d',lineL);
    root.selectAll('path.line-yoy').data([data.map(d=>d.yoy)]).join('path').attr('class','line-yoy').attr('d',lineL);

    // 互動
    const hover=root.selectAll('rect.hover').data(data, d=>d.ym);
    hover.enter().append('rect').attr('class','hover').attr('fill','transparent')
      .attr('x',d=>x(d.ym)).attr('width',x.bandwidth()).attr('y',0).attr('height',h)
      .on('mousemove',(evt,d)=>{ const ym=fmtYM(d.ym); tooltip.style.display='block'; tooltip.innerHTML=`<div><b>${stock.code||''} ${stock.name||''}</b>｜${ym}</div><div>合併營收：<b>${fmtMoney(d.amount||0)}</b></div><div>月增率 (MoM)：<b>${fmtPct(d.mom)}</b></div><div>年增率 (YoY)：<b>${fmtPct(d.yoy)}</b></div>`; tooltip.style.left=(evt.clientX)+'px'; tooltip.style.top=(evt.clientY)+'px'; })
      .on('mouseleave',()=>{ tooltip.style.display='none'; });
    hover.attr('x',d=>x(d.ym)).attr('width',x.bandwidth()).attr('height',h);

    // 需求 2：刪除底部白字 => 直接隱藏
    d3.select('#comboHint').style('display','none');
  }

  function onSearch(){ const input=$('#stockInput'); if(!input) return; const k=input.value.trim(); if(!k) return; const stock=findStock(k); if(!stock){ d3.select('#comboHint').style('display','none'); return; } renderCombo(stock); }

  document.addEventListener('DOMContentLoaded', async ()=>{ try{ await loadRevenue(); }catch(e){ console.error(e); d3.select('#comboHint').style('display','none'); } const btn=$('#runBtn'); if(btn){ btn.addEventListener('click', ()=>setTimeout(onSearch,0)); } window.addEventListener('resize', ()=>{ const input=$('#stockInput'); if(!input) return; const stock=findStock(input.value.trim()); if(stock) renderCombo(stock); }); });
})();
