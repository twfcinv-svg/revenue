// rev_combo_plugin.js — 不改動原 app.js：包裹 handleRun 並繪製「合併營收（柱）＋ MoM / YoY（線）」
(function(){
  function fmtYM(ym){ return ym ? (ym.slice(0,4)+"/"+ym.slice(4,6)) : ''; }
  function formatAmount(v){ if(v==null||!isFinite(v)) return '—'; const a=Math.abs(v); if(a>=1e12) return (v/1e12).toFixed(1)+'兆'; if(a>=1e8) return (v/1e8).toFixed(1)+'億'; if(a>=1e6) return (v/1e6).toFixed(1)+'M'; if(a>=1e3) return (v/1e3).toFixed(1)+'K'; return String(Math.round(v)); }
  function getRevenueValue(row, ym){
    try{
      const col = (COL_MAP && COL_MAP[ym] && COL_MAP[ym].Rev) ? COL_MAP[ym].Rev : null;
      if(!col) return null; let v = row[col]; if(v==null || v==='') return null;
      let mul = 1; if (/(\(|（)\s*(千|仟)\s*(\)|）)/.test(String(col))) mul = 1000;
      if (typeof v === 'string') { v = v.replace(/[ ,，]/g,'').replace(/[^\d\.\-]/g,'').trim(); }
      v = Number(v) * mul; return Number.isFinite(v) ? v : null;
    }catch(_){ return null; }
  }
  function renderRevComboChart(selfRow){
    const hint = document.getElementById('revComboHint');
    const svg  = d3.select('#revComboSvg');
    if(!svg.node()) return;
    svg.selectAll('*').remove();
    const seq = (typeof months!=='undefined' && Array.isArray(months)) ? months.slice().sort((a,b)=>a.localeCompare(b)) : [];
    const data = seq.map(ym => ({
      ym,
      rev: getRevenueValue(selfRow, ym),
      mom: (typeof getMetricValue==='function') ? getMetricValue(selfRow, ym, 'MoM') : null,
      yoy: (typeof getMetricValue==='function') ? getMetricValue(selfRow, ym, 'YoY') : null
    })).filter(d => d.rev!=null || d.mom!=null || d.yoy!=null);
    if (data.length === 0){ if(hint) hint.textContent='沒有可繪製的合併營收 / 月增 / 年增資料。'; return; } else { if(hint) hint.textContent=''; }

    const wrap = document.getElementById('revComboWrap');
    const W = Math.max(320, wrap.clientWidth);
    const H = Math.max(180, parseInt(getComputedStyle(wrap).height) || 320);
    const margin = { top: 10, right: 56, bottom: 26, left: 42 };
    const innerW = Math.max(1, W - margin.left - margin.right);
    const innerH = Math.max(1, H - margin.top - margin.bottom);

    svg.attr('width', W).attr('height', H);
    const g = svg.append('g').attr('transform', `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand().domain(data.map(d=>d.ym)).range([0, innerW]).paddingInner(0.15).paddingOuter(0.05);

    const pctVals = data.flatMap(d => [d.mom, d.yoy]).filter(v => v!=null && isFinite(v));
    const pad = 4; const pctMin = (pctVals.length? d3.min(pctVals) : -10) - pad; const pctMax = (pctVals.length? d3.max(pctVals) :  10) + pad;
    const yLeft = d3.scaleLinear().domain([pctMin, pctMax]).range([innerH, 0]).nice();

    const revVals = data.map(d=>d.rev).filter(v => v!=null && isFinite(v));
    const yRight = d3.scaleLinear().domain([0, (revVals.length? d3.max(revVals) : 0) * 1.08 + 1]).range([innerH, 0]).nice();

    g.append('g').attr('class','grid')
      .call(d3.axisLeft(yLeft).ticks(5).tickSize(-innerW).tickFormat(''))
      .selectAll('line').attr('opacity', 0.5);

    g.selectAll('.bar').data(data).enter().append('rect')
      .attr('class','bar')
      .attr('x', d => x(d.ym))
      .attr('y', d => yRight(d.rev ?? 0))
      .attr('width', Math.max(1, x.bandwidth()))
      .attr('height', d => innerH - yRight(d.rev ?? 0))
      .attr('fill', '#3b82f6');

    const lineGen = d3.line().defined(d => d!=null && isFinite(d[1]))
      .x((d) => x(d[0]) + x.bandwidth()/2)
      .y((d) => yLeft(d[1]));

    g.append('path').attr('class','line').attr('stroke','#f59e0b').attr('stroke-width',2)
      .attr('d', lineGen(data.map(d => [d.ym, d.yoy])));

    g.append('path').attr('class','line').attr('stroke','#06b6d4').attr('stroke-width',2)
      .attr('d', lineGen(data.map(d => [d.ym, d.mom])));

    const xAxis = d3.axisBottom(x)
      .tickValues(x.domain().filter((_,i) => data.length>36 ? i%3===0 : i%2===0))
      .tickFormat(ym => fmtYM(ym));

    g.append('g').attr('class','axis').attr('transform', `translate(0,${innerH})`).call(xAxis);
    g.append('g').attr('class','axis').call(d3.axisLeft(yLeft).ticks(6).tickFormat(v => v.toFixed(0)+'%'));
    g.append('g').attr('class','axis').attr('transform', `translate(${innerW},0)`).call(d3.axisRight(yRight).ticks(6).tickFormat(formatAmount));

    // tooltip
    let tip = d3.select('#revComboWrap').select('.tooltip');
    if (tip.empty()) tip = d3.select('#revComboWrap').append('div').attr('class','tooltip').style('opacity',0);
    const hoverRect = g.append('rect').attr('fill','transparent').attr('width',innerW).attr('height',innerH);
    hoverRect.on('mousemove', function(event){
      const [mx] = d3.pointer(event, this);
      const idx = Math.max(0, Math.min(data.length-1, Math.round((mx - x.step()/2)/x.step())));
      const d = data[idx]; if(!d) return;
      const html = `<div style="font-weight:600; margin-bottom:4px;">${fmtYM(d.ym)}</div>
        <div>合併營收：<b>${formatAmount(d.rev)}</b></div>
        <div>月增率 MoM：<b>${(d.mom!=null? d.mom.toFixed(1)+'%':'—')}</b></div>
        <div>年增率 YoY：<b>${(d.yoy!=null? d.yoy.toFixed(1)+'%':'—')}</b></div>`;
      tip.html(html).style('left', (d3.pointer(event, this)[0] + margin.left + 8)+'px')
         .style('top',  (d3.pointer(event, this)[1] + margin.top  + 8)+'px')
         .style('opacity',1);
    }).on('mouseleave', ()=> tip.style('opacity',0));
  }

  function wrapHandleRun(){
    if (typeof window.handleRun !== 'function') return false;
    const old = window.handleRun;
    window.handleRun = function(){
      old.apply(this, arguments);
      try{
        const raw = document.querySelector('#stockInput')?.value || '';
        let codeKey = (typeof normCode==='function') ? normCode(raw) : String(raw).trim();
        let rowSelf = (typeof byCode!=='undefined' && byCode.get) ? byCode.get(codeKey) : null;
        if(!rowSelf){
          const nameQ = (typeof normText==='function') ? normText(raw) : String(raw).trim();
          if (typeof byName!=='undefined' && byName.get) rowSelf = byName.get(nameQ);
          if(!rowSelf && Array.isArray(window.revenueRows)){
            rowSelf = revenueRows.find(r => (typeof normText==='function' ? normText(r['名稱']||r['公司名稱']||r['證券名稱']||'') : (r['名稱']||r['公司名稱']||r['證券名稱']||''))
              .startsWith(nameQ));
          }
        }
        if(rowSelf) renderRevComboChart(rowSelf);
      }catch(e){ console.error('[rev-combo] post handleRun error', e); }
    };
    return true;
  }

  function init(){
    // 延遲等待 app.js 初始化完成
    let retries = 0; const timer = setInterval(()=>{
      if (typeof d3!=='undefined' && typeof months!=='undefined' && typeof COL_MAP!=='undefined' && wrapHandleRun()){
        clearInterval(timer);
      } else if (++retries > 80){ // ~8秒
        clearInterval(timer);
      }
    }, 100);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
