// combo-chart.js updated with bar hover tooltip
(function(){
  const $ = (sel) => document.querySelector(sel);
  const tooltip = document.createElement('div');
  tooltip.className = 'tooltip';
  tooltip.style.display = 'none';
  document.body.appendChild(tooltip);

  function norm(s){ return String(s||'').trim(); }
  function fmtYM(ym){ return ym.slice(0,4)+'-'+ym.slice(4,6); }
  function fmtPct(v){ return (v==null)?'': d3.format('+.0f')(v)+'%'; }
  function fmtMoney(val){ if(val==null) return ''; const si=d3.format('.2s'); return si(val*1000).replace('G','B'); }

  document.addEventListener('DOMContentLoaded',()=>{
    window.renderCombo = function(stock){
      const svg = d3.select('#comboChart');
      const node = svg.node(); if(!node) return;
      const W = node.clientWidth || 960;
      const H = node.clientHeight || 380;
      svg.attr('viewBox',`0 0 ${W} ${H}`);
      const margin={top:10,right:60,bottom:36,left:56};
      const w=W-margin.left-margin.right;
      const h=H-margin.top-margin.bottom;
      const data=stock.series;
      const months=data.map(d=>d.ym);
      const x=d3.scaleBand().domain(months).range([0,w]).padding(0.15);
      const yR=d3.scaleLinear().domain([0,d3.max(data,d=>d.amount||0)*1.1]).range([h,0]);
      const maxPct=d3.max(data,d=>Math.max(Math.abs(d.mom||0),Math.abs(d.yoy||0)))||0;
      const yL=d3.scaleLinear().domain([-maxPct*1.2,maxPct*1.2]).range([h,0]);
      const root=svg.selectAll('g.root').data([null]).join('g').attr('class','root').attr('transform',`translate(${margin.left},${margin.top})`);

      root.selectAll('.x.axis').data([null]).join('g').attr('class','x axis').attr('transform',`translate(0,${h})`)
        .call(d3.axisBottom(x).tickFormat(fmtYM));
      root.selectAll('.y.axis-left').data([null]).join('g').attr('class','y axis-left')
        .call(d3.axisLeft(yL).tickFormat(d=>d+'%'));
      root.selectAll('.y.axis-right').data([null]).join('g').attr('class','y axis-right')
        .attr('transform',`translate(${w},0)`).call(d3.axisRight(yR).tickFormat(fmtMoney));

      const bars=root.selectAll('rect.bar').data(data, d=>d.ym);
      bars.enter().append('rect').attr('class','bar bar-fixed')
        .attr('x',d=>x(d.ym)).attr('width',x.bandwidth()).attr('y',h).attr('height',0)
        .on('mousemove',(evt,d)=>{
          tooltip.style.display='block';
          tooltip.innerHTML=`<div><b>${stock.code} ${stock.name}</b>｜${fmtYM(d.ym)}</div>
            <div>合併營收：<b>${fmtMoney(d.amount)}</b></div>
            <div>月增率 (MoM)：<b>${fmtPct(d.mom)}</b></div>
            <div>年增率 (YoY)：<b>${fmtPct(d.yoy)}</b></div>`;
          tooltip.style.left = evt.clientX + 'px';
          tooltip.style.top = evt.clientY + 'px';
        })
        .on('mouseleave',()=> tooltip.style.display='none')
        .merge(bars)
        .transition().duration(400)
        .attr('y',d=>yR(d.amount||0)).attr('height',d=>h-yR(d.amount||0));

      root.selectAll('path.line-mom').data([data.map(d=>d.mom)]).join('path')
        .attr('class','line-mom')
        .attr('d',d3.line().x((v,i)=>x(months[i])+x.bandwidth()/2).y(v=>yL(v)));
      root.selectAll('path.line-yoy').data([data.map(d=>d.yoy)]).join('path')
        .attr('class','line-yoy')
        .attr('d',d3.line().x((v,i)=>x(months[i])+x.bandwidth()/2).y(v=>yL(v)));
    }
  });
})();