
/* treemap-override-final.js */
(function(){
  function wait(){
    if(!window.handleRun){ return setTimeout(wait,100); }
    if(!window.upstreamAC || !window.downstreamHJ){ return setTimeout(wait,150); }
    const oldRun = window.handleRun;
    window.handleRun = function(){
      const raw=document.querySelector('#stockInput').value;
      const month=document.querySelector('#monthSelect').value;
      const metric=document.querySelector('#metricSelect').value;
      const colorMode=document.querySelector('#colorMode')?.value||'redPositive';
      const codeKey=raw.trim().replace(/\s+/g,'');
      const upstreamEdges   = window.upstreamAC.filter(x=>x.down===codeKey);
      const downstreamEdges = window.downstreamHJ.filter(x=>x.up===codeKey);
      const rowSelf = window.byCode.get(codeKey);
      if(!rowSelf){ alert('找不到此代號/名稱'); return; }
      requestAnimationFrame(()=>{
        window.renderResultChip(rowSelf,month,metric,colorMode);
        window.renderTreemap('upTreemap','upHint',upstreamEdges,'上游代號',month,metric,colorMode);
      });
      requestAnimationFrame(()=>{
        window.renderTreemap('downTreemap','downHint',downstreamEdges,'下游代號',month,metric,colorMode);
      });
    };
  }
  wait();
})();
