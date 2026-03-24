
/* treemap-override.js
 * Override Treemap downstream data to use H~J (downstreamHJ)
 */
(function(){
  function safeOverride(){
    if(!window.upstreamAC || !window.downstreamHJ){ setTimeout(safeOverride,120); return; }
    const oldRun = window.handleRun;
    if(!oldRun){ setTimeout(safeOverride,120); return; }
    window.handleRun = function(){
      const raw=document.querySelector('#stockInput').value;
      const month=document.querySelector('#monthSelect')?.value||'';
      const metric=document.querySelector('#metricSelect')?.value||'MoM';
      const colorMode=document.querySelector('#colorMode')?.value||'redPositive';
      let codeKey=raw.trim().replace(/\s+/g,'');
      const upstreamEdges = window.upstreamAC.filter(l=>l.down===codeKey);
      const downstreamEdges = window.downstreamHJ.filter(l=>l.up===codeKey);
      const rowSelf = window.byCode?.get(codeKey);
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
  safeOverride();
})();
