
/* treemap-override.js — C2 最終成功版 */
(function(){
  function wait(){
    if(!window.renderTreemap||!window.byCode){ return setTimeout(wait,120); }
    if(!window.upstreamAC||!window.downstreamHJ){ return setTimeout(wait,120); }

    window.handleRun=function(){
      const raw=document.querySelector('#stockInput').value;
      const codeKey=normCode(raw);
      const month=document.querySelector('#monthSelect')?.value;
      const metric=document.querySelector('#metricSelect')?.value;
      const colorMode=document.querySelector('#colorMode')?.value||'redPositive';
      const rowSelf=byCode.get(codeKey);
      if(!rowSelf){ alert('找不到此代號/名稱'); return; }

      const upstreamEdges=upstreamAC.filter(x=>x.down===codeKey);
      const downstreamEdges=downstreamHJ.filter(x=>x.up===codeKey);

      requestAnimationFrame(()=>{
        renderResultChip(rowSelf,month,metric,colorMode);
        renderTreemap('upTreemap','upHint',upstreamEdges,'上游代號',month,metric,colorMode);
      });
      requestAnimationFrame(()=>{
        renderTreemap('downTreemap','downHint',downstreamEdges,'下游代號',month,metric,colorMode);
      });
    };
  } wait();
})();
