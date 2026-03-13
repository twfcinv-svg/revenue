/* industry-chain-panel.js (v9.2 align) */
(function(){
  const $  = (s)=> document.querySelector(s);
  const RAF = (fn)=> requestAnimationFrame(()=> requestAnimationFrame(fn));
  const cssNudge = (name)=>{ const v = getComputedStyle(document.documentElement).getPropertyValue(name).trim(); const n = parseInt(v,10); return Number.isFinite(n)? n : 0; };
  const pxRound = (x)=> Math.round(x * (window.devicePixelRatio||1)) / (window.devicePixelRatio||1);
  function setFoldToAnchors(){
    const wrap   = document.querySelector('#icp-fold-wrap');
    const scroll = document.querySelector('#icp-scroll');
    const btn    = document.querySelector('#icp-expander');
    const fade   = document.querySelector('#icp-fade');
    const head   = document.querySelector('#combo-section .section-head');
    const chartW = document.querySelector('#combo-section .chart-wrap');
    const legend = document.querySelector('#combo-section .legend-row');
    if(!wrap||!scroll||!btn||!head||!chartW) return;
    const wrapRect   = wrap.getBoundingClientRect();
    const headRect   = head.getBoundingClientRect();
    const chartRect  = chartW.getBoundingClientRect();
    const legendRect = legend ? legend.getBoundingClientRect() : null;
    const headCS  = getComputedStyle(head);
    const chartCS = getComputedStyle(chartW);
    const headBottom = headRect.top + (parseFloat(headCS.paddingTop)||0) + (parseFloat(headCS.height)||headRect.height);
    const chartBottom = chartRect.bottom + (parseFloat(chartCS.marginBottom)||0);
    const legendBottom = legendRect ? legendRect.bottom : -Infinity;
    const nTop  = cssNudge('--icp-nudge-top');
    const nBot  = cssNudge('--icp-nudge-bottom');
    const topAnchor    = pxRound(headBottom - wrapRect.top) + nTop;
    const bottomAnchor = pxRound(Math.max(chartBottom, legendBottom) - wrapRect.top) + nBot;
    const firstBox = scroll.querySelector('.icp-box');
    let innerTopOffset = 0; if(firstBox){ const fb = firstBox.getBoundingClientRect(); const sc = scroll.getBoundingClientRect(); innerTopOffset = pxRound(fb.top - sc.top); }
    const isExpanded = btn.getAttribute('aria-expanded') === 'true';
    const marginTop = Math.max(0, topAnchor - innerTopOffset); scroll.style.marginTop = marginTop + 'px';
    const targetHeight = Math.max(0, bottomAnchor - (marginTop + innerTopOffset));
    if(!isExpanded){ scroll.style.maxHeight = targetHeight + 'px'; if(fade) fade.style.display='block'; }
    else{ scroll.style.maxHeight = scroll.scrollHeight + 'px'; if(fade) fade.style.display='none'; }
    const btnTop = pxRound(bottomAnchor - (btn.offsetHeight/2)); btn.style.top = (btnTop<0?0:btnTop) + 'px';
  }
  function toggleFold(){ const btn=document.getElementById('icp-expander'); const expanded=btn.getAttribute('aria-expanded')==='true'; btn.setAttribute('aria-expanded', String(!expanded)); btn.textContent = expanded ? '＋' : '－'; RAF(setFoldToAnchors); }
  function installObservers(){
    const head=document.querySelector('#combo-section .section-head'); const chart=document.querySelector('#combo-section .chart-wrap'); const legend=document.querySelector('#combo-section .legend-row');
    if(head){ const ro1=new ResizeObserver(()=> RAF(setFoldToAnchors)); ro1.observe(head); }
    if(chart){ const ro2=new ResizeObserver(()=> RAF(setFoldToAnchors)); ro2.observe(chart); }
    if(legend){ const ro3=new ResizeObserver(()=> RAF(setFoldToAnchors)); ro3.observe(legend); }
    window.addEventListener('resize', ()=> RAF(setFoldToAnchors));
  }
  async function firstAlign(){ try{ if(document.fonts && document.fonts.ready) await document.fonts.ready; }catch(e){} RAF(setFoldToAnchors); }
  document.addEventListener('DOMContentLoaded', ()=>{ const btn=document.getElementById('icp-expander'); if(btn) btn.addEventListener('click', toggleFold); installObservers(); firstAlign(); });
})();
