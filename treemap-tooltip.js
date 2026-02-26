
/* treemap-tooltip.js  讓上方 Treemap 的個股格子支援「模式 A：跟著游標漂浮」的資訊卡
 * 需求：滑到 Treemap 的「葉節點」格子時，顯示：代碼＋名稱＋目前指標值（MoM/YoY）
 * 相依：D3 v7、頁面已有 #upTreemap、#downTreemap，且外層容器 .treemap-wrap 存在。
 * 用法：在 index.html 於 app.js / app_query_link_B.js 載入之後，再載入本檔。
 */
(function(){
  const $ = (s)=> document.querySelector(s);
  const $$closest = (el, sel) => (el && el.closest) ? el.closest(sel) : null;

  // 建立 tooltip（掛到 .treemap-wrap 裡，避免受全站 CSS 影響）
  function createTooltipHost(svg){
    const wrap = $$closest(svg, '.treemap-wrap') || svg.parentNode || document.body;
    const style = getComputedStyle(wrap);
    if(style.position === 'static') wrap.style.position = 'relative';
    const tip = document.createElement('div');
    tip.className = 'treemap-tooltip';
    Object.assign(tip.style, {
      position: 'absolute', pointerEvents: 'none', display: 'none',
      padding: '8px 10px', fontSize: '12px', lineHeight: '1.4',
      color: '#fff', background: 'rgba(0,0,0,0.82)',
      border: '1px solid rgba(255,255,255,0.18)', borderRadius: '8px',
      boxShadow: '0 4px 14px rgba(0,0,0,.35)', zIndex: 1000, whiteSpace: 'nowrap'
    });
    wrap.appendChild(tip);
    return { wrap, tip };
  }

  function getMetricLabel(){
    const sel = $('#metricSelect');
    if(!sel) return '指標';
    const opt = sel.options[sel.selectedIndex];
    // 你的 index.html 裡 option 已寫明顯示文字
    return (opt && opt.textContent) ? opt.textContent : sel.value || '指標';
  }

  function fmtPct(v){
    if(v==null || isNaN(v)) return '';
    const f = d3.format('+.1f');
    return f(+v) + '%';
  }

  function extractDatumInfo(d){
    // 嘗試從 treemap node 或自訂欄位抓資訊（盡量容錯）
    const src = d && (d.data || d); // d3.treemap 的 node 會有 data
    const code = src && (src.code || src.個股 || src.stock || src.symbol || '');
    const name = src && (src.name || src.名稱 || src.title || '');
    // 目前 treemap 的數值通常是百分比（MoM/YoY），常放在 value 或 YoY/MoM 欄位
    let val = (d && d.value!=null) ? d.value : (src && (src.YoY!=null ? src.YoY : src.MoM));
    if(val==null && src && src.value!=null) val = src.value;
    return { code, name, val };
  }

  function placeTooltipNearMouse(evt, host){
    const rect = host.wrap.getBoundingClientRect();
    const mx = evt.clientX - rect.left;  // 容器內座標
    const my = evt.clientY - rect.top;
    const offX = 14, offY = 12;          // 右上角偏移

    const tip = host.tip;
    tip.style.display = 'block';
    tip.style.visibility = 'hidden';
    const tw = tip.offsetWidth || 160;
    const th = tip.offsetHeight || 80;

    let left = mx + offX;
    let top  = my - th - offY;
    if(left + tw > rect.width - 6) left = mx - tw - 8; // 右界
    if(top < 6) top = my + 12;                         // 上界
    if(left < 6) left = 6;
    if(top > rect.height - th - 6) top = rect.height - th - 6;

    tip.style.left = left + 'px';
    tip.style.top  = top  + 'px';
    tip.style.visibility = 'visible';
  }

  function bindOne(svgId){
    const svg = document.getElementById(svgId);
    if(!svg) return;
    const host = createTooltipHost(svg);

    const sel = d3.select(svg);

    function bind(){
      // 只綁「葉節點」：有 rect 且沒有 child g 的 group，或直接綁 rect 且有 data
      const cells = sel.selectAll('g').filter(function(){
        const g = d3.select(this);
        const hasRect = !g.select('rect').empty();
        const hasChildG = !g.select('g').empty();
        return hasRect && !hasChildG;
      });

      cells.style('pointer-events','all')
        .on('mousemove', function(evt, d){
          const info = extractDatumInfo(d || d3.select(this).datum());
          const title = (info.code?info.code+' ':'') + (info.name||'');
          const metric = getMetricLabel();
          host.tip.innerHTML = (
            '<div><b>'+ title +'</b></div>'+
            '<div>'+ metric +'：<b>'+ fmtPct(info.val) +'</b></div>'
          );
          placeTooltipNearMouse(evt, host);
        })
        .on('mouseleave', function(){ host.tip.style.display='none'; });

      // 也讓 rect 直接綁（以防 group 結構不同）
      sel.selectAll('rect')
        .style('pointer-events','all')
        .on('mousemove', function(evt, d){
          const info = extractDatumInfo(d || d3.select(this).datum());
          if(!info.code && !info.name && (this.parentNode)){
            // 嘗試從父層 g 中抓
            const pd = d3.select(this.parentNode).datum();
            const tmp = extractDatumInfo(pd);
            if(tmp.code||tmp.name||tmp.val!=null) Object.assign(info, tmp);
          }
          const title = (info.code?info.code+' ':'') + (info.name||'');
          const metric = getMetricLabel();
          host.tip.innerHTML = (
            '<div><b>'+ title +'</b></div>'+
            '<div>'+ metric +'：<b>'+ fmtPct(info.val) +'</b></div>'
          );
          placeTooltipNearMouse(evt, host);
        })
        .on('mouseleave', function(){ host.tip.style.display='none'; });
    }

    // 先綁一次
    bind();

    // 若 treemap 重繪（查詢或 RWD）→ 再次綁定
    const mo = new MutationObserver(()=>{ bind(); });
    mo.observe(svg, { childList:true, subtree:true });

    window.__rebindTreemapTooltip = window.__rebindTreemapTooltip || (()=>{ bind(); });
  }

  document.addEventListener('DOMContentLoaded', ()=>{
    bindOne('upTreemap');
    bindOne('downTreemap');

    const btn = $('#runBtn');
    if(btn) btn.addEventListener('click', ()=>{
      setTimeout(()=>{ if(window.__rebindTreemapTooltip) window.__rebindTreemapTooltip(); }, 50);
    });

    window.addEventListener('resize', ()=>{
      if(window.__rebindTreemapTooltip) window.__rebindTreemapTooltip();
    });
  });
})();
