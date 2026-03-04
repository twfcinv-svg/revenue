
/* treemap-tooltip.js  上下游 Treemap：游標停留於個股格子時，顯示「產業｜代碼 名稱｜表現(%)」
 * 內容格式（例）：
 *   封測設備
 *   7769 鴻勁
 *   +73.9%
 *
 * 設計：
 *  - Tooltip 以 .treemap-wrap 為定位容器，絕對定位跟著游標漂浮（模式 A）。
 *  - 產業、代碼、名稱、表現值皆做容錯：
 *     · code: data.code / 個股 / stock / symbol
 *     · name: data.name / 名稱 / title
 *     · industry: data.industry / 產業 / 產業別 / group / category；若無則向 parent.data.* 尋找
 *     · value: node.value 優先；否則 data.YoY / data.MoM；最後 data.value
 */
(function(){
  const $ = (s)=> document.querySelector(s);
  const $$closest = (el, sel) => (el && el.closest) ? el.closest(sel) : null;

  function ensureHost(svg){
    const wrap = $$closest(svg, '.treemap-wrap') || svg.parentNode || document.body;
    if(getComputedStyle(wrap).position === 'static') wrap.style.position = 'relative';
    let tip = wrap.querySelector(':scope > .treemap-tooltip');
    if(!tip){
      tip = document.createElement('div');
      tip.className = 'treemap-tooltip';
      Object.assign(tip.style, {
        position:'absolute', pointerEvents:'none', display:'none',
        padding:'8px 10px', fontSize:'12px', lineHeight:'1.5',
        color:'#fff', background:'rgba(0,0,0,0.82)',
        border:'1px solid rgba(255,255,255,0.18)', borderRadius:'8px',
        boxShadow:'0 4px 14px rgba(0,0,0,.35)', zIndex:1000, whiteSpace:'nowrap'
      });
      wrap.appendChild(tip);
    }
    return { wrap, tip };
  }

  function fmtPct(v){ if(v==null || isNaN(v)) return ''; return d3.format('+.1f')(+v) + '%'; }

  function getVal(d){
    if(d && typeof d.value === 'number') return d.value;
    const src = d && (d.data || d);
    if(src && typeof src.YoY === 'number') return src.YoY;
    if(src && typeof src.MoM === 'number') return src.MoM;
    if(src && typeof src.value === 'number') return src.value;
    return null;
  }

  function getCode(src){ return src.code || src['個股'] || src.stock || src.symbol || ''; }
  function getName(src){ return src.name || src['名稱'] || src.title || ''; }

  function getIndustry(d){
    const src = d && (d.data || d) || {};
    const keys = ['industry','產業','產業別','group','category'];
    for(const k of keys){ if(src[k]) return src[k]; }
    // 往父節點找（常見 treemap 分群）
    let p = d && d.parent;
    while(p){
      const pd = p.data || {};
      for(const k of ['industry','產業','產業別','group','category','name','title','label']){
        if(pd[k]) return pd[k];
      }
      p = p.parent;
    }
    return '';
  }

  function placeTip(evt, host){
    const rect = host.wrap.getBoundingClientRect();
    const mx = evt.clientX - rect.left;
    const my = evt.clientY - rect.top;
    const offX = 14, offY = 12; // 右上角偏移

    const tip = host.tip;
    tip.style.display = 'block';
    tip.style.visibility = 'hidden';
    const tw = tip.offsetWidth || 160, th = tip.offsetHeight || 72;

    let left = mx + offX;
    let top  = my - th - offY; // 優先放在游標上方

    // 邊界保護
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
    const host = ensureHost(svg);
    const sel = d3.select(svg);

    function renderTip(evt, datum){
      const nodeData = datum || {};
      const src = (nodeData.data || nodeData || {});
      const industry = getIndustry(nodeData) || '';
      const code = getCode(src);
      const name = getName(src);
      const val = getVal(nodeData);

      host.tip.innerHTML = (
        (industry?('<div><b>'+industry+'</b></div>'):'') +
        '<div>'+ (code?code+' ':'') + (name||'') + '</div>'+
        '<div><b>'+ fmtPct(val) +'</b></div>'
      );
      placeTip(evt, host);
    }

    function bind(){
      // 葉節點 group
      const cells = sel.selectAll('g').filter(function(){
        const g = d3.select(this);
        const hasRect = !g.select('rect').empty();
        const hasChildG = !g.select('g').empty();
        return hasRect && !hasChildG;
      });

      cells.style('pointer-events','all')
        .on('mousemove', function(evt, d){ renderTip(evt, d || d3.select(this).datum()); })
        .on('mouseleave', function(){ host.tip.style.display='none'; });

      // 防止不同結構：直接對所有 rect 再綁一次
      sel.selectAll('rect')
        .style('pointer-events','all')
        .on('mousemove', function(evt, d){
          let datum = d || d3.select(this).datum();
          if(!datum || (!datum.data && this.parentNode)){
            const pd = d3.select(this.parentNode).datum();
            if(pd) datum = pd;
          }
          renderTip(evt, datum);
        })
        .on('mouseleave', function(){ host.tip.style.display='none'; });
    }

    bind();
    const mo = new MutationObserver(()=>{ bind(); });
    mo.observe(svg, { childList:true, subtree:true });

    // 供外部在查詢/重繪後手動觸發
    window.__rebindTreemapTooltip = window.__rebindTreemapTooltip || (()=>{ bind(); });
  }

  document.addEventListener('DOMContentLoaded', ()=>{
    bindOne('upTreemap');
    bindOne('downTreemap');

    const btn = $('#runBtn');
    if(btn) btn.addEventListener('click', ()=>{
      // treemap 重繪後稍等再重綁
      setTimeout(()=>{ if(window.__rebindTreemapTooltip) window.__rebindTreemapTooltip(); }, 60);
    });

    window.addEventListener('resize', ()=>{
      if(window.__rebindTreemapTooltip) window.__rebindTreemapTooltip();
    });
  });
})();
