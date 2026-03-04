
/* treemap-tooltip.js  v2
 * 修正：數值改「優先讀取格子上顯示的百分比文字」，確保與方塊一致（例如 +74.5%）。
 * 若該文字不存在，再回退到 node.value / data.YoY / data.MoM。
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
    // 往父節點找（treemap 常見用 parent 表示產業分群）
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

  // 從當前 cell 的文字中直接抓百分比，確保與方塊一致
  function getDisplayPctFromCell(node){
    if(!node) return null;
    const g = $$closest(node, 'g') || node;
    const texts = g.querySelectorAll('text');
    let buf = '';
    texts.forEach(t => { buf += ' ' + (t.textContent || ''); });
    const m = buf.match(/([+-]?\d+(?:\.\d+)?)\s*%/);
    if(m) return m[0].replace(/\s+/g,''); // 保留原有正負與%號
    return null;
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

    function renderTip(evt, datum, node){
      const nodeData = datum || {};
      const src = (nodeData.data || nodeData || {});
      const industry = getIndustry(nodeData) || '';
      const code = getCode(src);
      const name = getName(src);

      // 1) 先讀 cell 上顯示的百分比文字
      let pctStr = getDisplayPctFromCell(node);
      // 2) 萬一沒有，才回退採資料欄位
      if(!pctStr){
        const val = getVal(nodeData);
        pctStr = fmtPct(val);
      }

      host.tip.innerHTML = (
        (industry?('<div><b>'+industry+'</b></div>'):'') +
        '<div>'+ (code?code+' ':'') + (name||'') + '</div>'+
        '<div><b>'+ pctStr +'</b></div>'
      );
      placeTip(evt, host);
    }

    function bind(){
      const cells = sel.selectAll('g').filter(function(){
        const g = d3.select(this);
        const hasRect = !g.select('rect').empty();
        const hasChildG = !g.select('g').empty();
        return hasRect && !hasChildG;
      });

      cells.style('pointer-events','all')
        .on('mousemove', function(evt, d){ renderTip(evt, d || d3.select(this).datum(), this); })
        .on('mouseleave', function(){ host.tip.style.display='none'; });

      sel.selectAll('rect')
        .style('pointer-events','all')
        .on('mousemove', function(evt, d){
          let datum = d || d3.select(this).datum();
          if(!datum || (!datum.data && this.parentNode)){
            const pd = d3.select(this.parentNode).datum();
            if(pd) datum = pd;
          }
          renderTip(evt, datum, this);
        })
        .on('mouseleave', function(){ host.tip.style.display='none'; });
    }

    bind();
    const mo = new MutationObserver(()=>{ bind(); });
    mo.observe(svg, { childList:true, subtree:true });

    window.__rebindTreemapTooltip = window.__rebindTreemapTooltip || (()=>{ bind(); });
  }

  document.addEventListener('DOMContentLoaded', ()=>{
    bindOne('upTreemap');
    bindOne('downTreemap');

    const btn = $('#runBtn');
    if(btn) btn.addEventListener('click', ()=>{
      setTimeout(()=>{ if(window.__rebindTreemapTooltip) window.__rebindTreemapTooltip(); }, 60);
    });

    window.addEventListener('resize', ()=>{
      if(window.__rebindTreemapTooltip) window.__rebindTreemapTooltip();
    });
  });
})();
