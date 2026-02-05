/* d3-treemap-autofont.js | 讓 D3 Treemap 的文字隨方塊面積自動縮放
 * 適用：你的 treemap 使用 D3 產出 <g> 內包含 <rect> 與 <text>，並在元素上綁了 d3 的 datum（含 x0,y0,x1,y1）。
 * 特色：
 *  - 自動掃描 #upTreemap 與 #downTreemap 兩個 SVG
 *  - 以節點像素面積占比分 5 段字級 (XL/L/M/S/XS)
 *  - XS 僅顯示代號；更小者直接隱藏 label（用 tooltip 看）
 *  - 綁 MutationObserver：當 treemap 重新繪製時自動重算
 */
(function(){
  var SVG_IDS = ['#upTreemap', '#downTreemap'];

  function bucketByRatio(r){ if(r>=0.08) return 'XL'; if(r>=0.03) return 'L'; if(r>=0.015) return 'M'; if(r>=0.006) return 'S'; return 'XS'; }
  function fontByBucket(b, isUpper){ var add=isUpper?2:0; switch(b){case 'XL':return 18+add;case 'L':return 16+add;case 'M':return 14+add;case 'S':return 12+add;default:return 10+add;} }
  function codeOnlyFrom(name){ if(!name) return ''; var m=/^\s*(\d{3,4})\b/.exec(String(name)); return m?m[1]:String(name); }

  function hasNodeDatum(el){ var d=el.__data__; return d && typeof d.x0==='number' && typeof d.y0==='number' && typeof d.x1==='number' && typeof d.y1==='number'; }
  function nodeArea(d){ var w=Math.max(0,d.x1-d.x0), h=Math.max(0,d.y1-d.y0); return w*h; }

  function scanNodes(svg){
    var nodes = [];
    // 掃描所有元素，找出綁有 treemap 布局座標的 datum
    var walker = document.createTreeWalker(svg, NodeFilter.SHOW_ELEMENT, null);
    var el; while(el = walker.nextNode()){
      if(hasNodeDatum(el)) nodes.push(el);
    }
    return nodes;
  }

  function findTextWithin(el){
    // 優先找同一 <g> 下面的文字；找不到就找子孫任何 <text>
    if(el.tagName && el.tagName.toLowerCase()==='g'){
      var t1 = el.querySelectorAll('text');
      if(t1 && t1.length) return Array.from(t1);
    }
    var p = el; while(p && p.nodeType===1 && p.tagName.toLowerCase()!=='svg'){
      var inSameG = p.querySelectorAll ? p.querySelectorAll('text') : null;
      if(inSameG && inSameG.length) return Array.from(inSameG);
      p = p.parentNode;
    }
    return [];
  }

  function applyForOneSvg(svg){
    var bbox = svg.getBoundingClientRect();
    var totalArea = Math.max(1, bbox.width * bbox.height);
    var nodeEls = scanNodes(svg);

    // 父層：有 children 的視為類股；無 children 視為個股
    var parents = nodeEls.filter(function(el){ var d=el.__data__; return d && d.children && d.children.length; });
    var leaves  = nodeEls.filter(function(el){ var d=el.__data__; return d && (!d.children || !d.children.length); });

    // 類股字級（以像素面積占整體畫布比）
    parents.forEach(function(el){
      var d=el.__data__; var r = nodeArea(d)/totalArea; var b=bucketByRatio(r); var f=fontByBucket(b, true);
      findTextWithin(el).forEach(function(t){ t.style.fontSize = f+'px'; t.style.lineHeight = Math.round(f*1.1)+'px'; t.style.pointerEvents='none'; });
    });

    // 個股字級（以各自面積占整體比）
    leaves.forEach(function(el){
      var d=el.__data__; var r = nodeArea(d)/totalArea; var b=bucketByRatio(r); var f=fontByBucket(b, false);
      var verySmall = (b==='XS' && r<0.0012); // 門檻可調
      var texts = findTextWithin(el);
      texts.forEach(function(t){
        if(verySmall){ t.style.display='none'; return; }
        t.style.display='';
        t.style.fontSize = f+'px';
        t.style.lineHeight = Math.round(f*1.1)+'px';
        t.style.pointerEvents='none';
        // XS 僅顯示代號（若文字包含代號）
        if(b==='XS'){
          var raw = t.textContent || '';
          var code = codeOnlyFrom(raw);
          if(code && code.length<=4) t.textContent = code; // 只留下代號
        }
      });
    });
  }

  function boot(){
    SVG_IDS.forEach(function(id){
      var svg = document.querySelector(id);
      if(!svg) return;
      // 初始執行
      applyForOneSvg(svg);
      // 監看 treemap 重繪
      var mo = new MutationObserver(function(){ applyForOneSvg(svg); });
      mo.observe(svg, { childList:true, subtree:true, characterData:true });
      // 視窗縮放也重算
      window.addEventListener('resize', function(){ applyForOneSvg(svg); });
    });
  }

  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
