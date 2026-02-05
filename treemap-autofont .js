/* treemap-autofont.js | 讓 Treemap 標籤字體隨面積縮放（父=類股、子=個股）
 * 不用改資料結構，直接在 option.series[*].data 回填 fontSize/顯示策略。
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) { define([], factory); }
  else if (typeof module === 'object' && module.exports) { module.exports = factory(); }
  else { root.applyAutoLabelSizing = factory(); }
}(typeof self !== 'undefined' ? self : this, function () {
  function bucketByRatio(r) { if (r>=0.08) return 'XL'; if (r>=0.03) return 'L'; if (r>=0.015) return 'M'; if (r>=0.006) return 'S'; return 'XS'; }
  function fontByBucket(b, isUpper) { var add=isUpper?2:0; switch(b){case 'XL':return 18+add;case 'L':return 16+add;case 'M':return 14+add;case 'S':return 12+add;default:return 10+add;} }
  function codeOnlyFrom(name){ if(!name) return ''; var m=/^\s*(\d{3,4})\b/.exec(String(name)); return m?m[1]:String(name); }
  function nodeWeight(v){ return Array.isArray(v) ? (Number(v[0])||0) : (Number(v)||0); }
  function ensure(obj, path, val){ var cur=obj; for(var i=0;i<path.length-1;i++){ if(!cur[path[i]]) cur[path[i]]={}; cur=cur[path[i]]; } if(cur[path[path.length-1]]==null) cur[path[path.length-1]]=val; return cur[path[path.length-1]]; }

  function applyToSeries(series){
    if(!series||series.type!=='treemap'||!Array.isArray(series.data)) return;
    var totalParentW=0; series.data.forEach(function(p){ totalParentW+=nodeWeight(p.value); }); if(totalParentW<=0) totalParentW=1;

    series.data.forEach(function(p){
      var pw=nodeWeight(p.value), pr=pw/totalParentW, pb=bucketByRatio(pr), pf=fontByBucket(pb,true);
      ensure(p,['upperLabel'],{}); p.upperLabel.show=true; p.upperLabel.fontSize=pf; p.upperLabel.color=p.upperLabel.color||'#e5e7eb';
      ensure(p,['label'],{}); p.label.fontSize=pf; p.label.color=p.label.color||'#e5e7eb'; p.label.overflow=p.label.overflow||'truncate';

      var children=Array.isArray(p.children)?p.children:[]; if(!children.length) return; var ct=0; children.forEach(function(c){ ct+=nodeWeight(c.value); }); if(ct<=0) ct=1;
      p.children=children.map(function(c){
        var cw=nodeWeight(c.value), cr=cw/ct, cb=bucketByRatio(cr), cf=fontByBucket(cb,false), verySmall=(cb==='XS' && cw<(ct*0.004));
        c.label=c.label||{}; c.label.show=!verySmall; c.label.fontSize=cf; c.label.color=c.label.color||'#e5e7eb'; c.label.overflow=c.label.overflow||'truncate'; c.label.lineHeight=Math.round(cf*1.1);
        var orig=c.label.formatter; c.label.formatter=function(params){ if(cb==='XS') return codeOnlyFrom(params.name); if(typeof orig==='function') return orig(params); return params.name; };
        return c;
      });
    });
    series.labelLayout = series.labelLayout || function(){ return { hideOverlap: true }; };
  }

  function applyAutoLabelSizing(option){ if(!option) return option; var list=Array.isArray(option.series)?option.series:(option.series?[option.series]:[]); list.forEach(applyToSeries); return option; }
  return applyAutoLabelSizing;
}));
