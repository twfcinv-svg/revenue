/* treemap-autofont-hook.js | 零侵入：自動攔截 chart.setOption，套用 applyAutoLabelSizing */
(function(){
  if(!(window.echarts && typeof echarts.init==='function')) return;
  var _init = echarts.init;
  echarts.init = function(dom, theme, opts){
    var chart = _init.call(echarts, dom, theme, opts);
    var _set = chart.setOption.bind(chart);
    chart.setOption = function(option){
      try { if (window.applyAutoLabelSizing) window.applyAutoLabelSizing(option); } catch(e) {}
      return _set.apply(null, arguments);
    };
    return chart;
  };
})();
