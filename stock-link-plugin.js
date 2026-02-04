/* stock-link-plugin.js | 為 D3 treemap 加上「點個股 → 開啟 FBS」
 * 使用方式：在 index.html 內、app.js 後面再引入：
 *   <script src="stock-link-plugin.js?v='+VER+'"></script>
 * 需求：D3 v6+（你的 index 已有 <script src="https://cdn.jsdelivr.net/npm/d3@7"></script>）
 */
(function(){
  const FBS_BASE = 'https://www.fbs.com.tw/MKT/Index?name=%EF%BC%AB%E7%B7%9A%E5%9C%96&stock='; // Ｊ線圖 + 代號
  const SVG_SELECTORS = ['#upTreemap', '#downTreemap'];

  function normText(x){ return (x==null? '' : String(x)).replace(/[\u3000\s]+/g,'').trim(); }

  function extractCodeFromDatum(d){
    if (!d) return null;
    // 嘗試常見欄位
    let code = d.data && (d.data['代號'] || d.data['股票代號'] || d.data['證券代號'] || d.data.code || d.data.ticker);
    if (!code) {
      // 從名稱最前面的四位數解析："2330 台積電"
      const nm = (d.data && (d.data.name || d.data['名稱'])) || '';
      const m = /^\s*(\d{4})\b/.exec(nm);
      if (m) code = m[1];
    }
    code = normText(code);
    return code || null;
  }

  function findDatumWithCode(node){
    while (node && node.nodeType === 1 && node.tagName.toLowerCase() !== 'svg') {
      const sel = d3.select(node);
      const datum = sel.datum();
      const code = extractCodeFromDatum(datum);
      if (code) return { datum, code };
      node = node.parentNode; // 往上找直到 svg
    }
    return null;
  }

  function handleClick(e){
    const hit = findDatumWithCode(e.target);
    if (!hit) return; // 非個股或找不到代號就忽略
    const url = FBS_BASE + encodeURIComponent(String(hit.code));
    window.open(url, '_blank');
  }

  function markPointerCursor(){
    SVG_SELECTORS.forEach(sel => {
      const svg = document.querySelector(sel);
      if (!svg) return;
      // 給所有 rect 指針游標；若只想標示葉節點，可在 app.js 製作專屬 class 再套這裡
      svg.querySelectorAll('rect').forEach(r => r.style.cursor = 'pointer');
    });
  }

  function boot(){
    SVG_SELECTORS.forEach(sel => {
      const svg = document.querySelector(sel);
      if (svg) svg.addEventListener('click', handleClick);
    });
    markPointerCursor();

    // 監看 treemap 重新繪製時自動補上指針游標
    SVG_SELECTORS.forEach(sel => {
      const svg = document.querySelector(sel);
      if (!svg) return;
      const mo = new MutationObserver(() => markPointerCursor());
      mo.observe(svg, { childList: true, subtree: true });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
