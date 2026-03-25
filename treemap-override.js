
/* treemap-override.js — 終極強制版，下游只使用 H~J */

(function(){

  console.log("Treemap override — 終極強制版載入");

  function applyOverride() {

    // 先清掉 runBtn 上所有舊的事件（app.js 的 handleRun）
    const btn = document.querySelector("#runBtn");
    const newBtn = btn.cloneNode(true);
    btn.parentNode.replaceChild(newBtn, btn);

    console.log("🔧 已移除 runBtn 所有既有事件（包含 app.js 的 handleRun）");

    // 重新綁定你自己的 handleRun（永遠不會被覆蓋）
    newBtn.addEventListener("click", function(){

      const raw = document.querySelector("#stockInput").value;
      const codeKey = normCode(raw);

      const month = document.querySelector("#monthSelect")?.value;
      const metric = document.querySelector("#metricSelect")?.value;
      const colorMode = document.querySelector("#colorMode")?.value || "redPositive";

      const rowSelf = byCode.get(codeKey);
      if (!rowSelf) { alert("找不到此代號/名稱"); return; }

      // 上游 A~C
      const upstreamEdges = upstreamAC.filter(x => x.down === codeKey);

      // 下游 H~J
      const downstreamEdges = downstreamHJ.filter(x => x.up === codeKey);

      console.log("🟢 強制使用 H~J，下游筆數 =", downstreamEdges.length);

      requestAnimationFrame(() => {
        renderResultChip(rowSelf, month, metric, colorMode);
        renderTreemap("upTreemap", "upHint", upstreamEdges, "上游代號", month, metric, colorMode);
      });

      requestAnimationFrame(() => {
        renderTreemap("downTreemap", "downHint", downstreamEdges, "下游代號", month, metric, colorMode);
      });
    });

    console.log("🚀 Treemap override — 已強制綁定專屬 handleRun()");
  }

  function wait() {
    // 確保 links-override 資料已載入
    if (!window.downstreamHJ || !window.upstreamAC) {
      return setTimeout(wait, 120);
    }

    // 確保 app.js 綁完事件，準備覆蓋
    const btn = document.querySelector("#runBtn");
    if (!btn) return setTimeout(wait, 100);

    console.log("Treemap override — 偵測到 runBtn，準備覆蓋");

    applyOverride();
  }

  wait();

})();
