
/* treemap-override.js — 終極版：直接覆蓋 Treemap 的資料來源，不經 handleRun */

(function(){

  console.log("Treemap override — FINAL version loaded");

  function waitCore() {

    // renderTreemap 必須存在
    if (typeof window.renderTreemap !== "function") {
      return setTimeout(waitCore, 80);
    }

    // upstreamAC / downstreamHJ 必須存在（可能為空但不能 undefined）
    if (!window.upstreamAC || !window.downstreamHJ) {
      return setTimeout(waitCore, 80);
    }

    console.log("Treemap override — core methods detected, patching handleRun...");

    // 替換 handleRun，不讓 app.js 覆蓋
    document.querySelector("#runBtn").addEventListener("click", function(){

      const raw = document.querySelector("#stockInput").value;
      const codeKey = normCode(raw);

      const month = document.querySelector("#monthSelect")?.value;
      const metric = document.querySelector("#metricSelect")?.value;
      const colorMode = document.querySelector("#colorMode")?.value || "redPositive";

      const rowSelf = byCode.get(codeKey);
      if (!rowSelf) {
        alert("找不到此代號/名稱");
        return;
      }

      // 上游：A~C
      const upstreamEdges = upstreamAC.filter(e => e.down === codeKey);

      // 下游：H~J（強制）
      let downstreamEdges = downstreamHJ.filter(e => e.up === codeKey);

      // 排除美股 .US
      downstreamEdges = downstreamEdges.filter(e => !e.down.endsWith(".US"));

      console.log("🔥 下游 H~J 筆數（排除 .US）=", downstreamEdges.length);

      requestAnimationFrame(() => {
        renderResultChip(rowSelf, month, metric, colorMode);
        renderTreemap("upTreemap", "upHint", upstreamEdges, "上游代號", month, metric, colorMode);
      });

      requestAnimationFrame(() => {
        renderTreemap("downTreemap", "downHint", downstreamEdges, "下游代號", month, metric, colorMode);
      });
    });

    console.log("Treemap override — patch completed (FINAL).");
  }

  waitCore();

})();
