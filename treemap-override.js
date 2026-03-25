
/* treemap-override.js — 強制覆蓋，下游只使用 H~J */

(function(){

  console.log("Treemap override — 強制啟動版載入");

  function wait() {

    // 等到 app.js 的 handleRun 已經準備好
    if (typeof window.handleRun !== "function") {
      return setTimeout(wait, 120);
    }

    // 等到 links-override 的資料建好（可能是空陣列但不能是 undefined）
    if (!window.downstreamHJ || !window.upstreamAC) {
      return setTimeout(wait, 120);
    }

    console.log("Treemap override — 已偵測 handleRun，立即覆蓋");

    // 強制覆蓋，不讓 app.js 後續再蓋回來
    const realHandleRun = window.handleRun;

    window.handleRun = function () {

      const raw = document.querySelector("#stockInput").value;
      const codeKey = normCode(raw);

      const month = document.querySelector("#monthSelect")?.value;
      const metric = document.querySelector("#metricSelect")?.value;
      const colorMode =
        document.querySelector("#colorMode")?.value || "redPositive";

      const rowSelf = byCode.get(codeKey);
      if (!rowSelf) {
        alert("找不到此代號/名稱");
        return;
      }

      // 上游 → A~C
      const upstreamEdges = upstreamAC.filter((x) => x.down === codeKey);

      // 下游 → **H~J，完全取代 linksByUp**
      const downstreamEdges = downstreamHJ.filter((x) => x.up === codeKey);

      console.log("🟢 下游使用 H~J 筆數 =", downstreamEdges.length);

      requestAnimationFrame(() => {
        renderResultChip(rowSelf, month, metric, colorMode);
        renderTreemap(
          "upTreemap",
          "upHint",
          upstreamEdges,
          "上游代號",
          month,
          metric,
          colorMode
        );
      });

      requestAnimationFrame(() => {
        renderTreemap(
          "downTreemap",
          "downHint",
          downstreamEdges,
          "下游代號",
          month,
          metric,
          colorMode
        );
      });
    };

    console.log("Treemap override — 覆蓋完成，開始使用 H~J");
  }

  wait();
})();
