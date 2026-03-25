
/* links-override.js — 安全版（不覆蓋 loadWorkbook，不破壞 byCode） */

(function(){

  console.log("links-override.js — safe version loaded");

  function waitForApp() {

    // 等 app.js 的 loadWorkbook 完成
    if (!window.byCode || window.byCode.size === 0) {
      return setTimeout(waitForApp, 100);
    }

    console.log("links-override: app.js 已載入，開始讀取 H~J 資料");

    // 手動讀取 Links sheet
    fetch(XLSX_FILE, { cache: "no-store" })
      .then(res => res.arrayBuffer())
      .then(buf => {
        const wb = XLSX.read(buf, { type: "array" });
        const ws = wb.Sheets[LINKS_SHEET];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

        window.upstreamAC = [];
        window.downstreamHJ = [];

        for (let i = 1; i < rows.length; i++) {
          const r = rows[i] || [];

          const A = r[0], B = r[1], C = r[2];
          if (A && B && C) {
            upstreamAC.push({
              up: normCode(A),
              down: normCode(B),
              type: normText(C)
            });
          }

          const H = r[7], I = r[8], J = r[9];
          if (H && I && J) {
            downstreamHJ.push({
              up: normCode(H),
              down: normCode(I),
              type: normText(J)
            });
          }
        }

        console.log("AC 筆數 =", upstreamAC.length);
        console.log("HJ 筆數 =", downstreamHJ.length);
      });
  }

  waitForApp();

})();
