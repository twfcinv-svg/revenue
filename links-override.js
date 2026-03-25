
/* links-override.js — 終極穩定版（完全獨立解析 Excel，無需依賴 app.js） */

(function(){

  console.log("links-override.js — 終極版載入");

  async function loadLinksManually() {

    // 直接讀取 data.xlsx，不依賴 app.js
    const res = await fetch(XLSX_FILE, { cache: "no-store" });
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const ws = wb.Sheets[LINKS_SHEET];
    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: null,
      blankrows: false,
    });

    // 建立 A~C 與 H~J 結構
    window.upstreamAC = [];
    window.downstreamHJ = [];

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i] || [];

      const A = r[0], B = r[1], C = r[2];
      if (A && B && C) {
        upstreamAC.push({
          up: normCode(A),
          down: normCode(B),
          type: normText(C),
        });
      }

      const H = r[7], I = r[8], J = r[9];
      if (H && I && J) {
        downstreamHJ.push({
          up: normCode(H),
          down: normCode(I),
          type: normText(J),
        });
      }
    }

    console.log("A~C 筆數 =", upstreamAC.length);
    console.log("H~J 筆數 =", downstreamHJ.length);

  }

  // 馬上執行
  loadLinksManually();

})();
