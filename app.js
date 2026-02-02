
/* app.js (cache-busting 版) — 讓 data.xlsx 一定抓最新
   概念：在請求 data.xlsx 時加上查詢參數 ?v=版本號，避免 GitHub Pages/瀏覽器快取 */

// 版本號策略：
// 1) 先讀 URL 上的 ?v=（你可手動帶入）
// 2) 若未帶，使用 5 分鐘一期的時間片，避免每次查詢都重抓，但 5 分鐘內更新就會換版本
const URL_VER = new URLSearchParams(location.search).get('v')
  || Math.floor(Date.now() / (1000*60*5));

const XLSX_FILE = new URL(`data.xlsx?v=${URL_VER}`, location.href).toString();

// 其餘邏輯請把原本 app.js 其餘程式維持不變，只需將上一版的：
//   const XLSX_FILE = new URL('data.xlsx', location.href).toString();
// 改成本段，並把 URL_VER 與 XLSX_FILE 放在檔案頂部即可。
