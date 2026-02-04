/* app_query_link_B.js | 右上角「查詢」區域：
 * - 點「查詢結果徽章 #resultChip」→ 以代號開啟 FBS（新分頁）
 * - 在輸入框 #stockInput 按 Enter → 以代號開啟 FBS（新分頁）
 * - 你也可以在 app.js 查詢完成後呼叫 window.setResultChipLink(code, name, extra)
 *   以明確指定代號與顯示文字（最穩定的 B 方法）。
 */
(function () {
  const FBS_BASE = 'https://www.fbs.com.tw/MKT/Index?name=%EF%BC%AB%E7%B7%9A%E5%9C%96&stock='; // Ｊ線圖 + 代號

  const $ = (sel) => document.querySelector(sel);
  const input = $('#stockInput');
  const chip  = $('#resultChip');
  const run   = $('#runBtn');

  function openFBS(code) {
    if (!code) return;
    const url = FBS_BASE + encodeURIComponent(String(code));
    window.open(url, '_blank');
  }

  function pickCodeFromText(text) {
    if (!text) return null;
    const s = String(text).trim();
    // 取前四位數（2330 台積電 → 2330；或純 2330）
    const m = /\b(\d{4})\b/.exec(s);
    return m ? m[1] : null;
  }

  function resolveCode() {
    // 1) chip 上若有 data-code（由 setResultChipLink 設定） → 優先
    if (chip && chip.dataset && chip.dataset.code) return chip.dataset.code;
    // 2) 從 chip 文字抓 4 碼
    if (chip) {
      const c1 = pickCodeFromText(chip.textContent || chip.innerText);
      if (c1) return c1;
    }
    // 3) 從輸入框抓 4 碼
    if (input) {
      const c2 = pickCodeFromText(input.value);
      if (c2) return c2;
    }
    return null;
  }

  // 讓 app.js 可以「明確指定」查詢結果徽章內容與代號（B 方法建議用）
  window.setResultChipLink = function setResultChipLink(code, name, extraText) {
    if (!chip) return;
    chip.dataset.code = code || '';
    const label = [code || '', name || ''].filter(Boolean).join(' ');
    chip.textContent = label + (extraText ? (' ｜ ' + extraText) : '');
    chip.style.cursor = 'pointer';
    chip.title = '點我開啟 FBS 個股頁（新分頁）';
  };

  function bindChipAndInput() {
    if (chip) {
      chip.style.cursor = 'pointer';
      chip.title = '點我開啟 FBS 個股頁（新分頁）';
      chip.addEventListener('click', () => openFBS(resolveCode()));
      // 若 app.js 會動態更新 chip 內容，監看後自動套 pointer 樣式
      const mo = new MutationObserver(() => { chip.style.cursor = 'pointer'; });
      mo.observe(chip, { childList: true, characterData: true, subtree: true });
    }

    if (input) {
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          openFBS(resolveCode());
        }
      });
    }

    if (run) {
      // 若你希望「按查詢」就同時開 FBS，改成 click 即可；這裡預設採用雙擊以避免影響原先一次點擊的查詢流程。
      run.addEventListener('dblclick', () => openFBS(resolveCode()));
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bindChipAndInput);
  } else {
    bindChipAndInput();
  }
})();
