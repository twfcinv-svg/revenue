
# 供應鏈月營收｜左側 Treemap（按關係類型）＋ 右側清單

- **左欄（上游）**：以 D3 Treemap 呈現、依 `關係類型` 分群。磁磚大小反映各股的 |MoM/YoY|，磁磚顏色依表現著色，
  並可切換「紅=好/綠=弱」與「綠=好/紅=弱」。
- **右欄（下游）**：維持原本的清單式呈現（不分關係類型），保留條帶強度與數值。
- **中間**：選定公司資訊。

> 專案為純前端，使用 SheetJS 載入 `data.xlsx` 以及 D3 v7 繪製 Treemap。把這些檔案放到同一層，部署到 GitHub Pages 即可。

## 使用
1. 將 `index.html`, `styles.css`, `app.js`, `data.xlsx` 放到 repo 根目錄。
2. GitHub Pages：Settings → Pages → Deploy from a branch → 選 `main` 與根目錄 `/`。

## Excel 規格
- **Revenue**：每家公司一列，欄位含 `個股`、`名稱`、`產業別`，與 `YYYYMM單月合併營收年成長(%)`、`YYYYMM單月合併營收月變動(%)`。
- **Links**：`上游代號`、`下游代號`、`關係類型`（本版本已不使用「權重」）。

## 顏色邏輯
- 右上控制可以切換：
  - **紅=好 / 綠=弱**（預設）
  - **綠=好 / 紅=弱**
- Treemap 以透明度呈現強度；清單以條帶寬度呈現強度。

## 授權
MIT License。
