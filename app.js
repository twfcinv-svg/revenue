
/* app.js quick patch — 改善「輸入代號沒反應」
 * 1) 正規化輸入代號（去零寬、全形空白、全形數字→半形）
 * 2) 若找不到代號，會嘗試用「中文名稱包含關鍵字」比對
 * 3) Links 與 Revenue 的代號一律字串化再比對
 * 4) 若仍找不到，狀態列顯示可用代號提示
 */

function z(s){return String(s==null?'':s)}
function toHalfWidth(str){ // 全形→半形
  return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0));
}
function normCode(s){
  return toHalfWidth(z(s))
    .replace(/[​-‍﻿]/g,'') // 零寬
    .replace(/[　]/g,' ')              // 全形空白
    .replace(/\s+/g,'')                    // 去所有空白
    .trim();
}

// 供整合：在你的原 app.js 中，將 handleRun 開頭改為：
function handleRun(){
  const raw = document.querySelector('#stockInput').value;
  let code = normCode(raw);
  const month = (document.querySelector('#monthSelect')?.value)||'';
  const metric = document.querySelector('#metricSelect')?.value||'MoM';
  const colorMode=document.querySelector('#colorMode')?.value||'redPositive';

  if(!code){ setStatus('請輸入股票代號或公司名稱'); return; }

  // 1) 代號精準比對
  if(!byCode.has(code)){
    // 2) 名稱模糊比對（中文名包含）
    const hit = revenueRows.find(r => z(r['名稱']).includes(raw.trim()));
    if(hit){ code = String(hit['個股']).trim(); }
  }

  if(!byCode.has(code)){
    const sample = Array.from(byCode.keys()).slice(0,10).join(', ');
    setStatus(`找不到此代號/名稱。「可用代號前 10 筆」：${sample} …`);
    return;
  }

  setStatus('');
  // 繼續你原本的渲染：
  renderResultChip(code, month, metric);
  const upstreamEdges = linksRows.filter(r => String(r['下游代號']).toString().trim() === code);
  renderTreemap('upTreemap','upHint', upstreamEdges, '上游代號', month, metric, colorMode);
  const downstreamEdges = linksRows.filter(r => String(r['上游代號']).toString().trim() === code);
  renderTreemap('downTreemap','downHint', downstreamEdges, '下游代號', month, metric, colorMode);
}
