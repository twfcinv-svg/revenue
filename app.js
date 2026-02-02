
/* app.js (patched) — 修復「月份無法選取」問題：
 * 1) 以正則解析欄名，容忍全形括號/百分比、空白、看不見字元
 * 2) 自動列出所有偵測到的月份；若 0 個，於狀態列顯示偵測到的原始欄名供排錯
 * 3) 容錯：忽略空白列、數字型代號、自動字串化
 */

const XLSX_FILE = 'data.xlsx';
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';

// 正規化：移除零寬字元/不可見空白/全形空白
function norm(s){
  return String(s==null?'':s)
    .replace(/[​-‍﻿]/g,'')    // 零寬
    .replace(/[　]/g,' ')                  // 全形空白
    .replace(/\s+/g,' ')                       // 多空白
    .trim();
}

// 允許下列樣式（半形或全形括號/百分比均可）：
// 202512單月合併營收年成長(%) / 202512單月合併營收月變動(%)
// 202512單月合併營收年成長（％） / 202512單月合併營收月變動（％）
const RX_YOY = /^(\d{6})\s*單月合併營收\s*年成長\s*[\(（]\s*%|％\s*[\)）]\s*$/;
const RX_MOM = /^(\d{6})\s*單月合併營收\s*月變動\s*[\(（]\s*%|％\s*[\)）]\s*$/;

let revenueRows = [], linksRows = [], byCode = new Map(), months = [];

window.addEventListener('DOMContentLoaded', async () => {
  setStatus('載入資料中…');
  try{
    await loadWorkbook(XLSX_FILE);
    initControls();
    setStatus(`資料就緒（偵測月份：${months.join(', ')||'無'}）`);
  }catch(e){
    console.error(e);
    setStatus('載入失敗：'+e.message);
  }
  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
  window.addEventListener('resize', ()=>{
    const code = document.querySelector('#stockInput')?.value?.trim();
    if(code && byCode.has(code)) handleRun();
  });
});

function setStatus(s){ const el=document.querySelector('#status'); if(el) el.textContent = s||''; }

async function loadWorkbook(url){
  const res = await fetch(url); if(!res.ok) throw new Error('HTTP '+res.status);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type:'array' });
  const wsRev = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];
  if(!wsRev || !wsLinks) throw new Error('找不到必要工作表 Revenue 或 Links');

  revenueRows = XLSX.utils.sheet_to_json(wsRev, { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });

  // 構建公司索引
  byCode.clear();
  for(const r of revenueRows){
    const c = String(r['個股'] ?? r['代號'] ?? '').trim();
    if(c) byCode.set(c, r);
  }

  // 從欄名解析月份
  const headers = Object.keys(revenueRows[0]||{}).map(norm);
  const foundYoY = new Set();
  const foundMoM = new Set();
  for(const h of headers){
    const hy = h.match(/^(\d{6})\s*單月合併營收\s*年成長\s*[\(（]?\s*(?:%|％)\s*[\)）]?\s*$/);
    if(hy) foundYoY.add(hy[1]);
    const hm = h.match(/^(\d{6})\s*單月合併營收\s*月變動\s*[\(（]?\s*(?:%|％)\s*[\)）]?\s*$/);
    if(hm) foundMoM.add(hm[1]);
  }
  months = Array.from(new Set([...foundYoY, ...foundMoM])).sort((a,b)=>b.localeCompare(a));

  // 若偵測不到月份，輸出提示以便你檢查欄名
  if(months.length===0){
    console.warn('未偵測到月份欄位。原始欄名：', Object.keys(revenueRows[0]||{}));
  }
}

function initControls(){
  const sel = document.querySelector('#monthSelect');
  if(!sel) return;
  sel.innerHTML = '';
  for(const m of months){
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = `${m.slice(0,4)}年${m.slice(4,6)}月`;
    sel.appendChild(opt);
  }
  if(!sel.value && months.length>0){ sel.value = months[0]; }
}

function getMetricValue(row, month, metric){
  if(!row) return null;
  // 嘗試多種括號/百分比寫法
  const patterns = [
    `${month}單月合併營收${metric==='YoY'?'年成長':'月變動'}(%)`,
    `${month}單月合併營收${metric==='YoY'?'年成長':'月變動'}（%）`,
    `${month}單月合併營收${metric==='YoY'?'年成長':'月變動'}（％）`,
    `${month}單月合併營收${metric==='YoY'?'年成長':'月變動'}(％)`
  ];
  let v=null;
  for(const p of patterns){ if(row[p]!=null && row[p]!=='' ){ v=row[p]; break; } }
  if(v==null) return null;
  v = Number(v);
  return Number.isFinite(v)?v:null;
}

function displayPct(v){ if(v==null||!isFinite(v)) return '—'; const s=v.toFixed(1)+'%'; return v>0?('+'+s):s; }

function safe(s){ return String(s??'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

// 下面保留你的 renderFocus / renderTreemap / handleRun 等函式，僅示範如何安全地取得 sel.value
function handleRun(){
  const code = document.querySelector('#stockInput').value.trim();
  const sel = document.querySelector('#monthSelect');
  const month = sel && sel.value ? sel.value : (months[0]||'');
  const metric = document.querySelector('#metricSelect').value;
  const colorMode = document.querySelector('#colorMode')?.value || 'redPositive';
  if(!code) return setStatus('請輸入股票代號');
  if(!byCode.has(code)) return setStatus(`找不到代號 ${code} 的公司於 Revenue 表`);
  setStatus('');
  // 在你的正式專案中，這裡呼叫 renderResultChip / renderTreemap ...
  console.log('OK run', {code, month, metric, colorMode});
}
