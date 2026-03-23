/* app.js — v3.14 (Treemap 讀取 Links A~C + H~J，並完整修復月份與營收解析) */

const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';

const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];

let revenueRows = [], months = [];
let byCode = new Map(), byName = new Map();
let linksByUp = new Map(), linksByDown = new Map();
let COL_MAP = {};

function norm(s){ return String(s||'').trim(); }
function normCode(s){ return String(s||'').replace(/\s+/g,'').trim(); }
function isUS(code){ return /\.US$/i.test(String(code||'')); }

window.addEventListener('DOMContentLoaded', async()=>{
  try{ await loadWorkbook(); initControls(); }
  catch(e){ console.error(e); alert(e.message); }
  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
});

async function loadWorkbook(){
  const res = await fetch(XLSX_FILE, {cache:'no-store'});
  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf,{type:'array'});

  const wsRev   = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];
  if(!wsRev || !wsLinks) throw new Error('找不到 Revenue 或 Links');

  // ---- 讀取 Revenue ----
  const rows = XLSX.utils.sheet_to_json(wsRev,{defval:null});
  revenueRows = rows;

  // 自動偵測月份欄位：找出所有 YYYYMM 開頭欄位
  const headers = Object.keys(rows[0]||{});
  const ymSet = new Set();
  for(const h of headers){
    const m = String(h).match(/^(20\d{2})(0[1-9]|1[0-2])/);
    if(m) ymSet.add(m[1]+m[2]);
  }
  months = Array.from(ymSet).sort().reverse();

  // 建立 COL_MAP
  for(const ym of months){
    COL_MAP[ym] = {
      amount: headers.find(h=>String(h).includes(ym) && /營收|金額/.test(h)) || null,
      mom:    headers.find(h=>String(h).includes(ym) && /MoM|月增|月變/.test(h)) || null,
      yoy:    headers.find(h=>String(h).includes(ym) && /YoY|年增|年成/.test(h)) || null,
    };
  }

  // 代號/名稱 map
  const sample = rows[0]||{};
  const codeKey = CODE_FIELDS.find(k=>k in sample) || CODE_FIELDS[0];
  const nameKey = NAME_FIELDS.find(k=>k in sample) || NAME_FIELDS[0];

  for(const r of revenueRows){
    const code = normCode(r[codeKey]);
    const name = norm(r[nameKey]);
    if(code) byCode.set(code,r);
    if(name) byName.set(name,r);
  }

  // ---- Treemap 讀取 Links（A~C + H~J）----
  const rowsA1 = XLSX.utils.sheet_to_json(wsLinks,{header:1,defval:null});
  const merged = [];

  for(let i=1;i<rowsA1.length;i++){
    const row = rowsA1[i] || [];
    const A = normCode(row[0]), B = normCode(row[1]), C = norm(row[2]);
    if(A||B||C) merged.push({up:A,down:B,type:C});
    const H = normCode(row[7]), I = normCode(row[8]), J = norm(row[9]);
    if(H||I||J) merged.push({up:H,down:I,type:J});
  }

  linksByUp.clear(); linksByDown.clear();
  for(const e of merged){
    if(e.up){
      if(!linksByUp.has(e.up)) linksByUp.set(e.up,[]);
      linksByUp.get(e.up).push(e);
    }
    if(e.down){
      if(!linksByDown.has(e.down)) linksByDown.set(e.down,[]);
      linksByDown.get(e.down).push(e);
    }
  }
}

function initControls(){
  const sel = document.querySelector('#monthSelect');
  sel.innerHTML='';
  for(const m of months){
    const o=document.createElement('option');
    o.value=m;
    o.textContent=`${m.slice(0,4)}年${m.slice(4,6)}月`;
    sel.appendChild(o);
  }
  if(months.length>0) sel.value=months[0];
}

function metricOf(row, ym, metric){
  const code = normCode(row['個股']||row['代號']||row['股票代碼']||'');
  if(isUS(code)) return null;
  const map = COL_MAP[ym]; if(!map) return null;
  const key = {MoM:'mom',YoY:'yoy'}[metric] || 'amount';
  const col = map[key]; if(!col) return null;
  let v = row[col]; if(v==null||v==='') return null;
  return Number(String(v).replace(/[%％]/g,''));
}

function handleRun(){ /* 你原本的方法不變，略 */ }
function renderTreemap(){ /* 你原本的方法不變，略 */ }
