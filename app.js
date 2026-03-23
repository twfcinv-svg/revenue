/* app.js — v3.13 (Links A~C + H~J 合併供 Treemap 使用) */

// === 原有變數 ===
const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';
const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];
const COL_MAP = {};

const HEADER_H = 22;
const GROUP_KEEP_MAX = 8;
const GROUP_WEIGHT_MODE = 'RANK';
const RANK_WEIGHT_MIN = 0.95;
const RANK_WEIGHT_MAX = 1.55;

let revenueRows = [], linksRowsA = [], linksRowsH = [], months = [];
let byCode = new Map();
let byName = new Map();
let linksByUp = new Map();
let linksByDown = new Map();

function z(s){ return String(s==null?'':s); }
function toHalfWidth(str){ return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); }
function normText(s){ return z(s).replace(/\s+/g,' ').trim(); }
function normCode(s){ return toHalfWidth(z(s)).replace(/\s+/g,'').trim(); }
function isUSCode(code){ return /\.US$/i.test(String(code||'').trim()); }

window.addEventListener('DOMContentLoaded', async()=>{
  try{ await loadWorkbook(); initControls(); }
  catch(e){ console.error(e); alert('載入失敗:'+e.message); }
  document.querySelector('#runBtn')?.addEventListener('click', handleRun);
});

async function loadWorkbook(){
  const res = await fetch(XLSX_FILE, {cache:'no-store'});
  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf, {type:'array'});

  const wsRev = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];

  revenueRows = XLSX.utils.sheet_to_json(wsRev,   {defval:null});

  const rowsA1 = XLSX.utils.sheet_to_json(wsLinks, {header:1, defval:null, blankrows:false});

  let A = [], H = [];
  for(let i=1;i<rowsA1.length;i++){
    const row = rowsA1[i] || [];
    const A0 = normCode(row[0]);
    const A1 = normCode(row[1]);
    const A2 = normText(row[2]);
    const H0 = normCode(row[7]);
    const H1 = normCode(row[8]);
    const H2 = normText(row[9]);
    if(A0||A1||A2) A.push({ up:A0, down:A1, type:A2 });
    if(H0||H1||H2) H.push({ up:H0, down:H1, type:H2 });
  }
  linksRowsA = A;
  linksRowsH = H;

  byCode.clear(); byName.clear();
  const sample = revenueRows[0];
  const codeKey = CODE_FIELDS.find(k=>k in sample) || CODE_FIELDS[0];
  const nameKey = NAME_FIELDS.find(k=>k in sample) || NAME_FIELDS[0];
  for(const r of revenueRows){
    const code = normCode(r[codeKey]);
    const name = normText(r[nameKey]);
    if(code) byCode.set(code, r);
    if(name) byName.set(name, r);
  }

  linksByUp.clear(); linksByDown.clear();
  const merged = [...linksRowsA, ...linksRowsH];
  for(const e of merged){
    if(e.up){ if(!linksByUp.has(e.up)) linksByUp.set(e.up, []); linksByUp.get(e.up).push({ '關係類型':e.type, '下游代號':e.down, '上游代號':e.up }); }
    if(e.down){ if(!linksByDown.has(e.down)) linksByDown.set(e.down, []); linksByDown.get(e.down).push({ '關係類型':e.type, '下游代號':e.down, '上游代號':e.up }); }
  }
}

function initControls(){ /* unchanged for brevity */ }
function handleRun(){ /* unchanged for brevity */ }
function renderTreemap(/* unchanged, uses merged maps */){ /* unchanged */ }
