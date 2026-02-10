/* assets/js/supplychain.js
 * 互動心智圖（供應鏈）元件
 * 需求：頁面已載入 D3 v7、XLSX（index.html 已有） 
 * 讀取 data.xlsx 的三張新表：ChainNodes / ChainLinks / ChainStocks
 * 並用 Revenue 表補上公司名稱（若 ChainStocks 名稱留空）
 */
(function () {
  const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
  const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
  const SHEET_NODES = 'ChainNodes';
  const SHEET_LINKS = 'ChainLinks';
  const SHEET_STOCKS = 'ChainStocks';
  const SHEET_REVENUE = 'Revenue';
  const WIDTH = 1000, HEIGHT = 620;
  const state = {
    inited: false,
    data: { nodes: [], links: [], stocksByStage: new Map(), stockToStages: new Map(), byCodeName: new Map() },
    svg: null, linkSel: null, nodeSel: null, highlightedStages: new Set(),
  };
  function updateSupplyChainByTicker(code) {
    if (!state.inited) return;
    const stages = state.data.stockToStages.get(code) || [];
    highlightStages(stages);
    if (stages.length > 0) renderStocksList(stages[0]);
    const el = document.getElementById('supplychain');
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  window.updateSupplyChainByTicker = updateSupplyChainByTicker;
  window.addEventListener('stockSelected', (e) => { if (e.detail && e.detail.code) updateSupplyChainByTicker(e.detail.code); });
  document.addEventListener('DOMContentLoaded', () => {
    const container = document.querySelector('#supplychain-vis');
    if (!container) return;
    const io = new IntersectionObserver(async entries => {
      const it = entries.find(x => x.isIntersecting);
      if (it && !state.inited) {
        try { const wb = await loadWorkbook(); buildData(wb); buildChart(container); state.inited = true; }
        catch (err) { console.error(err); container.innerHTML = `<div style="color:#b91c1c">載入供應鏈資料失敗：${err.message}</div>`; }
        finally { io.disconnect(); }
      }
    }, { threshold: 0.15 });
    io.observe(container);
  });
  async function loadWorkbook() {
    const res = await fetch(XLSX_FILE, { cache: 'no-store' });
    if (!res.ok) throw new Error('讀取 data.xlsx 失敗，HTTP '+res.status);
    const buf = await res.arrayBuffer();
    return XLSX.read(buf, { type: 'array' });
  }
  function buildData(wb) {
    const wsNodes = wb.Sheets[SHEET_NODES];
    const wsLinks = wb.Sheets[SHEET_LINKS];
    const wsStocks = wb.Sheets[SHEET_STOCKS];
    const wsRev   = wb.Sheets[SHEET_REVENUE];
    if (!wsNodes || !wsLinks || !wsStocks) { throw new Error(`缺少必要工作表：${[!wsNodes && SHEET_NODES, !wsLinks && SHEET_LINKS, !wsStocks && SHEET_STOCKS].filter(Boolean).join(', ')}`); }
    if (wsRev) {
      const rows = XLSX.utils.sheet_to_json(wsRev, { defval: null });
      const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
      const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];
      const sample = rows[0] || {};
      const codeKey = CODE_FIELDS.find(k => k in sample) || CODE_FIELDS[0];
      const nameKey = NAME_FIELDS.find(k => k in sample) || NAME_FIELDS[0];
      for (const r of rows) { const code = normCode(r[codeKey]); const name = normText(r[nameKey]); if (code) state.data.byCodeName.set(code, name); }
    }
    const nodes = XLSX.utils.sheet_to_json(wsNodes, { defval: null })
      .map(r => ({ id: Number(r['節點ID']), name: String(r['名稱'] || ''), order: Number(r['順序'] || r['節點ID']) }))
      .filter(d => Number.isFinite(d.id) && d.name);
    const links = XLSX.utils.sheet_to_json(wsLinks, { defval: null })
      .map(r => ({ source: Number(r['source']), target: Number(r['target']) }))
      .filter(e => Number.isFinite(e.source) && Number.isFinite(e.target));
    const stocksByStage = new Map();
    const stockToStages = new Map();
    XLSX.utils.sheet_to_json(wsStocks, { defval: null }).forEach(r => {
      const sid = Number(r['節點ID']);
      const code = normCode(r['個股']);
      let name  = normText(r['名稱'] || '');
      if (!name && code && state.data.byCodeName.has(code)) name = state.data.byCodeName.get(code);
      if (!Number.isFinite(sid) || !code) return;
      if (!stocksByStage.has(sid)) stocksByStage.set(sid, []);
      stocksByStage.get(sid).push({ code, name: name || '' });
      if (!stockToStages.has(code)) stockToStages.set(code, []);
      stockToStages.get(code).push(sid);
    });
    state.data.nodes = nodes;
    state.data.links = links;
    state.data.stocksByStage = stocksByStage;
    state.data.stockToStages = stockToStages;
  }
  function buildChart(container) {
    container.innerHTML = '';
    const svg = d3.select(container).append('svg')
      .attr('viewBox', `0 0 ${WIDTH} ${HEIGHT}`)
      .attr('preserveAspectRatio', 'xMidYMid meet')
      .attr('class', 'sc-svg');
    state.svg = svg;
    svg.append('defs').append('marker')
      .attr('id', 'sc-arrow')
      .attr('viewBox', '0 -5 10 10')
      .attr('refX', 22)
      .attr('refY', 0)
      .attr('markerWidth', 6)
      .attr('markerHeight', 6)
      .attr('orient', 'auto')
      .append('path')
      .attr('d', 'M0,-5L10,0L0,5')
      .attr('fill', '#bdbdbd');
    const g = svg.append('g').attr('class','sc-g');
    const linkSel = g.selectAll('.sc-link')
      .data(state.data.links)
      .enter().append('line')
      .attr('class','sc-link')
      .attr('stroke','#bdbdbd')
      .attr('stroke-width',1.5)
      .attr('marker-end','url(#sc-arrow)');
    state.linkSel = linkSel;
    const nodeSel = g.selectAll('.sc-node')
      .data(state.data.nodes)
      .enter().append('g')
      .attr('class','sc-node')
      .style('cursor','pointer')
      .on('click', (e, d) => { renderStocksList(d.id); highlightStages([d.id]); })
      .call(d3.drag()
        .on('start', (ev,d) => { if(!ev.active) sim.alphaTarget(0.3).restart(); d.fx=d.x; d.fy=d.y; })
        .on('drag',  (ev,d) => { d.fx=ev.x; d.fy=ev.y; })
        .on('end',   (ev,d) => { if(!ev.active) sim.alphaTarget(0); d.fx=null; d.fy=null; })
      );
    nodeSel.append('circle')
      .attr('r',22)
      .attr('fill','#ff8a8a')
      .attr('stroke','#cc6666')
      .attr('stroke-width',1.5);
    nodeSel.append('text')
      .attr('text-anchor','middle')
      .attr('dy',-32)
      .attr('class','sc-label')
      .text(d => d.name);
    state.nodeSel = nodeSel;
    const xScale = d3.scaleLinear()
      .domain(d3.extent(state.data.nodes, d => d.order))
      .range([80, WIDTH-80]);
    const sim = d3.forceSimulation(state.data.nodes)
      .force('link', d3.forceLink(state.data.links).id(d => d.id).distance(120))
      .force('charge', d3.forceManyBody().strength(-450))
      .force('center', d3.forceCenter(WIDTH/2, HEIGHT/2))
      .force('x', d3.forceX(d => xScale(d.order)).strength(0.45))
      .force('y', d3.forceY(HEIGHT/2).strength(0.06));
    sim.on('tick', () => {
      linkSel
        .attr('x1', d => d.source.x)
        .attr('y1', d => d.source.y)
        .attr('x2', d => d.target.x)
        .attr('y2', d => d.target.y);
      nodeSel.attr('transform', d => `translate(${d.x},${d.y})`);
    });
    const firstStage = state.data.nodes[0]?.id;
    if (firstStage != null) renderStocksList(firstStage);
  }
  function renderStocksList(stageId) {
    const listWrap = document.getElementById('supplychain-stock-list');
    const list = state.data.stocksByStage.get(stageId) || [];
    if (list.length === 0) { listWrap.innerHTML = `<div class="sc-empty">此節點目前沒有個股資料</div>`; return; }
    listWrap.innerHTML = `
      <h3>📌 相關個股</h3>
      <ul class="sc-stock-ul">
        ${list.map(s => `<li><span class="code">${s.code}</span> ${escapeHtml(s.name || '')}</li>`).join('')}
      </ul>
    `;
  }
  function highlightStages(stageIds) {
    state.highlightedStages = new Set(stageIds || []);
    state.nodeSel.select('circle')
      .attr('fill', d => state.highlightedStages.has(d.id) ? '#e74c3c' : '#ff8a8a')
      .attr('stroke', d => state.highlightedStages.has(d.id) ? '#b60e0e' : '#cc6666')
      .attr('stroke-width', d => state.highlightedStages.has(d.id) ? 3 : 1.5);
    state.linkSel
      .attr('stroke', d => (state.highlightedStages.has(d.source.id) || state.highlightedStages.has(d.target.id)) ? '#666' : '#bdbdbd')
      .attr('stroke-width', d => (state.highlightedStages.has(d.source.id) || state.highlightedStages.has(d.target.id)) ? 3 : 1.5);
  }
  function normText(s){ return String(s==null?'':s).replace(/[​-‍﻿]/g,'').replace(/　/g,' ').replace(/\s+/g,' ').trim(); }
  function toHalfWidth(str){ return String(str||'').replace(/[０-９Ａ-Ｚａ-ｚ]/g,ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); }
  function normCode(s){ return toHalfWidth(String(s||'')).replace(/[​-‍﻿]/g,'').replace(/\s+/g,'').trim(); }
  function escapeHtml(s){ return String(s).replace(/[&<>"']/g,c=>({'&':'&','<':'<','>':'>','"':'"',''':'''}[c])); }
})();
