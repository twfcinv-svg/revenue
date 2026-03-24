
/* links-override-final.js */
(function(){
  function waitLoad(){
    if(!window.loadWorkbook){ return setTimeout(waitLoad,100); }
    const old = window.loadWorkbook;
    window.loadWorkbook = async function(){
      await old();
      const res = await fetch(XLSX_FILE,{cache:'no-store'});
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf,{type:'array'});
      const ws = wb.Sheets[LINKS_SHEET];
      const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
      window.upstreamAC = [];
      window.downstreamHJ = [];
      for(let i=1;i<rows.length;i++){
        const r = rows[i]||[];
        const A=r[0],B=r[1],C=r[2];
        if(A&&B&&C){ upstreamAC.push({up:normCode(A),down:normCode(B),type:normText(C)}); }
        const H=r[7],I=r[8],J=r[9];
        if(H&&I&&J){ downstreamHJ.push({up:normCode(H),down:normCode(I),type:normText(J)}); }
      }
    };
  }
  waitLoad();
})();
