
// links-override.js — override for A–C / H–J Links separation
// This file injects custom logic AFTER app.js is loaded.

(function(){
    console.log('links-override.js loaded');

    // --- Override global parsed link sets ---
    window.upstreamAC = [];
    window.downstreamHJ = [];

    // Patch loadWorkbook to add AC/HJ parsing
    const oldLoad = window.loadWorkbook;
    window.loadWorkbook = async function(){
        const res = await fetch(XLSX_FILE, { cache:'no-store' });
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type:'array' });
        const wsLinks = wb.Sheets[LINKS_SHEET];
        const rows = XLSX.utils.sheet_to_json(wsLinks, { header:1, defval:null, blankrows:false });

        window.upstreamAC = [];
        window.downstreamHJ = [];

        for (let i = 1; i < rows.length; i++){
            const r = rows[i]; if (!r) continue;
            const A=r[0],B=r[1],C=r[2];
            if(A && B && C){ upstreamAC.push({ up:normCode(A), down:normCode(B), type:normText(C)}); }
            const H=r[7],I=r[8],J=r[9];
            if(H && I && J){ downstreamHJ.push({ up:normCode(H), down:normCode(I), type:normText(J)}); }
        }
        if(oldLoad) return oldLoad();
    };

    // Override handleRun edges construction
    const oldRun = window.handleRun;
    window.handleRun = function(){
        const raw = document.querySelector('#stockInput').value;
        const codeKey = normCode(raw);
        const upstreamEdges = upstreamAC.filter(l => l.down === codeKey);
        const downstreamEdges = downstreamHJ.filter(l => l.up === codeKey);
        window.__override_up = upstreamEdges;
        window.__override_down = downstreamEdges;
        return oldRun();
    };
})();
