import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

/* ── Global styles injected once ─────────────────────────────────────────── */
const GLOBAL_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', system-ui, sans-serif; background: #F7F6F3; color: #1a1a1a; }
  ::-webkit-scrollbar { width: 5px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: #D5D3CE; border-radius: 99px; }
  textarea, input, select, button { font-family: inherit; }
  input:focus, textarea:focus { outline: 2px solid #6B63D8; outline-offset: 1px; }
  button:focus-visible { outline: 2px solid #6B63D8; outline-offset: 2px; }
`;

/* ── Design tokens ────────────────────────────────────────────────────────── */
const D = {
  bg:       "#F7F6F3",
  surface:  "#FFFFFF",
  border:   "#E5E3DE",
  borderHover: "#C9C7C0",
  textPrimary:   "#1A1A1A",
  textSecondary: "#6B6960",
  textTertiary:  "#9E9B92",
  accent:   "#6B63D8",
  accentBg: "#EEEDFC",
  accentText: "#3D36A0",
  red:    { bg:"#FEF2F2", brd:"#FECACA", txt:"#991B1B", sub:"#DC2626" },
  amber:  { bg:"#FFFBEB", brd:"#FDE68A", txt:"#92400E", sub:"#D97706" },
  green:  { bg:"#F0FDF4", brd:"#BBF7D0", txt:"#14532D", sub:"#16A34A" },
  gray:   { bg:"#F9F9F7", brd:"#E5E3DE", txt:"#6B6960", sub:"#9E9B92" },
};

/* ── i18n ─────────────────────────────────────────────────────────────────── */
const T = {
  cs: {
    title: "Production Analyzer",
    tagline: "Ball Corporation — denní přehled výroby",
    subtitle_files: n => `${n} ${n===1?"soubor":n<5?"soubory":"souborů"} v historii`,
    subtitle_none: "Nahraj první soubor",
    upload_btn: "Nahrát xlsx",
    uploading: "Načítám…",
    drop_title: "Přetáhni xlsx soubor sem",
    drop_sub: "nebo klikni pro výběr souboru",
    tab_chat: "Chat",
    tab_comments: n => `Komentáře (${n})`,
    tab_history: n => `Historie (${n})`,
    no_comments: "Žádné komentáře v gap sheetu.",
    no_history: "Zatím žádná historie.",
    clear_history: "Vymazat historii",
    compare_btn: "Porovnat",
    analyzing: "Analyzuji…",
    input_placeholder: "Zeptej se na cokoliv…",
    send: "Odeslat",
    loaded_msg: (name, plants, comments) =>
      `Načteno: **${name}**\n${plants} závodů · ${comments} komentářů operátorů\n\nZeptej se nebo vyber rychlou otázku níže.`,
    welcome: "Nahraj xlsx soubor a zeptej se na cokoliv.\nKaždý soubor si pamatuji — postupem času uvidíme trendy a opakující se problémy.",
    history_welcome: n => `Mám uloženo ${n} ${n===1?"soubor":n<5?"soubory":"souborů"}. Nahraj nový nebo se rovnou zeptej.`,
    uploaded_on: "nahráno",
    quick: ["Co je největší problém teď?","Jaké problémy se opakují?","Jak si stojí efektivita?","Shrň komentáře manažerů","Porovnej závody vs. plán"],
    detail_prompt: p => `Řekni mi víc o závodě ${p} — co se tam děje a jaké jsou problémy?`,
    compare_prompt: f => `Porovnej soubor ${f} s ostatními — co se změnilo?`,
    error_load: "Soubor se nepodařilo načíst: ",
    error_ai: m => `Chyba: ${m}`,
    vs_ll: "vs LL",
    eff: "eff",
    comments_col: "Komentář",
    date_col: "Datum",
    plant_col: "Závod",
    line_col: "Linka",
  },
  en: {
    title: "Production Analyzer",
    tagline: "Ball Corporation — daily production overview",
    subtitle_files: n => `${n} file${n===1?"":"s"} in history`,
    subtitle_none: "Upload your first file",
    upload_btn: "Upload xlsx",
    uploading: "Loading…",
    drop_title: "Drop your xlsx file here",
    drop_sub: "or click to browse",
    tab_chat: "Chat",
    tab_comments: n => `Comments (${n})`,
    tab_history: n => `History (${n})`,
    no_comments: "No comments found in the gap sheet.",
    no_history: "No history yet.",
    clear_history: "Clear history",
    compare_btn: "Compare",
    analyzing: "Analyzing…",
    input_placeholder: "Ask anything about the report…",
    send: "Send",
    loaded_msg: (name, plants, comments) =>
      `Loaded: **${name}**\n${plants} plants · ${comments} operator comments\n\nAsk a question or pick a quick option below.`,
    welcome: "Upload an xlsx file and ask anything.\nI remember every file — over time we can spot trends and recurring issues.",
    history_welcome: n => `I have ${n} file${n===1?"":"s"} saved. Upload a new one or ask away.`,
    uploaded_on: "uploaded",
    quick: ["What's the biggest issue now?","Which problems keep recurring?","How is efficiency across plants?","Summarise manager comments","Compare plants vs. plan"],
    detail_prompt: p => `Tell me more about ${p} — what's going on and what are the issues?`,
    compare_prompt: f => `Compare file ${f} with the others — what changed?`,
    error_load: "Failed to load file: ",
    error_ai: m => `Error: ${m}`,
    vs_ll: "vs LL",
    eff: "eff",
    comments_col: "Comment",
    date_col: "Date",
    plant_col: "Plant",
    line_col: "Line",
  },
};


/* ── Storage ──────────────────────────────────────────────────────────────── */
const SK = "prod_hist_v2", LK = "prod_lang_v2";
const loadHistory = async () => { try { const r = localStorage.getItem(SK); return r ? JSON.parse(r) : []; } catch { return []; } };
const saveHistory = async h => { try { localStorage.setItem(SK, JSON.stringify(h)); } catch { /* ignore */ } };
const loadLang    = async () => { try { return localStorage.getItem(LK) || "cs"; } catch { return "cs"; } };
const saveLang    = async l => { try { localStorage.setItem(LK, l); } catch { /* ignore */ } };

/* ── Excel ────────────────────────────────────────────────────────────────── */
const excelDateToStr = n => !n || typeof n !== "number" ? null :
  new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10);

const COL_PLANT = {
  10:"Velim",11:"Velim",13:"Bellegarde",14:"Bellegarde",15:"Bellegarde",16:"Bellegarde",
  18:"Devizes",19:"Devizes",20:"Devizes",21:"Devizes",22:"Devizes",23:"Devizes",24:"Devizes",
  27:"Beaurepaire",28:"Beaurepaire",29:"Beaurepaire",30:"Beaurepaire",31:"Beaurepaire",32:"Beaurepaire",
  34:"Llinars",35:"Llinars",36:"Llinars",37:"Llinars",38:"Llinars",39:"Llinars",
  40:"Llinars",41:"Llinars",42:"Llinars",43:"Llinars",44:"Llinars",45:"Llinars",
};

function parseFile(wb, filename) {
  const sheets = {};
  wb.SheetNames.forEach(n => { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header:1, defval:"" }); });

  // Comments from gap sheet
  const comments = [];
  const gapWs = wb.Sheets["gap"];
  if (gapWs) {
    const keys = Object.keys(gapWs).filter(k => !k.startsWith("!"));
    const hdr = {};
    keys.forEach(a => { const m = a.match(/^([A-Z]+)3$/); if (m) hdr[XLSX.utils.decode_col(m[1])+1] = String(gapWs[a].v||""); });
    const rowDates = {};
    const mtdWs = wb.Sheets["MTD"];
    if (mtdWs) Object.keys(mtdWs).filter(k=>/^A\d+$/.test(k)).forEach(a => {
      const c = mtdWs[a]; if (c && typeof c.v==="number" && c.v>46000) rowDates[parseInt(a.slice(1))] = excelDateToStr(c.v);
    });
    keys.forEach(a => { const m = a.match(/^A(\d+)$/); const c = gapWs[a]; if (m && c && typeof c.v==="number" && c.v>46000) rowDates[parseInt(m[1])] = excelDateToStr(c.v); });
    keys.forEach(a => {
      const m = a.match(/^([A-Z]+)(\d+)$/); if (!m) return;
      const cell = gapWs[a]; if (!cell?.c) return;
      const col = XLSX.utils.decode_col(m[1])+1, row = parseInt(m[2]);
      const raw = cell.c.map(c=>c.t||"").join(" ").trim();
      let clean = raw.includes("Comment:") ? raw.split("Comment:").pop().trim() : raw;
      if (clean && !clean.includes("Your version of Excel"))
        comments.push({ date: rowDates[row]||null, plant: COL_PLANT[col]||null, line: hdr[col]||m[1], text: clean });
    });
  }

  // Plant summary from REPORT sheet
  const summary = [];
  const PLANTS         = ["Velim","Devizes","Bellegarde","Llinars","Lummen","SLP","Itupeva","Beaurepaire","Sherbrooke"];
  const AGGREGATE_ROWS = ["Cans","PG5","Slugs"];
  const PHC_PLANTS     = ["Velim","Devizes","Bellegarde","Llinars","Lummen","SLP","Itupeva"];
  const PHC_CAN_PLANTS = ["Velim","Devizes","Bellegarde","Itupeva"];
  if (sheets["REPORT"]) {
    const reportRows = sheets["REPORT"];
    reportRows.forEach(row => {
      const name = String(row[1]||"").trim();
      if (PLANTS.includes(name) || AGGREGATE_ROWS.includes(name))
        summary.push({
          plant:          name,
          production:     row[2],
          varAOP:         row[3],
          varLatestLL:    row[5],
          varPlanLL:      row[7],
          efficiency:     row[13],
          effVarAOP:      row[14],
          efficiencyTarget: EFFICIENCY_TARGETS[name] ?? null,
          spoilageHFI:    row[18],
          spoilageMfg:    row[19],
          totalSpoilage:  row[20],
          spoilageVarAOP: row[21],
          spoilageCans:   row[24],
          operatingDays:  row[59],
          isPHC:          PHC_PLANTS.includes(name),
          isPHCCan:       PHC_CAN_PLANTS.includes(name),
          isSlug:         ["Llinars","Lummen","SLP","Beaurepaire","Sherbrooke"].includes(name),
          isAggregate:    AGGREGATE_ROWS.includes(name),
        });
    });
  }

  // Monthly plan targets from LLPRODN sheet
  const monthlyPlan = {};
  if (sheets["LLPRODN"]) {
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    const planRow = sheets["LLPRODN"].find(r => String(r[0]||"").trim() === "Total Cans");
    if (planRow) months.forEach((m, i) => { monthlyPlan[m] = planRow[37 + i]; });
  }

  let period = null;
  if (sheets["SELECT"]) for (const row of sheets["SELECT"])
    if (String(row[1]).includes("INPUT YEAR")) { period = String(row[4]); break; }

  return { filename, summary, comments, period, monthlyPlan, uploadedAt: new Date().toISOString() };
}

/* ── Efficiency targets (AOP, June 2026) ─────────────────────────────────── */
const EFFICIENCY_TARGETS = {
  // PHC can plants — primary focus
  "Velim":       0.649,
  "Devizes":     0.562,
  "Bellegarde":  0.515,
  "Itupeva":     0.702,
  // PHC slug plants — info only
  "Llinars":     0.586,
  "Lummen":      0.550,
  "SLP":         0.479,
  // PHC aggregate rows
  "Cans":        0.585,  // all 7 PHC plants — KEY metric
  "PG5":         0.605,  // Cans + Slugs
  // Non-PHC slug plants
  "Beaurepaire": 0.805,
  "Sherbrooke":  0.797,
  "Slugs":       0.801,  // Beaurepaire + Sherbrooke
};

/* ── AI ───────────────────────────────────────────────────────────────────── */
const SYS = `You are a production analytics expert for Ball Corporation aerosol can manufacturing.

Plant structure:
PHC group = Velim (CZ), Devizes (UK), Bellegarde (FR), Itupeva (BR) [can plants] + Llinars (ES), Lummen (BE), SLP (MX) [slug plants].
Non-PHC: Beaurepaire (FR), Sherbrooke (CA) [slug plants].
Aggregate rows: Cans = all 7 PHC plants, PG5 = Cans + Slugs, Slugs = Beaurepaire + Sherbrooke.

June 2026 AOP efficiency targets:
PHC can plants (PRIMARY focus): Velim 64.9%, Devizes 56.2%, Bellegarde 51.5%, Itupeva 70.2%
PHC aggregate — KEY metric, always highlight: Cans 58.5%, PG5 60.5%
PHC slug plants (brief mention only): Llinars 58.6%, Lummen 55.0%, SLP 47.9%
Non-PHC (for info): Beaurepaire 80.5%, Sherbrooke 79.7%, Slugs total 80.1%

A plant is BELOW target if efficiency < its specific AOP target.
varAOP = variance vs AOP in millions of units (Mu) — negative = behind plan.
Total spoilage: below 12% good, above 20% serious.
Operator/manager comments come from Excel cell notes — they describe real failure reasons.
With multiple files, identify recurring patterns across time periods.
Be specific: name plant, metric, magnitude. Be concise.
IMPORTANT: If lang=cs → respond in Czech. If lang=en → respond in English.`;

function buildCtx(cur, hist) {
  let s = "";
  if (hist.length) {
    s += `=== HISTORY (${hist.length} files) ===\n`;
    hist.forEach(f => {
      s += `\n[${f.filename} | ${f.period||"?"} | ${f.uploadedAt?.slice(0,10)}]\n`;
      f.summary?.forEach(p => {
        const e   = typeof p.efficiency === "number" ? (p.efficiency * 100).toFixed(1) + "%" : "?";
        const tgt = typeof p.efficiencyTarget === "number" ? (p.efficiencyTarget * 100).toFixed(1) + "%" : "?";
        const vAOP = typeof p.varAOP === "number" ? (p.varAOP > 0 ? "+" : "") + p.varAOP.toFixed(2) + "M" : "?";
        const vLL  = typeof p.varLatestLL === "number" ? (p.varLatestLL > 0 ? "+" : "") + p.varLatestLL.toFixed(2) + "M" : "?";
        const sp  = typeof p.totalSpoilage === "number" ? (p.totalSpoilage * 100).toFixed(1) + "%" : "?";
        const flag = p.isAggregate ? " [AGG]" : p.isPHCCan ? " [PHC-can]" : p.isPHC ? " [PHC-slug]" : " [non-PHC]";
        s += `  ${p.plant}${flag}: prod ${typeof p.production === "number" ? p.production.toFixed(1) + "M" : "?"}, eff ${e} (target ${tgt}), vsAOP ${vAOP}, vsLL ${vLL}, spoilage ${sp}\n`;
      });
      if (f.comments?.length) {
        s += `  Comments (${f.comments.length}):\n`;
        f.comments.slice(0,20).forEach(c => s += `    [${c.date||"?"}] ${c.plant||"?"} ${c.line}: ${c.text.slice(0,120)}\n`);
        if (f.comments.length>20) s += `    ...+${f.comments.length-20} more\n`;
      }
    });
  }
  if (cur) {
    s += `\n=== CURRENT FILE: ${cur.filename} ===\n`;
    cur.summary?.forEach(p => {
      const e   = typeof p.efficiency === "number" ? (p.efficiency * 100).toFixed(1) + "%" : "?";
      const tgt = typeof p.efficiencyTarget === "number" ? (p.efficiencyTarget * 100).toFixed(1) + "%" : "?";
      const vAOP = typeof p.varAOP === "number" ? (p.varAOP > 0 ? "+" : "") + p.varAOP.toFixed(2) + "M" : "?";
      const vLL  = typeof p.varLatestLL === "number" ? (p.varLatestLL > 0 ? "+" : "") + p.varLatestLL.toFixed(2) + "M" : "?";
      const sp  = typeof p.totalSpoilage === "number" ? (p.totalSpoilage * 100).toFixed(1) + "%" : "?";
      const flag = p.isAggregate ? " [AGG]" : p.isPHCCan ? " [PHC-can]" : p.isPHC ? " [PHC-slug]" : " [non-PHC]";
      s += `${p.plant}${flag}: prod ${typeof p.production === "number" ? p.production.toFixed(1) + "M" : "?"}, eff ${e} (target ${tgt}), vsAOP ${vAOP}, vsLL ${vLL}, spoilage ${sp}\n`;
    });
    if (cur.monthlyPlan && Object.keys(cur.monthlyPlan).length) {
      const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
      const planStr = months.filter(m => cur.monthlyPlan[m] != null)
        .map(m => `${m}: ${typeof cur.monthlyPlan[m] === "number" ? cur.monthlyPlan[m].toFixed(1)+"M" : cur.monthlyPlan[m]}`).join(", ");
      s += `PHC Total Cans AOP monthly plan: ${planStr}\n`;
    }
    if (cur.comments?.length) {
      s += `\nGap sheet comments (${cur.comments.length} total):\n`;
      cur.comments.forEach(c => s += `  [${c.date||"?"}] ${c.plant||"?"} ${c.line}: ${c.text.slice(0,150)}\n`);
    }
  }
  return s;
}

async function askAI(question, cur, hist, chatHist, lang) {
  const ctx = buildCtx(cur, hist);
  const msgs = [
    ...chatHist.slice(-6).map(m => ({ role:m.role, content: m.role==="user"?`${ctx}\nlang=${lang}\nQuestion: ${m.text}`:m.text })),
    { role:"user", content:`${ctx}\nlang=${lang}\nQuestion: ${question}` }
  ];
  const res = await fetch("/api/chat", {
    method:"POST",
    headers:{ "Content-Type":"application/json" },
    body: JSON.stringify({ model:"claude-haiku-4-5-20251001", max_tokens:1000, system:SYS, messages:msgs })
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.find(b=>b.type==="text")?.text || "—";
}

/* ── Helpers ──────────────────────────────────────────────────────────────── */
function plantStatus(p) {
  const v = typeof p.varLatestLL === "number" ? p.varLatestLL : 0;
  const e = typeof p.efficiency === "number" ? p.efficiency : 0;
  const target = typeof p.efficiencyTarget === "number" ? p.efficiencyTarget : 0.55;
  if (v < -1.5 || e < target - 0.12) return "red";
  if (v < -0.5 || e < target - 0.04) return "amber";
  return "green";
}

function renderText(text) {
  return text.split("\n").map((line, i) => {
    const parts = line.split(/(\*\*.*?\*\*)/g).map((p, j) =>
      p.startsWith("**") ? <strong key={j} style={{ fontWeight:600 }}>{p.slice(2,-2)}</strong> : p
    );
    return <span key={i}>{parts}{i < text.split("\n").length-1 && <br/>}</span>;
  });
}

/* ── App ──────────────────────────────────────────────────────────────────── */
export default function App() {
  const [lang, setLang]         = useState("cs");
  const [cur, setCur]           = useState(null);
  const [hist, setHist]         = useState([]);
  const [msgs, setMsgs]         = useState([]);
  const [input, setInput]       = useState("");
  const [loading, setLoading]   = useState(false);
  const [parsing, setParsing]   = useState(false);
  const [tab, setTab]           = useState("chat");
  const [dragOver, setDragOver] = useState(false);
  const fileRef  = useRef();
  const bottomRef = useRef();
  const ready    = useRef(false);
  const t = T[lang];

  // Inject global CSS once
  useEffect(() => {
    if (!document.getElementById("pa-global")) {
      const s = document.createElement("style"); s.id = "pa-global"; s.textContent = GLOBAL_CSS; document.head.appendChild(s);
    }
  }, []);

  useEffect(() => {
    if (ready.current) return; ready.current = true;
    Promise.all([loadHistory(), loadLang()]).then(([h, l]) => {
      setHist(h); setLang(l);
      setMsgs([{ role:"assistant", text: h.length>0 ? T[l].history_welcome(h.length) : T[l].welcome }]);
    });
  }, []);

  const handleFile = useCallback(async f => {
    if (!f) return;
    setParsing(true);
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type:"array" });
      const parsed = parseFile(wb, f.name);
      setCur(parsed);
      const newH = [...hist.filter(h=>h.filename!==f.name), parsed].slice(-10);
      setHist(newH); await saveHistory(newH);
      setMsgs(p => [...p, { role:"assistant", text: t.loaded_msg(f.name, parsed.summary?.length||0, parsed.comments?.length||0) }]);
      setTab("chat");
    } catch(e) { setMsgs(p => [...p, { role:"assistant", text: t.error_load+e.message }]); }
    setParsing(false);
  }, [hist, t]);

  const send = useCallback(async q => {
    if (!q.trim()||loading) return;
    setInput("");
    setMsgs(p => [...p, { role:"user", text:q }]);
    setLoading(true);
    try {
      const reply = await askAI(q, cur, hist, msgs, lang);
      setMsgs(p => [...p, { role:"assistant", text:reply }]);
    } catch(e) { setMsgs(p => [...p, { role:"assistant", text: t.error_ai(e.message) }]); }
    setLoading(false);
    setTimeout(() => bottomRef.current?.scrollIntoView({ behavior:"smooth" }), 80);
  }, [loading, cur, hist, msgs, lang, t]);

  const switchLang = l => { setLang(l); saveLang(l); };
  const plants = cur?.summary||[];
  const comments = cur?.comments||[];

  /* ── Render ─────────────────────────────────────────────────────────────── */
  return (
    <div style={{ height:"100vh", display:"flex", flexDirection:"column", background:D.bg, overflow:"hidden" }}>

      {/* Top bar */}
      <div style={{ background:D.surface, borderBottom:`1px solid ${D.border}`, padding:"0 24px", height:56, display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0, zIndex:10 }}>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          <div style={{ width:32, height:32, borderRadius:8, background:D.accentBg, display:"flex", alignItems:"center", justifyContent:"center" }}>
            <span style={{ fontSize:16 }}>📊</span>
          </div>
          <div>
            <div style={{ fontSize:15, fontWeight:600, color:D.textPrimary, lineHeight:1.2 }}>{t.title}</div>
            <div style={{ fontSize:11, color:D.textTertiary }}>{t.tagline}</div>
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          {/* Lang toggle */}
          <div style={{ display:"flex", background:D.bg, border:`1px solid ${D.border}`, borderRadius:8, overflow:"hidden" }}>
            {["cs","en"].map(l => (
              <button key={l} onClick={() => switchLang(l)}
                style={{ border:"none", padding:"5px 14px", cursor:"pointer", fontSize:12, fontWeight: lang===l?600:400,
                  background: lang===l ? D.surface : "transparent",
                  color: lang===l ? D.textPrimary : D.textSecondary,
                  borderRight: l==="cs" ? `1px solid ${D.border}` : "none" }}>
                {l.toUpperCase()}
              </button>
            ))}
          </div>
          {/* Upload button */}
          <button onClick={() => fileRef.current?.click()}
            style={{ display:"flex", alignItems:"center", gap:6, border:`1px solid ${D.border}`, borderRadius:8,
              background: D.surface, color: D.textPrimary, padding:"6px 14px", cursor:"pointer", fontSize:13, fontWeight:500,
              transition:"border-color .15s", boxShadow:"0 1px 2px rgba(0,0,0,.05)" }}
            onMouseEnter={e=>e.currentTarget.style.borderColor=D.borderHover}
            onMouseLeave={e=>e.currentTarget.style.borderColor=D.border}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
            </svg>
            {parsing ? t.uploading : t.upload_btn}
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display:"none" }} onChange={e=>handleFile(e.target.files[0])} />
        </div>
      </div>

      {/* Main */}
      <div style={{ flex:1, overflow:"hidden", display:"flex", flexDirection:"column", maxWidth:900, width:"100%", margin:"0 auto", padding:"20px 20px 16px" }}>

        {/* Drop zone */}
        {!cur && hist.length===0 && (
          <div onDrop={e=>{e.preventDefault();setDragOver(false);handleFile(e.dataTransfer.files[0]);}}
            onDragOver={e=>{e.preventDefault();setDragOver(true);}}
            onDragLeave={()=>setDragOver(false)}
            onClick={() => fileRef.current?.click()}
            style={{ border:`2px dashed ${dragOver?D.accent:D.borderHover}`, borderRadius:16, padding:"56px 32px",
              textAlign:"center", cursor:"pointer", background: dragOver?D.accentBg:D.surface,
              transition:"all .2s", marginBottom:24 }}>
            <div style={{ fontSize:40, marginBottom:12 }}>📂</div>
            <div style={{ fontSize:16, fontWeight:600, color:D.textPrimary, marginBottom:6 }}>{t.drop_title}</div>
            <div style={{ fontSize:13, color:D.textSecondary }}>{t.drop_sub}</div>
          </div>
        )}

        {/* File info bar */}
        {cur && (
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", background:D.surface,
            border:`1px solid ${D.border}`, borderRadius:10, padding:"10px 16px", marginBottom:16 }}>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <span style={{ fontSize:20 }}>📄</span>
              <div>
                <div style={{ fontSize:13, fontWeight:600, color:D.textPrimary }}>{cur.filename}</div>
                <div style={{ fontSize:11, color:D.textSecondary }}>{cur.period||"?"} · {cur.summary?.length||0} plants · {cur.comments?.length||0} comments</div>
              </div>
            </div>
            <button onClick={() => fileRef.current?.click()}
              style={{ fontSize:12, color:D.textSecondary, border:`1px solid ${D.border}`, borderRadius:6,
                background:"transparent", padding:"4px 10px", cursor:"pointer" }}>
              {t.upload_btn}
            </button>
          </div>
        )}

        {/* Plant cards */}
        {plants.length > 0 && (
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(140px,1fr))", gap:10, marginBottom:20 }}>
            {plants.map(p => {
              const s = plantStatus(p); const c = D[s];
              const eff = typeof p.efficiency==="number"?(p.efficiency*100).toFixed(0)+"%":"?";
              const v   = typeof p.varLatestLL==="number"?(p.varLatestLL>0?"+":"")+p.varLatestLL.toFixed(1)+"M":"?";
              const spoil = typeof p.totalSpoilage==="number"?(p.totalSpoilage*100).toFixed(1)+"%":null;
              const tgt = typeof p.efficiencyTarget==="number" && p.efficiencyTarget>0
                ? "cíl "+( p.efficiencyTarget*100).toFixed(0)+"%" : null;
              return (
                <button key={p.plant} onClick={() => send(t.detail_prompt(p.plant))}
                  style={{ background:c.bg, border:`1.5px solid ${c.brd}`, borderRadius:10, padding:"12px 14px",
                    cursor:"pointer", textAlign:"left", transition:"transform .1s, box-shadow .1s",
                    boxShadow:"0 1px 3px rgba(0,0,0,.05)" }}
                  onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-1px)";e.currentTarget.style.boxShadow="0 4px 12px rgba(0,0,0,.1)";}}
                  onMouseLeave={e=>{e.currentTarget.style.transform="";e.currentTarget.style.boxShadow="0 1px 3px rgba(0,0,0,.05)";}}>
                  <div style={{ fontSize:12, fontWeight:600, color:c.txt, marginBottom:6 }}>{p.plant}</div>
                  <div style={{ fontSize:13, fontWeight:700, color:c.sub }}>{eff}</div>
                  {tgt && <div style={{ fontSize:10, color:c.txt, marginTop:1, opacity:.7 }}>{tgt}</div>}
                  <div style={{ fontSize:11, color:c.txt, marginTop:2, opacity:.8 }}>{v} {t.vs_ll}</div>
                  {spoil && <div style={{ fontSize:10, color:c.sub, marginTop:3, opacity:.7 }}>⚠ {spoil}</div>}
                </button>
              );
            })}
          </div>
        )}

        {/* Tabs */}
        {(cur || hist.length > 0) && (
          <div style={{ display:"flex", gap:0, marginBottom:0, borderBottom:`1px solid ${D.border}` }}>
            {[["chat", t.tab_chat],["comments", t.tab_comments(comments.length)],["history", t.tab_history(hist.length)]].map(([id,label]) => (
              <button key={id} onClick={() => setTab(id)}
                style={{ border:"none", borderBottom: tab===id?`2px solid ${D.accent}`:"2px solid transparent",
                  background:"transparent", padding:"10px 18px 11px", fontSize:13, cursor:"pointer",
                  color: tab===id ? D.accent : D.textSecondary,
                  fontWeight: tab===id ? 600 : 400, transition:"color .15s", marginBottom:-1 }}>
                {label}
              </button>
            ))}
          </div>
        )}

        {/* Chat panel */}
        {tab === "chat" && (
          <div style={{ flex:1, display:"flex", flexDirection:"column", background:D.surface, border:`1px solid ${D.border}`, borderTop:"none",
            borderRadius:"0 0 12px 12px", overflow:"hidden" }}>
            {/* Messages */}
            <div style={{ flex:1, overflowY:"auto", padding:"20px 20px 8px" }}>
              {msgs.map((m,i) => (
                <div key={i} style={{ marginBottom:16, display:"flex", gap:10, alignItems:"flex-start",
                  flexDirection: m.role==="user"?"row-reverse":"row" }}>
                  {/* Avatar */}
                  <div style={{ width:28, height:28, borderRadius:8, flexShrink:0, display:"flex", alignItems:"center", justifyContent:"center", fontSize:14,
                    background: m.role==="user" ? D.accentBg : D.bg, border:`1px solid ${D.border}` }}>
                    {m.role==="user" ? "👤" : "🤖"}
                  </div>
                  {/* Bubble */}
                  <div style={{ maxWidth:"80%", borderRadius: m.role==="user"?"12px 4px 12px 12px":"4px 12px 12px 12px",
                    padding:"10px 14px", fontSize:13, lineHeight:1.65,
                    background: m.role==="user" ? D.accentBg : D.bg,
                    color: m.role==="user" ? D.accentText : D.textPrimary,
                    border:`1px solid ${m.role==="user"?D.accentBg:D.border}` }}>
                    {renderText(m.text)}
                  </div>
                </div>
              ))}
              {loading && (
                <div style={{ display:"flex", gap:10, marginBottom:16 }}>
                  <div style={{ width:28, height:28, borderRadius:8, background:D.bg, border:`1px solid ${D.border}`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:14 }}>🤖</div>
                  <div style={{ borderRadius:"4px 12px 12px 12px", padding:"10px 14px", background:D.bg, border:`1px solid ${D.border}`, fontSize:13, color:D.textSecondary }}>
                    <span style={{ animation:"pulse 1s infinite" }}>{t.analyzing}</span>
                  </div>
                </div>
              )}
              <div ref={bottomRef} />
            </div>
            {/* Quick questions */}
            <div style={{ padding:"8px 20px 10px", display:"flex", flexWrap:"wrap", gap:6, flexShrink:0, borderTop:`1px solid ${D.border}` }}>
              {t.quick.map(q => (
                <button key={q} onClick={() => send(q)} disabled={loading}
                  style={{ border:`1px solid ${D.border}`, borderRadius:20, background:D.bg,
                    color:D.textSecondary, fontSize:12, padding:"5px 12px", cursor:"pointer",
                    transition:"all .15s", opacity: loading ? 0.5 : 1 }}
                  onMouseEnter={e=>{e.currentTarget.style.background=D.accentBg;e.currentTarget.style.color=D.accentText;e.currentTarget.style.borderColor=D.accent;}}
                  onMouseLeave={e=>{e.currentTarget.style.background=D.bg;e.currentTarget.style.color=D.textSecondary;e.currentTarget.style.borderColor=D.border;}}>
                  {q}
                </button>
              ))}
            </div>
            {/* Input */}
            <div style={{ padding:"10px 16px 16px", display:"flex", gap:8, flexShrink:0 }}>
              <input value={input} onChange={e=>setInput(e.target.value)}
                onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&send(input)}
                placeholder={t.input_placeholder} disabled={loading}
                style={{ flex:1, border:`1px solid ${D.border}`, borderRadius:10, padding:"10px 14px",
                  fontSize:13, color:D.textPrimary, background:D.bg,
                  transition:"border-color .15s" }} />
              <button onClick={() => send(input)} disabled={loading||!input.trim()}
                style={{ border:"none", borderRadius:10, background: loading||!input.trim() ? D.border : D.accent,
                  color: loading||!input.trim() ? D.textTertiary : "#fff",
                  padding:"10px 16px", cursor: loading||!input.trim()?"not-allowed":"pointer",
                  fontSize:13, fontWeight:500, transition:"background .15s" }}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
                </svg>
              </button>
            </div>
          </div>
        )}

        {/* Comments panel */}
        {tab === "comments" && (
          <div style={{ flex:1, display:"flex", flexDirection:"column", background:D.surface, border:`1px solid ${D.border}`, borderTop:"none", borderRadius:"0 0 12px 12px", overflow:"hidden" }}>
            {comments.length === 0
              ? <div style={{ padding:"40px", textAlign:"center", color:D.textSecondary, fontSize:13 }}>{t.no_comments}</div>
              : <div style={{ flex:1, overflowY:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse" }}>
                    <thead>
                      <tr style={{ background:D.bg }}>
                        {[t.date_col, t.plant_col, t.line_col, t.comments_col].map(h => (
                          <th key={h} style={{ padding:"10px 16px", fontSize:11, fontWeight:600, color:D.textSecondary,
                            textAlign:"left", borderBottom:`1px solid ${D.border}`, letterSpacing:".04em", textTransform:"uppercase" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {comments.map((c,i) => (
                        <tr key={i} style={{ borderBottom:`1px solid ${D.border}` }}
                          onMouseEnter={e=>e.currentTarget.style.background=D.bg}
                          onMouseLeave={e=>e.currentTarget.style.background=""}>
                          <td style={{ padding:"10px 16px", fontSize:12, color:D.textSecondary, whiteSpace:"nowrap", verticalAlign:"top" }}>{c.date||"—"}</td>
                          <td style={{ padding:"10px 16px", fontSize:12, fontWeight:500, color:D.textPrimary, whiteSpace:"nowrap", verticalAlign:"top" }}>{c.plant||"—"}</td>
                          <td style={{ padding:"10px 16px", fontSize:12, color:D.textSecondary, whiteSpace:"nowrap", verticalAlign:"top" }}>{c.line}</td>
                          <td style={{ padding:"10px 16px", fontSize:13, color:D.textPrimary, lineHeight:1.55, verticalAlign:"top" }}>{c.text}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
            }
          </div>
        )}

        {/* History panel */}
        {tab === "history" && (
          <div style={{ flex:1, display:"flex", flexDirection:"column", background:D.surface, border:`1px solid ${D.border}`, borderTop:"none", borderRadius:"0 0 12px 12px", overflow:"hidden" }}>
            {hist.length === 0
              ? <div style={{ padding:"40px", textAlign:"center", color:D.textSecondary, fontSize:13 }}>{t.no_history}</div>
              : <>
                  <div style={{ flex:1, overflowY:"auto" }}>
                    {[...hist].reverse().map((h,i) => (
                      <div key={i} style={{ padding:"14px 20px", borderBottom:`1px solid ${D.border}` }}>
                        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:12, marginBottom:10 }}>
                          <div>
                            <div style={{ fontSize:13, fontWeight:600, color:D.textPrimary }}>{h.filename}</div>
                            <div style={{ fontSize:11, color:D.textSecondary, marginTop:2 }}>
                              {h.period||"?"} · {t.uploaded_on} {h.uploadedAt?.slice(0,10)} · {h.comments?.length||0} comments
                            </div>
                          </div>
                          <button onClick={() => send(t.compare_prompt(h.filename))}
                            style={{ border:`1px solid ${D.border}`, borderRadius:6, background:"transparent",
                              color:D.textSecondary, fontSize:11, padding:"4px 10px", cursor:"pointer", whiteSpace:"nowrap",
                              fontWeight:500 }}>
                            {t.compare_btn}
                          </button>
                        </div>
                        <div style={{ display:"flex", flexWrap:"wrap", gap:5 }}>
                          {h.summary?.map(p => {
                            const bad = (typeof p.varLatestLL==="number"&&p.varLatestLL<-0.5)||(typeof p.efficiency==="number"&&p.efficiency<0.47);
                            const eff = typeof p.efficiency==="number"?(p.efficiency*100).toFixed(0)+"%":"?";
                            const v   = typeof p.varLatestLL==="number"?(p.varLatestLL>0?"+":"")+p.varLatestLL.toFixed(1)+"M":"?";
                            const c   = bad ? D.red : D.gray;
                            return (
                              <span key={p.plant} style={{ fontSize:11, padding:"3px 8px", borderRadius:5,
                                background:c.bg, color:c.txt, border:`1px solid ${c.brd}`, fontWeight:500 }}>
                                {p.plant} · {eff} · {v}
                              </span>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                  </div>
                  <div style={{ padding:"12px 20px", borderTop:`1px solid ${D.border}`, display:"flex", justifyContent:"flex-end" }}>
                    <button onClick={async()=>{setHist([]);await saveHistory([]);}}
                      style={{ border:`1px solid ${D.border}`, borderRadius:6, background:"transparent",
                        color:D.textSecondary, fontSize:12, padding:"5px 12px", cursor:"pointer" }}>
                      {t.clear_history}
                    </button>
                  </div>
                </>
            }
          </div>
        )}
      </div>
    </div>
  );
}