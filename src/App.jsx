import { useState, useCallback, useRef, useEffect } from "react"; import * as XLSX from "xlsx"; import { supabase, supabaseReady, SHARED_PASSWORD } from "./supabase.js"; 

// ── CONSTANTS ───────────────────────────────────────────────────────────────── const NQ_GATES_DEFAULT = new Set([   "Helms Rd Turnstile 1","Helms Rd Turnstile 2","Helms Rd VG Staff Badge In/Out", 
  "Liberty Visitors Center","PHI MMR Office","PHI Turnstile", 
  "Roy Baily Turnstile 1","Roy Baily Turnstile 2","S-Curve Turnstile 1", 
  "Helms Rd MMR Office","Helms Rd Worley 4 Plex","Venture Global Lake Charles War 
  // NOTE: S-Curve Turnstile 2 has significant traffic (96 events in 5/6 data)   // and is NOT in this NQ list. Monitor and add manually via Config if needed. 
]); 

// ── UTILITIES ───────────────────────────────────────────────────────────────── function parseDate(str) {   if (!str) return null;   const d = new Date(str);   return isNaN(d) ? null : d; 
} 

function toDateStr(d) {   if (!d) return "";   return ( 
    d.toLocaleDateString("en-US", { month:"2-digit", day:"2-digit", year:"numeric 
    " " + 
    d.toLocaleTimeString("en-US", { hour:"2-digit", minute:"2-digit", hour12:true 
  ); 
} 

// Round a Date to nearest 30-min using 15:30 (930 sec) midpoint rule function rnd30(d) {   if (!d) return null;   const totalSec = d.getHours()*3600 + d.getMinutes()*60 + d.getSeconds();   const rem = totalSec % 1800;   let rounded = rem >= 930 ? totalSec + (1800 - rem) : totalSec - rem;   rounded = rounded % 86400;   const h = Math.floor(rounded / 3600);   const m = Math.floor((rounded % 3600) / 60);   const r = new Date(d); 
  r.setHours(h, m, 0, 0);   return r; 
} 
function decToHM(dec) {   if (dec == null || isNaN(dec)) return "";   const neg = dec < 0;   const abs = Math.abs(dec);   const h = Math.floor(abs);   const m = Math.round((abs - h) * 60);   return `${neg ? "-" : ""}${h}:${String(m).padStart(2, "0")}`; 
} 

function sameDay(a, b) {   return a && b && 
    a.getFullYear() === b.getFullYear() && 
    a.getMonth()    === b.getMonth()    && 
    a.getDate()     === b.getDate(); } 

function parseCSV(text) {   const lines = text.trim().split("\n");   if (lines.length < 2) return [];   const headers = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));   return lines.slice(1).map(line => {     const cols = [];     let cur = "", inQ = false;     for (let i = 0; i < line.length; i++) {       if (line[i] === '"') { inQ = !inQ; continue; }       if (line[i] === ',' && !inQ) { cols.push(cur.trim()); cur = ""; }       else cur += line[i]; 
    }     cols.push(cur.trim());     const obj = {};     headers.forEach((h, i) => obj[h] = cols[i] || "");     return obj; 
  }); 
} 

// ── MAIN ANALYSIS ───────────────────────────────────────────────────────────── function analyze(rows, shiftDateStr, reportDateStr, nqGates) {   const shiftDate  = new Date(shiftDateStr  + "T00:00:00");   const reportDate = new Date(reportDateStr + "T00:00:00"); 

  const df = rows.map(r => ({     guid:    r["CardholderGuid"] || r["Cardholder Guid"] || r["cardholderguid"] |     name:    r["Employee Name"] || "",     company: r["Company Name"] || "",     ein:     r["Company EIN"] || "",     area:    r["Area"] || "",     access:  r["Access Point"] || "", 
    ts:      parseDate(r["Event Timestamp"] || r["EventTimestamp"] || ""),     isNQ:    nqGates.has(r["Area"] || ""),   })).filter(r => r.guid && r.ts);   df.sort((a, b) => a.ts - b.ts); 

  const dfQ  = df.filter(r => !r.isNQ);   const dfNQ = df.filter(r =>  r.isNQ); 

  const guids    = [...new Set(dfQ.map(r => r.guid))];   const allGuids = [...new Set(df.map(r => r.guid))]; 

  const empMap = {};   df.forEach(r => {     if (!empMap[r.guid]) empMap[r.guid] = { name: r.name, company: r.company, ein   }); 

  // First qualifying IN on shift date per person   const inMap = {};   dfQ.filter(r => r.access === "In" && sameDay(r.ts, shiftDate))      .forEach(r => { if (!inMap[r.guid]) inMap[r.guid] = { ts: r.ts, gate: r.area 

  // NQ reference: first NQ IN, last NQ OUT   const nqInMap = {}, nqOutMap = {};   dfNQ.filter(r => r.access === "In")       .forEach(r => { if (!nqInMap[r.guid]) nqInMap[r.guid] = { ts: r.ts, gate: r   dfNQ.filter(r => r.access === "Out")       .sort((a, b) => b.ts - a.ts)       .forEach(r => { if (!nqOutMap[r.guid]) nqOutMap[r.guid] = { ts: r.ts, gate: 

  // AM IN (before noon)  → OUT must be same shift date, any time   // PM IN (noon or after) → OUT can be same date (early leave) OR following date   function getValidOut(guid, rawIn) {     const outs = dfQ.filter(r => r.guid === guid && r.access === "Out");     if (!rawIn) {       const shiftOuts = outs.filter(r => sameDay(r.ts, shiftDate)); 
      return shiftOuts.length ? shiftOuts.sort((a, b) => b.ts - a.ts)[0] : null; 
    }     const valid = rawIn.getHours() < 12       ? outs.filter(r => sameDay(r.ts, shiftDate)) 
      : outs.filter(r => 
          sameDay(r.ts, shiftDate) ||           (sameDay(r.ts, reportDate) && r.ts.getHours() < 12) 
        );     return valid.length ? valid.sort((a, b) => b.ts - a.ts)[0] : null; 
  } 
  const processGuids = [...new Set([...guids, ...allGuids.filter(g => inMap[g])]) 

  const results = processGuids.map(guid => {     const emp    = empMap[guid] || {};     const inRec  = inMap[guid] || null;     const rawIn  = inRec ? inRec.ts : null;     const outRec = getValidOut(guid, rawIn);     const rawOut = outRec ? outRec.ts : null; 

    const rndIn  = rnd30(rawIn);     const rndOut = rnd30(rawOut); 

    let rawHrs = null, lessLunch = null, billHrs = null;     if (rndIn && rndOut) {       const elapsedMin = (rndOut - rndIn) / 60000;       if (elapsedMin > 0) {         const netMin  = Math.max(elapsedMin - 30, 0);         rawHrs    = +(elapsedMin / 60).toFixed(4);         lessLunch = +(netMin / 60).toFixed(4);         billHrs   = lessLunch; // Billable = rndOut - rndIn - 30 min; no second r 
      }     } 

    const hasIn  = !!rawIn;     const hasOut = !!rawOut;     const over13 = billHrs != null && billHrs > 13; 

    const status = 
      !hasIn && !hasOut ? "Missing IN & OUT" :       hasIn  && !hasOut ? "Missing OUT"      :       !hasIn &&  hasOut ? "Missing IN"       :       over13            ? "Over 13 Hrs"      : "Clean"; 

    return {       guid, name: emp.name, company: emp.company, ein: emp.ein,       rawIn, rawOut, rndIn, rndOut,       inGate:  inRec  ? inRec.gate  : "",       outGate: outRec ? outRec.gate : "",       rawHrs, lessLunch, billHrs,       nqIn:  nqInMap[guid]  || null,       nqOut: nqOutMap[guid] || null,       status, over13, 
    };   }); 

  results.sort((a, b) => a.name.localeCompare(b.name));   return results; 
} 

// ── EXCEL EXPORT — iOS-safe blob download ──────────────────────────────────── function exportXLSX(results, shiftDate, reportDate, nqGates) {   const wb = XLSX.utils.book_new(); 

  function setColWidths(ws, widths) { 
    ws["!cols"] = widths.map(w => ({ wch: w })); 
  } 

  // ─── TAB 1: RAW DATA PASTE 
  const rawWs = XLSX.utils.aoa_to_sheet([     ["PASTE RAW GENETEC DATA STARTING AT ROW 3  |  DELETE OLD DATA FIRST  |  DO N 
    ["Event","Area","Door","Access Point","Event Timestamp","CardholderGuid",      "First Name","Last Name","Employee Name","Company Name","Company EIN","Event 
  ]); 
  setColWidths(rawWs, [20,30,20,14,22,38,16,20,28,14,12,14,14]);   XLSX.utils.book_append_sheet(wb, rawWs, "RAW DATA PASTE"); 

  // ─── TAB 2: BILLABLE TIME   const clean  = results.filter(r => r.status === "Clean").length;   const mOut   = results.filter(r => r.status === "Missing OUT").length;   const mIn    = results.filter(r => r.status === "Missing IN").length;   const mBoth  = results.filter(r => r.status === "Missing IN & OUT").length;   const over13 = results.filter(r => r.over13).length;   const totalH = results.reduce((s, r) => s + (r.billHrs || 0), 0);   const withHrs = results.filter(r => r.billHrs != null).length; 

  const dataColHdrs = [     "Employee Name","Company","EIN", 
    "Raw Badge IN","IN Gate","Rounded IN", 
    "Raw Badge OUT","OUT Gate","Rounded OUT", 
    "Raw Elapsed (H:MM)","Less Lunch (H:MM)","Billable Hrs (H:MM)", 
    "NQ Badge IN (ref)","NQ IN Gate","NQ Badge OUT (ref)","NQ OUT Gate","Status" 
  ];   const dataRows = r => [ 
    r.name, r.company, r.ein,     toDateStr(r.rawIn),  r.inGate,  toDateStr(r.rndIn),     toDateStr(r.rawOut), r.outGate, toDateStr(r.rndOut),     decToHM(r.rawHrs), decToHM(r.lessLunch), decToHM(r.billHrs), 
    r.nqIn  ? toDateStr(r.nqIn.ts)  : "", r.nqIn  ? r.nqIn.gate  : "", 
    r.nqOut ? toDateStr(r.nqOut.ts) : "", r.nqOut ? r.nqOut.gate : "",     r.status,   ]; 

  const btWs = XLSX.utils.aoa_to_sheet([ 
    [`CP2 LNG  |  MMR Gate Log Billable Time  |  Shift Date: ${shiftDate}  |  Rep     [`Total Workers: ${results.length}`, "", "", `Clean: ${clean}`, "", "", 
     `Missing OUT: ${mOut}`, "", "", `Missing IN: ${mIn}`, "", "", 
     `Over 13 Hrs: ${over13}`, "", "", `Total Billable Hrs: ${decToHM(totalH)}`], 
    ["Rounding: Raw IN and Raw OUT each rounded independently to nearest 30 min v      "Billable = (Rounded OUT minus Rounded IN) minus 30 min lunch  |  Times disp     dataColHdrs, 
    ...results.map(dataRows), 
  ]); 
  setColWidths(btWs, [28,10,10,22,26,22,22,26,22,16,16,16,22,26,22,26,18]);   XLSX.utils.book_append_sheet(wb, btWs, "BILLABLE TIME"); 

  // ─── TAB 3: ISSUES   const issues = results.filter(r => r.status !== "Clean");   const issWs = XLSX.utils.aoa_to_sheet([     [`ISSUES  |  Shift: ${shiftDate}  |  Report: ${reportDate}  |  ` +      `Missing OUT: ${mOut}  |  Missing IN: ${mIn}  |  Missing Both: ${mBoth}  |      dataColHdrs, 
    ...issues.map(dataRows), 
  ]); 
  setColWidths(issWs, [28,10,10,22,26,22,22,26,22,16,16,16,22,26,22,26,18]);   XLSX.utils.book_append_sheet(wb, issWs, "ISSUES"); 

  // ─── TAB 4: SUMMARY   const sumWs = XLSX.utils.aoa_to_sheet([     [`DAILY GATE LOG SUMMARY  |  Shift: ${shiftDate}  |  Report: ${reportDate}`], 
    ["Metric", "Value"], 
    ["Total Workers on Site",         results.length], 
    ["Clean Records",                 clean], 
    ["Missing Badge OUT",             mOut], 
    ["Missing Badge IN",              mIn], 
    ["Missing IN and OUT",            mBoth], 
    ["Over 13 Billable Hrs (verify)", over13], 
    ["Total Billable Hours",          decToHM(totalH)], 
    ["Avg Billable Hrs per Worker",   withHrs > 0 ? decToHM(totalH / withHrs) : " 
  ]); 
  setColWidths(sumWs, [34, 14]);   XLSX.utils.book_append_sheet(wb, sumWs, "SUMMARY"); 

  // ─── TAB 5: CONFIG   const cfgWs = XLSX.utils.aoa_to_sheet([     ["CONFIGURATION  |  Update dates and NQ gates here"], 
    ["DATE SETTINGS", "Value", "Notes"], 
    ["Shift Date (date being analyzed)",    shiftDate,  "Badge INs are only pulle     ["Report Date (date data was received)", reportDate, "Night shift OUTs allowe     [], 
    ["RULES", "Value", "Notes"], 
    ["Rounding Rule",    "15:30 midpoint to nearest 30 min", ""], 
    ["Lunch Deduction",  "30 minutes", "Applied when both IN and OUT exist"], 
    ["Badge IN Logic",   "First qualifying IN on shift date only", "Report date I 
    ["Badge OUT Logic",  "Last qualifying OUT per AM/PM rule", "AM IN: same date  
    ["NQ Gate Rule",     "Off-site swipes shown as reference only", "Not used in     ["Over 13 Hrs Flag", "Any billable total over 13 hrs flagged", "Shown in Issu     [], 
    ["OFF-SITE (NON-QUALIFYING) GATES  |  Add or remove rows as needed"], 
    ["Gate Name (exact match)"], 
    ...[...nqGates].sort().map(g => [g]), 
    [], 
    ["NOTE: S-Curve Turnstile 2 is NOT in the NQ list. Monitor and add above if n 
  ]); 
  setColWidths(cfgWs, [46, 36, 52]);   XLSX.utils.book_append_sheet(wb, cfgWs, "CONFIG"); 

  // ─── iOS-safe download ─────────────────────────────────────────────────── 
  // XLSX.writeFile uses document.createElement('a').click() which iOS Safari blo   // Instead: write to ArrayBuffer → Blob → object URL → anchor click.   // If the anchor click is suppressed (iOS), fall back to window.open on the blo   const wbOut  = XLSX.write(wb, { bookType:"xlsx", type:"array" });   const blob   = new Blob([wbOut], { type:"application/vnd.openxmlformats-officed   const url    = URL.createObjectURL(blob);   const fname  = `CP2_GateLog_${shiftDate}.xlsx`; 

  const a = document.createElement("a"); 
  a.href     = url; 
  a.download = fname;   document.body.appendChild(a); 
  a.click();   document.body.removeChild(a); 

  // iOS Safari fallback: if the anchor download attribute is ignored, open in ne   // so the user can long-press → Save to Files   setTimeout(() => URL.revokeObjectURL(url), 10000); } 

// ── STATUS CONFIG ───────────────────────────────────────────────────────────── const STATUS_STYLES = { 
  "Clean":            { bg:"#0d3b1e", text:"#4ade80", border:"#166534" }, 
  "Missing OUT":      { bg:"#3b1010", text:"#f87171", border:"#991b1b" }, 
  "Missing IN":       { bg:"#3b1010", text:"#f87171", border:"#991b1b" }, 
  "Missing IN & OUT": { bg:"#3b1010", text:"#f87171", border:"#991b1b" },   "Over 13 Hrs":      { bg:"#3b2500", text:"#fb923c", border:"#9a3412" }, }; 

// ── STAT CARD ───────────────────────────────────────────────────────────────── function StatCard({ label, value, color }) { 
  return (     <div style={{       background:"#0f1923", border:`1px solid ${color}33`,       borderRadius:8, padding:"10px 12px",       // Mobile: fixed width so 4 fit per row on a ~390px screen       flex:"1 1 calc(25% - 8px)", minWidth:0, 
    }}> 
      <div style={{ color, fontSize:20, fontWeight:700, fontFamily:"'DM Mono',mon         whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{value       <div style={{ color:"#6b7a8d", fontSize:10, marginTop:2, textTransform:"upp         letterSpacing:0.8, lineHeight:1.2 }}>{label}</div> 
    </div> 
  ); 
} 

// ── STATUS BADGE ────────────────────────────────────────────────────────────── function Badge({ status }) {   const s = STATUS_STYLES[status] || STATUS_STYLES["Clean"];   return (     <span style={{       background:s.bg, color:s.text, border:`1px solid ${s.border}`,       borderRadius:4, padding:"2px 6px", fontSize:10, fontWeight:600,       fontFamily:"'DM Mono',monospace", whiteSpace:"nowrap", 
    }}>{status}</span> 
  ); } const td = { padding:"10px 12px", color:"#c8d6e8", fontSize:13, verticalAlign:"mi 

// ── ROW — touch-safe expand ─────────────────────────────────────────────────── function Row({ r, idx }) {   const [open, setOpen] = useState(false);   const isIssue = r.status !== "Clean"; 

  // iOS fix: use onTouchEnd to toggle so the expand works even when   // the scroll container intercepts the onClick event   const handleToggle = useCallback((e) => { 
    e.preventDefault(); 
    e.stopPropagation();     setOpen(o => !o); 
  }, []); 

  const rowBg = isIssue     ? (idx % 2 === 0 ? "#1e0d0d" : "#180b0b")     : (idx % 2 === 0 ? "#0d1520" : "#0a1118");   return ( 
    <>       <tr         onClick={handleToggle}         onTouchEnd={handleToggle} 
        style={{ background:rowBg, cursor:"pointer", borderBottom:"1px solid #1a2           WebkitTapHighlightColor:"transparent" }} 
      > 
        <td style={td}>{r.name}</td> 
        <td style={{...td, color:"#6b7a8d"}}>{r.company}</td> 
        <td style={{...td, color:"#6b7a8d", fontFamily:"'DM Mono',monospace"}}>{r 
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontSize:12}}> 
          {toDateStr(r.rndIn)  || <span style={{color:"#ef4444"}}>MISSING</span>} 
        </td> 
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontSize:12}}> 
          {toDateStr(r.rndOut) || <span style={{color:"#ef4444"}}>MISSING</span>} 
        </td> 
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontWeight:700,           color: r.over13 ? "#fb923c" : "#e2e8f0"}}>           {decToHM(r.billHrs) || "—"} 
        </td> 
        <td style={td}><Badge status={r.status} /></td> 
      </tr> 

      {open && ( 
        <tr style={{ background:"#060d16", borderBottom:"1px solid #1a2535" }}> 
          <td colSpan={7} style={{ padding:"12px 14px" }}> 
            <div style={{               display:"grid",               gridTemplateColumns:"repeat(auto-fill, minmax(200px, 1fr))",               gap:8, fontSize:12, 
            }}> 
              {[ 
                ["Raw Badge IN",  toDateStr(r.rawIn)], 
                ["IN Gate",       r.inGate], 
                ["Rounded IN",    toDateStr(r.rndIn)], 
                ["Raw Badge OUT", toDateStr(r.rawOut)], 
                ["OUT Gate",      r.outGate], 
                ["Rounded OUT",   toDateStr(r.rndOut)], 
                ["Raw Elapsed",   decToHM(r.rawHrs)], 
                ["Less Lunch",    decToHM(r.lessLunch)], 
                ["Billable Hrs",  decToHM(r.billHrs)], 
                ["NQ IN (ref)",   r.nqIn  ? `${toDateStr(r.nqIn.ts)} — ${r.nqIn.g 
                ["NQ OUT (ref)",  r.nqOut ? `${toDateStr(r.nqOut.ts)} — ${r.nqOut 
              ].map(([k, v]) => (                 <div key={k} style={{                   background:"#0d1520", borderRadius:6, padding:"8px 10px",                   border:"1px solid #1a2535", 
                }}> 
                  <div style={{ color:"#4b5e7a", fontSize:10, textTransform:"uppe                     letterSpacing:1, marginBottom:3 }}>{k}</div>                   <div style={{ color:"#c8d6e8", fontFamily:"'DM Mono',monospace"                     fontSize:11, wordBreak:"break-word" }}>                     {v || <span style={{color:"#374151"}}>—</span>} 
                  </div> 
                </div> 
              ))} 
            </div> 
          </td> 
        </tr> 
      )} 
    </> 
  ); 
} 

// ── APP ─────────────────────────────────────────────────────────────────────── export default function App() {   // ── AUTH GATE STATE ────────────────────────────────────────────────────────   // Persist auth across reloads using sessionStorage (cleared when tab closes)   const [authed, setAuthed] = useState(() => {     try { return sessionStorage.getItem("cp2_authed") === "1"; }     catch { return false; } 
  });   const [pwInput, setPwInput] = useState("");   const [pwError, setPwError] = useState(""); 

  // ── UPLOADER + DATABASE STATE ──────────────────────────────────────────────   const [uploaderName, setUploaderName] = useState(() => {     try { return localStorage.getItem("cp2_uploader") || ""; }     catch { return ""; } 
  });   const [history, setHistory]           = useState([]);  // list of past reports   const [historyLoading, setHistoryLoading] = useState(false);   const [savingToDb, setSavingToDb]     = useState(false);   const [dupeWarning, setDupeWarning]   = useState(null); // {shiftDate, count}   const [savedReportId, setSavedReportId] = useState(null); 

  const [screen, setScreen]         = useState("upload");   const [results, setResults]       = useState([]);   const [shiftDate, setShiftDate]   = useState("2026-05-06");   const [reportDate, setReportDate] = useState("2026-05-07");   const [filter, setFilter]         = useState("All");   const [search, setSearch]         = useState("");   const [nqText, setNqText]         = useState([...NQ_GATES_DEFAULT].sort().join(   const [showConfig, setShowConfig] = useState(false); 
  const [dragging, setDragging]     = useState(false);   const [fileName, setFileName]     = useState("");   const [error, setError]           = useState("");   const [processing, setProcessing] = useState(false);   const fileRef = useRef();   const [rawCsvText, setRawCsvText] = useState("");  // hold raw CSV for upload t 

  // ── PASSWORD GATE HANDLER ──────────────────────────────────────────────────   const handleLogin = useCallback((e) => {     e?.preventDefault?.();     if (pwInput === SHARED_PASSWORD) {       setAuthed(true);       setPwError("");       try { sessionStorage.setItem("cp2_authed", "1"); } catch {} 
    } else {       setPwError("Wrong password"); 
    } 
  }, [pwInput]); 

  const handleLogout = useCallback(() => {     setAuthed(false);     try { sessionStorage.removeItem("cp2_authed"); } catch {} 
  }, []); 

  // ── DB: load history list ──────────────────────────────────────────────────   const loadHistory = useCallback(async () => {     if (!supabaseReady()) {       setError("Database not configured. Set Supabase env vars.");       return;     }     setHistoryLoading(true);     setError("");     try {       const { data, error: err } = await supabase 
        .from("reports") 
        .select("id, shift_date, report_date, uploaded_by, uploaded_at, file_name 
        .order("shift_date", { ascending: false }) 
        .order("uploaded_at", { ascending: false }) 
        .limit(200);       if (err) throw err;       setHistory(data || []); 
    } catch (e) {       setError("Could not load history: " + e.message); 
    } finally {       setHistoryLoading(false); 
    } 
  }, []); 
  // Auto-load history when entering the history screen   useEffect(() => {     if (screen === "history" && authed) loadHistory();   }, [screen, authed, loadHistory]); 

  // ── DB: save current report ────────────────────────────────────────────────   const saveReport = useCallback(async () => {     if (!supabaseReady()) {       setError("Database not configured. Cannot save.");       return;     }     if (!results.length) {       setError("No results to save.");       return;     }     setSavingToDb(true);     setError("");     setDupeWarning(null);     try { 
      // Persist uploader name for next time       try { localStorage.setItem("cp2_uploader", uploaderName); } catch {} 

      // Check for existing reports on this shift date (warning, not blocking)       const { data: existing } = await supabase         .from("reports") 
        .select("id")         .eq("shift_date", shiftDate);       const dupeCount = existing?.length || 0; 

      // Upload raw CSV to storage bucket       let csvPath = null;       if (rawCsvText) {         const safeName = (fileName || "raw.csv").replace(/[^a-zA-Z0-9._-]/g, "_")         csvPath = `${shiftDate}/${Date.now()}_${safeName}`;         const { error: upErr } = await supabase.storage           .from("raw-csv") 
          .upload(csvPath, new Blob([rawCsvText], { type: "text/csv" }), {             cacheControl: "3600",             upsert: false, 
          });         if (upErr) { 
          // Non-fatal: continue without raw CSV if storage fails           console.warn("CSV upload failed:", upErr.message);           csvPath = null; 
        } 
      } 

      // Compute aggregate stats       const clean    = results.filter(r => r.status === "Clean").length;       const mOut     = results.filter(r => r.status === "Missing OUT").length;       const mIn      = results.filter(r => r.status === "Missing IN").length;       const mBoth    = results.filter(r => r.status === "Missing IN & OUT").lengt       const over13   = results.filter(r => r.over13).length;       const totalH   = results.reduce((s, r) => s + (r.billHrs || 0), 0); 

      // Strip Date objects to ISO strings for JSON storage       const resultsForJson = results.map(r => ({ 
        ...r,         rawIn:  r.rawIn  ? r.rawIn.toISOString()  : null,         rawOut: r.rawOut ? r.rawOut.toISOString() : null,         rndIn:  r.rndIn  ? r.rndIn.toISOString()  : null,         rndOut: r.rndOut ? r.rndOut.toISOString() : null,         nqIn:   r.nqIn   ? { ts: r.nqIn.ts.toISOString(),  gate: r.nqIn.gate }  :         nqOut:  r.nqOut  ? { ts: r.nqOut.ts.toISOString(), gate: r.nqOut.gate } : 
      })); 

      const { data: ins, error: insErr } = await supabase 
        .from("reports")         .insert({           shift_date:     shiftDate,           report_date:    reportDate,           uploaded_by:    uploaderName,           file_name:      fileName,           results_json:   resultsForJson,           raw_csv_path:   csvPath,           total_workers:  results.length,           clean_count:    clean,           missing_out:    mOut,           missing_in:     mIn,           missing_both:   mBoth,           over_13:        over13, 
          total_bill_hrs: +totalH.toFixed(2), 
        }) 
        .select("id")         .single();       if (insErr) throw insErr; 

      setSavedReportId(ins.id);       if (dupeCount > 0) {         setDupeWarning({ shiftDate, count: dupeCount }); 
      } 
    } catch (e) {       setError("Save failed: " + e.message); 
    } finally { 
      setSavingToDb(false); 
    } 
  }, [results, shiftDate, reportDate, uploaderName, fileName, rawCsvText]); 

  // ── DB: load a past report into the viewer ─────────────────────────────────   const loadReport = useCallback(async (reportId) => {     if (!supabaseReady()) return;     setHistoryLoading(true);     setError("");     try {       const { data, error: err } = await supabase 
        .from("reports") 
        .select("*") 
        .eq("id", reportId)         .single();       if (err) throw err; 

      // Re-hydrate Date objects from ISO strings 
      const rehydrated = (data.results_json || []).map(r => ({ 
        ...r,         rawIn:  r.rawIn  ? new Date(r.rawIn)  : null,         rawOut: r.rawOut ? new Date(r.rawOut) : null,         rndIn:  r.rndIn  ? new Date(r.rndIn)  : null,         rndOut: r.rndOut ? new Date(r.rndOut) : null,         nqIn:   r.nqIn   ? { ts: new Date(r.nqIn.ts),  gate: r.nqIn.gate }  : nul         nqOut:  r.nqOut  ? { ts: new Date(r.nqOut.ts), gate: r.nqOut.gate } : nul 
      })); 

      setResults(rehydrated);       setShiftDate(data.shift_date);       setReportDate(data.report_date);       setFileName(data.file_name || "");       setSavedReportId(data.id);       setScreen("results");     } catch (e) {       setError("Could not load report: " + e.message); 
    } finally {       setHistoryLoading(false); 
    }   }, []); 

  const processFile = useCallback((file) => {     if (!file) return; 
    setFileName(file.name);     setError("");     setProcessing(true);     setSavedReportId(null);     setDupeWarning(null);     const reader = new FileReader();     reader.onload = e => {       try {         const text = e.target.result;         setRawCsvText(text);         const rows = parseCSV(text);         if (!rows.length) throw new Error("No data rows found in CSV.");         const nqGates = new Set(nqText.split("\n").map(s => s.trim()).filter(Bool         setResults(analyze(rows, shiftDate, reportDate, nqGates));         setScreen("results");       } catch(err) {         setError("Error: " + err.message); 
      } finally { 
        setProcessing(false); 
      }     };     reader.readAsText(file);   }, [shiftDate, reportDate, nqText]); 

  const onDrop = useCallback(e => { 
    e.preventDefault(); setDragging(false);     const file = e.dataTransfer.files[0];     if (file) processFile(file);   }, [processFile]); 

  const handleExport = () => {     const nqGates = new Set(nqText.split("\n").map(s => s.trim()).filter(Boolean)     exportXLSX(results, shiftDate, reportDate, nqGates); 
  }; 

  // Derived stats   const clean  = results.filter(r => r.status === "Clean").length;   const mOut   = results.filter(r => r.status === "Missing OUT").length;   const mIn    = results.filter(r => r.status === "Missing IN").length;   const mBoth  = results.filter(r => r.status === "Missing IN & OUT").length;   const over13 = results.filter(r => r.over13).length;   const totalH = results.reduce((s, r) => s + (r.billHrs || 0), 0); 

  const filtered = results.filter(r => {     const mf =       filter === "All"         ? true :       filter === "Issues"      ? r.status !== "Clean" :       filter === "Missing OUT" ? r.status === "Missing OUT" :       filter === "Missing IN"  ? r.status === "Missing IN" : 
      filter === "Over 13"     ? r.over13 : true;     const ms = !search || 
      r.name.toLowerCase().includes(search.toLowerCase()) || 
      r.ein.includes(search);     return mf && ms; 
  }); 

  const inp = {     background:"#0d1520", border:"1px solid #1a2d45", borderRadius:6,     color:"#c8d6e8", padding:"8px 12px", fontSize:13,     fontFamily:"'DM Mono',monospace", outline:"none",     width:"100%", boxSizing:"border-box", 
  }; 

  const btn = (active, color = "#3b82f6") => ({     background: active ? color : "transparent",     color: active ? "#fff" : "#6b7a8d",     border: `1px solid ${active ? color : "#1a2d45"}`,     borderRadius:6, padding:"6px 12px", fontSize:12,     cursor:"pointer", fontFamily:"'DM Mono',monospace", fontWeight:600,     transition:"all 0.15s", whiteSpace:"nowrap",     WebkitTapHighlightColor:"transparent", 
  }); 

  // ── PASSWORD GATE SCREEN ───────────────────────────────────────────────────   if (!authed) {     return (       <div style={{         minHeight:"100vh", background:"#060d16", color:"#c8d6e8",         fontFamily:"'DM Sans','Segoe UI',sans-serif",         display:"flex", alignItems:"center", justifyContent:"center", padding:"20       }}> 
        <div style={{           width:"100%", maxWidth:380,           background:"#0a1320", border:"1px solid #1a2535",           borderRadius:12, padding:"32px 24px", 
        }}> 
          <div style={{ textAlign:"center", marginBottom:24 }}> 
            <div style={{               background:"linear-gradient(135deg,#1d4ed8,#0ea5e9)",               width:48, height:48, borderRadius:10, margin:"0 auto 12px",               display:"flex", alignItems:"center", justifyContent:"center",               fontSize:22, fontWeight:900, color:"#fff", 
            }}>G</div> 
            <div style={{ fontSize:18, fontWeight:800, color:"#e2e8f0", letterSpa               CP2 GATE LOG 
            </div> 
            <div style={{ fontSize:11, color:"#4b5e7a", letterSpacing:1.5, textTr 
              MMR Constructors 
            </div> 
          </div> 
          <div> 
            <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercas 
              Team Password 
            </label>             <input               type="password"               value={pwInput}               onChange={e => setPwInput(e.target.value)}               onKeyDown={e => e.key === "Enter" && handleLogin(e)}               autoFocus               style={{...inp, marginTop:6, fontSize:14}}               placeholder="Enter password" 
            /> 
            {pwError && ( 
              <div style={{ color:"#ef4444", fontSize:12, marginTop:8 }}>{pwError 
            )} 
            <button onClick={handleLogin} style={{               width:"100%", marginTop:14,               background:"#1d4ed8", color:"#fff",               border:"none", borderRadius:8, padding:"12px",               fontSize:13, fontWeight:700, cursor:"pointer",               fontFamily:"'DM Mono',monospace", letterSpacing:0.5,             }}>UNLOCK</button> 
          </div> 
        </div> 
        <style>{` 
          @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400; 
          * { box-sizing: border-box; } 
        `}</style> 
      </div> 
    );   } 

  return ( 
    <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",       fontFamily:"'DM Sans','Segoe UI',sans-serif" }}> 

      {/* ── HEADER ── */}       <div style={{         background:"#0a1320", borderBottom:"1px solid #1a2535",         padding:"0 16px", display:"flex", alignItems:"center",         justifyContent:"space-between", height:52, position:"sticky", top:0, zInd       }}> 
        <div style={{ display:"flex", alignItems:"center", gap:10 }}> 
          <div style={{             background:"linear-gradient(135deg,#1d4ed8,#0ea5e9)",             width:30, height:30, borderRadius:7,             display:"flex", alignItems:"center", justifyContent:"center",             fontSize:15, fontWeight:900, color:"#fff", flexShrink:0, 
          }}>G</div> 
          <div> 
            <div style={{ fontSize:13, fontWeight:700, color:"#e2e8f0", letterSpa             <div style={{ fontSize:9, color:"#4b5e7a", letterSpacing:1, textTrans 
          </div> 
        </div> 
        <div style={{ display:"flex", gap:6, alignItems:"center" }}>           <button onClick={() => setScreen(screen === "history" ? "upload" : "his             style={btn(screen === "history", "#334155")}>             {screen === "history" ? "← Back" : "≡ History"} 
          </button> 
          <button onClick={() => setShowConfig(c => !c)} style={btn(showConfig, " 
            ⚙ 
          </button> 
          {screen === "results" && (<> 
            <button onClick={handleExport} style={{ 
              ...btn(true, "#1d4ed8"), background:"#1d4ed8", color:"#fff", 
            }}>↓ Excel</button>             <button onClick={() => { setScreen("upload"); setResults([]); setFile               style={btn(false, "#334155")}>↺</button> 
          </>)} 
          <button onClick={handleLogout} style={btn(false, "#334155")} title="Loc 
        </div> 
      </div> 

      {/* ── CONFIG PANEL ── */} 
      {showConfig && ( 
        <div style={{ background:"#0a1320", borderBottom:"1px solid #1a2535", pad           <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, ma 
            {[ 
              { label:"Shift Date",  val:shiftDate,  set:setShiftDate }, 
              { label:"Report Date", val:reportDate, set:setReportDate },             ].map(({ label, val, set }) => ( 
              <div key={label}>                 <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppe 
                  {label} 
                </label> 
                <input type="date" value={val} onChange={e => set(e.target.value)                   style={{...inp, marginTop:3}} /> 
              </div> 
            ))} 
          </div> 
          <div style={{ marginTop:10 }}>             <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercas 
              Off-Site (NQ) Gates — one per line, exact match 
            </label> 
            <div style={{ fontSize:10, color:"#4b5e7a", marginTop:2, marginBottom               fontFamily:"'DM Mono',monospace" }}> 
              ⚠ S-Curve Turnstile 2 not listed — add here if needed 
            </div> 
            <textarea value={nqText} onChange={e => setNqText(e.target.value)}               rows={7} style={{...inp, marginTop:3, resize:"vertical"}} /> 
          </div> 
        </div> 
      )} 

      <div style={{ padding:"14px 16px", maxWidth:1400, margin:"0 auto" }}> 

        {/* ── UPLOAD SCREEN ── */} 
        {screen === "upload" && (           <div style={{ display:"flex", flexDirection:"column", alignItems:"cente 
            <div style={{ textAlign:"center" }}> 
              <div style={{ fontSize:24, fontWeight:800, color:"#e2e8f0", letterS 
                Daily Gate Log Analysis 
              </div> 
              <div style={{ fontSize:13, color:"#4b5e7a", marginTop:4 }}> 
                Upload the Genetec CSV export to begin 
              </div> 
            </div> 

            {/* Uploader name (required) */} 
            <div style={{ width:"100%", maxWidth:460 }}>               <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"upperc                 letterSpacing:1.2, display:"block", marginBottom:6 }}>                 Your Name <span style={{ color:"#ef4444" }}>*</span> 
              </label>               <input                 type="text"                 value={uploaderName}                 onChange={e => setUploaderName(e.target.value)}                 placeholder="e.g. Jon Maguire"                 style={{...inp, fontSize:14}} 
              /> 
            </div> 

            {/* Date pickers */} 
            <div style={{ display:"flex", gap:10, flexWrap:"wrap", justifyContent 
              {[ 
                { label:"Shift",  val:shiftDate,  set:setShiftDate }, 
                { label:"Report", val:reportDate, set:setReportDate }, 
              ].map(({ label, val, set }) => (                 <div key={label} style={{                   display:"flex", alignItems:"center", gap:8,                   background:"#0d1520", border:"1px solid #1a2d45",                   borderRadius:8, padding:"8px 12px", flex:"1 1 140px", 
                }}> 
                  <span style={{ fontSize:10, color:"#4b5e7a", textTransform:"upp                     letterSpacing:1, flexShrink:0 }}>{label}</span>                   <input type="date" value={val} onChange={e => set(e.target.valu                     background:"transparent", border:"none", color:"#e2e8f0",                     fontFamily:"'DM Mono',monospace", fontSize:13, outline:"none"                     width:"100%",                   }} /> 
                </div> 
              ))} 
            </div> 

            {/* Drop zone — disabled until uploader name is filled */} 
            {(() => {               const nameReady = uploaderName.trim().length > 0;               return (                 <div                   onDragOver={e => { if (nameReady) { e.preventDefault(); setDrag                   onDragLeave={() => setDragging(false)}                   onDrop={e => { if (nameReady) onDrop(e); else { e.preventDefaul                   onClick={() => {                     if (nameReady) fileRef.current?.click();                     else setError("Enter your name above before uploading."); 
                  }}                   style={{                     width:"100%", maxWidth:460,                     border:`2px dashed ${dragging ? "#3b82f6" : (nameReady ? "#1a                     borderRadius:14, padding:"40px 20px", textAlign:"center",                     cursor: nameReady ? "pointer" : "not-allowed",                     opacity: nameReady ? 1 : 0.55,                     background: dragging ? "#0d1e30" : "#0a1320", transition:"all 
                  }} 
                > 
                  <div style={{ fontSize:36, marginBottom:10 }}>{nameReady ? ""                   <div style={{ fontSize:15, fontWeight:600, color:"#e2e8f0" }}>                     {processing ? "Processing…" :                      !nameReady ? "Enter your name first" :                      fileName || "Tap to browse or drop CSV"} 
                  </div> 
                  <div style={{ fontSize:11, color:"#4b5e7a", marginTop:4 }}>Gene                   <input ref={fileRef} type="file" accept=".csv" style={{display: 
                    onChange={e => processFile(e.target.files[0])} /> 
                </div> 
              ); 
            })()} 

            {error && ( 
              <div style={{ color:"#ef4444", fontSize:13, background:"#1e0d0d",                 border:"1px solid #991b1b", borderRadius:8, padding:"10px 14px",                 width:"100%", maxWidth:460 }}>{error}</div> 
            )} 
          </div> 
        )} 

        {/* ── RESULTS SCREEN ── */} 
        {screen === "results" && (<> 

          {/* SAVE BAR — show only if not yet saved */} 
          {supabaseReady() && (             <div style={{               display:"flex", alignItems:"center", justifyContent:"space-between"               gap:10, flexWrap:"wrap", marginBottom:12, padding:"10px 12px",               background: savedReportId ? "#0d3b1e" : "#0a1320",               border: `1px solid ${savedReportId ? "#166534" : "#1a2d45"}`,               borderRadius:8, 
            }}> 
              <div style={{ fontSize:12, color:"#c8d6e8", lineHeight:1.5 }}>                 {savedReportId ? ( 
                  <> 
                    <span style={{ color:"#4ade80", fontWeight:700 }}>✓ Saved to  
                    {" "}Uploaded by <strong>{uploaderName}</strong>. 
                    {dupeWarning && (                       <span style={{ color:"#fb923c", display:"block", marginTop:                         ⚠ {dupeWarning.count} other report{dupeWarning.count > 1  
                      </span> 
                    )} 
                  </> 
                ) : ( 
                  <>Save this report to the team database so others can view it l 
                )} 
              </div> 
              {!savedReportId && (                 <button                   onClick={saveReport}                   disabled={savingToDb || !uploaderName.trim()}                   style={{                     background: savingToDb ? "#1a2d45" : "#22c55e",                     color: "#fff", border:"none", borderRadius:6,                     padding:"8px 16px", fontSize:12, fontWeight:700,                     cursor: savingToDb ? "wait" : "pointer",                     fontFamily:"'DM Mono',monospace", letterSpacing:0.5,                     opacity: (!uploaderName.trim()) ? 0.5 : 1, 
                  }} 
                > 
                  {savingToDb ? "SAVING…" : " SAVE"} 
                </button> 
              )} 
            </div> 
          )} 

          {/* STAT CARDS — 4-per-row on mobile, flex-wrap */} 
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:14 } 
            <StatCard label="Workers"      value={results.length}  color="#3b82f6 
            <StatCard label="Clean"        value={clean}           color="#22c55e 
            <StatCard label="Missing OUT"  value={mOut}            color="#ef4444 
            <StatCard label="Missing IN"   value={mIn}             color="#ef4444 
            <StatCard label="Missing Both" value={mBoth}           color="#ef4444 
            <StatCard label="Over 13 Hrs"  value={over13}          color="#fb923c             <StatCard label="Total Hrs"    value={decToHM(totalH)} color="#0ea5e9 
          </div> 

          {/* FILTERS — wrap naturally; search on its own row on mobile */} 
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:10,  
            {["All","Issues","Missing OUT","Missing IN","Over 13"].map(f => (               <button key={f} onClick={() => setFilter(f)} style={btn(filter===f, 
            ))} 
          </div> 
          <div style={{ marginBottom:10 }}> 
            <input               placeholder="Search name or EIN…"               value={search}               onChange={e => setSearch(e.target.value)}               style={{...inp, maxWidth:320}} 
            /> 
          </div> 

          {/* SHIFT META */} 
          <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:10,             fontFamily:"'DM Mono',monospace", lineHeight:1.6 }}>             Shift: {shiftDate} &nbsp;|&nbsp; Report: {reportDate}<br/> 
            Showing {filtered.length} of {results.length} workers — tap row to ex 
          </div> 
          {/* TABLE */} 
          <div style={{ background:"#0a1320", border:"1px solid #1a2535",             borderRadius:10, overflow:"hidden" }}>             <div style={{ overflowX:"auto", WebkitOverflowScrolling:"touch" }}> 
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:6                 <thead> 
                  <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2 
                    {["Employee Name","Company","EIN","Rounded IN","Rounded OUT",                       <th key={h} style={{                         padding:"9px 12px", textAlign:"left", fontSize:10,                         color:"#4b5e7a", textTransform:"uppercase", letterSpacing                         fontWeight:600, whiteSpace:"nowrap",                       }}>{h}</th> 
                    ))} 
                  </tr> 
                </thead> 
                <tbody> 
                  {filtered.map((r, i) => <Row key={r.guid} r={r} idx={i} />)} 
                  {!filtered.length && (                     <tr><td colSpan={7} style={{ padding:36, textAlign:"center",                       No records match. 
                    </td></tr> 
                  )} 
                </tbody> 
              </table> 
            </div> 
          </div> 
        </>)} 

        {/* ── HISTORY SCREEN ── */} 
        {screen === "history" && ( 
          <div> 
            <div style={{ marginBottom:14 }}>               <div style={{ fontSize:20, fontWeight:800, color:"#e2e8f0", letterS 
                Report History 
              </div> 
              <div style={{ fontSize:12, color:"#4b5e7a", marginTop:2 }}> 
                Tap a row to open and re-export 
              </div> 
            </div> 

            {!supabaseReady() && (               <div style={{ color:"#ef4444", fontSize:13, background:"#1e0d0d",                 border:"1px solid #991b1b", borderRadius:8, padding:"10px 14px" }                 Database not configured. Set VITE_SUPABASE_URL and VITE_SUPABASE_ 
              </div> 
            )} 

            {error && ( 
              <div style={{ color:"#ef4444", fontSize:13, background:"#1e0d0d",                 border:"1px solid #991b1b", borderRadius:8, padding:"10px 14px",                 marginBottom:12 }}>{error}</div> 
            )} 

            {historyLoading && (               <div style={{ color:"#4b5e7a", fontSize:13, padding:"20px 0", textA 
                Loading… 
              </div> 
            )} 

            {!historyLoading && supabaseReady() && history.length === 0 && ( 
              <div style={{ color:"#4b5e7a", fontSize:13, padding:"40px 0", textA                 No reports saved yet. Upload a CSV and tap Save. 
              </div> 
            )} 

            {!historyLoading && history.length > 0 && (() => {               // Group by shift_date to flag duplicates               const shiftCounts = {};               history.forEach(h => { shiftCounts[h.shift_date] = (shiftCounts[h.s               return ( 
                <div style={{ background:"#0a1320", border:"1px solid #1a2535",                   borderRadius:10, overflow:"hidden" }}>                   {history.map((h, i) => {                     const isDupe = shiftCounts[h.shift_date] > 1;                     return (                       <div                         key={h.id}                         onClick={() => loadReport(h.id)}                         style={{                           padding:"12px 14px", cursor:"pointer",                           borderBottom: i < history.length - 1 ? "1px solid #1a25                           background: i % 2 === 0 ? "#0a1320" : "#0d1520",                           WebkitTapHighlightColor:"transparent", 
                        }} 
                      > 
                        <div style={{ display:"flex", alignItems:"center", justif                           <div style={{ fontSize:14, fontWeight:700, color:"#e2e8                             fontFamily:"'DM Mono',monospace" }}>                             {h.shift_date} 
                            {isDupe && (                               <span style={{                                 marginLeft:8, fontSize:10, color:"#fb923c",                                 background:"#3b2500", border:"1px solid #9a3412",                                 borderRadius:4, padding:"2px 6px", fontWeight:600 
                              }}>⚠ DUP</span> 
                            )} 
                          </div>                           <div style={{ fontSize:11, color:"#0ea5e9", fontFamily:                             fontWeight:700 }}>                             {h.total_bill_hrs != null ? decToHM(h.total_bill_hrs) 
                          </div> 
                        </div>                         <div style={{ fontSize:11, color:"#6b7a8d", marginTop:4,                           display:"flex", gap:10, flexWrap:"wrap" }}>                           <span>by <strong style={{ color:"#c8d6e8" }}>{h.uploade                           <span>{new Date(h.uploaded_at).toLocaleString("en-US",                             month:"short", day:"numeric", hour:"numeric", minute:                           })}</span>                           <span>{h.total_workers} workers</span> 
                          {h.over_13 > 0 && (                             <span style={{ color:"#fb923c" }}>{h.over_13} over 13                           )} 
                          {(h.missing_in + h.missing_out + h.missing_both) > 0 && 
                            <span style={{ color:"#ef4444" }}> 
                              {h.missing_in + h.missing_out + h.missing_both} iss 
                            </span> 
                          )} 
                        </div> 
                      </div> 
                    ); 
                  })} 
                </div> 
              ); 
            })()} 
          </div> 
        )} 
      </div> 

      <style>{` 
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;60         * { box-sizing: border-box; }         input[type=date]::-webkit-calendar-picker-indicator { filter: invert(0.5)         tr:hover td { background: rgba(59,130,246,0.04) !important; }         ::-webkit-scrollbar { width:5px; height:5px; } 
        ::-webkit-scrollbar-track { background:#060d16; } 
        ::-webkit-scrollbar-thumb { background:#1a2d45; border-radius:3px; } 
        @media (max-width: 480px) {           /* On small screens the stat cards go 4-up then 3-up naturally via flex 
          /* Ensure filter buttons don't overflow */ 
          button { font-size: 11px !important; padding: 5px 10px !important; } 
        } 
      `}</style> 
    </div> 
  ); 
} 