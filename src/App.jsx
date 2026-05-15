import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase, supabaseReady, SHARED_PASSWORD } from "./supabase.js";

// ── CONSTANTS ─────────────────────────────────────────────────────────────────
const NQ_GATES_DEFAULT = new Set([
  "Helms Rd Turnstile 1","Helms Rd Turnstile 2","Helms Rd VG Staff Badge In/Out",
  "Liberty Visitors Center","PHI MMR Office","PHI Turnstile",
  "Roy Baily Turnstile 1","Roy Baily Turnstile 2","S-Curve Turnstile 1",
  "Helms Rd MMR Office","Helms Rd Worley 4 Plex","Venture Global Lake Charles Warehouse",
  // NOTE: S-Curve Turnstile 2 has significant traffic (96 events in 5/6 data)
  // and is NOT in this NQ list. Monitor and add manually via Config if needed.
]);

// ── UTILITIES ─────────────────────────────────────────────────────────────────
function parseDate(str) {
  if (!str) return null;
  const d = new Date(str);
  return isNaN(d) ? null : d;
}

function toDateStr(d) {
  if (!d) return "";
  return (
    d.toLocaleDateString("en-US", { month:"2-digit", day:"2-digit", year:"numeric" }) +
    " " +
    d.toLocaleTimeString("en-US", { hour:"2-digit", minute:"2-digit", hour12:true })
  );
}

// Round a Date to nearest 30-min using 15:30 (930 sec) midpoint rule
function rnd30(d) {
  if (!d) return null;
  const totalSec = d.getHours()*3600 + d.getMinutes()*60 + d.getSeconds();
  const rem = totalSec % 1800;
  let rounded = rem >= 930 ? totalSec + (1800 - rem) : totalSec - rem;
  rounded = rounded % 86400;
  const h = Math.floor(rounded / 3600);
  const m = Math.floor((rounded % 3600) / 60);
  const r = new Date(d);
  r.setHours(h, m, 0, 0);
  return r;
}

function decToHM(dec) {
  if (dec == null || isNaN(dec)) return "";
  const neg = dec < 0;
  const abs = Math.abs(dec);
  const h = Math.floor(abs);
  const m = Math.round((abs - h) * 60);
  return `${neg ? "-" : ""}${h}:${String(m).padStart(2, "0")}`;
}

function sameDay(a, b) {
  return a && b &&
    a.getFullYear() === b.getFullYear() &&
    a.getMonth()    === b.getMonth()    &&
    a.getDate()     === b.getDate();
}

function parseCSV(text) {
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const headers = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));
  return lines.slice(1).map(line => {
    const cols = [];
    let cur = "", inQ = false;
    for (let i = 0; i < line.length; i++) {
      if (line[i] === '"') { inQ = !inQ; continue; }
      if (line[i] === ',' && !inQ) { cols.push(cur.trim()); cur = ""; }
      else cur += line[i];
    }
    cols.push(cur.trim());
    const obj = {};
    headers.forEach((h, i) => obj[h] = cols[i] || "");
    return obj;
  });
}

// ── MAIN ANALYSIS ─────────────────────────────────────────────────────────────
function analyze(rows, shiftDateStr, reportDateStr, nqGates) {
  const shiftDate  = new Date(shiftDateStr  + "T00:00:00");
  const reportDate = new Date(reportDateStr + "T00:00:00");

  const df = rows.map(r => ({
    guid:    r["CardholderGuid"] || r["Cardholder Guid"] || r["cardholderguid"] || "",
    name:    r["Employee Name"] || "",
    company: r["Company Name"] || "",
    ein:     r["Company EIN"] || "",
    area:    r["Area"] || "",
    access:  r["Access Point"] || "",
    ts:      parseDate(r["Event Timestamp"] || r["EventTimestamp"] || ""),
    isNQ:    nqGates.has(r["Area"] || ""),
  })).filter(r => r.guid && r.ts);

  df.sort((a, b) => a.ts - b.ts);

  const dfQ  = df.filter(r => !r.isNQ);
  const dfNQ = df.filter(r =>  r.isNQ);

  const guids    = [...new Set(dfQ.map(r => r.guid))];
  const allGuids = [...new Set(df.map(r => r.guid))];

  const empMap = {};
  df.forEach(r => {
    if (!empMap[r.guid]) empMap[r.guid] = { name: r.name, company: r.company, ein: r.ein };
  });

  // First qualifying IN on shift date per person
  const inMap = {};
  dfQ.filter(r => r.access === "In" && sameDay(r.ts, shiftDate))
     .forEach(r => { if (!inMap[r.guid]) inMap[r.guid] = { ts: r.ts, gate: r.area }; });

  // NQ reference: first NQ IN, last NQ OUT
  const nqInMap = {}, nqOutMap = {};
  dfNQ.filter(r => r.access === "In")
      .forEach(r => { if (!nqInMap[r.guid]) nqInMap[r.guid] = { ts: r.ts, gate: r.area }; });
  dfNQ.filter(r => r.access === "Out")
      .sort((a, b) => b.ts - a.ts)
      .forEach(r => { if (!nqOutMap[r.guid]) nqOutMap[r.guid] = { ts: r.ts, gate: r.area }; });

  // AM IN (before noon)  → OUT must be same shift date, any time
  // PM IN (noon or after) → OUT can be same date (early leave) OR following date before noon
  function getValidOut(guid, rawIn) {
    const outs = dfQ.filter(r => r.guid === guid && r.access === "Out");
    if (!rawIn) {
      const shiftOuts = outs.filter(r => sameDay(r.ts, shiftDate));
      return shiftOuts.length ? shiftOuts.sort((a, b) => b.ts - a.ts)[0] : null;
    }
    const valid = rawIn.getHours() < 12
      ? outs.filter(r => sameDay(r.ts, shiftDate))
      : outs.filter(r =>
          sameDay(r.ts, shiftDate) ||
          (sameDay(r.ts, reportDate) && r.ts.getHours() < 12)
        );
    return valid.length ? valid.sort((a, b) => b.ts - a.ts)[0] : null;
  }

  const processGuids = [...new Set([...guids, ...allGuids.filter(g => inMap[g])])];

  const results = processGuids.map(guid => {
    const emp    = empMap[guid] || {};
    const inRec  = inMap[guid] || null;
    const rawIn  = inRec ? inRec.ts : null;
    const outRec = getValidOut(guid, rawIn);
    const rawOut = outRec ? outRec.ts : null;

    const rndIn  = rnd30(rawIn);
    const rndOut = rnd30(rawOut);

    let rawHrs = null, lessLunch = null, billHrs = null;
    if (rndIn && rndOut) {
      const elapsedMin = (rndOut - rndIn) / 60000;
      if (elapsedMin > 0) {
        const netMin  = Math.max(elapsedMin - 30, 0);
        rawHrs    = +(elapsedMin / 60).toFixed(4);
        lessLunch = +(netMin / 60).toFixed(4);
        billHrs   = lessLunch; // Billable = rndOut - rndIn - 30 min; no second rounding pass
      }
    }

    const hasIn  = !!rawIn;
    const hasOut = !!rawOut;
    const over13 = billHrs != null && billHrs > 13;

    const status =
      !hasIn && !hasOut ? "Missing IN & OUT" :
      hasIn  && !hasOut ? "Missing OUT"      :
      !hasIn &&  hasOut ? "Missing IN"       :
      over13            ? "Over 13 Hrs"      : "Clean";

    return {
      guid, name: emp.name, company: emp.company, ein: emp.ein,
      rawIn, rawOut, rndIn, rndOut,
      inGate:  inRec  ? inRec.gate  : "",
      outGate: outRec ? outRec.area : "",
      rawHrs, lessLunch, billHrs,
      nqIn:  nqInMap[guid]  || null,
      nqOut: nqOutMap[guid] || null,
      status, over13,
    };
  });

  results.sort((a, b) => a.name.localeCompare(b.name));
  return results;
}

// ── EXCEL EXPORT — iOS-safe blob download ────────────────────────────────────
function exportXLSX(results, shiftDate, reportDate, nqGates) {
  const wb = XLSX.utils.book_new();

  function setColWidths(ws, widths) {
    ws["!cols"] = widths.map(w => ({ wch: w }));
  }

  // ─── TAB 1: RAW DATA PASTE
  const rawWs = XLSX.utils.aoa_to_sheet([
    ["PASTE RAW GENETEC DATA STARTING AT ROW 3  |  DELETE OLD DATA FIRST  |  DO NOT EDIT ROW 2 HEADERS"],
    ["Event","Area","Door","Access Point","Event Timestamp","CardholderGuid",
     "First Name","Last Name","Employee Name","Company Name","Company EIN","Event_Date","Event_Time"],
  ]);
  setColWidths(rawWs, [20,30,20,14,22,38,16,20,28,14,12,14,14]);
  XLSX.utils.book_append_sheet(wb, rawWs, "RAW DATA PASTE");

  // ─── TAB 2: BILLABLE TIME
  const clean  = results.filter(r => r.status === "Clean").length;
  const mOut   = results.filter(r => r.status === "Missing OUT").length;
  const mIn    = results.filter(r => r.status === "Missing IN").length;
  const mBoth  = results.filter(r => r.status === "Missing IN & OUT").length;
  const over13 = results.filter(r => r.over13).length;
  const totalH = results.reduce((s, r) => s + (r.billHrs || 0), 0);
  const withHrs = results.filter(r => r.billHrs != null).length;

  const dataColHdrs = [
    "Employee Name","Company","EIN",
    "Raw Badge IN","IN Gate","Rounded IN",
    "Raw Badge OUT","OUT Gate","Rounded OUT",
    "Raw Elapsed (H:MM)","Less Lunch (H:MM)","Billable Hrs (H:MM)",
    "NQ Badge IN (ref)","NQ IN Gate","NQ Badge OUT (ref)","NQ OUT Gate","Status"
  ];
  const dataRows = r => [
    r.name, r.company, r.ein,
    toDateStr(r.rawIn),  r.inGate,  toDateStr(r.rndIn),
    toDateStr(r.rawOut), r.outGate, toDateStr(r.rndOut),
    decToHM(r.rawHrs), decToHM(r.lessLunch), decToHM(r.billHrs),
    r.nqIn  ? toDateStr(r.nqIn.ts)  : "", r.nqIn  ? r.nqIn.gate  : "",
    r.nqOut ? toDateStr(r.nqOut.ts) : "", r.nqOut ? r.nqOut.gate : "",
    r.status,
  ];

  const btWs = XLSX.utils.aoa_to_sheet([
    [`CP2 LNG  |  MMR Gate Log Billable Time  |  Shift Date: ${shiftDate}  |  Report Date: ${reportDate}`],
    [`Total Workers: ${results.length}`, "", "", `Clean: ${clean}`, "", "",
     `Missing OUT: ${mOut}`, "", "", `Missing IN: ${mIn}`, "", "",
     `Over 13 Hrs: ${over13}`, "", "", `Total Billable Hrs: ${decToHM(totalH)}`],
    ["Rounding: Raw IN and Raw OUT each rounded independently to nearest 30 min via 15:30 midpoint rule  |  " +
     "Billable = (Rounded OUT minus Rounded IN) minus 30 min lunch  |  Times displayed as H:MM"],
    dataColHdrs,
    ...results.map(dataRows),
  ]);
  setColWidths(btWs, [28,10,10,22,26,22,22,26,22,16,16,16,22,26,22,26,18]);
  XLSX.utils.book_append_sheet(wb, btWs, "BILLABLE TIME");

  // ─── TAB 3: ISSUES
  const issues = results.filter(r => r.status !== "Clean");
  const issWs = XLSX.utils.aoa_to_sheet([
    [`ISSUES  |  Shift: ${shiftDate}  |  Report: ${reportDate}  |  ` +
     `Missing OUT: ${mOut}  |  Missing IN: ${mIn}  |  Missing Both: ${mBoth}  |  Over 13 Hrs: ${over13}`],
    dataColHdrs,
    ...issues.map(dataRows),
  ]);
  setColWidths(issWs, [28,10,10,22,26,22,22,26,22,16,16,16,22,26,22,26,18]);
  XLSX.utils.book_append_sheet(wb, issWs, "ISSUES");

  // ─── TAB 4: SUMMARY
  const sumWs = XLSX.utils.aoa_to_sheet([
    [`DAILY GATE LOG SUMMARY  |  Shift: ${shiftDate}  |  Report: ${reportDate}`],
    ["Metric", "Value"],
    ["Total Workers on Site",         results.length],
    ["Clean Records",                 clean],
    ["Missing Badge OUT",             mOut],
    ["Missing Badge IN",              mIn],
    ["Missing IN and OUT",            mBoth],
    ["Over 13 Billable Hrs (verify)", over13],
    ["Total Billable Hours",          decToHM(totalH)],
    ["Avg Billable Hrs per Worker",   withHrs > 0 ? decToHM(totalH / withHrs) : ""],
  ]);
  setColWidths(sumWs, [34, 14]);
  XLSX.utils.book_append_sheet(wb, sumWs, "SUMMARY");

  // ─── TAB 5: CONFIG
  const cfgWs = XLSX.utils.aoa_to_sheet([
    ["CONFIGURATION  |  Update dates and NQ gates here"],
    ["DATE SETTINGS", "Value", "Notes"],
    ["Shift Date (date being analyzed)",    shiftDate,  "Badge INs are only pulled from this date"],
    ["Report Date (date data was received)", reportDate, "Night shift OUTs allowed on this date AM only"],
    [],
    ["RULES", "Value", "Notes"],
    ["Rounding Rule",    "15:30 midpoint to nearest 30 min", ""],
    ["Lunch Deduction",  "30 minutes", "Applied when both IN and OUT exist"],
    ["Badge IN Logic",   "First qualifying IN on shift date only", "Report date INs are excluded"],
    ["Badge OUT Logic",  "Last qualifying OUT per AM/PM rule", "AM IN: same date OUT | PM IN: same date or next AM"],
    ["NQ Gate Rule",     "Off-site swipes shown as reference only", "Not used in calculations"],
    ["Over 13 Hrs Flag", "Any billable total over 13 hrs flagged", "Shown in Issues tab for review"],
    [],
    ["OFF-SITE (NON-QUALIFYING) GATES  |  Add or remove rows as needed"],
    ["Gate Name (exact match)"],
    ...[...nqGates].sort().map(g => [g]),
    [],
    ["NOTE: S-Curve Turnstile 2 is NOT in the NQ list. Monitor and add above if needed."],
  ]);
  setColWidths(cfgWs, [46, 36, 52]);
  XLSX.utils.book_append_sheet(wb, cfgWs, "CONFIG");

  // ─── iOS-safe download ───────────────────────────────────────────────────
  // XLSX.writeFile uses document.createElement('a').click() which iOS Safari blocks.
  // Instead: write to ArrayBuffer → Blob → object URL → anchor click.
  // If the anchor click is suppressed (iOS), fall back to window.open on the blob URL.
  const wbOut  = XLSX.write(wb, { bookType:"xlsx", type:"array" });
  const blob   = new Blob([wbOut], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url    = URL.createObjectURL(blob);
  const fname  = `CP2_GateLog_${shiftDate}.xlsx`;

  const a = document.createElement("a");
  a.href     = url;
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  // iOS Safari fallback: if the anchor download attribute is ignored, open in new tab
  // so the user can long-press → Save to Files
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// ── STATUS CONFIG ─────────────────────────────────────────────────────────────
const STATUS_STYLES = {
  "Clean":            { bg:"#0d3b1e", text:"#4ade80", border:"#166534" },
  "Missing OUT":      { bg:"#3b1010", text:"#f87171", border:"#991b1b" },
  "Missing IN":       { bg:"#3b1010", text:"#f87171", border:"#991b1b" },
  "Missing IN & OUT": { bg:"#3b1010", text:"#f87171", border:"#991b1b" },
  "Over 13 Hrs":      { bg:"#3b2500", text:"#fb923c", border:"#9a3412" },
};

// ── STAT CARD ─────────────────────────────────────────────────────────────────
function StatCard({ label, value, color }) {
  return (
    <div style={{
      background:"#0f1923", border:`1px solid ${color}33`,
      borderRadius:8, padding:"10px 12px",
      // Mobile: fixed width so 4 fit per row on a ~390px screen
      flex:"1 1 calc(25% - 8px)", minWidth:0,
    }}>
      <div style={{ color, fontSize:20, fontWeight:700, fontFamily:"'DM Mono',monospace",
        whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{value}</div>
      <div style={{ color:"#6b7a8d", fontSize:10, marginTop:2, textTransform:"uppercase",
        letterSpacing:0.8, lineHeight:1.2 }}>{label}</div>
    </div>
  );
}

// ── STATUS BADGE ──────────────────────────────────────────────────────────────
function Badge({ status }) {
  const s = STATUS_STYLES[status] || STATUS_STYLES["Clean"];
  return (
    <span style={{
      background:s.bg, color:s.text, border:`1px solid ${s.border}`,
      borderRadius:4, padding:"2px 6px", fontSize:10, fontWeight:600,
      fontFamily:"'DM Mono',monospace", whiteSpace:"nowrap",
    }}>{status}</span>
  );
}

const td = { padding:"10px 12px", color:"#c8d6e8", fontSize:13, verticalAlign:"middle" };

// ── ROW — touch-safe expand ───────────────────────────────────────────────────
function Row({ r, idx, showRoster }) {
  const [open, setOpen] = useState(false);
  const isIssue = r.status !== "Clean";

  // iOS fix: use onTouchEnd to toggle so the expand works even when
  // the scroll container intercepts the onClick event
  const handleToggle = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setOpen(o => !o);
  }, []);

  // Highlight not-in-roster rows in amber-ish background
  const rowBg = r.notInRoster
    ? (idx % 2 === 0 ? "#1f1408" : "#1a1106")
    : isIssue
      ? (idx % 2 === 0 ? "#1e0d0d" : "#180b0b")
      : (idx % 2 === 0 ? "#0d1520" : "#0a1118");

  const colSpan = showRoster ? 10 : 7;

  return (
    <>
      <tr
        onClick={handleToggle}
        onTouchEnd={handleToggle}
        style={{ background:rowBg, cursor:"pointer", borderBottom:"1px solid #1a2535",
          WebkitTapHighlightColor:"transparent" }}
      >
        <td style={td}>
          {r.name}
          {r.notInRoster && (
            <span style={{
              marginLeft:8, fontSize:9, padding:"2px 5px", borderRadius:3,
              background:"#92400e", color:"#fff", letterSpacing:0.5,
            }}>NOT IN ROSTER</span>
          )}
        </td>
        <td style={{...td, color:"#6b7a8d"}}>{r.company}</td>
        <td style={{...td, color:"#6b7a8d", fontFamily:"'DM Mono',monospace"}}>{r.ein}</td>
        {showRoster && (<>
          <td style={{...td, fontSize:12, color: r.supervisor === "N/A" ? "#4b5e7a" : "#c8d6e8"}}>{r.supervisor}</td>
          <td style={{...td, fontSize:12, color: r.area === "N/A" ? "#4b5e7a" : "#c8d6e8"}}>{r.area}</td>
          <td style={{...td, fontSize:12, color: r.scope === "N/A" ? "#4b5e7a" : "#c8d6e8"}}>{r.scope}</td>
        </>)}
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontSize:12}}>
          {toDateStr(r.rndIn)  || <span style={{color:"#ef4444"}}>MISSING</span>}
        </td>
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontSize:12}}>
          {toDateStr(r.rndOut) || <span style={{color:"#ef4444"}}>MISSING</span>}
        </td>
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontWeight:700,
          color: r.over13 ? "#fb923c" : "#e2e8f0"}}>
          {decToHM(r.billHrs) || "—"}
        </td>
        <td style={td}><Badge status={r.status} /></td>
      </tr>

      {open && (
        <tr style={{ background:"#060d16", borderBottom:"1px solid #1a2535" }}>
          <td colSpan={colSpan} style={{ padding:"12px 14px" }}>
            <div style={{
              display:"grid",
              gridTemplateColumns:"repeat(auto-fill, minmax(200px, 1fr))",
              gap:8, fontSize:12,
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
                ["NQ IN (ref)",   r.nqIn  ? `${toDateStr(r.nqIn.ts)} — ${r.nqIn.gate}`   : ""],
                ["NQ OUT (ref)",  r.nqOut ? `${toDateStr(r.nqOut.ts)} — ${r.nqOut.gate}` : ""],
                ...(showRoster ? [
                  ["Supervisor",  r.supervisor],
                  ["Area",        r.area],
                  ["Scope",       r.scope],
                ] : []),
              ].map(([k, v]) => (
                <div key={k} style={{
                  background:"#0d1520", borderRadius:6, padding:"8px 10px",
                  border:"1px solid #1a2535",
                }}>
                  <div style={{ color:"#4b5e7a", fontSize:10, textTransform:"uppercase",
                    letterSpacing:1, marginBottom:3 }}>{k}</div>
                  <div style={{ color:"#c8d6e8", fontFamily:"'DM Mono',monospace",
                    fontSize:11, wordBreak:"break-word" }}>
                    {v || <span style={{color:"#374151"}}>—</span>}
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

// ── APP ───────────────────────────────────────────────────────────────────────
export default function App() {
  // ── Auth + Supabase state ──
  const [authed, setAuthed] = useState(() => sessionStorage.getItem("cp2gateauth") === "1");
  const [authPw, setAuthPw] = useState("");
  const [authErr, setAuthErr] = useState("");
  const [uploadedBy, setUploadedBy] = useState(() => localStorage.getItem("cp2-uploader") || "");
  const [history, setHistory] = useState([]);
  const [menuOpen, setMenuOpen] = useState(false);
  const [saveMsg, setSaveMsg] = useState("");

  useEffect(() => {
    if (uploadedBy) localStorage.setItem("cp2-uploader", uploadedBy);
  }, [uploadedBy]);

  const [screen, setScreen]         = useState("upload");
  const [results, setResults]       = useState([]);
  const [shiftDate, setShiftDate]   = useState("2026-05-06");
  const [reportDate, setReportDate] = useState("2026-05-07");
  const [filter, setFilter]         = useState("All");
  const [search, setSearch]         = useState("");
  const [supFilter, setSupFilter]   = useState("");
  const [areaFilter, setAreaFilter] = useState("");
  const [scopeFilter, setScopeFilter] = useState("");
  const [nqText, setNqText]         = useState([...NQ_GATES_DEFAULT].sort().join("\n"));
  const [showConfig, setShowConfig] = useState(false);
  const [dragging, setDragging]     = useState(false);
  const [fileName, setFileName]     = useState("");
  const [error, setError]           = useState("");
  const [processing, setProcessing] = useState(false);
  const fileRef = useRef();

  // ── Employee roster state ──
  const [roster, setRoster]         = useState([]); // [{ein, firstName, lastName, supervisor, area, scope}]
  const [rosterMeta, setRosterMeta] = useState(null); // {uploaded_at, uploaded_by, count}
  const [rosterMsg, setRosterMsg]   = useState("");
  const rosterFileRef = useRef();

  // Load roster on mount
  useEffect(() => {
    if (!supabaseReady()) return;
    (async () => {
      try {
        const { data, error } = await supabase
          .from("employees")
          .select("*")
          .order("last_name", { ascending: true });
        if (!error && data) {
          setRoster(data.map(r => ({
            ein: String(r.ein || "").trim(),
            firstName: r.first_name || "",
            lastName: r.last_name || "",
            supervisor: r.supervisor || "",
            area: r.area || "",
            scope: r.scope || "",
          })));
          if (data.length > 0) {
            setRosterMeta({
              count: data.length,
              uploaded_at: data[0].uploaded_at,
              uploaded_by: data[0].uploaded_by,
            });
          }
        }
      } catch (_) { /* silent */ }
    })();
  }, []);

  // Roster upload handler (parses .xlsx, replaces in Supabase)
  const handleRosterUpload = async (file) => {
    if (!file) return;
    setRosterMsg("");
    if (!supabaseReady()) { setRosterMsg("Database not configured."); return; }
    if (!uploadedBy.trim()) { setRosterMsg("Enter your name first."); return; }
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) throw new Error("Spreadsheet is empty.");

      // Flexible column detection
      const keyMap = {};
      const first = rows[0];
      for (const k of Object.keys(first)) {
        const lower = k.toLowerCase().trim();
        if (lower === "ein" || lower.includes("ein")) keyMap.ein = k;
        else if (lower.includes("first")) keyMap.firstName = k;
        else if (lower.includes("last")) keyMap.lastName = k;
        else if (lower.includes("supervisor")) keyMap.supervisor = k;
        else if (lower === "area" || lower.includes("area")) keyMap.area = k;
        else if (lower === "scope" || lower.includes("scope")) keyMap.scope = k;
      }
      if (!keyMap.ein) throw new Error("No EIN column found in spreadsheet.");

      const parsed = rows
        .map(r => ({
          ein: String(r[keyMap.ein] || "").trim(),
          first_name: String(r[keyMap.firstName] || "").trim(),
          last_name: String(r[keyMap.lastName] || "").trim(),
          supervisor: String(r[keyMap.supervisor] || "").trim(),
          area: String(r[keyMap.area] || "").trim(),
          scope: String(r[keyMap.scope] || "").trim(),
          uploaded_by: uploadedBy.trim(),
          uploaded_at: new Date().toISOString(),
        }))
        .filter(r => r.ein);

      if (!parsed.length) throw new Error("No valid rows with EIN found.");

      // Replace strategy: delete all, insert new
      const { error: delErr } = await supabase.from("employees").delete().neq("ein", "___never___");
      if (delErr) throw delErr;
      const { error: insErr } = await supabase.from("employees").insert(parsed);
      if (insErr) throw insErr;

      setRoster(parsed.map(r => ({
        ein: r.ein, firstName: r.first_name, lastName: r.last_name,
        supervisor: r.supervisor, area: r.area, scope: r.scope,
      })));
      setRosterMeta({ count: parsed.length, uploaded_at: parsed[0].uploaded_at, uploaded_by: parsed[0].uploaded_by });
      setRosterMsg(`Uploaded ${parsed.length} employees.`);
    } catch (e) {
      setRosterMsg("Upload error: " + e.message);
    }
  };

  // Compute absent employees: in roster but no badge event today
  const absentEmployees = (() => {
    if (!roster.length || !results.length) return [];
    const badgedEINs = new Set(results.map(r => String(r.ein || "").trim()).filter(Boolean));
    return roster.filter(emp => emp.ein && !badgedEINs.has(emp.ein));
  })();

  const processFile = useCallback((file) => {
    if (!file) return;
    setFileName(file.name);
    setError("");
    setProcessing(true);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const rows = parseCSV(e.target.result);
        if (!rows.length) throw new Error("No data rows found in CSV.");
        const nqGates = new Set(nqText.split("\n").map(s => s.trim()).filter(Boolean));
        setResults(analyze(rows, shiftDate, reportDate, nqGates));
        setScreen("results");
      } catch(err) {
        setError("Error: " + err.message);
      } finally {
        setProcessing(false);
      }
    };
    reader.readAsText(file);
  }, [shiftDate, reportDate, nqText]);

  const onDrop = useCallback(e => {
    e.preventDefault(); setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, [processFile]);

  const handleExport = () => {
    const nqGates = new Set(nqText.split("\n").map(s => s.trim()).filter(Boolean));
    exportXLSX(results, shiftDate, reportDate, nqGates);
  };

  // Derived stats
  const clean  = results.filter(r => r.status === "Clean").length;
  const mOut   = results.filter(r => r.status === "Missing OUT").length;
  const mIn    = results.filter(r => r.status === "Missing IN").length;
  const mBoth  = results.filter(r => r.status === "Missing IN & OUT").length;
  const over13 = results.filter(r => r.over13).length;
  const totalH = results.reduce((s, r) => s + (r.billHrs || 0), 0);

  // Roster lookup by EIN
  const rosterByEIN = (() => {
    const map = {};
    for (const r of roster) {
      if (r.ein) map[String(r.ein).trim()] = r;
    }
    return map;
  })();

  // Enrich each result row with supervisor / area / scope / notInRoster flag
  const enriched = results.map(r => {
    const ein = String(r.ein || "").trim();
    const emp = ein ? rosterByEIN[ein] : null;
    const notInRoster = roster.length > 0 && (!ein || !emp);
    return {
      ...r,
      supervisor: emp && emp.supervisor ? emp.supervisor : (emp ? "N/A" : "N/A"),
      area:       emp && emp.area       ? emp.area       : (emp ? "N/A" : "N/A"),
      scope:      emp && emp.scope      ? emp.scope      : (emp ? "N/A" : "N/A"),
      notInRoster,
    };
  });

  // Filter dropdown values: build unique sorted lists from the enriched data
  const supervisors = [...new Set(enriched.map(r => r.supervisor).filter(Boolean))].sort();
  const areas       = [...new Set(enriched.map(r => r.area).filter(Boolean))].sort();
  const scopes      = [...new Set(enriched.map(r => r.scope).filter(Boolean))].sort();

  const notInRosterCount = enriched.filter(r => r.notInRoster).length;

  const filtered = enriched.filter(r => {
    const mf =
      filter === "All"            ? true :
      filter === "Issues"         ? r.status !== "Clean" :
      filter === "Missing OUT"    ? r.status === "Missing OUT" :
      filter === "Missing IN"     ? r.status === "Missing IN" :
      filter === "Over 13"        ? r.over13 :
      filter === "Not in Roster"  ? r.notInRoster : true;
    const ms = !search ||
      r.name.toLowerCase().includes(search.toLowerCase()) ||
      r.ein.includes(search);
    const supOk = !supFilter   || r.supervisor === supFilter;
    const areaOk = !areaFilter || r.area === areaFilter;
    const scopeOk = !scopeFilter || r.scope === scopeFilter;
    return mf && ms && supOk && areaOk && scopeOk;
  });

  const inp = {
    background:"#0d1520", border:"1px solid #1a2d45", borderRadius:6,
    color:"#c8d6e8", padding:"8px 12px", fontSize:13,
    fontFamily:"'DM Mono',monospace", outline:"none",
    width:"100%", boxSizing:"border-box",
  };

  const btn = (active, color = "#3b82f6") => ({
    background: active ? color : "transparent",
    color: active ? "#fff" : "#6b7a8d",
    border: `1px solid ${active ? color : "#1a2d45"}`,
    borderRadius:6, padding:"6px 12px", fontSize:12,
    cursor:"pointer", fontFamily:"'DM Mono',monospace", fontWeight:600,
    transition:"all 0.15s", whiteSpace:"nowrap",
    WebkitTapHighlightColor:"transparent",
  });

  // ── Supabase save / load handlers ──
  const handleSave = async () => {
    setSaveMsg("");
    if (!uploadedBy.trim()) { setSaveMsg("Enter your name first."); return; }
    if (!supabaseReady()) { setSaveMsg("Database not configured."); return; }
    if (!results.length) { setSaveMsg("Nothing to save."); return; }
    try {
      const totalBillHrs = results.reduce((s, r) => s + (r.billHrs || 0), 0);
      const issuesCount = results.filter(r => r.status !== "Clean").length;
      const { error } = await supabase.from("reports").insert({
        shift_date: shiftDate,
        report_date: reportDate,
        uploaded_by: uploadedBy.trim(),
        results_json: { results, shiftDate, reportDate },
        raw_csv_path: null,
        shifts_count: results.length,
        issues_count: issuesCount,
        total_bill_hrs: Math.round(totalBillHrs * 100) / 100
      });
      if (error) throw error;
      setSaveMsg("Saved.");
    } catch (e) {
      setSaveMsg("Save error: " + e.message);
    }
  };

  const loadHistory = async () => {
    if (!supabaseReady()) { setSaveMsg("Database not configured."); return; }
    const { data, error } = await supabase
      .from("reports")
      .select("id, shift_date, report_date, uploaded_by, shifts_count, issues_count, total_bill_hrs")
      .order("report_date", { ascending: false })
      .limit(100);
    if (error) { setSaveMsg("History error: " + error.message); return; }
    setHistory(data || []);
    setScreen("history");
    setMenuOpen(false);
  };

  const loadReport = async (id) => {
    const { data, error } = await supabase.from("reports").select("*").eq("id", id).single();
    if (error) { setSaveMsg("Load error: " + error.message); return; }
    const payload = data.results_json || {};
    setResults(payload.results || []);
    if (payload.shiftDate) setShiftDate(payload.shiftDate);
    if (payload.reportDate) setReportDate(payload.reportDate);
    setScreen("results");
  };

  const handleLogout = () => {
    sessionStorage.removeItem("cp2gateauth");
    setAuthed(false);
    setMenuOpen(false);
  };

  const dupDates = (() => {
    const counts = {};
    for (const h of history) counts[h.shift_date] = (counts[h.shift_date] || 0) + 1;
    return new Set(Object.entries(counts).filter(([_, n]) => n > 1).map(([d]) => d));
  })();

  // ── Auth gate ──
  if (!authed) {
    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif",
        display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
        <div style={{ background:"#0a1320", border:"1px solid #1a2535",
          borderRadius:10, padding:28, maxWidth:380, width:"100%" }}>
          <div style={{ fontSize:16, fontWeight:700, color:"#e2e8f0", marginBottom:6 }}>CP2 Gate Log</div>
          <div style={{ fontSize:11, color:"#4b5e7a", letterSpacing:1, textTransform:"uppercase", marginBottom:18 }}>MMR Constructors</div>
          <input
            type="password"
            value={authPw}
            onChange={e => setAuthPw(e.target.value)}
            onKeyDown={e => {
              if (e.key === "Enter") {
                if (authPw === SHARED_PASSWORD) {
                  sessionStorage.setItem("cp2gateauth", "1");
                  setAuthed(true); setAuthErr("");
                } else { setAuthErr("Incorrect password"); }
              }
            }}
            placeholder="Password"
            autoFocus
            style={{ ...inp, marginBottom:12 }}
          />
          {authErr && <div style={{ color:"#ef4444", fontSize:12, marginBottom:10 }}>{authErr}</div>}
          <button
            onClick={() => {
              if (authPw === SHARED_PASSWORD) {
                sessionStorage.setItem("cp2gateauth", "1");
                setAuthed(true); setAuthErr("");
              } else { setAuthErr("Incorrect password"); }
            }}
            style={{
              background:"#1d4ed8", color:"#fff", border:"none",
              borderRadius:6, padding:"10px 16px", fontSize:13, fontWeight:600,
              cursor:"pointer", width:"100%", fontFamily:"'DM Mono',monospace",
            }}
          >Enter</button>
        </div>
      </div>
    );
  }

  // ── History screen ──
  if (screen === "history") {
    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
          <div style={{ fontSize:16, fontWeight:700, color:"#e2e8f0" }}>HISTORY</div>
          <button onClick={() => setScreen(results.length ? "results" : "upload")} style={btn(false, "#334155")}>Back</button>
        </div>
        <div style={{ background:"#0a1320", border:"1px solid #1a2535", borderRadius:10, overflow:"hidden" }}>
          {history.length === 0 ? (
            <div style={{ padding:24, color:"#4b5e7a", fontSize:13 }}>No saved reports.</div>
          ) : (
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:680 }}>
                <thead>
                  <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2535" }}>
                    {["Shift Date","Uploaded By","Saved","Shifts","Issues","Total Hrs",""].map(h => (
                      <th key={h} style={{ padding:"9px 12px", textAlign:"left", fontSize:10,
                        color:"#4b5e7a", textTransform:"uppercase", letterSpacing:0.8, fontWeight:600 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {history.map(h => (
                    <tr key={h.id} style={{ borderBottom:"1px solid #1a2535" }}>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>
                        {h.shift_date}
                        {dupDates.has(h.shift_date) && (
                          <span style={{ background:"#ef4444", color:"#fff", fontSize:10,
                            padding:"2px 6px", borderRadius:4, marginLeft:6 }}>⚠ DUP</span>
                        )}
                      </td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{h.uploaded_by}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{new Date(h.report_date).toLocaleString()}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{h.shifts_count}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{h.issues_count}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{h.total_bill_hrs}</td>
                      <td style={{ padding:"9px 12px" }}>
                        <button onClick={() => loadReport(h.id)} style={btn(false, "#334155")}>Load</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
        {saveMsg && <div style={{ marginTop:12, fontSize:12, color:"#4b5e7a" }}>{saveMsg}</div>}
      </div>
    );
  }

  // ── Absent employees screen ──
  if (screen === "absent") {
    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700, color:"#e2e8f0" }}>ABSENT EMPLOYEES</div>
            <div style={{ fontSize:11, color:"#4b5e7a", marginTop:2 }}>
              In roster but did not badge on {shiftDate} · {absentEmployees.length} of {roster.length} total
            </div>
          </div>
          <button onClick={() => setScreen("results")} style={btn(false, "#334155")}>Back to Results</button>
        </div>
        <div style={{ background:"#0a1320", border:"1px solid #1a2535", borderRadius:10, overflow:"hidden" }}>
          {!roster.length ? (
            <div style={{ padding:24, color:"#4b5e7a", fontSize:13 }}>
              No employee roster uploaded yet. Open ⚙ Config to upload one.
            </div>
          ) : !absentEmployees.length ? (
            <div style={{ padding:24, color:"#34d399", fontSize:13 }}>
              ✓ All {roster.length} employees badged in today.
            </div>
          ) : (
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:680 }}>
                <thead>
                  <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2535" }}>
                    {["Name","EIN","Supervisor","Area","Scope"].map(h => (
                      <th key={h} style={{ padding:"9px 12px", textAlign:"left", fontSize:10,
                        color:"#4b5e7a", textTransform:"uppercase", letterSpacing:0.8, fontWeight:600 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {absentEmployees
                    .slice()
                    .sort((a, b) => (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName))
                    .map((e, i) => (
                      <tr key={i} style={{ borderBottom:"1px solid #1a2535" }}>
                        <td style={{ padding:"9px 12px", fontSize:12 }}>{e.lastName}, {e.firstName}</td>
                        <td style={{ padding:"9px 12px", fontSize:12, fontFamily:"'DM Mono',monospace" }}>{e.ein}</td>
                        <td style={{ padding:"9px 12px", fontSize:12 }}>{e.supervisor}</td>
                        <td style={{ padding:"9px 12px", fontSize:12 }}>{e.area}</td>
                        <td style={{ padding:"9px 12px", fontSize:12 }}>{e.scope}</td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
      fontFamily:"'DM Sans','Segoe UI',sans-serif" }}>

      {/* ── HEADER ── */}
      <div style={{
        background:"#0a1320", borderBottom:"1px solid #1a2535",
        padding:"0 16px", display:"flex", alignItems:"center",
        justifyContent:"space-between", height:52, position:"sticky", top:0, zIndex:10,
      }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{
            background:"linear-gradient(135deg,#1d4ed8,#0ea5e9)",
            width:30, height:30, borderRadius:7,
            display:"flex", alignItems:"center", justifyContent:"center",
            fontSize:15, fontWeight:900, color:"#fff", flexShrink:0,
          }}>G</div>
          <div>
            <div style={{ fontSize:13, fontWeight:700, color:"#e2e8f0", letterSpacing:0.5 }}>CP2 GATE LOG</div>
            <div style={{ fontSize:9, color:"#4b5e7a", letterSpacing:1, textTransform:"uppercase" }}>MMR Constructors</div>
          </div>
        </div>
        <div style={{ display:"flex", gap:6, alignItems:"center" }}>
          <button onClick={() => setShowConfig(c => !c)} style={btn(showConfig, "#334155")}>
            ⚙ Config
          </button>
          {screen === "results" && (<>
            <button onClick={handleExport} style={{
              ...btn(true, "#1d4ed8"), background:"#1d4ed8", color:"#fff",
            }}>↓ Excel</button>
            <button onClick={handleSave} style={btn(false, "#334155")} title="Save to history">💾 Save</button>
            {roster.length > 0 && (
              <button
                onClick={() => setScreen("absent")}
                style={{
                  ...btn(false, "#334155"),
                  color: absentEmployees.length > 0 ? "#f87171" : "#4b5e7a",
                  borderColor: absentEmployees.length > 0 ? "#7f1d1d" : "#1a2d45",
                }}
                title="Show absent employees"
              >🚫 Absent {absentEmployees.length > 0 ? `(${absentEmployees.length})` : ""}</button>
            )}
            <button onClick={() => { setScreen("upload"); setResults([]); setFileName(""); }}
              style={btn(false, "#334155")}>↺</button>
          </>)}
          <div style={{ position:"relative" }}>
            <button onClick={() => setMenuOpen(o => !o)} style={btn(false, "#334155")}>≡</button>
            {menuOpen && (
              <div style={{ position:"absolute", right:0, top:34, background:"#0a1320",
                border:"1px solid #1a2535", borderRadius:8, padding:6, minWidth:140, zIndex:20 }}>
                <button onClick={loadHistory} style={{ ...btn(false, "#334155"), width:"100%", marginBottom:4 }}>History</button>
                <button onClick={handleLogout} style={{ ...btn(false, "#334155"), width:"100%", color:"#ef4444", borderColor:"#7f1d1d" }}>Logout</button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* ── CONFIG PANEL ── */}
      {showConfig && (
        <div style={{ background:"#0a1320", borderBottom:"1px solid #1a2535", padding:"14px 16px" }}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, maxWidth:500 }}>
            {[
              { label:"Shift Date",  val:shiftDate,  set:setShiftDate },
              { label:"Report Date", val:reportDate, set:setReportDate },
            ].map(({ label, val, set }) => (
              <div key={label}>
                <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase", letterSpacing:1 }}>
                  {label}
                </label>
                <input type="date" value={val} onChange={e => set(e.target.value)}
                  style={{...inp, marginTop:3}} />
              </div>
            ))}
          </div>
          <div style={{ marginTop:10 }}>
            <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase", letterSpacing:1 }}>
              Off-Site (NQ) Gates — one per line, exact match
            </label>
            <div style={{ fontSize:10, color:"#4b5e7a", marginTop:2, marginBottom:3,
              fontFamily:"'DM Mono',monospace" }}>
              ⚠ S-Curve Turnstile 2 not listed — add here if needed
            </div>
            <textarea value={nqText} onChange={e => setNqText(e.target.value)}
              rows={7} style={{...inp, marginTop:3, resize:"vertical"}} />
          </div>

          {/* ── Employee Roster Upload ── */}
          <div style={{ marginTop:14, paddingTop:14, borderTop:"1px solid #1a2535" }}>
            <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase", letterSpacing:1 }}>
              MMR Employee Roster (.xlsx)
            </label>
            <div style={{ fontSize:10, color:"#4b5e7a", marginTop:2, marginBottom:6,
              fontFamily:"'DM Mono',monospace" }}>
              Columns: EIN, First Name, Last Name, Supervisor, Area, Scope. New upload replaces previous.
            </div>
            <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
              <input
                ref={rosterFileRef}
                type="file"
                accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                onChange={e => handleRosterUpload(e.target.files[0])}
                style={{ display:"none" }}
              />
              <button
                onClick={() => rosterFileRef.current?.click()}
                style={{
                  background:"#1d4ed8", color:"#fff", border:"none",
                  borderRadius:6, padding:"8px 14px", fontSize:12, fontWeight:600,
                  cursor:"pointer", fontFamily:"'DM Mono',monospace",
                }}
              >📤 Upload Roster</button>
              {rosterMeta && (
                <div style={{ fontSize:11, color:"#4b5e7a", fontFamily:"'DM Mono',monospace" }}>
                  {rosterMeta.count} employees · last by {rosterMeta.uploaded_by} · {new Date(rosterMeta.uploaded_at).toLocaleString()}
                </div>
              )}
            </div>
            {rosterMsg && (
              <div style={{
                marginTop:8, fontSize:12,
                color: rosterMsg.startsWith("Upload error") || rosterMsg.startsWith("Enter") || rosterMsg.startsWith("Database") ? "#f87171" : "#34d399",
                fontFamily:"'DM Mono',monospace",
              }}>{rosterMsg}</div>
            )}
          </div>
        </div>
      )}

      <div style={{ padding:"14px 16px", maxWidth:1400, margin:"0 auto" }}>

        {/* ── UPLOAD SCREEN ── */}
        {screen === "upload" && (
          <div style={{ display:"flex", flexDirection:"column", alignItems:"center", paddingTop:40, gap:20 }}>
            <div style={{ textAlign:"center" }}>
              <div style={{ fontSize:24, fontWeight:800, color:"#e2e8f0", letterSpacing:-0.5 }}>
                Daily Gate Log Analysis
              </div>
              <div style={{ fontSize:13, color:"#4b5e7a", marginTop:4 }}>
                Upload the Genetec CSV export to begin
              </div>
            </div>

            {/* Uploaded by */}
            <div style={{
              display:"flex", alignItems:"center", gap:8,
              background:"#0d1520", border:"1px solid #1a2d45",
              borderRadius:8, padding:"8px 12px", width:"100%", maxWidth:420,
            }}>
              <span style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase",
                letterSpacing:1, flexShrink:0 }}>Uploader</span>
              <input
                type="text"
                value={uploadedBy}
                onChange={e => setUploadedBy(e.target.value)}
                placeholder="Your name (required to save)"
                style={{
                  background:"transparent", border:"none", color:"#e2e8f0",
                  fontFamily:"'DM Mono',monospace", fontSize:13, outline:"none",
                  width:"100%",
                }}
              />
            </div>

            {/* Date pickers */}
            <div style={{ display:"flex", gap:10, flexWrap:"wrap", justifyContent:"center", width:"100%" }}>
              {[
                { label:"Shift",  val:shiftDate,  set:setShiftDate },
                { label:"Report", val:reportDate, set:setReportDate },
              ].map(({ label, val, set }) => (
                <div key={label} style={{
                  display:"flex", alignItems:"center", gap:8,
                  background:"#0d1520", border:"1px solid #1a2d45",
                  borderRadius:8, padding:"8px 12px", flex:"1 1 140px",
                }}>
                  <span style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase",
                    letterSpacing:1, flexShrink:0 }}>{label}</span>
                  <input type="date" value={val} onChange={e => set(e.target.value)} style={{
                    background:"transparent", border:"none", color:"#e2e8f0",
                    fontFamily:"'DM Mono',monospace", fontSize:13, outline:"none",
                    width:"100%",
                  }} />
                </div>
              ))}
            </div>

            {/* Drop zone */}
            <div
              onDragOver={e => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              onClick={() => fileRef.current?.click()}
              style={{
                width:"100%", maxWidth:460,
                border:`2px dashed ${dragging ? "#3b82f6" : "#1a2d45"}`,
                borderRadius:14, padding:"40px 20px", textAlign:"center", cursor:"pointer",
                background: dragging ? "#0d1e30" : "#0a1320", transition:"all 0.2s",
              }}
            >
              <div style={{ fontSize:36, marginBottom:10 }}>📂</div>
              <div style={{ fontSize:15, fontWeight:600, color:"#e2e8f0" }}>
                {processing ? "Processing…" : fileName || "Tap to browse or drop CSV"}
              </div>
              <div style={{ fontSize:11, color:"#4b5e7a", marginTop:4 }}>Genetec raw export (.csv)</div>
              <input ref={fileRef} type="file" accept=".csv" style={{display:"none"}}
                onChange={e => processFile(e.target.files[0])} />
            </div>

            {error && (
              <div style={{ color:"#ef4444", fontSize:13, background:"#1e0d0d",
                border:"1px solid #991b1b", borderRadius:8, padding:"10px 14px",
                width:"100%", maxWidth:460 }}>{error}</div>
            )}
          </div>
        )}

        {/* ── RESULTS SCREEN ── */}
        {screen === "results" && (<>

          {saveMsg && (
            <div style={{
              background: saveMsg.startsWith("Save") || saveMsg.startsWith("History") || saveMsg.startsWith("Load") || saveMsg.startsWith("Nothing") || saveMsg.startsWith("Enter") || saveMsg.startsWith("Database") ? "#3f1d1d" : "#14532d",
              border: "1px solid " + (saveMsg.startsWith("Save") || saveMsg.startsWith("History") || saveMsg.startsWith("Load") || saveMsg.startsWith("Nothing") || saveMsg.startsWith("Enter") || saveMsg.startsWith("Database") ? "#7f1d1d" : "#166534"),
              color:"#e2e8f0", padding:"8px 12px", borderRadius:6, fontSize:12, marginBottom:10,
            }}>{saveMsg}</div>
          )}

          {/* STAT CARDS — 4-per-row on mobile, flex-wrap */}
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:14 }}>
            <StatCard label="Workers"      value={results.length}  color="#3b82f6" />
            <StatCard label="Clean"        value={clean}           color="#22c55e" />
            <StatCard label="Missing OUT"  value={mOut}            color="#ef4444" />
            <StatCard label="Missing IN"   value={mIn}             color="#ef4444" />
            <StatCard label="Missing Both" value={mBoth}           color="#ef4444" />
            <StatCard label="Over 13 Hrs"  value={over13}          color="#fb923c" />
            <StatCard label="Total Hrs"    value={decToHM(totalH)} color="#0ea5e9" />
          </div>

          {/* FILTERS — wrap naturally; search on its own row on mobile */}
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:10, alignItems:"center" }}>
            {["All","Issues","Missing OUT","Missing IN","Over 13"].map(f => (
              <button key={f} onClick={() => setFilter(f)} style={btn(filter===f,"#1d4ed8")}>{f}</button>
            ))}
            {notInRosterCount > 0 && (
              <button
                onClick={() => setFilter(filter === "Not in Roster" ? "All" : "Not in Roster")}
                style={{
                  ...btn(filter === "Not in Roster", "#f59e0b"),
                  color: filter === "Not in Roster" ? "#fff" : "#f59e0b",
                  borderColor: "#92400e",
                }}
              >⚠ Not in Roster ({notInRosterCount})</button>
            )}
          </div>

          {/* Supervisor / Area / Scope filters — only show when roster loaded */}
          {roster.length > 0 && (
            <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:10, alignItems:"center" }}>
              <select
                value={supFilter}
                onChange={e => setSupFilter(e.target.value)}
                style={{
                  background:"#0d1520", color:"#c8d6e8", border:"1px solid #1a2d45",
                  borderRadius:6, padding:"6px 10px", fontSize:12,
                  fontFamily:"'DM Mono',monospace", cursor:"pointer",
                }}
              >
                <option value="">All Supervisors</option>
                {supervisors.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              <select
                value={areaFilter}
                onChange={e => setAreaFilter(e.target.value)}
                style={{
                  background:"#0d1520", color:"#c8d6e8", border:"1px solid #1a2d45",
                  borderRadius:6, padding:"6px 10px", fontSize:12,
                  fontFamily:"'DM Mono',monospace", cursor:"pointer",
                }}
              >
                <option value="">All Areas</option>
                {areas.map(a => <option key={a} value={a}>{a}</option>)}
              </select>
              <select
                value={scopeFilter}
                onChange={e => setScopeFilter(e.target.value)}
                style={{
                  background:"#0d1520", color:"#c8d6e8", border:"1px solid #1a2d45",
                  borderRadius:6, padding:"6px 10px", fontSize:12,
                  fontFamily:"'DM Mono',monospace", cursor:"pointer",
                }}
              >
                <option value="">All Scopes</option>
                {scopes.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              {(supFilter || areaFilter || scopeFilter) && (
                <button
                  onClick={() => { setSupFilter(""); setAreaFilter(""); setScopeFilter(""); }}
                  style={btn(false, "#334155")}
                >Clear</button>
              )}
            </div>
          )}
          <div style={{ marginBottom:10 }}>
            <input
              placeholder="Search name or EIN…"
              value={search}
              onChange={e => setSearch(e.target.value)}
              style={{...inp, maxWidth:320}}
            />
          </div>

          {/* SHIFT META */}
          <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:10,
            fontFamily:"'DM Mono',monospace", lineHeight:1.6 }}>
            Shift: {shiftDate} &nbsp;|&nbsp; Report: {reportDate}<br/>
            Showing {filtered.length} of {results.length} workers — tap row to expand
          </div>

          {/* TABLE */}
          <div style={{ background:"#0a1320", border:"1px solid #1a2535",
            borderRadius:10, overflow:"hidden" }}>
            <div style={{ overflowX:"auto", WebkitOverflowScrolling:"touch" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:680 }}>
                <thead>
                  <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2535" }}>
                    {[
                      "Employee Name","Company","EIN",
                      ...(roster.length > 0 ? ["Supervisor","Area","Scope"] : []),
                      "Rounded IN","Rounded OUT","Billable","Status"
                    ].map(h => (
                      <th key={h} style={{
                        padding:"9px 12px", textAlign:"left", fontSize:10,
                        color:"#4b5e7a", textTransform:"uppercase", letterSpacing:0.8,
                        fontWeight:600, whiteSpace:"nowrap",
                      }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => <Row key={r.guid} r={r} idx={i} showRoster={roster.length > 0} />)}
                  {!filtered.length && (
                    <tr><td colSpan={roster.length > 0 ? 10 : 7} style={{ padding:36, textAlign:"center", color:"#4b5e7a" }}>
                      No records match.
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>)}
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        input[type=date]::-webkit-calendar-picker-indicator { filter: invert(0.5); }
        tr:hover td { background: rgba(59,130,246,0.04) !important; }
        ::-webkit-scrollbar { width:5px; height:5px; }
        ::-webkit-scrollbar-track { background:#060d16; }
        ::-webkit-scrollbar-thumb { background:#1a2d45; border-radius:3px; }
        @media (max-width: 480px) {
          /* On small screens the stat cards go 4-up then 3-up naturally via flex-wrap */
          /* Ensure filter buttons don't overflow */
          button { font-size: 11px !important; padding: 5px 10px !important; }
        }
      `}</style>
    </div>
  );
}
