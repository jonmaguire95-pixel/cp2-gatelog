import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase, supabaseReady, authStart, writePasswordStep, takeCodeIntent } from "./supabase.js";
import LoginScreen from "./LoginScreen.jsx";
import { filesFromZip } from "./unzip.js";

import {
  NQ_GATES_DEFAULT, RULES_DEFAULT, NIGHT_CUT_H, FLAG_STATUSES, REVIEW_STATUSES,
  parseDate, toDateStr, minutesToClock, extraRulesSummary, midpointToStr, rulesSummary, toYMD, addDays,
  afternoonCounts, detectShiftDate, decToHM, parseCSV, GATE_FILE_RE, parseWorkbook, swipeCount, NO_SWIPES,
  shortDate, weekdayOf, dayLabel, weekStart, weekDays, newestPerDate, groupDays, analyze,
  gateFlagArea, summarizeByArea, statusCounts,
} from "./daylogic.js";

// ── FILE READING ──────────────────────────────────────────────────────────────
async function readGateFile(file) {
  if (/\.csv$/i.test(file.name)) return parseCSV(await file.text());
  return parseWorkbook(await file.arrayBuffer());
}

// A week of exports can come as one .zip. The Genetec files inside it are
// pulled out and handled exactly like files dropped on their own. Anything
// wrong with a zip stops the whole drop, so a day can't go missing quietly.
const ZIP_RE = /\.zip$/i;

function zipMessage(name, e) {
  const fix = "Unzip it and drop the files instead.";
  switch (e && e.code) {
    case "browser": return `This browser can't open .zip files. ${fix}`;
    case "locked":  return `${name} is password protected. ${fix}`;
    case "method":  return `${name} uses a compression type the site can't read. ${fix}`;
    case "toobig":  return `${name} is too large to open here. ${fix}`;
    default:        return `${name} couldn't be opened. It may be damaged or not a real zip. ${fix}`;
  }
}

async function expandZips(files) {
  const out = [];
  for (const f of files) {
    if (!ZIP_RE.test(f.name)) { out.push(f); continue; }
    let inner;
    try { inner = await filesFromZip(f, GATE_FILE_RE); }
    catch (e) { throw new Error(zipMessage(f.name, e)); }
    if (!inner.length) throw new Error(`No Genetec files (.xlsx or .csv) found in ${f.name}.`);
    out.push(...inner);
  }
  return out;
}


// ── EXCEL EXPORT: iOS safe blob download ────────────────────────────────────
// XLSX.writeFile uses document.createElement('a').click() which iOS Safari blocks.
// Instead: write to ArrayBuffer → Blob → object URL → anchor click.
// Excel shows who made a file and with what program under File, Info. Every
// export carries Jon Maguire as author and Excel as the program.
function stampExcelApp(buf) {
  try {
    const zip = XLSX.CFB.read(new Uint8Array(buf), { type: "array" });
    const entry = XLSX.CFB.find(zip, "/docProps/app.xml");
    if (!entry || !entry.content) return buf;
    const xml = new TextDecoder().decode(entry.content)
      .replace("<Application>SheetJS</Application>", "<Application>Microsoft Excel</Application>");
    entry.content = new TextEncoder().encode(xml);
    entry.size = entry.content.length;
    return XLSX.CFB.write(zip, { fileType: "zip", type: "array" });
  } catch (_) {
    return buf;
  }
}

function downloadWorkbook(wb, fname) {
  wb.Props = { ...(wb.Props || {}), Author: "Jon Maguire", LastAuthor: "Jon Maguire" };
  const wbOut = stampExcelApp(XLSX.write(wb, { bookType:"xlsx", type:"array" }));
  const blob  = new Blob([wbOut], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url   = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href     = url;
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// Adds the Excel header filter arrows to a sheet. Rows and columns are zero
// indexed. The range runs from the header row through the last data row only,
// so totals rows below the data stay out of the filter.
function setAutoFilter(ws, headerRow, lastRow, lastCol) {
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({
    s: { r: headerRow, c: 0 },
    e: { r: Math.max(lastRow, headerRow), c: lastCol },
  }) };
}

// Roster columns in the exports: Supervisor and Area once a roster is loaded,
// plus Scope when the roster has one. Rows carry the values the screen shows.
function rosterColumns(rosterInfo) {
  if (!rosterInfo || !rosterInfo.loaded) return { hdrs: [], widths: [], vals: () => [] };
  const keys = ["supervisor", "area", ...(rosterInfo.hasScope ? ["scope"] : [])];
  return {
    hdrs: keys.map(k => k[0].toUpperCase() + k.slice(1)),
    widths: keys.map(() => 24),
    vals: r => r.notInRoster ? keys.map(k => (k === "area" ? "Not in roster" : "")) : keys.map(k => (r[k] && r[k] !== "N/A" ? r[k] : "")),
  };
}

// The Excel codes for a day that has no hours counted
const DAY_CODES = { "Missing IN": "MISS", "Missing OUT": "MISS", "Missing IN & OUT": "MISS", "Late IN": "LATE IN", "Late OUT": "LATE OUT" };
const dayCell = d => !d ? "" : d.billHrs != null ? decToHM(d.billHrs) : (DAY_CODES[d.status] || "0:00");

function exportXLSX(results, shiftDate, reportDate, nqGates, rules = RULES_DEFAULT, area = null, rosterInfo = null) {
  const wb = XLSX.utils.book_new();
  const rc = rosterColumns(rosterInfo);

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
  const c = statusCounts(results);
  const totalH = results.reduce((s, r) => s + (r.billHrs || 0), 0);
  const withHrs = results.filter(r => r.billHrs != null).length;

  const dataColHdrs = [
    "Employee Name","Company","EIN", ...rc.hdrs,
    "Raw Badge IN","IN Gate","Rounded IN",
    "Raw Badge OUT","OUT Gate","Rounded OUT",
    "Raw Elapsed (H:MM)","Less Lunch (H:MM)","Billable Hrs (H:MM)",
    "NQ Badge IN (ref)","NQ IN Gate","NQ Badge OUT (ref)","NQ OUT Gate","Status","Note"
  ];
  const dataRows = r => [
    r.name, r.company, r.ein, ...rc.vals(r),
    toDateStr(r.rawIn),  r.inGate,  toDateStr(r.rndIn),
    toDateStr(r.rawOut), r.outGate, toDateStr(r.rndOut),
    decToHM(r.rawHrs), decToHM(r.lessLunch), decToHM(r.billHrs),
    r.nqIn  ? toDateStr(r.nqIn.ts)  : "", r.nqIn  ? r.nqIn.gate  : "",
    r.nqOut ? toDateStr(r.nqOut.ts) : "", r.nqOut ? r.nqOut.gate : "",
    r.status, r.note || "",
  ];
  const colW = [28,10,10, ...rc.widths, 22,26,22,22,26,22,16,16,16,22,26,22,26,18,60];

  const btWs = XLSX.utils.aoa_to_sheet([
    [`CP2 LNG  |  MMR Gate Log Billable Time  |  Shift Date: ${shiftDate}  |  Report Date: ${reportDate}`],
    [`Total Workers: ${results.length}`, "", "", `Clean: ${c.Clean}`, "", "",
     `Missing OUT: ${c["Missing OUT"]}`, "", "", `Missing IN: ${c["Missing IN"]}`, "", "",
     `Missing Both: ${c["Missing IN & OUT"]}`, "", "", `Late IN: ${c["Late IN"]}`, "", "",
     `Over 13 Hrs: ${c.over13}`, "", "", `Total Billable Hrs: ${decToHM(totalH)}`],
    [(rules.roundTo > 0
       ? `Rounding: Raw IN and Raw OUT each rounded independently to nearest ${rules.roundTo} min via ${midpointToStr(rules.midpointSec)} midpoint rule`
       : "Rounding: none, exact badge times used") +
     `  |  Billable = (Rounded OUT minus Rounded IN) minus ${rules.lunchMin} min lunch  |  Times displayed as H:MM` +
     (extraRulesSummary(rules) ? `  |  ${extraRulesSummary(rules)}` : "")],
    dataColHdrs,
    ...results.map(dataRows),
  ]);
  setColWidths(btWs, colW);
  setAutoFilter(btWs, 3, 3 + results.length, dataColHdrs.length - 1);
  XLSX.utils.book_append_sheet(wb, btWs, "BILLABLE TIME");

  // ─── TAB 3: ISSUES
  const issues = results.filter(r => r.status !== "Clean");
  const issWs = XLSX.utils.aoa_to_sheet([
    [`ISSUES  |  Shift: ${shiftDate}  |  Report: ${reportDate}  |  ` +
     `Missing OUT: ${c["Missing OUT"]}  |  Missing IN: ${c["Missing IN"]}  |  Missing Both: ${c["Missing IN & OUT"]}  |  ` +
     `For review: Late IN ${c["Late IN"]}, Late OUT ${c["Late OUT"]}, Split Shift ${c["Split Shift"]}, Over 13 Hrs ${c.over13}`],
    dataColHdrs,
    ...issues.map(dataRows),
  ]);
  setColWidths(issWs, colW);
  setAutoFilter(issWs, 1, 1 + issues.length, dataColHdrs.length - 1);
  XLSX.utils.book_append_sheet(wb, issWs, "ISSUES");

  // ─── TAB 4: ISSUES BY AREA
  if (area) {
    XLSX.utils.book_append_sheet(wb,
      areaSheet(`ISSUES BY AREA  |  Shift: ${shiftDate}  |  Report: ${reportDate}`, area.rows, area.basis),
      "ISSUES BY AREA");
  }

  // ─── TAB 5: SUMMARY
  const sumRows = [
    ["Total Workers on Site",         results.length],
    ["Clean Records",                 c.Clean],
    ["Missing Badge OUT",             c["Missing OUT"]],
    ["Missing Badge IN",              c["Missing IN"]],
    ["Missing IN and OUT",            c["Missing IN & OUT"]],
    ["Late IN (review)",              c["Late IN"]],
    ["Late OUT (review)",             c["Late OUT"]],
    ["Split Shift (review)",          c["Split Shift"]],
    ["Over 13 Billable Hrs (verify)", c.over13],
    ["Total Billable Hours",          decToHM(totalH)],
    ["Avg Billable Hrs per Worker",   withHrs > 0 ? decToHM(totalH / withHrs) : ""],
  ];
  const sumWs = XLSX.utils.aoa_to_sheet([
    [`DAILY GATE LOG SUMMARY  |  Shift: ${shiftDate}  |  Report: ${reportDate}`],
    ["Metric", "Value"],
    ...sumRows,
  ]);
  setColWidths(sumWs, [34, 14]);
  setAutoFilter(sumWs, 1, 1 + sumRows.length, 1);
  XLSX.utils.book_append_sheet(wb, sumWs, "SUMMARY");

  // ─── TAB 6: CONFIG
  const cfgWs = XLSX.utils.aoa_to_sheet([
    ["CONFIGURATION  |  Update dates and NQ gates here"],
    ["DATE SETTINGS", "Value", "Notes"],
    ["Shift Date (date being analyzed)",    shiftDate,  "Badge INs are only pulled from this date"],
    ["Report Date (date data was received)", reportDate, `Night shift OUTs count on this date until ${NIGHT_CUT_H}:00 AM`],
    [],
    ["RULES", "Value", "Notes"],
    ["Rounding Rule",    rulesSummary(rules), "Editable in the app Config panel"],
    ["Lunch Deduction",  `${rules.lunchMin} minutes`, "Applied when both IN and OUT exist"],
    ["Lunch Threshold",  rules.lunchOverHrs > 0 ? `Only days over ${rules.lunchOverHrs} hrs` : "Every day",
     "Rounded span must be strictly longer than the threshold for lunch to come off"],
    ["Earliest Counted IN", rules.earliestInMin > 0 ? minutesToClock(rules.earliestInMin) : "Off",
     "Badge INs before this time count as exactly this time"],
    ["IN Window",        rules.inWindowHrs > 0 ? `${rules.inWindowHrs} hrs after first badge` : "Off",
     "An onsite IN later than this after the first badge of the day is flagged Late IN, with no hours counted"],
    ["Badge IN Logic",   "First onsite IN on shift date", "Report date INs are excluded"],
    ["Badge OUT Logic",  "Last onsite OUT after the IN", `AM IN: OUT on the shift date | PM IN: shift date or next day by ${NIGHT_CUT_H}:00 AM`],
    ["Back IN after an OUT", "Same shift", "Badged OUT, came back IN, never badged OUT on site after: Missing OUT"],
    ["NQ Gate Rule",     "Off site swipes shown as reference only", "Not used in calculations"],
    ["Over 13 Hrs Flag", "Any billable total over 13 hrs flagged", "Shown in Issues tab for review"],
    [],
    ["OFF SITE (NON QUALIFYING) GATES  |  Add or remove rows as needed"],
    ["Gate Name (exact match)"],
    ...[...nqGates].sort().map(g => [g]),
    [],
    ...(nqGates.has("S-Curve Turnstile 2") ? [] :
      [["NOTE: S-Curve Turnstile 2 is NOT in the NQ list. Monitor and add via Config if needed."]]),
  ]);
  setColWidths(cfgWs, [46, 36, 52]);
  XLSX.utils.book_append_sheet(wb, cfgWs, "CONFIG");

  downloadWorkbook(wb, `CP2_GateLog_${shiftDate}.xlsx`);
}

// ── WEEKLY ROLLUP EXPORT ──────────────────────────────────────────────────────
// rollup: { emps: [{name, company, ein, supervisor, area, scope, notInRoster, days:{date:{billHrs,status}}, daysWorked, issueDays, total}], dates, dayCounts }
// dayMeta: [{shift_date, total_workers, total_bill_hrs}] for the picked reports
function exportWeeklyXLSX(rollup, dayMeta, rules = RULES_DEFAULT, area = null, rosterInfo = null) {
  const wb = XLSX.utils.book_new();
  const { emps, dates } = rollup;
  const first = dates[0], last = dates[dates.length - 1];
  const fmtD = d => `${weekdayOf(d)} ${d.slice(5,7)}/${d.slice(8,10)}`;
  const rc = rosterColumns(rosterInfo);

  function setColWidths(ws, widths) {
    ws["!cols"] = widths.map(w => ({ wch: w }));
  }

  // ─── TAB 1: WEEKLY ROLLUP
  const grandTotal = emps.reduce((s, e) => s + e.total, 0);
  const hdrs = ["Employee Name","Company","EIN", ...rc.hdrs, ...dates.map(fmtD),"Days Worked","Issue Days","Total Hrs (H:MM)"];
  const empRow = e => [
    e.name, e.company, e.ein, ...rc.vals(e),
    ...dates.map(d => dayCell(e.days[d])),
    e.daysWorked, e.issueDays, decToHM(e.total),
  ];
  const dayTotals = dates.map(d =>
    decToHM(emps.reduce((s, e) => s + ((e.days[d] && e.days[d].billHrs) || 0), 0)));
  const pad = rc.hdrs.map(() => "");

  const wkWs = XLSX.utils.aoa_to_sheet([
    [`CP2 LNG  |  MMR Gate Log Weekly Rollup  |  ${first} to ${last}  |  ${dates.length} day(s)  |  ${emps.length} employees`],
    [(rules.roundTo > 0
       ? `Rounding: nearest ${rules.roundTo} min via ${midpointToStr(rules.midpointSec)} midpoint`
       : "Rounding: none") +
     `  |  Lunch: ${rules.lunchMin} min  |  MISS = badge record incomplete that day (0 hrs counted)  |  LATE IN, LATE OUT = flagged for review (0 hrs counted)  |  Times as H:MM` +
     (extraRulesSummary(rules) ? `  |  ${extraRulesSummary(rules)}` : "")],
    hdrs,
    ...emps.map(empRow),
    [],
    ["DAILY TOTALS","","", ...pad, ...dayTotals, "", "", decToHM(grandTotal)],
  ]);
  setColWidths(wkWs, [28,10,10, ...rc.widths, ...dates.map(() => 10), 12, 11, 15]);
  setAutoFilter(wkWs, 2, 2 + emps.length, hdrs.length - 1);
  XLSX.utils.book_append_sheet(wb, wkWs, "WEEKLY ROLLUP");

  // ─── TAB 2: DAILY SUMMARY (counted from each saved day's results)
  const sumHdr = ["Shift Date","Workers","Clean","Missing OUT","Missing IN","Missing Both","Late IN","Late OUT","Split Shift","Over 13 Hrs","Total Billable Hrs"];
  const dsWs = XLSX.utils.aoa_to_sheet([
    [`DAILY SUMMARY  |  ${first} to ${last}`],
    sumHdr,
    ...dayMeta.map(d => {
      const k = (rollup.dayCounts && rollup.dayCounts[d.shift_date]) || {};
      return [
        `${weekdayOf(d.shift_date)} ${d.shift_date}`, d.total_workers, k.Clean ?? "", k["Missing OUT"] ?? "", k["Missing IN"] ?? "",
        k["Missing IN & OUT"] ?? "", k["Late IN"] ?? "", k["Late OUT"] ?? "", k["Split Shift"] ?? "", k.over13 ?? "",
        decToHM(Number(d.total_bill_hrs) || 0),
      ];
    }),
    [],
    ["TOTAL", "", "", "", "", "", "", "", "", "",
     decToHM(dayMeta.reduce((s, d) => s + (Number(d.total_bill_hrs) || 0), 0))],
  ]);
  setColWidths(dsWs, [16,10,10,12,12,13,10,10,11,12,18]);
  setAutoFilter(dsWs, 1, 1 + dayMeta.length, sumHdr.length - 1);
  XLSX.utils.book_append_sheet(wb, dsWs, "DAILY SUMMARY");

  // ─── TAB 3: ISSUES BY AREA
  if (area) {
    XLSX.utils.book_append_sheet(wb,
      areaSheet(`ISSUES BY AREA  |  ${first} to ${last}  |  ${dates.length} day(s)`, area.rows, area.basis,
        "Each person counts once per day flagged."),
      "ISSUES BY AREA");
  }

  downloadWorkbook(wb, `CP2_GateLog_Weekly_${first}_to_${last}.xlsx`);
}

// ── STATUS CONFIG ─────────────────────────────────────────────────────────────
const STATUS_STYLES = {
  "Clean":            { bg:"#0d3b1e", text:"#4ade80", border:"#166534" },
  "Missing OUT":      { bg:"#3b1010", text:"#f87171", border:"#991b1b" },
  "Missing IN":       { bg:"#3b1010", text:"#f87171", border:"#991b1b" },
  "Missing IN & OUT": { bg:"#3b1010", text:"#f87171", border:"#991b1b" },
  "Over 13 Hrs":      { bg:"#3b2500", text:"#fb923c", border:"#9a3412" },
  // Flagged for review: shown with the times seen, no hours counted (Split
  // Shift counts its main shift)
  "Late IN":          { bg:"#3a2f05", text:"#facc15", border:"#a16207" },
  "Late OUT":         { bg:"#3a2f05", text:"#facc15", border:"#a16207" },
  "Split Shift":      { bg:"#3a2f05", text:"#facc15", border:"#a16207" },
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

// ── ISSUES BY AREA ────────────────────────────────────────────────────────────
// Every missing badge flag is placed in an area. With a roster loaded that's
// the person's work area from the roster. Without one it's the gate area where
// the unmatched swipe happened: the IN gate for a Missing OUT, the OUT gate for
// a Missing IN, and the last place they badged for a Missing IN & OUT.

const AREA_BASIS_TEXT = {
  roster: "Grouped by each person's work area from the roster.",
  gate: "Grouped by gate area, where the unmatched badge happened: the IN gate for a Missing OUT, the OUT gate for a Missing IN, the last place badged for Missing Both. Upload the roster and it switches to the roster's work areas.",
};

// Employee numbers are matched without regard to case or stray spaces
const einKey = v => String(v == null ? "" : v).trim().toUpperCase();

// Roster columns, found by header name so extra or reordered columns are fine
function rosterKeyMap(keys) {
  const km = {};
  for (const k of keys) {
    const lower = String(k).toLowerCase().trim();
    if (!km.ein && (/\bein\b/.test(lower) || lower.includes("employee number") || lower.includes("employee #") || lower.includes("emp #") || lower.includes("emp no"))) km.ein = k;
    else if (!km.firstName && lower.includes("first")) km.firstName = k;
    else if (!km.lastName && lower.includes("last")) km.lastName = k;
    else if (!km.supervisor && (lower.includes("supervisor") || lower.includes("foreman"))) km.supervisor = k;
    else if (!km.area && lower.includes("area")) km.area = k;
    else if (!km.scope && lower.includes("scope")) km.scope = k;
  }
  return km;
}

// Blank roster template, or the current roster in the same layout so it can be
// edited and uploaded again
const ROSTER_HEADERS = ["EIN", "Area", "Supervisor", "First Name", "Last Name"];
function exportRosterTemplate(roster = []) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ROSTER_HEADERS,
    ...roster.map(e => [e.ein, e.area, e.supervisor, e.firstName, e.lastName]),
  ]);
  ws["!cols"] = [{ wch: 14 }, { wch: 30 }, { wch: 26 }, { wch: 18 }, { wch: 22 }];
  setAutoFilter(ws, 0, Math.max(roster.length, 1), ROSTER_HEADERS.length - 1);
  XLSX.utils.book_append_sheet(wb, ws, "ROSTER");
  const how = XLSX.utils.aoa_to_sheet([
    ["CP2 Gate Log roster"],
    [],
    ["One row per employee on the ROSTER tab. Uploading a roster replaces the previous one."],
    [],
    ["Column", "What to enter"],
    ["EIN", "Required. The employee number exactly as it shows in the Genetec data (Company EIN), for example 143909 or MMR4502."],
    ["Area", "The work area the employee is assigned to. Issues by Area groups by this column once the roster is uploaded."],
    ["Supervisor", "Optional. Leave blank if the employee has no assigned supervisor."],
    ["First Name, Last Name", "Optional. Shown on the Absent list."],
  ]);
  how["!cols"] = [{ wch: 24 }, { wch: 110 }];
  XLSX.utils.book_append_sheet(wb, how, "HOW TO FILL IN");
  downloadWorkbook(wb, roster.length
    ? `CP2_GateLog_Roster_${toYMD(new Date())}.xlsx`
    : "CP2_GateLog_Roster_Template.xlsx");
}

function areaSheet(title, rows, basis, countNote) {
  const hdr = ["Area", "Missing IN", "Missing OUT", "Missing IN & OUT", "Total Flags", "Share of Flags (%)"];
  const all = rows.reduce((s, r) => s + r.total, 0);
  const sum = k => rows.reduce((s, r) => s + r[k], 0);
  const ws = XLSX.utils.aoa_to_sheet([
    [title],
    [AREA_BASIS_TEXT[basis] + (countNote ? `  ${countNote}` : "")],
    hdr,
    ...rows.map(r => [r.area, r.missingIn, r.missingOut, r.missingBoth, r.total,
      all ? Math.round((r.total / all) * 1000) / 10 : 0]),
    [],
    ["TOTAL", sum("missingIn"), sum("missingOut"), sum("missingBoth"), all, all ? 100 : 0],
  ]);
  ws["!cols"] = [{ wch: 40 }, { wch: 12 }, { wch: 13 }, { wch: 17 }, { wch: 12 }, { wch: 18 }];
  setAutoFilter(ws, 2, 2 + rows.length, hdr.length - 1);
  return ws;
}

function AreaTable({ title, rows, basis, note }) {
  const [showAll, setShowAll] = useState(false);
  if (!rows.length) return null;
  const all = rows.reduce((s, r) => s + r.total, 0);
  const top = rows[0].total || 1;
  const shown = showAll ? rows : rows.slice(0, 8);
  const th = { padding:"8px 12px", textAlign:"left", fontSize:10, color:"#4b5e7a",
    textTransform:"uppercase", letterSpacing:0.8, fontWeight:600, whiteSpace:"nowrap" };
  const num = { padding:"8px 12px", fontSize:12, fontFamily:"'DM Mono',monospace", whiteSpace:"nowrap" };
  return (
    <div data-testid="area-table" style={{ background:"#0a1320", border:"1px solid #1a2535", borderRadius:10, overflow:"hidden", marginBottom:14 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"baseline", flexWrap:"wrap", gap:6,
        padding:"10px 12px", borderBottom:"1px solid #1a2535" }}>
        <div style={{ fontSize:12, fontWeight:700, color:"#e2e8f0", letterSpacing:0.6 }}>{title}</div>
        <div style={{ fontSize:11, color:"#6b7a8d", fontFamily:"'DM Mono',monospace" }}>
          {all} flag{all === 1 ? "" : "s"} in {rows.length} area{rows.length === 1 ? "" : "s"}
        </div>
      </div>
      <div style={{ overflowX:"auto", WebkitOverflowScrolling:"touch" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:600 }}>
          <thead>
            <tr style={{ background:"#060d16", borderBottom:"1px solid #1a2535" }}>
              {["Area", "Missing IN", "Missing OUT", "Missing Both", "Total", ""].map((h, i) => <th key={i} style={th}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {shown.map((r, i) => (
              <tr key={r.area} style={{ borderBottom:"1px solid #1a2535", background: i === 0 ? "#1a1010" : "transparent" }}>
                <td style={{ padding:"8px 12px", fontSize:12, color: i === 0 ? "#fca5a5" : "#c8d6e8", fontWeight: i === 0 ? 700 : 400 }}>{r.area}</td>
                <td style={num}>{r.missingIn}</td>
                <td style={num}>{r.missingOut}</td>
                <td style={num}>{r.missingBoth}</td>
                <td style={{ ...num, fontWeight:700, color:"#e2e8f0" }}>{r.total}</td>
                <td style={{ padding:"8px 12px", width:"28%" }}>
                  <div style={{ height:6, borderRadius:3, background:"#1a2535" }}>
                    <div style={{ height:6, borderRadius:3, background: i === 0 ? "#ef4444" : "#b45353",
                      width: `${Math.max(2, Math.round((r.total / top) * 100))}%` }} />
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, padding:"8px 12px" }}>
        <div style={{ fontSize:10, color:"#4b5e7a", lineHeight:1.5, flex:"1 1 280px" }}>
          {AREA_BASIS_TEXT[basis]}{note ? ` ${note}` : ""}
        </div>
        {rows.length > 8 && (
          <button onClick={() => setShowAll(v => !v)} style={{
            background:"transparent", color:"#6b7a8d", border:"1px solid #1a2d45", borderRadius:6,
            padding:"5px 10px", fontSize:11, cursor:"pointer", fontFamily:"'DM Mono',monospace" }}>
            {showAll ? "Show top 8" : `Show all ${rows.length} areas`}
          </button>
        )}
      </div>
    </div>
  );
}

const td = { padding:"10px 12px", color:"#c8d6e8", fontSize:13, verticalAlign:"middle" };

// ── ROW: touch safe expand ───────────────────────────────────────────────────
function Row({ r, idx, showRoster, showScope }) {
  const [open, setOpen] = useState(false);
  const isIssue = r.status !== "Clean";
  const isReview = REVIEW_STATUSES.has(r.status) && r.status !== "Over 13 Hrs";

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
    : isReview
      ? (idx % 2 === 0 ? "#1d1a0a" : "#181507")
    : isIssue
      ? (idx % 2 === 0 ? "#1e0d0d" : "#180b0b")
      : (idx % 2 === 0 ? "#0d1520" : "#0a1118");

  const colSpan = 7 + (showRoster ? (showScope ? 3 : 2) : 0);

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
          {showScope && <td style={{...td, fontSize:12, color: r.scope === "N/A" ? "#4b5e7a" : "#c8d6e8"}}>{r.scope}</td>}
        </>)}
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontSize:12}}>
          {toDateStr(r.rndIn)  || <span style={{color:"#ef4444"}}>MISSING</span>}
        </td>
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontSize:12}}>
          {toDateStr(r.rndOut) || <span style={{color:"#ef4444"}}>MISSING</span>}
        </td>
        <td style={{...td, fontFamily:"'DM Mono',monospace", fontWeight:700,
          color: r.over13 ? "#fb923c" : "#e2e8f0"}}>
          {decToHM(r.billHrs) || ""}
        </td>
        <td style={td} className={r.note ? "gl-note" : undefined} title={r.note || undefined}>
          <Badge status={r.status} />
        </td>
      </tr>

      {open && (
        <tr style={{ background:"#060d16", borderBottom:"1px solid #1a2535" }}>
          <td colSpan={colSpan} style={{ padding:"12px 14px" }}>
            {r.note && (
              <div data-testid="row-note" style={{ background:"#1d1a0a", border:"1px solid #a16207", borderRadius:6,
                padding:"8px 10px", marginBottom:8, color:"#fde68a", fontSize:12, lineHeight:1.5 }}>{r.note}</div>
            )}
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
                ["NQ IN (ref)",   r.nqIn  ? `${toDateStr(r.nqIn.ts)} at ${r.nqIn.gate}`   : ""],
                ["NQ OUT (ref)",  r.nqOut ? `${toDateStr(r.nqOut.ts)} at ${r.nqOut.gate}` : ""],
                ...(showRoster ? [
                  ["Supervisor",  r.supervisor],
                  ["Area",        r.area],
                  ...(showScope ? [["Scope", r.scope]] : []),
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
                    {v || <span style={{color:"#374151"}}>none</span>}
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

// ── SAVED COPIES ──────────────────────────────────────────────────────────────
// A day's headline numbers, to compare a new upload with a copy already saved
function dayPreview(res) {
  const c = statusCounts(res);
  return { workers: res.length, flags: c.flags, review: c.review, hrs: res.reduce((s, r) => s + (r.billHrs || 0), 0) };
}
const bulkDone = f => /^(Saved|Skipped)/.test(f.status || "");
const savedFlags = c => (c.missing_in || 0) + (c.missing_out || 0) + (c.missing_both || 0);
const savedWhen = c => c.uploaded_at
  ? new Date(c.uploaded_at).toLocaleString("en-US", { month: "numeric", day: "numeric", hour: "numeric", minute: "2-digit" })
  : "";
const countOf = (n, one, many) => `${n} ${n === 1 ? one : many}`;
const dayLine = (workers, flags, hrs) =>
  `${countOf(workers, "worker", "workers")}, ${countOf(flags, "missing badge flag", "missing badge flags")}, ${decToHM(hrs)} hrs`;
const copyLine = c => dayLine(c.total_workers || 0, savedFlags(c), Number(c.total_bill_hrs) || 0);
const previewLine = p => dayLine(p.workers, p.flags, p.hrs);

// ── APP ───────────────────────────────────────────────────────────────────────
export default function App() {
  // ── Auth: everyone signs in with their own @mmrgrp.com account (Supabase
  // Auth). The database only answers to confirmed MMR accounts, so nothing in
  // this page grants access by itself.
  const [session, setSession]       = useState(null);
  const [authReady, setAuthReady]   = useState(false);
  const [pwStep, setPwStep]         = useState(null);  // "create" or "reset" after an emailed code or link
  const [authNotice, setAuthNotice] = useState(null);  // problem with an email link
  const appOpen = authReady && !!session && !pwStep;
  const accountEmail = (session && session.user && session.user.email) || "";

  useEffect(() => {
    let alive = true;
    authStart.then(start => {
      if (!alive) return;
      setSession(start.session);
      setPwStep(start.step);
      setAuthNotice(start.notice);
      setAuthReady(true);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((event, s) => {
      if (!alive || event === "INITIAL_SESSION") return;
      setSession(s || null);
      if (!s) { setPwStep(null); writePasswordStep(null); return; }
      // Signed in by typing an emailed code: show the password screen the code
      // was for. A reset always sets a new password. A new account chooses
      // one, unless it already has one (then it goes straight in).
      if (event === "SIGNED_IN" || event === "PASSWORD_RECOVERY") {
        const intent = takeCodeIntent();
        if (intent) {
          const meta = (s.user && s.user.user_metadata) || {};
          const step = intent === "forgot" || event === "PASSWORD_RECOVERY" ? "reset" : meta.password_set ? null : "create";
          writePasswordStep(step);
          setPwStep(step);
          setAuthNotice(null);
        }
      }
    });
    return () => { alive = false; sub.subscription.unsubscribe(); };
  }, []);

  const handlePasswordSaved = () => {
    writePasswordStep(null);
    setPwStep(null);
    setAuthNotice(null);
  };

  const [uploadedBy, setUploadedBy] = useState("");
  const [history, setHistory] = useState([]);
  const [menuOpen, setMenuOpen] = useState(false);
  const [saveMsg, setSaveMsg] = useState("");
  // One click saves a report once: the button locks while saving and stays
  // "Saved" for that report, so a double click can't make a duplicate
  const [saveBusy, setSaveBusy] = useState(false);
  const [savedResults, setSavedResults] = useState(null);
  const saveLock = useRef(false);
  // Clearing data: a delete needs a second click, clearing every report needs
  // the word CLEAR typed, and a pending confirmation times out on its own.
  const [delPending, setDelPending] = useState(null); // report id waiting for its second click
  const [clearText, setClearText]   = useState("");
  const [clearBusy, setClearBusy]   = useState(false);
  const [rosterClearPending, setRosterClearPending] = useState(false);

  // The uploader name is remembered per account on this computer and starts
  // out as the account's email address.
  useEffect(() => {
    if (!appOpen || !accountEmail) return;
    let saved = "";
    try { saved = localStorage.getItem("cp2-uploader:" + accountEmail) || ""; } catch (_) { /* storage blocked */ }
    setUploadedBy(saved || accountEmail);
  }, [appOpen, accountEmail]);

  useEffect(() => {
    if (!accountEmail || !uploadedBy) return;
    try { localStorage.setItem("cp2-uploader:" + accountEmail, uploadedBy); } catch (_) { /* storage blocked */ }
  }, [uploadedBy, accountEmail]);

  const [screen, setScreen]         = useState("upload");
  const [results, setResults]       = useState([]);
  // Default to yesterday's shift, today's report, the normal morning workflow
  const [shiftDate, setShiftDate]   = useState(() => toYMD(addDays(new Date(), -1)));
  const [reportDate, setReportDate] = useState(() => toYMD(new Date()));
  const [filter, setFilter]         = useState("All");
  const [search, setSearch]         = useState("");
  const [supFilter, setSupFilter]   = useState("");
  const [areaFilter, setAreaFilter] = useState("");
  const [scopeFilter, setScopeFilter] = useState("");
  const [nqText, setNqText]         = useState([...NQ_GATES_DEFAULT].sort().join("\n"));
  // Every gate name this device has seen, from uploads, saved reports, or the
  // shared config. Feeds the gate checklist in the Config panel.
  const [seenGates, setSeenGates]   = useState(() => {
    try {
      const cached = JSON.parse(localStorage.getItem("cp2-config") || "null");
      return cached && Array.isArray(cached.knownGates) ? cached.knownGates : [];
    } catch (_) { return []; }
  });
  const [newGate, setNewGate]       = useState("");
  const [rules, setRules]           = useState(RULES_DEFAULT);
  const [cfgMeta, setCfgMeta]       = useState(null); // {updated_by, updated_at} from database
  const [cfgMsg, setCfgMsg]         = useState("");
  const [showConfig, setShowConfig] = useState(false);
  const [dragging, setDragging]     = useState(false);
  const [fileName, setFileName]     = useState("");
  const [error, setError]           = useState("");
  const [processing, setProcessing] = useState(false);
  const fileRef = useRef();

  // ── Bulk upload state ──
  const [bulkFiles, setBulkFiles]   = useState([]); // [{name, rows, shiftDate, reportDate, status}]
  const [bulkBusy, setBulkBusy]     = useState(false);
  const [knownDates, setKnownDates] = useState(new Set()); // shift_dates already saved in DB
  const [savedCopies, setSavedCopies] = useState({});      // shift_date -> reports already saved for it
  const [resaveAsk, setResaveAsk]   = useState(null);      // a single day that's already saved, waiting for a choice
  const [dupPending, setDupPending] = useState(null);      // History: duplicate date waiting for its second click
  // A question about saving one day goes away when a different day is opened
  useEffect(() => { setResaveAsk(null); }, [results]);
  const [bulkFrom, setBulkFrom]     = useState("");        // zip name(s) the files came from

  // ── Weekly rollup state ──
  const [weeklyReports, setWeeklyReports] = useState([]); // report metadata rows
  const [selIds, setSelIds]         = useState(new Set());
  const [rollup, setRollup]         = useState(null); // {emps, dates}
  const [rollupMeta, setRollupMeta] = useState([]);   // day meta for selected reports
  const [weeklyMsg, setWeeklyMsg]   = useState("");
  const [weeklyNote, setWeeklyNote] = useState("");   // which week opened after a save
  const [weeklySearch, setWeeklySearch] = useState("");
  const [weeklyBusy, setWeeklyBusy] = useState(false);

  // ── Employee roster state ──
  const [roster, setRoster]         = useState([]); // [{ein, firstName, lastName, supervisor, area, scope}]
  const [rosterMeta, setRosterMeta] = useState(null); // {uploaded_at, uploaded_by, count}
  const [rosterMsg, setRosterMsg]   = useState("");
  const rosterMsgBad = /^(Upload error|Clear error|Enter|Database)/.test(rosterMsg);
  const rosterFileRef = useRef();
  const rosterFileRef2 = useRef();

  // Fold new gate names into the seen list, keeping it sorted and deduped
  const mergeSeenGates = useCallback((names) => {
    setSeenGates(prev => {
      const merged = new Set(prev);
      let changed = false;
      for (const n of names || []) {
        const t = String(n || "").trim();
        if (t && !merged.has(t)) { merged.add(t); changed = true; }
      }
      return changed ? [...merged].sort() : prev;
    });
  }, []);

  // Keep the seen gate list cached on this device
  useEffect(() => {
    try {
      const cached = JSON.parse(localStorage.getItem("cp2-config") || "null") || {};
      cached.knownGates = seenGates;
      localStorage.setItem("cp2-config", JSON.stringify(cached));
    } catch (_) { /* ignore */ }
  }, [seenGates]);

  // Load roster once signed in
  useEffect(() => {
    if (!appOpen || !supabaseReady()) return;
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
  }, [appOpen]);

  // Load shared config (NQ gates + rounding rules) once signed in.
  // Database is the source of truth for the team; localStorage is a fallback cache.
  useEffect(() => {
    if (!appOpen) return;
    try {
      const cached = JSON.parse(localStorage.getItem("cp2-config") || "null");
      if (cached) {
        if (Array.isArray(cached.nqGates)) setNqText(cached.nqGates.join("\n"));
        if (cached.rules) setRules({ ...RULES_DEFAULT, ...cached.rules });
      }
    } catch (_) { /* ignore bad cache */ }
    if (!supabaseReady()) return;
    (async () => {
      try {
        const { data, error } = await supabase.from("app_config").select("*");
        if (error || !data) return;
        const byKey = {};
        data.forEach(r => { byKey[r.key] = r; });
        let latest = null;
        if (byKey.nq_gates && Array.isArray(byKey.nq_gates.value)) {
          setNqText(byKey.nq_gates.value.join("\n"));
          latest = byKey.nq_gates;
        }
        if (byKey.rules && byKey.rules.value) {
          setRules({ ...RULES_DEFAULT, ...byKey.rules.value });
          if (!latest || byKey.rules.updated_at > latest.updated_at) latest = byKey.rules;
        }
        if (byKey.known_gates && Array.isArray(byKey.known_gates.value)) {
          mergeSeenGates(byKey.known_gates.value);
        }
        if (latest) setCfgMeta({ updated_by: latest.updated_by, updated_at: latest.updated_at });
        localStorage.setItem("cp2-config", JSON.stringify({
          nqGates: byKey.nq_gates ? byKey.nq_gates.value : null,
          rules: byKey.rules ? byKey.rules.value : null,
          knownGates: byKey.known_gates && Array.isArray(byKey.known_gates.value)
            ? byKey.known_gates.value : seenGates,
        }));
      } catch (_) { /* silent, defaults still apply */ }
    })();
  }, [appOpen]);

  // Save shared config for the whole team
  const handleSaveConfig = async () => {
    setCfgMsg("");
    const gates = nqText.split("\n").map(s => s.trim()).filter(Boolean);
    const cleanRules = {
      roundTo: Math.max(0, Number(rules.roundTo) || 0),
      midpointSec: Math.max(0, Number(rules.midpointSec) || 0),
      lunchMin: Math.max(0, Number(rules.lunchMin) || 0),
      earliestInMin: Math.max(0, Number(rules.earliestInMin) || 0),
      lunchOverHrs: Math.max(0, Number(rules.lunchOverHrs) || 0),
      inWindowHrs: Math.max(0, Number(rules.inWindowHrs) || 0),
    };
    // The shared seen gate list keeps the checklist consistent for the whole team
    const knownUnion = [...new Set([...seenGates, ...gates])].sort();
    setSeenGates(knownUnion);
    localStorage.setItem("cp2-config", JSON.stringify({ nqGates: gates, rules: cleanRules, knownGates: knownUnion }));
    if (!supabaseReady()) { setCfgMsg("Saved on this device only, database not configured."); return; }
    try {
      const by = uploadedBy.trim() || "unknown";
      const now = new Date().toISOString();
      const { error } = await supabase.from("app_config").upsert([
        { key: "nq_gates",    value: gates,      updated_by: by, updated_at: now },
        { key: "rules",       value: cleanRules, updated_by: by, updated_at: now },
        { key: "known_gates", value: knownUnion, updated_by: by, updated_at: now },
      ]);
      if (error) throw error;
      setCfgMeta({ updated_by: by, updated_at: now });
      setCfgMsg("Config saved, applies for everyone.");
    } catch (e) {
      setCfgMsg("Config save error: " + e.message);
    }
  };

  // Add a gate by name. It lands in the checklist already checked (off site),
  // since that is the only reason to add one by hand. Uncheck to keep it qualifying.
  const addGate = () => {
    const t = newGate.trim();
    if (!t) return;
    mergeSeenGates([t]);
    const next = currentGates();
    next.add(t);
    setNqText([...next].sort().join("\n"));
    setNewGate("");
  };

  // Roster upload handler (.xlsx or .csv, replaces the roster in the database).
  // EIN is required; area, supervisor, names and scope can be blank.
  const handleRosterUpload = async (file) => {
    if (!file) return;
    setRosterMsg("");
    if (!supabaseReady()) { setRosterMsg("Database not configured."); return; }
    if (!uploadedBy.trim()) { setRosterMsg("Enter your name first."); return; }
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", raw: /\.csv$/i.test(file.name) });

      // Use the first tab with an EIN column (the template's ROSTER tab)
      let rows = [], keyMap = null;
      for (const name of wb.SheetNames) {
        const list = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" });
        if (!list.length) continue;
        const km = rosterKeyMap(Object.keys(list[0]));
        if (km.ein) { rows = list; keyMap = km; break; }
      }
      if (!keyMap) throw new Error("No EIN column found. Start from the blank template.");

      const cell = (r, k) => (k ? String(r[k] == null ? "" : r[k]).trim() : "");
      const byEin = new Map();
      for (const r of rows) {
        const ein = cell(r, keyMap.ein);
        if (!ein) continue;
        byEin.set(ein.toUpperCase(), {   // a repeated EIN keeps its last row
          ein,
          first_name: cell(r, keyMap.firstName),
          last_name: cell(r, keyMap.lastName),
          supervisor: cell(r, keyMap.supervisor),
          area: cell(r, keyMap.area),
          scope: cell(r, keyMap.scope),
          uploaded_by: uploadedBy.trim(),
          uploaded_at: new Date().toISOString(),
        });
      }
      const parsed = [...byEin.values()];

      if (!parsed.length) throw new Error("No rows with an EIN found.");

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
  const currentGates = () => new Set(nqText.split("\n").map(s => s.trim()).filter(Boolean));

  const processFile = useCallback(async (file) => {
    if (!file) return;
    setFileName(file.name);
    setError("");
    setProcessing(true);
    try {
      const rows = await readGateFile(file);
      if (!rows.length) throw new Error("No data rows found in the file.");
      if (!swipeCount(rows)) throw new Error(NO_SWIPES);
      mergeSeenGates(rows.map(r => r["Area"]));
      const nqGates = currentGates();
      // Keep the chosen shift date when the file covers that day's afternoon.
      // Otherwise use the day the file was pulled for, read from the data.
      let sd = shiftDate, rd = reportDate;
      const picked = afternoonCounts(rows)[sd] ? null : detectShiftDate(rows, nqGates);
      if (picked && picked !== sd) {
        sd = picked;
        rd = toYMD(addDays(new Date(picked + "T00:00:00"), 1));
        setShiftDate(sd);
        setReportDate(rd);
      }
      setResults(analyze(rows, sd, rd, nqGates, rules));
      setScreen("results");
    } catch (err) {
      setError("Error: " + err.message);
    } finally {
      setProcessing(false);
    }
  }, [shiftDate, reportDate, nqText, rules, mergeSeenGates]);

  // Multiple files → bulk screen. Each file's shift date is detected; files
  // for the same shift date are combined with duplicate swipes removed, then
  // every day is saved to the database in one pass. `from` names the zip(s)
  // the files came out of, if any.
  const prepareBulk = useCallback((files, from = "") => {
    setError("");
    setProcessing(true);
    setBulkFrom(from);
    const nqGates = currentGates();
    Promise.all([...files].map(f =>
      readGateFile(f)
        .then(rows => ({ name: f.name, rows,
          problem: !rows.length ? "No data rows" : !swipeCount(rows) ? NO_SWIPES : "" }))
        .catch(err => ({ name: f.name, rows: [], problem: "Read error: " + err.message }))
    )).then(read => {
      for (const f of read) if (!f.problem) mergeSeenGates(f.rows.map(r => r["Area"]));
      // Days are grouped by the shift date in each file's own swipes, and each
      // day also takes the other files' swipes from noon the day before
      // through noon the next morning (see groupDays)
      // Each day is run once here so a day that's already saved can be
      // compared with the saved copy before anything is written
      const all = groupDays(read, nqGates, toYMD(addDays(new Date(), -1))).map(f =>
        f.rows.length && f.shiftDate
          ? { ...f, preview: dayPreview(analyze(f.rows, f.shiftDate, f.reportDate, nqGates, rules)), dup: "replace" }
          : f);
      setBulkFiles(all);
      setScreen("bulk");
      setProcessing(false);
    });
    // Load the days already in the database so a day saved before gets reviewed
    if (supabaseReady()) loadSavedCopies();
  }, [nqText, rules, mergeSeenGates]);

  const handleFiles = useCallback(async (fileList) => {
    // Copy the list right away: the browser empties it once the event is over
    const dropped = [...(fileList || [])].filter(Boolean);
    let files = dropped.filter(f => GATE_FILE_RE.test(f.name) || ZIP_RE.test(f.name));
    if (!files.length) {
      if (dropped.length) setError("Use the Genetec export: .xlsx or .csv files, or a .zip of them.");
      return;
    }
    const zips = files.filter(f => ZIP_RE.test(f.name)).map(f => f.name);
    if (zips.length) {
      setError("");
      setProcessing(true);
      try {
        files = await expandZips(files);
      } catch (e) {
        setError(e.message);
        setProcessing(false);
        return;
      }
    }
    if (files.length === 1) processFile(files[0]);
    else prepareBulk(files, zips.join(", "));
  }, [processFile, prepareBulk]);

  const onDrop = useCallback(e => {
    e.preventDefault(); setDragging(false);
    handleFiles(e.dataTransfer.files);
  }, [handleFiles]);

  const handleExport = () => {
    exportXLSX(enriched, shiftDate, reportDate, currentGates(), rules, { rows: dailyAreaRows, basis: areaBasis }, rosterInfo);
  };

  // Derived stats
  const counts = statusCounts(results);
  const clean  = counts.Clean;
  const mOut   = counts["Missing OUT"];
  const mIn    = counts["Missing IN"];
  const mBoth  = counts["Missing IN & OUT"];
  const over13 = counts.over13;
  const reviewToday = counts["Late IN"] + counts["Late OUT"] + counts["Split Shift"];
  const totalH = results.reduce((s, r) => s + (r.billHrs || 0), 0);

  // Roster lookup by EIN
  const rosterByEIN = (() => {
    const map = {};
    for (const r of roster) {
      if (r.ein) map[einKey(r.ein)] = r;
    }
    return map;
  })();

  // The roster's Scope column only shows when some row in the roster has one
  const hasScope = roster.some(r => r.scope);
  const rosterInfo = { loaded: roster.length > 0, hasScope };

  // Genetec gives a few IDs to more than one person. When an ID belongs to more
  // than one badge holder in the data on screen, a roster row only applies to
  // the one whose name matches it.
  const sharedIds = rows => {
    const names = new Map();
    for (const r of rows) {
      const k = einKey(r.ein);
      if (!k) continue;
      if (!names.has(k)) names.set(k, new Set());
      names.get(k).add(r.name);
    }
    return new Set([...names].filter(([, n]) => n.size > 1).map(([k]) => k));
  };
  const nameFits = (emp, badgeName) => {
    const last = new Set(String(badgeName || "").toUpperCase().split(",")[0].split(/\s+/).filter(Boolean));
    return String(emp.lastName || "").toUpperCase().split(/\s+/).some(t => t && last.has(t));
  };
  const rosterEmp = (rawEin, name, shared) => {
    const ein = einKey(rawEin);
    const emp = ein ? rosterByEIN[ein] : null;
    if (emp && shared && shared.has(ein) && emp.lastName && !nameFits(emp, name)) return null;
    return emp || null;
  };

  // Supervisor / area / scope from the roster, and whether the person is on it
  const rosterFields = (rawEin, name, shared) => {
    const ein = einKey(rawEin);
    const emp = rosterEmp(rawEin, name, shared);
    return {
      supervisor: emp && emp.supervisor ? emp.supervisor : "N/A",
      area:       emp && emp.area       ? emp.area       : "N/A",
      scope:      emp && emp.scope      ? emp.scope      : "N/A",
      notInRoster: roster.length > 0 && (!ein || !emp),
    };
  };

  // Enrich each result row with supervisor / area / scope / notInRoster flag
  const sharedToday = sharedIds(results);
  const enriched = results.map(r => ({ ...r, ...rosterFields(r.ein, r.name, sharedToday) }));

  // Roster employees with no badge on the shift date
  const absentEmployees = (() => {
    if (!roster.length || !results.length) return [];
    const badged = new Set(enriched.filter(r => !r.notInRoster).map(r => einKey(r.ein)));
    return roster.filter(emp => emp.ein && !badged.has(einKey(emp.ein)));
  })();

  // Filter dropdown values: build unique sorted lists from the enriched data
  const supervisors = [...new Set(enriched.map(r => r.supervisor).filter(Boolean))].sort();
  const areas       = [...new Set(enriched.map(r => r.area).filter(Boolean))].sort();
  const scopes      = [...new Set(enriched.map(r => r.scope).filter(Boolean))].sort();

  const notInRosterCount = enriched.filter(r => r.notInRoster).length;

  // Issues by area: roster work areas when a roster is loaded, gate areas otherwise
  const areaBasis = roster.length > 0 ? "roster" : "gate";
  const rosterAreaOf = (ein, name, shared) => {
    const emp = rosterEmp(ein, name, shared);
    return emp ? (emp.area || "No area in roster") : "Not in roster";
  };
  const areaOfWith = shared => areaBasis === "roster" ? (it => rosterAreaOf(it.ein, it.name, shared)) : gateFlagArea;
  const areaOf = areaOfWith(sharedToday);
  const dailyAreaRows = summarizeByArea(results, areaOf);

  const filtered = enriched.filter(r => {
    const mf =
      filter === "All"            ? true :
      filter === "Issues"         ? r.status !== "Clean" :
      filter === "Missing OUT"    ? r.status === "Missing OUT" :
      filter === "Missing IN"     ? r.status === "Missing IN" :
      filter === "Late IN"        ? r.status === "Late IN" :
      filter === "Review"         ? r.status === "Late OUT" || r.status === "Split Shift" :
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
  // Writes the columns that actually exist in the reports table
  // (the old build wrote shifts_count/issues_count, which don't, so every save failed).
  const saveReport = async (res, sDate, rDate, fname) => {
    const totalBillHrs = res.reduce((s, r) => s + (r.billHrs || 0), 0);
    const { error } = await supabase.from("reports").insert({
      shift_date: sDate,
      report_date: rDate,
      uploaded_by: uploadedBy.trim(),
      file_name: fname || null,
      results_json: { results: res, shiftDate: sDate, reportDate: rDate },
      raw_csv_path: null,
      total_workers: res.length,
      clean_count:  res.filter(r => r.status === "Clean").length,
      missing_out:  res.filter(r => r.status === "Missing OUT").length,
      missing_in:   res.filter(r => r.status === "Missing IN").length,
      missing_both: res.filter(r => r.status === "Missing IN & OUT").length,
      over_13:      res.filter(r => r.over13).length,
      total_bill_hrs: Math.round(totalBillHrs * 100) / 100,
    });
    if (error) throw error;
  };

  // Every saved report for each shift date, newest first
  const loadSavedCopies = async () => {
    const { data } = await supabase.from("reports").select(reportListSelect).order("uploaded_at", { ascending: false });
    if (!data) return {};
    const by = {};
    for (const r of data) (by[r.shift_date] ||= []).push(r);
    setKnownDates(new Set(Object.keys(by)));
    setSavedCopies(by);
    return by;
  };

  // Removes the earlier saved copies of a day once its new copy is saved.
  // Returns a note for the status line.
  const removeCopies = async (copies) => {
    const ids = copies.map(c => c.id);
    if (!ids.length) return "";
    const { error, count } = await supabase.from("reports").delete({ count: "exact" }).in("id", ids);
    if (error || count !== ids.length) {
      return `. The earlier copy was not removed (${error ? error.message : "the database did not allow it"}); remove it in History`;
    }
    return ids.length === 1 ? ". Replaced the earlier copy" : `. Replaced the ${ids.length} earlier copies`;
  };

  // Save the open day. A day that's already saved is shown next to its saved
  // copy first, and saved only once a choice is made: replace the saved copy,
  // or keep both.
  const handleSave = async (choice) => {
    if (choice === "cancel") { setResaveAsk(null); return; }
    if (saveLock.current || savedResults === results) return;
    setSaveMsg("");
    if (!uploadedBy.trim()) { setSaveMsg("Enter your name first."); return; }
    if (!supabaseReady()) { setSaveMsg("Database not configured."); return; }
    if (!results.length) { setSaveMsg("Nothing to save."); return; }
    saveLock.current = true;
    setSaveBusy(true);
    try {
      let existing = resaveAsk && choice ? resaveAsk.existing : null;
      if (!choice) {
        const { data, error } = await supabase.from("reports").select(reportListSelect)
          .eq("shift_date", shiftDate).order("uploaded_at", { ascending: false });
        if (error) throw error;
        if (data && data.length) { setResaveAsk({ existing: data, date: shiftDate }); return; }
      }
      await saveReport(results, shiftDate, reportDate, fileName);
      const note = !existing || !existing.length ? ""
        : choice === "replace" ? await removeCopies(existing)
        : existing.length === 1 ? ". Kept the earlier copy too" : `. Kept the ${existing.length} earlier copies too`;
      setResaveAsk(null);
      setSavedResults(results);
      setSaveMsg("Saved" + (note || ".") + (note && !note.endsWith(".") ? "." : ""));
      loadSavedCopies();
    } catch (e) {
      setSaveMsg("Save error: " + e.message);
    } finally {
      saveLock.current = false;
      setSaveBusy(false);
    }
  };

  // Process every prepared bulk file: analyze with current gates/rules, save each day
  const runBulk = async () => {
    if (!uploadedBy.trim()) { setBulkFiles(f => f.map(x => x.status ? x : { ...x, status: "Enter your name first." })); return; }
    if (!supabaseReady()) { setBulkFiles(f => f.map(x => x.status ? x : { ...x, status: "Database not configured." })); return; }
    if (saveLock.current) return;
    saveLock.current = true;
    setBulkBusy(true);
    const nqGates = currentGates();
    const out = [];
    // Saved copies as of now, so a choice made on screen applies to what's
    // actually in the database
    const copiesNow = await loadSavedCopies();
    for (const f of bulkFiles) {
      // A day already saved from this screen isn't saved a second time
      if (!f.rows.length || !f.shiftDate || bulkDone(f)) { out.push(f); continue; }
      const copies = copiesNow[f.shiftDate] || [];
      if (copies.length && f.dup === "skip") { out.push({ ...f, status: "Skipped: already saved" }); continue; }
      try {
        const res = analyze(f.rows, f.shiftDate, f.reportDate, nqGates, rules);
        await saveReport(res, f.shiftDate, f.reportDate, f.name);
        const note = !copies.length ? ""
          : f.dup !== "keep" ? await removeCopies(copies)
          : copies.length === 1 ? ". Kept the earlier copy too" : `. Kept the ${copies.length} earlier copies too`;
        out.push({ ...f, status: `Saved: ${res.length} workers, ${decToHM(res.reduce((s, r) => s + (r.billHrs || 0), 0))} hrs${note}` });
      } catch (e) {
        out.push({ ...f, status: "Save error: " + e.message });
      }
      setBulkFiles([...out, ...bulkFiles.slice(out.length)]);
    }
    setBulkFiles(out);
    await loadSavedCopies();
    saveLock.current = false;
    setBulkBusy(false);
  };

  const reportListSelect = "id, shift_date, report_date, uploaded_by, uploaded_at, file_name, total_workers, clean_count, missing_out, missing_in, missing_both, over_13, total_bill_hrs";

  const loadHistory = async () => {
    if (!supabaseReady()) { setSaveMsg("Database not configured."); return; }
    const { data, error } = await supabase
      .from("reports")
      .select(reportListSelect)
      .order("shift_date", { ascending: false })
      .limit(100);
    if (error) { setSaveMsg("History error: " + error.message); return; }
    setHistory(data || []);
    setDelPending(null);
    setClearText("");
    setScreen("history");
    setMenuOpen(false);
  };

  // ── Clear data ──
  // Deletes run under the database's row policies for signed in team members.
  // A count of 0 means the database removed nothing, so nothing comes off the
  // screen either and the message says so.
  const refreshKnownDates = async () => {
    const { data } = await supabase.from("reports").select("shift_date");
    if (data) setKnownDates(new Set(data.map(r => r.shift_date)));
  };

  const deleteReport = async (id) => {
    setSaveMsg("");
    if (!supabaseReady()) { setSaveMsg("Database not configured."); return; }
    setClearBusy(true);
    try {
      const { error, count } = await supabase.from("reports").delete({ count: "exact" }).eq("id", id);
      if (error) throw error;
      if (!count) throw new Error("the database did not allow the delete");
      setHistory(h => h.filter(r => r.id !== id));
      setWeeklyReports(w => w.filter(r => r.id !== id));
      setSelIds(s => { const next = new Set(s); next.delete(id); return next; });
      setRollup(null);
      setSaveMsg("Report deleted.");
      refreshKnownDates();
    } catch (e) {
      setSaveMsg("Delete error: " + e.message);
    } finally {
      setDelPending(null);
      setClearBusy(false);
    }
  };

  // A day saved more than once: keep the newest copy, remove the older ones
  const removeOlderCopies = async (date) => {
    setSaveMsg("");
    const copies = history.filter(h => h.shift_date === date)
      .sort((a, b) => String(b.uploaded_at || "").localeCompare(String(a.uploaded_at || "")));
    const older = copies.slice(1).map(c => c.id);
    if (!older.length) return;
    setClearBusy(true);
    try {
      const { error, count } = await supabase.from("reports").delete({ count: "exact" }).in("id", older);
      if (error) throw error;
      if (count !== older.length) throw new Error("the database did not allow the delete");
      setHistory(h => h.filter(r => !older.includes(r.id)));
      setWeeklyReports(w => w.filter(r => !older.includes(r.id)));
      setSelIds(sel => { const next = new Set(sel); older.forEach(id => next.delete(id)); return next; });
      setRollup(null);
      setSaveMsg(`${dayLabel(date)}: kept the copy saved ${savedWhen(copies[0])}, removed ${older.length} older cop${older.length === 1 ? "y" : "ies"}.`);
      refreshKnownDates();
    } catch (e) {
      setSaveMsg("Delete error: " + e.message);
    } finally {
      setDupPending(null);
      setClearBusy(false);
    }
  };

  const clearAllReports = async () => {
    setSaveMsg("");
    if (!supabaseReady()) { setSaveMsg("Database not configured."); return; }
    setClearBusy(true);
    try {
      const { error, count } = await supabase.from("reports").delete({ count: "exact" }).not("id", "is", null);
      if (error) throw error;
      setHistory([]); setWeeklyReports([]); setSelIds(new Set()); setRollup(null); setKnownDates(new Set());
      setSaveMsg(`Cleared ${count || 0} saved report${count === 1 ? "" : "s"}.`);
    } catch (e) {
      setSaveMsg("Clear error: " + e.message);
    } finally {
      setClearText("");
      setClearBusy(false);
    }
  };

  const clearRoster = async () => {
    setRosterMsg("");
    if (!supabaseReady()) { setRosterMsg("Database not configured."); setRosterClearPending(false); return; }
    try {
      const { error, count } = await supabase.from("employees").delete({ count: "exact" }).neq("ein", "___never___");
      if (error) throw error;
      setRoster([]);
      setRosterMeta(null);
      setRosterMsg(`Roster cleared, ${count || 0} employee${count === 1 ? "" : "s"} removed.`);
    } catch (e) {
      setRosterMsg("Clear error: " + e.message);
    } finally {
      setRosterClearPending(false);
    }
  };

  // A confirmation that is not acted on goes back to normal after a few seconds
  useEffect(() => {
    if (delPending == null && !rosterClearPending && dupPending == null) return;
    const t = setTimeout(() => { setDelPending(null); setRosterClearPending(false); setDupPending(null); }, 8000);
    return () => clearTimeout(t);
  }, [delPending, rosterClearPending, dupPending]);

  // Saved results have their Date fields serialized to strings, so revive them
  const reviveResults = (list) => (list || []).map(r => ({
    ...r,
    rawIn:  r.rawIn  ? new Date(r.rawIn)  : null,
    rawOut: r.rawOut ? new Date(r.rawOut) : null,
    rndIn:  r.rndIn  ? new Date(r.rndIn)  : null,
    rndOut: r.rndOut ? new Date(r.rndOut) : null,
    nqIn:   r.nqIn  ? { ...r.nqIn,  ts: new Date(r.nqIn.ts)  } : null,
    nqOut:  r.nqOut ? { ...r.nqOut, ts: new Date(r.nqOut.ts) } : null,
  }));

  const loadReport = async (id) => {
    const { data, error } = await supabase.from("reports").select("*").eq("id", id).single();
    if (error) { setSaveMsg("Load error: " + error.message); return; }
    const payload = data.results_json || {};
    const revived = reviveResults(payload.results);
    setResults(revived);
    mergeSeenGates(revived.flatMap(r => [
      r.inGate, r.outGate,
      r.nqIn ? r.nqIn.gate : "", r.nqOut ? r.nqOut.gate : "",
    ]));
    if (payload.shiftDate) setShiftDate(payload.shiftDate);
    if (payload.reportDate) setReportDate(payload.reportDate);
    setScreen("results");
  };

  // ── Weekly rollup handlers ──
  // Loads the saved report list and opens the weekly screen. Returns the list
  // so a caller can pick days from it straight away.
  const loadWeekly = async () => {
    setWeeklyMsg("");
    setWeeklyNote("");
    setMenuOpen(false);
    if (!supabaseReady()) { setWeeklyMsg("Database not configured."); setScreen("weekly"); return []; }
    const { data, error } = await supabase
      .from("reports")
      .select(reportListSelect)
      .order("shift_date", { ascending: false })
      .limit(200);
    if (error) { setWeeklyMsg("Load error: " + error.message); }
    setWeeklyReports(data || []);
    setRollup(null);
    setScreen("weekly");
    return data || [];
  };

  const toggleSel = (id) => {
    setSelIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
    setRollup(null);
    setWeeklyNote("");
  };

  // Every saved day in the list, the newest copy of a day saved twice
  const selectAll = () => {
    setSelIds(new Set(newestPerDate(weeklyReports).map(r => r.id)));
    setRollup(null);
    setWeeklyNote("");
  };

  // Quick select a Monday to Sunday week: 0 = this week, -1 = last week
  const selectWeek = (offset) => {
    const days = weekDays(toYMD(addDays(new Date(weekStart(toYMD(new Date())) + "T00:00:00"), offset * 7)));
    setSelIds(new Set(newestPerDate(weeklyReports, days).map(r => r.id)));
    setRollup(null);
    setWeeklyNote("");
  };

  // After a bulk save: open the rollup on the Monday to Sunday week (by shift
  // date) holding most of the days just saved, with the newest saved report
  // for each day of that week picked, and build it. Saved days that belong to
  // another week, and days of this week with nothing saved, are named on
  // screen so nothing drops out without notice.
  const openWeekAfterSave = async (savedDates) => {
    const reports = await loadWeekly();
    if (!savedDates.length || !reports.length) return;
    const count = {};
    for (const d of savedDates) { const m = weekStart(d); count[m] = (count[m] || 0) + 1; }
    const mon = Object.keys(count).sort((a, b) => count[b] - count[a] || b.localeCompare(a))[0];
    const days = weekDays(mon);
    const inWeek = newestPerDate(reports, days);
    const ids = new Set(inWeek.map(r => r.id));
    setSelIds(ids);
    const have = new Set(inWeek.map(r => r.shift_date));
    const missing = days.filter(d => !have.has(d));
    const other = [...new Set(savedDates.filter(d => !days.includes(d)))].sort();
    const notes = [`Week of ${dayLabel(days[0])} to ${dayLabel(days[6])}.`];
    if (missing.length) notes.push(`Not saved yet: ${missing.map(dayLabel).join(", ")}.`);
    if (other.length) notes.push(`Also saved, in another week: ${other.map(dayLabel).join(", ")}.`);
    setWeeklyNote(notes.join(" "));
    await buildRollupFor(ids, reports);
  };

  const buildRollup = () => buildRollupFor(selIds, weeklyReports);

  const buildRollupFor = async (ids, reports) => {
    setWeeklyMsg("");
    const chosen = reports.filter(r => ids.has(r.id));
    if (!chosen.length) { setWeeklyMsg("Select at least one day first."); return; }
    setWeeklyBusy(true);
    try {
      const { data, error } = await supabase
        .from("reports")
        .select("id, shift_date, uploaded_at, results_json")
        .in("id", [...ids]);
      if (error) throw error;
      // One copy per shift date: when two saved copies of a day are picked,
      // only the later saved one counts, for the people and for the totals
      const sorted = newestPerDate(data || []).sort((a, b) => a.shift_date.localeCompare(b.shift_date));
      const byEmp = {};
      const dayCounts = {};
      const dates = [...new Set(sorted.map(r => r.shift_date))].sort();
      for (const rep of sorted) {
        const res = (rep.results_json && rep.results_json.results) || [];
        dayCounts[rep.shift_date] = statusCounts(res);
        for (const r of res) {
          const key = r.guid || `${r.ein}|${r.name}`;
          const e = byEmp[key] || (byEmp[key] = {
            name: r.name || "", company: r.company || "", ein: r.ein || "",
            days: {}, daysWorked: 0, issueDays: 0, total: 0,
          });
          e.days[rep.shift_date] = { billHrs: r.billHrs, status: r.status, over13: r.over13, flagArea: gateFlagArea(r), note: r.note || "" };
        }
      }
      const emps = Object.values(byEmp).map(e => {
        let daysWorked = 0, issueDays = 0, total = 0;
        for (const d of Object.values(e.days)) {
          if (d.billHrs != null) { daysWorked++; total += d.billHrs; }
          if (d.status !== "Clean") issueDays++;
        }
        return { ...e, daysWorked, issueDays, total };
      }).sort((a, b) => a.name.localeCompare(b.name));
      setRollup({ emps, dates, dayCounts });
      setRollupMeta(newestPerDate(chosen).sort((a, b) => a.shift_date.localeCompare(b.shift_date)));
    } catch (e) {
      setWeeklyMsg("Rollup error: " + e.message);
    } finally {
      setWeeklyBusy(false);
    }
  };

  const handleLogout = async () => {
    setMenuOpen(false);
    // Local scope ends the session in this browser only; the person stays
    // signed in on their other devices.
    try { await supabase.auth.signOut({ scope: "local" }); } catch (_) { /* already signed out */ }
    writePasswordStep(null);
    setSession(null); setPwStep(null);
    setResults([]); setFileName(""); setScreen("upload");
    setHistory([]); setWeeklyReports([]); setRollup(null); setSelIds(new Set());
    setRoster([]); setRosterMeta(null); setUploadedBy("");
  };

  const dupDates = (() => {
    const counts = {};
    for (const h of history) counts[h.shift_date] = (counts[h.shift_date] || 0) + 1;
    return new Set(Object.entries(counts).filter(([_, n]) => n > 1).map(([d]) => d));
  })();

  // ── Auth gate ──
  if (!authReady) {
    return <div style={{ minHeight:"100vh", background:"#07080b" }} />;
  }
  if (!session || pwStep) {
    return (
      <LoginScreen
        step={session ? pwStep : null}
        email={accountEmail}
        notice={authNotice}
        onPasswordSaved={handlePasswordSaved}
        onSignOut={handleLogout}
      />
    );
  }

  // ── Header on every screen ──
  // The same bar sits at the top of every screen and stays put while the page
  // scrolls. The four main buttons keep their places on every screen; the
  // report's own buttons (Excel, Save, Absent) sit to their left while a day's
  // report is open. On a phone the buttons wrap to a second row.
  const onMain = screen === "upload" || screen === "results";
  const goUpload = () => {
    setMenuOpen(false);
    setScreen("upload"); setResults([]); setFileName(""); setBulkFiles([]);
  };
  const openConfig = () => {
    setMenuOpen(false);
    if (onMain) setShowConfig(c => !c);
    else { setScreen(results.length ? "results" : "upload"); setShowConfig(true); }
  };
  const dayOpen = results.length > 0 && (screen === "results" || screen === "absent");
  const savedThis = savedResults === results && results.length > 0;
  // Fonts, box sizing and the small screen button size, on every screen
  const pageHeader = (<>
    <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        input[type=date]::-webkit-calendar-picker-indicator { filter: invert(0.5); }
        tr:hover td { background: rgba(59,130,246,0.04) !important; }
        ::-webkit-scrollbar { width:5px; height:5px; }
        ::-webkit-scrollbar-track { background:#060d16; }
        ::-webkit-scrollbar-thumb { background:#1a2d45; border-radius:3px; }
        .gl-note::after { content:"ⓘ"; color:#6b7a8d; font-size:11px; margin-left:6px; }
        @media (max-width: 480px) {
          /* On small screens the stat cards go 4-up then 3-up naturally via flex-wrap */
          /* Ensure filter buttons don't overflow */
          button { font-size: 11px !important; padding: 5px 10px !important; }
        }
    `}</style>
    <div data-testid="page-header" style={{
      background:"#0a1320", borderBottom:"1px solid #1a2535",
      padding:"8px 16px", display:"flex", flexWrap:"wrap", alignItems:"center", rowGap:8, columnGap:10,
      justifyContent:"space-between", minHeight:52, boxSizing:"border-box", position:"sticky", top:0, zIndex:30,
      margin: onMain ? 0 : "-20px -20px 16px",
    }}>
      <div style={{ display:"flex", alignItems:"center", gap:10, flexShrink:0 }}>
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
      {/* wrap-reverse keeps the main buttons on the top line when a phone
          screen pushes the report buttons onto a line of their own */}
      <div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap-reverse", justifyContent:"flex-end", flex:"1 1 auto" }}>
        {dayOpen && (<div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap", justifyContent:"flex-end" }}>
          <button onClick={handleExport} style={{
            ...btn(true, "#1d4ed8"), background:"#1d4ed8", color:"#fff",
          }}>↓ Excel</button>
          <button onClick={() => handleSave()} disabled={saveBusy || savedThis}
            style={{ ...btn(false, "#334155"), cursor: saveBusy || savedThis ? "default" : "pointer", opacity: saveBusy || savedThis ? 0.75 : 1 }}
            title={savedThis ? "This report is saved" : "Save to history"}>
            {savedThis ? "✓ Saved" : saveBusy ? "Saving…" : "💾 Save"}
          </button>
          {roster.length > 0 && (
            <button
              onClick={() => setScreen("absent")}
              style={{
                ...btn(screen === "absent", "#334155"),
                color: absentEmployees.length > 0 ? "#f87171" : "#4b5e7a",
                borderColor: absentEmployees.length > 0 ? "#7f1d1d" : "#1a2d45",
              }}
              title="Show absent employees"
            >🚫 Absent {absentEmployees.length > 0 ? `(${absentEmployees.length})` : ""}</button>
          )}
        </div>)}
        <div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap", justifyContent:"flex-end" }}>
        <button onClick={() => { if (screen !== "weekly") loadWeekly(); }} style={btn(screen === "weekly", "#334155")}>📅 Weekly</button>
        <button onClick={openConfig} style={btn(onMain && showConfig, "#334155")}>⚙ Config</button>
        <button onClick={goUpload} style={btn(screen === "upload", "#334155")} title="Back to the file upload box">File Upload Box</button>
        <div style={{ position:"relative" }}>
          <button onClick={() => setMenuOpen(o => !o)} style={btn(false, "#334155")} aria-label="Menu">≡</button>
          {menuOpen && (
            <div style={{ position:"absolute", right:0, top:34, background:"#0a1320",
              border:"1px solid #1a2535", borderRadius:8, padding:6, minWidth:200, zIndex:40 }}>
              <div style={{ padding:"4px 6px 8px", fontSize:10, color:"#4b5e7a", lineHeight:1.5, wordBreak:"break-all" }}>
                Signed in as<br /><span style={{ color:"#c8d6e8", fontSize:11 }}>{accountEmail}</span>
              </div>
              <button onClick={loadHistory} style={{ ...btn(false, "#334155"), width:"100%", marginBottom:4 }}>History</button>
              <button onClick={loadWeekly} style={{ ...btn(false, "#334155"), width:"100%", marginBottom:4 }}>Weekly Rollup</button>
              <button onClick={handleLogout} style={{ ...btn(false, "#334155"), width:"100%", color:"#ef4444", borderColor:"#7f1d1d" }}>Logout</button>
            </div>
          )}
        </div>
        </div>
      </div>
    </div>
  </>);

  // ── History screen ──
  if (screen === "history") {
    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
        {pageHeader}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, gap:10, flexWrap:"wrap" }}>
          <div style={{ fontSize:16, fontWeight:700, color:"#e2e8f0" }}>HISTORY</div>
          <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
            {history.length > 0 && (() => {
              const armed = clearText.trim().toUpperCase() === "CLEAR" && !clearBusy;
              return (<>
                <input
                  value={clearText}
                  onChange={e => setClearText(e.target.value)}
                  placeholder="type CLEAR to enable"
                  aria-label="Type CLEAR to enable clearing every saved report"
                  style={{ ...inp, width:170, padding:"6px 10px", fontSize:12 }}
                />
                <button
                  onClick={clearAllReports}
                  disabled={!armed}
                  title="Removes every saved report for the whole team"
                  style={{ ...btn(armed, "#b91c1c"), color: armed ? "#fff" : "#4b5e7a",
                    cursor: armed ? "pointer" : "default", opacity: armed ? 1 : 0.7 }}
                >Clear all saved reports</button>
              </>);
            })()}
            <button onClick={() => setScreen(results.length ? "results" : "upload")} style={btn(false, "#334155")}>Back</button>
          </div>
        </div>
        {/* Days saved more than once: compare the copies and keep one */}
        {dupDates.size > 0 && (
          <div data-testid="dup-review" style={{ background:"#1f1408", border:"1px solid #92400e", borderRadius:10,
            padding:"10px 12px", marginBottom:14, color:"#fde68a", fontSize:12, lineHeight:1.6 }}>
            <div style={{ fontWeight:700, marginBottom:4 }}>
              {dupDates.size} day{dupDates.size === 1 ? " is" : "s are"} saved more than once. The Weekly Rollup uses the newest copy. Review and remove the older ones:
            </div>
            {[...dupDates].sort().map(d => {
              const copies = history.filter(h => h.shift_date === d)
                .sort((a, b) => String(b.uploaded_at || "").localeCompare(String(a.uploaded_at || "")));
              const armed = dupPending === d;
              return (
                <div key={d} style={{ borderTop:"1px solid #3b2a12", paddingTop:6, marginTop:6 }}>
                  <div style={{ fontWeight:700, color:"#fef3c7" }}>{dayLabel(d)} ({d})</div>
                  {copies.map((c, i) => (
                    <div key={c.id} style={{ fontFamily:"'DM Mono',monospace", fontSize:11 }}>
                      {i === 0 ? "Newest" : "Older"}: saved {savedWhen(c)} by {c.uploaded_by}, {copyLine(c)}
                      {c.file_name ? <span style={{ color:"#a8a29e" }}> · {c.file_name}</span> : null}
                    </div>
                  ))}
                  <div style={{ display:"flex", gap:6, marginTop:6, flexWrap:"wrap" }}>
                    {armed ? (<>
                      <button onClick={() => removeOlderCopies(d)} disabled={clearBusy} style={btn(true, "#b91c1c")}>
                        Confirm: remove {copies.length - 1} older cop{copies.length - 1 === 1 ? "y" : "ies"}
                      </button>
                      <button onClick={() => setDupPending(null)} style={btn(false, "#334155")}>Keep all</button>
                    </>) : (
                      <button onClick={() => setDupPending(d)} style={{ ...btn(false, "#334155"), color:"#fbbf24", borderColor:"#92400e" }}>
                        Keep newest, remove older
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        <div style={{ background:"#0a1320", border:"1px solid #1a2535", borderRadius:10, overflow:"hidden" }}>
          {history.length === 0 ? (
            <div style={{ padding:24, color:"#4b5e7a", fontSize:13 }}>No saved reports.</div>
          ) : (
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:680 }}>
                <thead>
                  <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2535" }}>
                    {["Shift Date","Uploaded By","Saved","Workers","Issues","Total Hrs",""].map(h => (
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
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{h.uploaded_at ? new Date(h.uploaded_at).toLocaleString() : ""}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{h.total_workers}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{(h.total_workers || 0) - (h.clean_count || 0)}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{decToHM(Number(h.total_bill_hrs) || 0)}</td>
                      <td style={{ padding:"9px 12px", whiteSpace:"nowrap" }}>
                        <button onClick={() => loadReport(h.id)} style={btn(false, "#334155")}>Load</button>
                        {delPending === h.id ? (<>
                          <button onClick={() => deleteReport(h.id)} disabled={clearBusy}
                            style={{ ...btn(true, "#b91c1c"), marginLeft:6 }}>Confirm delete</button>
                          <button onClick={() => setDelPending(null)}
                            style={{ ...btn(false, "#334155"), marginLeft:6 }}>Keep</button>
                        </>) : (
                          <button onClick={() => setDelPending(h.id)} title="Remove this saved report"
                            style={{ ...btn(false, "#334155"), marginLeft:6, color:"#ef4444", borderColor:"#7f1d1d" }}>Delete</button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
        {saveMsg && (
          <div style={{ marginTop:12, fontSize:12, fontFamily:"'DM Mono',monospace",
            color: /error/i.test(saveMsg) ? "#f87171" : "#34d399" }}>{saveMsg}</div>
        )}
        <div style={{ marginTop:10, fontSize:11, color:"#4b5e7a", lineHeight:1.5 }}>
          Delete removes a saved report for the whole team. A deleted day can be saved again by uploading its file.
          Clear all removes every saved report, including days not shown in this list.
        </div>
      </div>
    );
  }

  // ── Bulk upload screen (multiple daily CSVs → database) ──
  if (screen === "bulk") {
    const savedDates = bulkFiles.filter(f => (f.status || "").startsWith("Saved")).map(f => f.shiftDate);
    const savedCount = savedDates.length;
    const allSaved = bulkFiles.some(bulkDone) && bulkFiles.every(f => !f.rows.length || !f.shiftDate || bulkDone(f));
    const replacing = bulkFiles.filter(f => f.rows.length && f.shiftDate && !bulkDone(f) && (savedCopies[f.shiftDate] || []).length && f.dup === "replace").length;
    const dayCount = bulkFiles.filter(f => f.shiftDate).length;
    const fileCount = bulkFiles.reduce((s, f) => s + (f.fileCount || 1), 0);
    const plural = (n, word) => `${n} ${word}${n === 1 ? "" : "s"}`;
    // A changed date reruns the day so the comparison with a saved copy stays right
    const withPreview = x => x.rows.length && /^\d{4}-\d{2}-\d{2}$/.test(x.shiftDate) && /^\d{4}-\d{2}-\d{2}$/.test(x.reportDate)
      ? { ...x, preview: dayPreview(analyze(x.rows, x.shiftDate, x.reportDate, currentGates(), rules)) }
      : { ...x, preview: null };
    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
        {pageHeader}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700, color:"#e2e8f0" }}>SAVE MULTIPLE DAYS</div>
            <div data-testid="bulk-summary" style={{ fontSize:11, color:"#4b5e7a", marginTop:2, wordBreak:"break-word" }}>
              {plural(dayCount, "day")} from {plural(fileCount, "file")}{bulkFrom ? ` in ${bulkFrom}` : ""} · shift dates come from the swipe data, check them before saving
            </div>
          </div>
          <button onClick={() => { setScreen("upload"); setBulkFiles([]); }} style={btn(false, "#334155")}>Back</button>
        </div>

        <div style={{
          display:"flex", alignItems:"center", gap:8, marginBottom:12,
          background:"#0d1520", border:"1px solid #1a2d45",
          borderRadius:8, padding:"8px 12px", maxWidth:420,
        }}>
          <span style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase", letterSpacing:1, flexShrink:0 }}>Uploader</span>
          <input
            type="text" value={uploadedBy} onChange={e => setUploadedBy(e.target.value)}
            placeholder="Your name (required to save)"
            style={{ background:"transparent", border:"none", color:"#e2e8f0",
              fontFamily:"'DM Mono',monospace", fontSize:13, outline:"none", width:"100%" }}
          />
        </div>

        <div style={{ background:"#0a1320", border:"1px solid #1a2535", borderRadius:10, overflow:"hidden", marginBottom:14 }}>
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:720 }}>
              <thead>
                <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2535" }}>
                  {["File","Shift Date","Report Date","Rows","Status"].map(h => (
                    <th key={h} style={{ padding:"9px 12px", textAlign:"left", fontSize:10,
                      color:"#4b5e7a", textTransform:"uppercase", letterSpacing:0.8, fontWeight:600 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {bulkFiles.map((f, i) => (
                  <tr key={i} style={{ borderBottom:"1px solid #1a2535" }}>
                    <td style={{ padding:"9px 12px", fontSize:12, fontFamily:"'DM Mono',monospace", wordBreak:"break-all" }}>{f.name}</td>
                    <td style={{ padding:"9px 12px" }}>
                      <input type="date" value={f.shiftDate}
                        onChange={e => setBulkFiles(fs => fs.map((x, j) => j === i
                          ? withPreview({ ...x, shiftDate: e.target.value,
                              reportDate: e.target.value ? toYMD(addDays(new Date(e.target.value + "T00:00:00"), 1)) : x.reportDate })
                          : x))}
                        style={{ ...inp, width:150 }} />
                      {knownDates.has(f.shiftDate) && !bulkDone(f) && (
                        <div style={{ color:"#fb923c", fontSize:10, marginTop:3 }}>⚠ saved before, see Status</div>
                      )}
                    </td>
                    <td style={{ padding:"9px 12px" }}>
                      <input type="date" value={f.reportDate}
                        onChange={e => setBulkFiles(fs => fs.map((x, j) => j === i ? withPreview({ ...x, reportDate: e.target.value }) : x))}
                        style={{ ...inp, width:150 }} />
                    </td>
                    <td style={{ padding:"9px 12px", fontSize:12, fontFamily:"'DM Mono',monospace" }}>{f.rows.length}</td>
                    <td style={{ padding:"9px 12px", fontSize:12,
                      color: bulkDone(f) ? "#34d399" : f.status ? "#f87171" : "#4b5e7a" }}>
                      {f.status || "Ready"}
                      {f.note && <div style={{ color:"#6b7a8d", fontSize:10, marginTop:3 }}>{f.note}</div>}
                      {/* A day saved before: compare and choose before anything is written */}
                      {!bulkDone(f) && f.preview && (savedCopies[f.shiftDate] || []).length > 0 && (
                        <div data-testid="resave-review" style={{ marginTop:6, background:"#1f1408", border:"1px solid #92400e",
                          borderRadius:6, padding:"6px 8px", color:"#fde68a", fontSize:11, lineHeight:1.5 }}>
                          {(savedCopies[f.shiftDate] || []).map(c => (
                            <div key={c.id}>Saved {savedWhen(c)} by {c.uploaded_by}: {copyLine(c)}</div>
                          ))}
                          <div>This upload: {previewLine(f.preview)}</div>
                          <select value={f.dup || "replace"} aria-label={`What to do with the saved copy of ${f.shiftDate}`}
                            onChange={e => setBulkFiles(fs => fs.map((x, j) => j === i ? { ...x, dup: e.target.value } : x))}
                            style={{ ...inp, marginTop:4, padding:"4px 8px", fontSize:11, width:"auto" }}>
                            <option value="replace">Replace the saved copy</option>
                            <option value="keep">Keep both</option>
                            <option value="skip">Don't save this day</option>
                          </select>
                        </div>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
          <button
            onClick={runBulk}
            disabled={bulkBusy || allSaved}
            style={{
              background: bulkBusy || allSaved ? "#1a2d45" : "#1d4ed8", color:"#fff", border:"none",
              borderRadius:6, padding:"10px 18px", fontSize:13, fontWeight:600,
              cursor: bulkBusy || allSaved ? "default" : "pointer", fontFamily:"'DM Mono',monospace",
            }}
          >{bulkBusy ? "Saving…" : allSaved ? "✓ All Saved" : `Process & Save All${replacing ? ` (replaces ${replacing} saved day${replacing === 1 ? "" : "s"})` : ""}`}</button>
          {savedCount > 0 && (
            <button onClick={() => openWeekAfterSave(savedDates)} disabled={bulkBusy}
              style={{ ...btn(true, "#0e7490"), background:"#0e7490", color:"#fff" }}>
              📅 Go to Weekly Rollup →
            </button>
          )}
        </div>
      </div>
    );
  }

  // ── Weekly rollup screen ──
  if (screen === "weekly") {
    const selReports = weeklyReports.filter(r => selIds.has(r.id));
    const dateCounts = {};
    selReports.forEach(r => { dateCounts[r.shift_date] = (dateCounts[r.shift_date] || 0) + 1; });
    const dupSelected = Object.values(dateCounts).some(n => n > 1);
    // Selected days that fall in more than one Monday to Sunday week
    const selWeeks = {};
    selReports.forEach(r => { (selWeeks[weekStart(r.shift_date)] ||= new Set()).add(r.shift_date); });
    const weekKeys = Object.keys(selWeeks).sort();
    const fmtD = d => `${weekdayOf(d)} ${d.slice(5,7)}/${d.slice(8,10)}`;
    const showRosterCols = roster.length > 0;
    const sharedWeek = rollup ? sharedIds(rollup.emps) : new Set();
    const rollEmps = rollup ? rollup.emps.map(e => ({ ...e, ...rosterFields(e.ein, e.name, sharedWeek) })) : [];
    const shownEmps = rollup
      ? rollEmps.filter(e => !weeklySearch ||
          e.name.toLowerCase().includes(weeklySearch.toLowerCase()) ||
          (e.ein || "").includes(weeklySearch))
      : [];
    const grandTotal = rollup ? rollup.emps.reduce((s, e) => s + e.total, 0) : 0;
    const issueCells = rollup ? rollup.emps.reduce((s, e) => s + e.issueDays, 0) : 0;
    const weeklyAreaRows = rollup
      ? summarizeByArea(
          rollup.emps.flatMap(e => Object.values(e.days).map(d => ({ status: d.status, flagArea: d.flagArea, ein: e.ein, name: e.name }))),
          areaOfWith(sharedWeek))
      : [];

    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
        {pageHeader}
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700, color:"#e2e8f0" }}>WEEKLY ROLLUP</div>
            <div style={{ fontSize:11, color:"#4b5e7a", marginTop:2 }}>
              Pick the saved days to include, then build the rollup
            </div>
          </div>
          <button onClick={() => setScreen(results.length ? "results" : "upload")} style={btn(false, "#334155")}>Back</button>
        </div>

        {weeklyMsg && (
          <div style={{ background:"#3f1d1d", border:"1px solid #7f1d1d", color:"#e2e8f0",
            padding:"8px 12px", borderRadius:6, fontSize:12, marginBottom:10 }}>{weeklyMsg}</div>
        )}
        {weeklyNote && (
          <div data-testid="weekly-note" style={{ background:"#0d1520", border:"1px solid #1a2d45", color:"#c8d6e8",
            padding:"8px 12px", borderRadius:6, fontSize:12, marginBottom:10, lineHeight:1.5 }}>{weeklyNote}</div>
        )}

        {/* Day picker */}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:10, alignItems:"center" }}>
          <button onClick={() => selectWeek(0)}  style={btn(false, "#1d4ed8")}>This Week</button>
          <button onClick={() => selectWeek(-1)} style={btn(false, "#1d4ed8")}>Last Week</button>
          <button onClick={selectAll} style={btn(false, "#1d4ed8")}
            title="Every saved day in the list (the newest copy of a day saved twice)">Select All</button>
          {selIds.size > 0 && (
            <button onClick={() => { setSelIds(new Set()); setRollup(null); setWeeklyNote(""); }} style={btn(false, "#334155")}>Clear</button>
          )}
          <span style={{ fontSize:11, color:"#4b5e7a", fontFamily:"'DM Mono',monospace" }}>
            {selIds.size} day(s) selected
          </span>
        </div>

        {dupSelected && (
          <div style={{ background:"#3b2500", border:"1px solid #9a3412", color:"#fb923c",
            padding:"8px 12px", borderRadius:6, fontSize:12, marginBottom:10 }}>
            ⚠ Two reports with the same shift date are selected. The later saved one counts for that day.
          </div>
        )}
        {weekKeys.length > 1 && (
          <div data-testid="week-span" style={{ background:"#3b2500", border:"1px solid #9a3412", color:"#fb923c",
            padding:"8px 12px", borderRadius:6, fontSize:12, marginBottom:10, lineHeight:1.5 }}>
            ⚠ The selected days are in {weekKeys.length} different weeks (Monday to Sunday):{" "}
            {weekKeys.map(w => `week of ${dayLabel(w)}: ${[...selWeeks[w]].sort().map(dayLabel).join(", ")}`).join("; ")}.
            {" "}A night shift is dated the day it starts, so Sunday night belongs to the week that ends that Sunday.
          </div>
        )}

        <div style={{ background:"#0a1320", border:"1px solid #1a2535", borderRadius:10, overflow:"hidden", marginBottom:14 }}>
          {weeklyReports.length === 0 ? (
            <div style={{ padding:24, color:"#4b5e7a", fontSize:13 }}>
              No saved daily reports yet. Upload CSVs (drop several at once) and save them first.
            </div>
          ) : (
            <div style={{ overflowX:"auto", maxHeight:300, overflowY:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:680 }}>
                <thead>
                  <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2535" }}>
                    {["","Shift Date","File","Uploaded By","Workers","Total Hrs"].map((h, i) => (
                      <th key={i} style={{ padding:"9px 12px", textAlign:"left", fontSize:10,
                        color:"#4b5e7a", textTransform:"uppercase", letterSpacing:0.8, fontWeight:600,
                        position:"sticky", top:0, background:"#060d16" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {weeklyReports.map(r => (
                    <tr key={r.id}
                      onClick={() => toggleSel(r.id)}
                      style={{ borderBottom:"1px solid #1a2535", cursor:"pointer",
                        background: selIds.has(r.id) ? "#0d1e30" : "transparent" }}>
                      <td style={{ padding:"9px 12px" }}>
                        <input type="checkbox" readOnly checked={selIds.has(r.id)} style={{ cursor:"pointer" }} />
                      </td>
                      <td style={{ padding:"9px 12px", fontSize:12, fontFamily:"'DM Mono',monospace", whiteSpace:"nowrap" }}>
                        <span style={{ color: weekdayOf(r.shift_date) === "Sun" ? "#fbbf24" : "#6b7a8d", marginRight:6 }}>{weekdayOf(r.shift_date)}</span>
                        {r.shift_date}
                      </td>
                      <td style={{ padding:"9px 12px", fontSize:11, fontFamily:"'DM Mono',monospace",
                        color:"#6b7a8d", wordBreak:"break-all" }}>{r.file_name || ""}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{r.uploaded_by}</td>
                      <td style={{ padding:"9px 12px", fontSize:12 }}>{r.total_workers}</td>
                      <td style={{ padding:"9px 12px", fontSize:12, fontFamily:"'DM Mono',monospace" }}>
                        {decToHM(Number(r.total_bill_hrs) || 0)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:16 }}>
          <button
            onClick={buildRollup}
            disabled={weeklyBusy || !selIds.size}
            style={{
              background: weeklyBusy || !selIds.size ? "#1a2d45" : "#1d4ed8", color:"#fff", border:"none",
              borderRadius:6, padding:"10px 18px", fontSize:13, fontWeight:600,
              cursor: weeklyBusy || !selIds.size ? "default" : "pointer", fontFamily:"'DM Mono',monospace",
            }}
          >{weeklyBusy ? "Building…" : "Build Rollup"}</button>
          {rollup && (
            <button
              onClick={() => exportWeeklyXLSX({ ...rollup, emps: rollEmps }, rollupMeta, rules, { rows: weeklyAreaRows, basis: areaBasis }, rosterInfo)}
              style={{ ...btn(true, "#1d4ed8"), background:"#0e7490", borderColor:"#0e7490", color:"#fff" }}
            >↓ Excel</button>
          )}
        </div>

        {rollup && (<>
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:14 }}>
            <StatCard label="Employees"   value={rollup.emps.length}    color="#3b82f6" />
            <StatCard label="Days"        value={rollup.dates.length}   color="#0ea5e9" />
            <StatCard label="Issue Days"  value={issueCells}            color="#ef4444" />
            <StatCard label="Total Hrs"   value={decToHM(grandTotal)}   color="#22c55e" />
          </div>

          <AreaTable
            title="ISSUES BY AREA, SELECTED DAYS"
            rows={weeklyAreaRows}
            basis={areaBasis}
            note="Each person counts once per day flagged."
          />

          <div style={{ marginBottom:10 }}>
            <input
              placeholder="Search name or EIN…"
              value={weeklySearch}
              onChange={e => setWeeklySearch(e.target.value)}
              style={{...inp, maxWidth:320}}
            />
          </div>

          <div style={{ background:"#0a1320", border:"1px solid #1a2535", borderRadius:10, overflow:"hidden" }}>
            <div style={{ overflowX:"auto", WebkitOverflowScrolling:"touch" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:680 }}>
                <thead>
                  <tr style={{ background:"#060d16", borderBottom:"2px solid #1a2535" }}>
                    {["Employee Name","Company","EIN", ...(showRosterCols ? ["Supervisor","Area", ...(hasScope ? ["Scope"] : [])] : []),
                      ...rollup.dates.map(fmtD), "Days","Total"].map((h, i) => (
                      <th key={i} style={{ padding:"9px 12px", textAlign:"left", fontSize:10,
                        color:"#4b5e7a", textTransform:"uppercase", letterSpacing:0.8,
                        fontWeight:600, whiteSpace:"nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {shownEmps.map((e, i) => (
                    <tr key={i} style={{ borderBottom:"1px solid #1a2535",
                      background: i % 2 === 0 ? "#0d1520" : "#0a1118" }}>
                      <td style={{ ...td, whiteSpace:"nowrap" }}>{e.name}</td>
                      <td style={{ ...td, color:"#6b7a8d" }}>{e.company}</td>
                      <td style={{ ...td, color:"#6b7a8d", fontFamily:"'DM Mono',monospace" }}>{e.ein}</td>
                      {showRosterCols && (<>
                        <td style={{ ...td, fontSize:12, whiteSpace:"nowrap" }}>{e.notInRoster || e.supervisor === "N/A" ? "" : e.supervisor}</td>
                        <td style={{ ...td, fontSize:12, color: e.notInRoster ? "#f59e0b" : "#c8d6e8", whiteSpace:"nowrap" }}>
                          {e.notInRoster ? "Not in roster" : e.area === "N/A" ? "" : e.area}</td>
                        {hasScope && <td style={{ ...td, fontSize:12 }}>{e.notInRoster || e.scope === "N/A" ? "" : e.scope}</td>}
                      </>)}
                      {rollup.dates.map(d => {
                        const day = e.days[d];
                        const review = day && day.billHrs == null && (day.status === "Late IN" || day.status === "Late OUT");
                        return (
                          <td key={d} title={day && day.note ? day.note : undefined} style={{ ...td, fontFamily:"'DM Mono',monospace", fontSize:12,
                            whiteSpace:"nowrap",
                            color: !day ? "#374151" : review ? "#facc15" : day.billHrs == null && day.status !== "Clean" ? "#ef4444"
                              : day.over13 ? "#fb923c" : day.status === "Split Shift" ? "#facc15" : "#c8d6e8" }}>
                            {dayCell(day)}
                          </td>
                        );
                      })}
                      <td style={{ ...td, fontFamily:"'DM Mono',monospace" }}>{e.daysWorked}</td>
                      <td style={{ ...td, fontFamily:"'DM Mono',monospace", fontWeight:700, color:"#e2e8f0" }}>
                        {decToHM(e.total)}
                      </td>
                    </tr>
                  ))}
                  {!shownEmps.length && (
                    <tr><td colSpan={5 + rollup.dates.length + (showRosterCols ? (hasScope ? 3 : 2) : 0)} style={{ padding:36, textAlign:"center", color:"#4b5e7a" }}>
                      No records match.
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>)}
      </div>
    );
  }

  // ── Absent employees screen ──
  if (screen === "absent") {
    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
        {pageHeader}
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

      {pageHeader}

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
              Gate List (checked = off site, excluded from calculations)
            </label>
            <div style={{ fontSize:10, color:"#4b5e7a", marginTop:2, marginBottom:3,
              fontFamily:"'DM Mono',monospace" }}>
              Gate names collect here automatically from uploaded CSVs. Unchecked gates count as qualifying.
            </div>
            {(() => {
              const nqSet = currentGates();
              const allGates = [...new Set([...seenGates, ...nqSet])].sort();
              return (
                <div style={{ background:"#0d1520", border:"1px solid #1a2d45", borderRadius:6,
                  marginTop:3, maxHeight:230, overflowY:"auto", padding:"4px 10px" }}>
                  {!allGates.length && (
                    <div style={{ fontSize:11, color:"#4b5e7a", padding:"8px 0" }}>
                      No gates yet. Upload a CSV or add a gate below.
                    </div>
                  )}
                  {allGates.map(g => {
                    const excluded = nqSet.has(g);
                    return (
                      <label key={g} style={{ display:"flex", alignItems:"center", gap:8,
                        padding:"4px 0", cursor:"pointer", fontSize:12,
                        fontFamily:"'DM Mono',monospace",
                        color: excluded ? "#fb923c" : "#c8d6e8" }}>
                        <input
                          type="checkbox"
                          checked={excluded}
                          onChange={() => {
                            // Keep the gate visible in the list even after unchecking
                            mergeSeenGates([g]);
                            const next = new Set(nqSet);
                            if (excluded) next.delete(g); else next.add(g);
                            setNqText([...next].sort().join("\n"));
                          }}
                          style={{ cursor:"pointer", flexShrink:0 }}
                        />
                        <span style={{ wordBreak:"break-all" }}>{g}</span>
                        <span style={{ marginLeft:"auto", flexShrink:0, fontSize:9, letterSpacing:0.5,
                          color: excluded ? "#fb923c" : "#34d399" }}>
                          {excluded ? "OFF SITE" : "QUALIFYING"}
                        </span>
                      </label>
                    );
                  })}
                </div>
              );
            })()}
            <div style={{ display:"flex", gap:8, marginTop:8, maxWidth:420 }}>
              <input
                value={newGate}
                onChange={e => setNewGate(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") addGate(); }}
                placeholder="Add a gate name (exact match)"
                style={inp}
              />
              <button onClick={addGate} style={{ ...btn(true, "#334155"), background:"#334155",
                color:"#fff", flexShrink:0 }}>Add</button>
            </div>
            <div style={{ fontSize:10, color:"#4b5e7a", marginTop:6, fontFamily:"'DM Mono',monospace" }}>
              {currentGates().size} gate(s) excluded. Checked gates apply to the next upload on this device. Save Config for Team to apply for everyone.
              {!currentGates().has("S-Curve Turnstile 2") &&
                " Note: S-Curve Turnstile 2 counts as qualifying by design. Check it if that should change."}
            </div>
          </div>

          {/* ── Rounding rules ── */}
          <div style={{ marginTop:14, paddingTop:14, borderTop:"1px solid #1a2535" }}>
            <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase", letterSpacing:1 }}>
              Rounding Rules
            </label>
            <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginTop:6, alignItems:"flex-end" }}>
              <div style={{ flex:"1 1 150px", maxWidth:200 }}>
                <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:3 }}>Round to nearest</div>
                <select
                  value={rules.roundTo}
                  onChange={e => {
                    const roundTo = Number(e.target.value);
                    // Sensible midpoint default when the interval changes: exact half
                    setRules(r => ({ ...r, roundTo,
                      midpointSec: roundTo === 30 ? 930 : Math.floor(roundTo * 60 / 2) }));
                  }}
                  style={{ ...inp, cursor:"pointer" }}
                >
                  <option value={15}>15 minutes</option>
                  <option value={30}>30 minutes</option>
                  <option value={60}>60 minutes</option>
                  <option value={0}>No rounding</option>
                </select>
              </div>
              {rules.roundTo > 0 && (
                <div style={{ flex:"1 1 150px", maxWidth:200 }}>
                  <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:3 }}>
                    Round up at (min:sec past mark)
                  </div>
                  <input
                    type="text"
                    value={midpointToStr(rules.midpointSec)}
                    onChange={e => {
                      const m = e.target.value.match(/^(\d{1,2}):?(\d{0,2})$/);
                      if (m) {
                        const sec = Number(m[1]) * 60 + Number(m[2] || 0);
                        setRules(r => ({ ...r, midpointSec: sec }));
                      }
                    }}
                    style={inp}
                  />
                </div>
              )}
              <div style={{ flex:"1 1 150px", maxWidth:200 }}>
                <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:3 }}>Lunch deduction (min)</div>
                <input
                  type="number" min="0" step="5"
                  value={rules.lunchMin}
                  onChange={e => setRules(r => ({ ...r, lunchMin: Math.max(0, Number(e.target.value) || 0) }))}
                  style={inp}
                />
              </div>
              <div style={{ flex:"1 1 150px", maxWidth:200 }}>
                <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:3 }}>Lunch only over (hrs, 0 = every day)</div>
                <input
                  type="number" min="0" step="0.5"
                  value={rules.lunchOverHrs}
                  onChange={e => setRules(r => ({ ...r, lunchOverHrs: Math.max(0, Number(e.target.value) || 0) }))}
                  style={inp}
                />
              </div>
              <div style={{ flex:"1 1 150px", maxWidth:200 }}>
                <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:3 }}>Earliest counted IN (blank = off)</div>
                <input
                  type="time"
                  value={rules.earliestInMin > 0
                    ? `${String(Math.floor(rules.earliestInMin / 60)).padStart(2, "0")}:${String(rules.earliestInMin % 60).padStart(2, "0")}`
                    : ""}
                  onChange={e => {
                    const m = e.target.value.match(/^(\d{2}):(\d{2})$/);
                    setRules(r => ({ ...r, earliestInMin: m ? Number(m[1]) * 60 + Number(m[2]) : 0 }));
                  }}
                  style={inp}
                />
              </div>
              <div style={{ flex:"1 1 150px", maxWidth:200 }}>
                <div style={{ fontSize:10, color:"#4b5e7a", marginBottom:3 }}>IN window after first badge (hrs, 0 = off)</div>
                <input
                  type="number" min="0" step="0.5"
                  value={rules.inWindowHrs}
                  onChange={e => setRules(r => ({ ...r, inWindowHrs: Math.max(0, Number(e.target.value) || 0) }))}
                  style={inp}
                />
              </div>
            </div>
            <div style={{ fontSize:10, color:"#4b5e7a", marginTop:6, fontFamily:"'DM Mono',monospace" }}>
              Current rule: {rulesSummary(rules)} · lunch {rules.lunchMin} min.{extraRulesSummary(rules) && ` ${extraRulesSummary(rules)}.`} Applies to new uploads. Saved reports keep the numbers they were saved with.
            </div>
          </div>

          {/* ── Save config for the team ── */}
          <div style={{ marginTop:12, display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}>
            <button
              onClick={handleSaveConfig}
              style={{
                background:"#1d4ed8", color:"#fff", border:"none",
                borderRadius:6, padding:"8px 14px", fontSize:12, fontWeight:600,
                cursor:"pointer", fontFamily:"'DM Mono',monospace",
              }}
            >💾 Save Config for Team</button>
            {cfgMeta && !cfgMsg && (
              <span style={{ fontSize:11, color:"#4b5e7a", fontFamily:"'DM Mono',monospace" }}>
                last saved by {cfgMeta.updated_by} · {new Date(cfgMeta.updated_at).toLocaleString()}
              </span>
            )}
            {cfgMsg && (
              <span style={{ fontSize:12, fontFamily:"'DM Mono',monospace",
                color: cfgMsg.includes("error") ? "#f87171" : "#34d399" }}>{cfgMsg}</span>
            )}
          </div>

          {/* ── Employee Roster Upload ── */}
          <div style={{ marginTop:14, paddingTop:14, borderTop:"1px solid #1a2535" }}>
            <label style={{ fontSize:10, color:"#4b5e7a", textTransform:"uppercase", letterSpacing:1 }}>
              MMR Employee Roster (.xlsx or .csv)
            </label>
            <div style={{ fontSize:10, color:"#4b5e7a", marginTop:2, marginBottom:6,
              fontFamily:"'DM Mono',monospace", lineHeight:1.5 }}>
              Columns: EIN (required), Area, Supervisor (can be blank), First Name, Last Name.
              A new upload replaces the previous roster. Upload the roster and it switches to the roster's work areas.
              Clear Roster removes the roster for the whole team.
            </div>
            <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
              <input
                ref={rosterFileRef}
                data-testid="roster-file-config"
                type="file"
                accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                onChange={e => { handleRosterUpload(e.target.files[0]); e.target.value = ""; }}
                style={{ display:"none" }}
              />
              <button onClick={() => exportRosterTemplate()} style={btn(false, "#334155")}>↓ Blank Template</button>
              {roster.length > 0 && (
                <button onClick={() => exportRosterTemplate(roster)} style={btn(false, "#334155")}>↓ Current Roster</button>
              )}
              <button
                onClick={() => rosterFileRef.current?.click()}
                style={{
                  background:"#1d4ed8", color:"#fff", border:"none",
                  borderRadius:6, padding:"8px 14px", fontSize:12, fontWeight:600,
                  cursor:"pointer", fontFamily:"'DM Mono',monospace",
                }}
              >📤 Upload Roster</button>
              {roster.length > 0 && (rosterClearPending ? (<>
                <button onClick={clearRoster} style={btn(true, "#b91c1c")}>Confirm clear roster</button>
                <button onClick={() => setRosterClearPending(false)} style={btn(false, "#334155")}>Keep</button>
              </>) : (
                <button onClick={() => setRosterClearPending(true)} title="Remove the roster for the whole team"
                  style={{ ...btn(false, "#334155"), color:"#ef4444", borderColor:"#7f1d1d" }}>Clear Roster</button>
              ))}
              {rosterMeta && (
                <div style={{ fontSize:11, color:"#4b5e7a", fontFamily:"'DM Mono',monospace" }}>
                  {rosterMeta.count} employees · last by {rosterMeta.uploaded_by} · {new Date(rosterMeta.uploaded_at).toLocaleString()}
                </div>
              )}
            </div>
            {rosterMsg && (
              <div style={{
                marginTop:8, fontSize:12,
                color: rosterMsgBad ? "#f87171" : "#34d399",
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
                Upload the Genetec export (.xlsx or .csv) to begin
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
                {processing ? "Processing…" : fileName || "Tap to browse or drop Genetec files"}
              </div>
              <div style={{ fontSize:11, color:"#4b5e7a", marginTop:4, lineHeight:1.5 }}>
                Genetec raw export (.xlsx or .csv), or a .zip of a week's files. Drop several days
                at once to save them all; two files for the same day are combined automatically.
              </div>
              <input ref={fileRef} data-testid="gate-files" type="file"
                accept=".xlsx,.xls,.xlsm,.csv,.zip,application/zip,application/x-zip-compressed" multiple style={{display:"none"}}
                onChange={e => { handleFiles(e.target.files); e.target.value = ""; }} />
            </div>

            {error && (
              <div style={{ color:"#ef4444", fontSize:13, background:"#1e0d0d",
                border:"1px solid #991b1b", borderRadius:8, padding:"10px 14px",
                width:"100%", maxWidth:460 }}>{error}</div>
            )}

            {/* Employee roster: blank template out, filled roster in */}
            <div data-testid="roster-card" style={{ width:"100%", maxWidth:460, background:"#0a1320",
              border:"1px solid #1a2535", borderRadius:10, padding:"12px 14px" }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"baseline", gap:8, flexWrap:"wrap" }}>
                <div style={{ fontSize:12, fontWeight:700, color:"#e2e8f0", letterSpacing:0.4 }}>Employee roster</div>
                <div style={{ fontSize:11, fontFamily:"'DM Mono',monospace", color: roster.length ? "#34d399" : "#6b7a8d" }}>
                  {roster.length ? `${roster.length} employees loaded` : "No roster loaded"}
                </div>
              </div>
              <div style={{ fontSize:11, color:"#6b7a8d", marginTop:6, lineHeight:1.5 }}>
                Upload the roster and it switches to the roster's work areas. Download the blank template,
                fill in each employee number with their area and/or supervisor (supervisor can be blank), then upload it.
              </div>
              <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginTop:10 }}>
                <button onClick={() => exportRosterTemplate()} style={btn(false, "#334155")}>↓ Blank Template</button>
                {roster.length > 0 && (
                  <button onClick={() => exportRosterTemplate(roster)} style={btn(false, "#334155")}>↓ Current Roster</button>
                )}
                <button onClick={() => rosterFileRef2.current?.click()}
                  style={{ ...btn(true, "#1d4ed8"), background:"#1d4ed8", color:"#fff" }}>↑ Upload Roster</button>
                <input ref={rosterFileRef2} data-testid="roster-file" type="file" accept=".xlsx,.xls,.csv" style={{ display:"none" }}
                  onChange={e => { handleRosterUpload(e.target.files[0]); e.target.value = ""; }} />
                {roster.length > 0 && (rosterClearPending ? (<>
                  <button onClick={clearRoster} style={btn(true, "#b91c1c")}>Confirm clear roster</button>
                  <button onClick={() => setRosterClearPending(false)} style={btn(false, "#334155")}>Keep</button>
                </>) : (
                  <button onClick={() => setRosterClearPending(true)} title="Remove the roster for the whole team"
                    style={{ ...btn(false, "#334155"), color:"#ef4444", borderColor:"#7f1d1d" }}>Clear Roster</button>
                ))}
              </div>
              {rosterMsg && (
                <div style={{ marginTop:8, fontSize:12, fontFamily:"'DM Mono',monospace",
                  color: rosterMsgBad ? "#f87171" : "#34d399" }}>
                  {rosterMsg}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ── RESULTS SCREEN ── */}
        {screen === "results" && (<>

          {saveMsg && (() => {
            const bad = /error|not removed|^Nothing|^Enter|^Database/i.test(saveMsg);
            return (
              <div data-testid="save-msg" style={{
                background: bad ? "#3f1d1d" : "#14532d",
                border: "1px solid " + (bad ? "#7f1d1d" : "#166534"),
                color:"#e2e8f0", padding:"8px 12px", borderRadius:6, fontSize:12, marginBottom:10,
              }}>{saveMsg}</div>
            );
          })()}

          {/* This day is saved already: compare, then replace the saved copy or keep both */}
          {resaveAsk && resaveAsk.date === shiftDate && savedResults !== results && (
            <div data-testid="resave-ask" style={{ background:"#1f1408", border:"1px solid #92400e", color:"#fde68a",
              padding:"10px 12px", borderRadius:6, fontSize:12, marginBottom:10, lineHeight:1.6 }}>
              <div style={{ fontWeight:700 }}>{dayLabel(shiftDate)} is already saved.</div>
              {resaveAsk.existing.map(c => (
                <div key={c.id}>Saved {savedWhen(c)} by {c.uploaded_by}: {copyLine(c)}</div>
              ))}
              <div>This upload: {previewLine(dayPreview(results))}</div>
              <div style={{ display:"flex", gap:8, marginTop:8, flexWrap:"wrap" }}>
                <button onClick={() => handleSave("replace")} disabled={saveBusy}
                  style={{ ...btn(true, "#1d4ed8") }}>Replace the saved copy</button>
                <button onClick={() => handleSave("keep")} disabled={saveBusy} style={btn(false, "#334155")}>Keep both</button>
                <button onClick={() => handleSave("cancel")} disabled={saveBusy} style={btn(false, "#334155")}>Cancel</button>
              </div>
            </div>
          )}

          {/* STAT CARDS: 4 per row on mobile, flex wrap */}
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:14 }}>
            <StatCard label="Workers"      value={results.length}  color="#3b82f6" />
            <StatCard label="Clean"        value={clean}           color="#22c55e" />
            <StatCard label="Missing OUT"  value={mOut}            color="#ef4444" />
            <StatCard label="Missing IN"   value={mIn}             color="#ef4444" />
            <StatCard label="Missing Both" value={mBoth}           color="#ef4444" />
            <StatCard label="Late IN"      value={counts["Late IN"]} color="#facc15" />
            {counts["Late OUT"] > 0 && <StatCard label="Late OUT" value={counts["Late OUT"]} color="#facc15" />}
            {counts["Split Shift"] > 0 && <StatCard label="Split Shift" value={counts["Split Shift"]} color="#facc15" />}
            <StatCard label="Over 13 Hrs"  value={over13}          color="#fb923c" />
            <StatCard label="Total Hrs"    value={decToHM(totalH)} color="#0ea5e9" />
          </div>

          <AreaTable title="ISSUES BY AREA" rows={dailyAreaRows} basis={areaBasis}
            note={reviewToday > 0
              ? `Late IN, Late OUT and Split Shift (${reviewToday} today) are flagged for review, not counted here as missing badges.`
              : ""} />

          {/* FILTERS: wrap naturally; search on its own row on mobile */}
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:10, alignItems:"center" }}>
            {["All","Issues","Missing OUT","Missing IN","Late IN","Over 13",
              ...(counts["Late OUT"] + counts["Split Shift"] > 0 ? ["Review"] : [])].map(f => (
              <button key={f} onClick={() => setFilter(f)} style={btn(filter===f,"#1d4ed8")}
                title={f === "Review" ? "Late OUT and Split Shift" : undefined}>{f}</button>
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

          {/* Supervisor / Area / Scope filters: only show when roster loaded */}
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
              {hasScope && <select
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
              </select>}
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
            Showing {filtered.length} of {results.length} workers, tap a row to expand
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
                      ...(roster.length > 0 ? ["Supervisor","Area", ...(hasScope ? ["Scope"] : [])] : []),
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
                  {filtered.map((r, i) => <Row key={r.guid} r={r} idx={i} showRoster={roster.length > 0} showScope={hasScope} />)}
                  {!filtered.length && (
                    <tr><td colSpan={7 + (roster.length > 0 ? (hasScope ? 3 : 2) : 0)} style={{ padding:36, textAlign:"center", color:"#4b5e7a" }}>
                      No records match.
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>)}
      </div>

    </div>
  );
}
