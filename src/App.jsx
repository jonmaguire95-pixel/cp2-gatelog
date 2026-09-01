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

// ── ROUNDING RULES (configurable via Config panel, stored in database) ────────
// Defaults preserve original behavior: nearest 30 min via 15:30 (930 sec) midpoint,
// 30 min lunch deduction.
// earliestInMin: INs before this time of day (minutes from midnight) count as
//   exactly this time. 0 turns the rule off.
// lunchOverHrs: lunch comes off only when the rounded span is strictly longer
//   than this many hours. 0 means lunch comes off every day.
// inWindowHrs: a qualifying IN only counts within this many hours of the
//   person's first badge anywhere on site that day. 0 turns the rule off.
const RULES_DEFAULT = { roundTo: 30, midpointSec: 930, lunchMin: 30, earliestInMin: 0, lunchOverHrs: 0, inWindowHrs: 0 };

function minutesToClock(min) {
  const h = Math.floor(min / 60), m = min % 60;
  const ap = h >= 12 ? "PM" : "AM";
  const h12 = h % 12 || 12;
  return `${h12}:${String(m).padStart(2, "0")} ${ap}`;
}

// The optional rules, summarized only when they are switched on
function extraRulesSummary(rules) {
  const parts = [];
  if (rules.earliestInMin > 0) parts.push(`INs before ${minutesToClock(rules.earliestInMin)} count as ${minutesToClock(rules.earliestInMin)}`);
  if (rules.lunchOverHrs > 0) parts.push(`lunch only over ${rules.lunchOverHrs} hrs`);
  if (rules.inWindowHrs > 0) parts.push(`IN counted only within ${rules.inWindowHrs} hrs of first badge`);
  return parts.join("  |  ");
}

function midpointToStr(sec) {
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return `${m}:${String(s).padStart(2, "0")}`;
}

function rulesSummary(rules) {
  return rules.roundTo > 0
    ? `${midpointToStr(rules.midpointSec)} midpoint to nearest ${rules.roundTo} min`
    : "No rounding (exact badge times)";
}

// Round a Date per the configured rule: nearest {roundTo} min, rounding up when the
// remainder is at or past {midpointSec}. roundTo=0 disables rounding (seconds zeroed).
function roundTime(d, rules) {
  if (!d) return null;
  const r = new Date(d);
  if (!rules || !rules.roundTo) { r.setSeconds(0, 0); return r; }
  const interval = rules.roundTo * 60;
  const totalSec = d.getHours()*3600 + d.getMinutes()*60 + d.getSeconds();
  const rem = totalSec % interval;
  let rounded = rem >= rules.midpointSec ? totalSec + (interval - rem) : totalSec - rem;
  rounded = rounded % 86400;
  const h = Math.floor(rounded / 3600);
  const m = Math.floor((rounded % 3600) / 60);
  r.setHours(h, m, 0, 0);
  return r;
}

function toYMD(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

function addDays(d, n) {
  const r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

// Detect the shift date of a raw CSV: the most common calendar date among
// qualifying IN events (falls back to all events if none qualify).
function detectShiftDate(rows, nqGates) {
  const counts = {};
  let best = null, bestN = 0;
  const tally = (r) => {
    const ts = parseDate(r["Event Timestamp"] || r["EventTimestamp"] || "");
    if (!ts) return;
    const d = toYMD(ts);
    counts[d] = (counts[d] || 0) + 1;
    if (counts[d] > bestN) { best = d; bestN = counts[d]; }
  };
  const qIns = rows.filter(r =>
    (r["Access Point"] || "") === "In" && !nqGates.has(r["Area"] || ""));
  (qIns.length ? qIns : rows).forEach(tally);
  return best;
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
function analyze(rows, shiftDateStr, reportDateStr, nqGates, rules = RULES_DEFAULT) {
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

  // Optional window rule: a qualifying IN only counts within {inWindowHrs}
  // hours of the person's first badge anywhere on site that day, NQ included
  const firstAnyIn = {};
  if (rules.inWindowHrs > 0) {
    df.filter(r => r.access === "In" && sameDay(r.ts, shiftDate))
      .forEach(r => { if (!firstAnyIn[r.guid]) firstAnyIn[r.guid] = r.ts; });
  }

  // First qualifying IN on shift date per person
  const inMap = {};
  dfQ.filter(r => r.access === "In" && sameDay(r.ts, shiftDate))
     .forEach(r => {
       if (rules.inWindowHrs > 0) {
         const anchor = firstAnyIn[r.guid];
         if (anchor && (r.ts - anchor) / 3600000 > rules.inWindowHrs) return;
       }
       if (!inMap[r.guid]) inMap[r.guid] = { ts: r.ts, gate: r.area };
     });

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

    // Optional early start floor: an IN before the earliest counted time bills
    // from exactly that time instead of rounding (a 5:40 IN counts as 6:00)
    let rndIn = roundTime(rawIn, rules);
    if (rawIn && rules.earliestInMin > 0) {
      const todSec = rawIn.getHours() * 3600 + rawIn.getMinutes() * 60 + rawIn.getSeconds();
      if (todSec < rules.earliestInMin * 60) {
        rndIn = new Date(rawIn);
        rndIn.setHours(Math.floor(rules.earliestInMin / 60), rules.earliestInMin % 60, 0, 0);
      }
    }
    const rndOut = roundTime(rawOut, rules);

    let rawHrs = null, lessLunch = null, billHrs = null;
    if (rndIn && rndOut) {
      const elapsedMin = (rndOut - rndIn) / 60000;
      if (elapsedMin > 0) {
        // Lunch comes off every day unless a threshold is set, then only on
        // days whose rounded span is strictly longer than the threshold
        const takeLunch = !rules.lunchOverHrs || elapsedMin / 60 > rules.lunchOverHrs;
        const netMin  = Math.max(elapsedMin - (takeLunch ? rules.lunchMin : 0), 0);
        rawHrs    = +(elapsedMin / 60).toFixed(4);
        lessLunch = +(netMin / 60).toFixed(4);
        billHrs   = lessLunch; // Billable = rndOut - rndIn - lunch; no second rounding pass
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
// XLSX.writeFile uses document.createElement('a').click() which iOS Safari blocks.
// Instead: write to ArrayBuffer → Blob → object URL → anchor click.
function downloadWorkbook(wb, fname) {
  const wbOut = XLSX.write(wb, { bookType:"xlsx", type:"array" });
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

function exportXLSX(results, shiftDate, reportDate, nqGates, rules = RULES_DEFAULT) {
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
    [(rules.roundTo > 0
       ? `Rounding: Raw IN and Raw OUT each rounded independently to nearest ${rules.roundTo} min via ${midpointToStr(rules.midpointSec)} midpoint rule`
       : "Rounding: none — exact badge times used") +
     `  |  Billable = (Rounded OUT minus Rounded IN) minus ${rules.lunchMin} min lunch  |  Times displayed as H:MM` +
     (extraRulesSummary(rules) ? `  |  ${extraRulesSummary(rules)}` : "")],
    dataColHdrs,
    ...results.map(dataRows),
  ]);
  setColWidths(btWs, [28,10,10,22,26,22,22,26,22,16,16,16,22,26,22,26,18]);
  setAutoFilter(btWs, 3, 3 + results.length, dataColHdrs.length - 1);
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
  setAutoFilter(issWs, 1, 1 + issues.length, dataColHdrs.length - 1);
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
  setAutoFilter(sumWs, 1, 9, 1);
  XLSX.utils.book_append_sheet(wb, sumWs, "SUMMARY");

  // ─── TAB 5: CONFIG
  const cfgWs = XLSX.utils.aoa_to_sheet([
    ["CONFIGURATION  |  Update dates and NQ gates here"],
    ["DATE SETTINGS", "Value", "Notes"],
    ["Shift Date (date being analyzed)",    shiftDate,  "Badge INs are only pulled from this date"],
    ["Report Date (date data was received)", reportDate, "Night shift OUTs allowed on this date AM only"],
    [],
    ["RULES", "Value", "Notes"],
    ["Rounding Rule",    rulesSummary(rules), "Editable in the app Config panel"],
    ["Lunch Deduction",  `${rules.lunchMin} minutes`, "Applied when both IN and OUT exist"],
    ["Lunch Threshold",  rules.lunchOverHrs > 0 ? `Only days over ${rules.lunchOverHrs} hrs` : "Every day",
     "Rounded span must be strictly longer than the threshold for lunch to come off"],
    ["Earliest Counted IN", rules.earliestInMin > 0 ? minutesToClock(rules.earliestInMin) : "Off",
     "Badge INs before this time count as exactly this time"],
    ["IN Window",        rules.inWindowHrs > 0 ? `${rules.inWindowHrs} hrs after first badge` : "Off",
     "Qualifying IN counts only within the window after the first badge of the day"],
    ["Badge IN Logic",   "First qualifying IN on shift date only", "Report date INs are excluded"],
    ["Badge OUT Logic",  "Last qualifying OUT per AM/PM rule", "AM IN: same date OUT | PM IN: same date or next AM"],
    ["NQ Gate Rule",     "Off-site swipes shown as reference only", "Not used in calculations"],
    ["Over 13 Hrs Flag", "Any billable total over 13 hrs flagged", "Shown in Issues tab for review"],
    [],
    ["OFF-SITE (NON-QUALIFYING) GATES  |  Add or remove rows as needed"],
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
// rollup: { emps: [{name, company, ein, days:{date:{billHrs,status}}, daysWorked, issueDays, total}], dates: [..] }
// dayMeta: [{shift_date, total_workers, clean_count, missing_out, missing_in, missing_both, over_13, total_bill_hrs}]
function exportWeeklyXLSX(rollup, dayMeta, rules = RULES_DEFAULT) {
  const wb = XLSX.utils.book_new();
  const { emps, dates } = rollup;
  const first = dates[0], last = dates[dates.length - 1];
  const fmtD = d => `${d.slice(5,7)}/${d.slice(8,10)}`;

  function setColWidths(ws, widths) {
    ws["!cols"] = widths.map(w => ({ wch: w }));
  }

  // ─── TAB 1: WEEKLY ROLLUP
  const grandTotal = emps.reduce((s, e) => s + e.total, 0);
  const hdrs = ["Employee Name","Company","EIN", ...dates.map(fmtD), "Days Worked","Issue Days","Total Hrs (H:MM)"];
  const empRow = e => [
    e.name, e.company, e.ein,
    ...dates.map(d => {
      const day = e.days[d];
      if (!day) return "";
      if (day.billHrs != null) return decToHM(day.billHrs);
      return "MISS";
    }),
    e.daysWorked, e.issueDays, decToHM(e.total),
  ];
  const dayTotals = dates.map(d =>
    decToHM(emps.reduce((s, e) => s + ((e.days[d] && e.days[d].billHrs) || 0), 0)));

  const wkWs = XLSX.utils.aoa_to_sheet([
    [`CP2 LNG  |  MMR Gate Log Weekly Rollup  |  ${first} to ${last}  |  ${dates.length} day(s)  |  ${emps.length} employees`],
    [(rules.roundTo > 0
       ? `Rounding: nearest ${rules.roundTo} min via ${midpointToStr(rules.midpointSec)} midpoint`
       : "Rounding: none") +
     `  |  Lunch: ${rules.lunchMin} min  |  MISS = badge record incomplete that day (0 hrs counted)  |  Times as H:MM` +
     (extraRulesSummary(rules) ? `  |  ${extraRulesSummary(rules)}` : "")],
    hdrs,
    ...emps.map(empRow),
    [],
    ["DAILY TOTALS","","", ...dayTotals, "", "", decToHM(grandTotal)],
  ]);
  setColWidths(wkWs, [28,10,10, ...dates.map(() => 9), 12, 11, 15]);
  setAutoFilter(wkWs, 2, 2 + emps.length, hdrs.length - 1);
  XLSX.utils.book_append_sheet(wb, wkWs, "WEEKLY ROLLUP");

  // ─── TAB 2: DAILY SUMMARY
  const dsWs = XLSX.utils.aoa_to_sheet([
    [`DAILY SUMMARY  |  ${first} to ${last}`],
    ["Shift Date","Workers","Clean","Missing OUT","Missing IN","Missing Both","Over 13 Hrs","Total Billable Hrs"],
    ...dayMeta.map(d => [
      d.shift_date, d.total_workers, d.clean_count, d.missing_out, d.missing_in,
      d.missing_both, d.over_13, decToHM(Number(d.total_bill_hrs) || 0),
    ]),
    [],
    ["TOTAL", "", "", "", "", "", "",
     decToHM(dayMeta.reduce((s, d) => s + (Number(d.total_bill_hrs) || 0), 0))],
  ]);
  setColWidths(dsWs, [12,10,10,12,12,13,12,18]);
  setAutoFilter(dsWs, 1, 1 + dayMeta.length, 7);
  XLSX.utils.book_append_sheet(wb, dsWs, "DAILY SUMMARY");

  downloadWorkbook(wb, `CP2_GateLog_Weekly_${first}_to_${last}.xlsx`);
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
  // Default to yesterday's shift, today's report — the normal morning workflow
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

  // ── Weekly rollup state ──
  const [weeklyReports, setWeeklyReports] = useState([]); // report metadata rows
  const [selIds, setSelIds]         = useState(new Set());
  const [rollup, setRollup]         = useState(null); // {emps, dates}
  const [rollupMeta, setRollupMeta] = useState([]);   // day meta for selected reports
  const [weeklyMsg, setWeeklyMsg]   = useState("");
  const [weeklySearch, setWeeklySearch] = useState("");
  const [weeklyBusy, setWeeklyBusy] = useState(false);

  // ── Employee roster state ──
  const [roster, setRoster]         = useState([]); // [{ein, firstName, lastName, supervisor, area, scope}]
  const [rosterMeta, setRosterMeta] = useState(null); // {uploaded_at, uploaded_by, count}
  const [rosterMsg, setRosterMsg]   = useState("");
  const rosterFileRef = useRef();

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

  // Load shared config (NQ gates + rounding rules) on mount.
  // Database is the source of truth for the team; localStorage is a fallback cache.
  useEffect(() => {
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
      } catch (_) { /* silent — defaults still apply */ }
    })();
  }, []);

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
    if (!supabaseReady()) { setCfgMsg("Saved on this device only — database not configured."); return; }
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
      setCfgMsg("Config saved — applies for everyone.");
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

  const currentGates = () => new Set(nqText.split("\n").map(s => s.trim()).filter(Boolean));

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
        mergeSeenGates(rows.map(r => r["Area"]));
        const nqGates = currentGates();
        setResults(analyze(rows, shiftDate, reportDate, nqGates, rules));
        setScreen("results");
      } catch(err) {
        setError("Error: " + err.message);
      } finally {
        setProcessing(false);
      }
    };
    reader.readAsText(file);
  }, [shiftDate, reportDate, nqText, rules, mergeSeenGates]);

  // Multiple files → bulk screen: each file parsed, shift date auto-detected,
  // then every day is saved to the database in one pass.
  const prepareBulk = useCallback((files) => {
    setError("");
    setProcessing(true);
    const nqGates = currentGates();
    Promise.all([...files].map(f =>
      f.text().then(text => {
        const rows = parseCSV(text);
        if (!rows.length) return { name: f.name, rows: [], shiftDate: "", reportDate: "", status: "No data rows" };
        mergeSeenGates(rows.map(r => r["Area"]));
        const sd = detectShiftDate(rows, nqGates) || toYMD(addDays(new Date(), -1));
        return { name: f.name, rows, shiftDate: sd, reportDate: toYMD(addDays(new Date(sd + "T00:00:00"), 1)), status: "" };
      }).catch(err => ({ name: f.name, rows: [], shiftDate: "", reportDate: "", status: "Read error: " + err.message }))
    )).then(parsed => {
      parsed.sort((a, b) => a.shiftDate.localeCompare(b.shiftDate));
      setBulkFiles(parsed);
      setScreen("bulk");
      setProcessing(false);
    });
    // Load shift dates already in the database so duplicates get flagged
    if (supabaseReady()) {
      supabase.from("reports").select("shift_date").then(({ data }) => {
        if (data) setKnownDates(new Set(data.map(r => r.shift_date)));
      });
    }
  }, [nqText, mergeSeenGates]);

  const handleFiles = useCallback((fileList) => {
    const files = [...(fileList || [])].filter(f => f && f.name.toLowerCase().endsWith(".csv"));
    if (!files.length) return;
    if (files.length === 1) processFile(files[0]);
    else prepareBulk(files);
  }, [processFile, prepareBulk]);

  const onDrop = useCallback(e => {
    e.preventDefault(); setDragging(false);
    handleFiles(e.dataTransfer.files);
  }, [handleFiles]);

  const handleExport = () => {
    exportXLSX(results, shiftDate, reportDate, currentGates(), rules);
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
  // Writes the columns that actually exist in the reports table
  // (the old build wrote shifts_count/issues_count, which don't — every save failed).
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

  const handleSave = async () => {
    setSaveMsg("");
    if (!uploadedBy.trim()) { setSaveMsg("Enter your name first."); return; }
    if (!supabaseReady()) { setSaveMsg("Database not configured."); return; }
    if (!results.length) { setSaveMsg("Nothing to save."); return; }
    try {
      await saveReport(results, shiftDate, reportDate, fileName);
      setSaveMsg("Saved.");
    } catch (e) {
      setSaveMsg("Save error: " + e.message);
    }
  };

  // Process every prepared bulk file: analyze with current gates/rules, save each day
  const runBulk = async () => {
    if (!uploadedBy.trim()) { setBulkFiles(f => f.map(x => x.status ? x : { ...x, status: "Enter your name first." })); return; }
    if (!supabaseReady()) { setBulkFiles(f => f.map(x => x.status ? x : { ...x, status: "Database not configured." })); return; }
    setBulkBusy(true);
    const nqGates = currentGates();
    const out = [];
    for (const f of bulkFiles) {
      if (!f.rows.length || !f.shiftDate) { out.push(f); continue; }
      try {
        const res = analyze(f.rows, f.shiftDate, f.reportDate, nqGates, rules);
        await saveReport(res, f.shiftDate, f.reportDate, f.name);
        out.push({ ...f, status: `Saved — ${res.length} workers, ${decToHM(res.reduce((s, r) => s + (r.billHrs || 0), 0))} hrs` });
      } catch (e) {
        out.push({ ...f, status: "Save error: " + e.message });
      }
      setBulkFiles([...out, ...bulkFiles.slice(out.length)]);
    }
    setBulkFiles(out);
    setKnownDates(k => new Set([...k, ...out.filter(f => f.status.startsWith("Saved")).map(f => f.shiftDate)]));
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
    setScreen("history");
    setMenuOpen(false);
  };

  // Saved results have their Date fields serialized to strings — revive them
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
  const loadWeekly = async () => {
    setWeeklyMsg("");
    setMenuOpen(false);
    if (!supabaseReady()) { setWeeklyMsg("Database not configured."); setScreen("weekly"); return; }
    const { data, error } = await supabase
      .from("reports")
      .select(reportListSelect)
      .order("shift_date", { ascending: false })
      .limit(200);
    if (error) { setWeeklyMsg("Load error: " + error.message); }
    setWeeklyReports(data || []);
    setRollup(null);
    setScreen("weekly");
  };

  const toggleSel = (id) => {
    setSelIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
    setRollup(null);
  };

  // Quick-select a Monday–Sunday week: 0 = this week, -1 = last week
  const selectWeek = (offset) => {
    const today = new Date();
    const dow = (today.getDay() + 6) % 7; // Monday = 0
    const mon = addDays(today, -dow + offset * 7);
    const dates = new Set([...Array(7)].map((_, i) => toYMD(addDays(mon, i))));
    setSelIds(new Set(weeklyReports.filter(r => dates.has(r.shift_date)).map(r => r.id)));
    setRollup(null);
  };

  const buildRollup = async () => {
    setWeeklyMsg("");
    const chosen = weeklyReports.filter(r => selIds.has(r.id));
    if (!chosen.length) { setWeeklyMsg("Select at least one day first."); return; }
    setWeeklyBusy(true);
    try {
      const { data, error } = await supabase
        .from("reports")
        .select("id, shift_date, results_json")
        .in("id", [...selIds]);
      if (error) throw error;
      const sorted = (data || []).slice().sort((a, b) => a.shift_date.localeCompare(b.shift_date));
      const byEmp = {};
      const dates = [...new Set(sorted.map(r => r.shift_date))].sort();
      for (const rep of sorted) {
        const res = (rep.results_json && rep.results_json.results) || [];
        for (const r of res) {
          const key = r.guid || `${r.ein}|${r.name}`;
          const e = byEmp[key] || (byEmp[key] = {
            name: r.name || "", company: r.company || "", ein: r.ein || "",
            days: {}, daysWorked: 0, issueDays: 0, total: 0,
          });
          // Same date in two selected reports: the later-saved report wins
          e.days[rep.shift_date] = { billHrs: r.billHrs, status: r.status, over13: r.over13 };
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
      setRollup({ emps, dates });
      setRollupMeta(chosen.slice().sort((a, b) => a.shift_date.localeCompare(b.shift_date)));
    } catch (e) {
      setWeeklyMsg("Rollup error: " + e.message);
    } finally {
      setWeeklyBusy(false);
    }
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

  // ── Bulk upload screen (multiple daily CSVs → database) ──
  if (screen === "bulk") {
    const savedCount = bulkFiles.filter(f => f.status.startsWith("Saved")).length;
    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700, color:"#e2e8f0" }}>SAVE MULTIPLE DAYS</div>
            <div style={{ fontSize:11, color:"#4b5e7a", marginTop:2 }}>
              {bulkFiles.length} file(s) · shift dates auto-detected — adjust before saving
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
                          ? { ...x, shiftDate: e.target.value,
                              reportDate: e.target.value ? toYMD(addDays(new Date(e.target.value + "T00:00:00"), 1)) : x.reportDate }
                          : x))}
                        style={{ ...inp, width:150 }} />
                      {knownDates.has(f.shiftDate) && !f.status.startsWith("Saved") && (
                        <div style={{ color:"#fb923c", fontSize:10, marginTop:3 }}>⚠ already in database — saving again creates a duplicate</div>
                      )}
                    </td>
                    <td style={{ padding:"9px 12px" }}>
                      <input type="date" value={f.reportDate}
                        onChange={e => setBulkFiles(fs => fs.map((x, j) => j === i ? { ...x, reportDate: e.target.value } : x))}
                        style={{ ...inp, width:150 }} />
                    </td>
                    <td style={{ padding:"9px 12px", fontSize:12, fontFamily:"'DM Mono',monospace" }}>{f.rows.length}</td>
                    <td style={{ padding:"9px 12px", fontSize:12,
                      color: f.status.startsWith("Saved") ? "#34d399" : f.status ? "#f87171" : "#4b5e7a" }}>
                      {f.status || "Ready"}
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
            disabled={bulkBusy}
            style={{
              background: bulkBusy ? "#1a2d45" : "#1d4ed8", color:"#fff", border:"none",
              borderRadius:6, padding:"10px 18px", fontSize:13, fontWeight:600,
              cursor: bulkBusy ? "default" : "pointer", fontFamily:"'DM Mono',monospace",
            }}
          >{bulkBusy ? "Saving…" : "Process & Save All"}</button>
          {savedCount > 0 && (
            <button onClick={loadWeekly} style={{ ...btn(true, "#0e7490"), background:"#0e7490", color:"#fff" }}>
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
    const fmtD = d => `${d.slice(5,7)}/${d.slice(8,10)}`;
    const shownEmps = rollup
      ? rollup.emps.filter(e => !weeklySearch ||
          e.name.toLowerCase().includes(weeklySearch.toLowerCase()) ||
          (e.ein || "").includes(weeklySearch))
      : [];
    const grandTotal = rollup ? rollup.emps.reduce((s, e) => s + e.total, 0) : 0;
    const issueCells = rollup ? rollup.emps.reduce((s, e) => s + e.issueDays, 0) : 0;

    return (
      <div style={{ minHeight:"100vh", background:"#060d16", color:"#c8d6e8",
        fontFamily:"'DM Sans','Segoe UI',sans-serif", padding:20 }}>
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

        {/* Day picker */}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:10, alignItems:"center" }}>
          <button onClick={() => selectWeek(0)}  style={btn(false, "#1d4ed8")}>This Week</button>
          <button onClick={() => selectWeek(-1)} style={btn(false, "#1d4ed8")}>Last Week</button>
          {selIds.size > 0 && (
            <button onClick={() => { setSelIds(new Set()); setRollup(null); }} style={btn(false, "#334155")}>Clear</button>
          )}
          <span style={{ fontSize:11, color:"#4b5e7a", fontFamily:"'DM Mono',monospace" }}>
            {selIds.size} day(s) selected
          </span>
        </div>

        {dupSelected && (
          <div style={{ background:"#3b2500", border:"1px solid #9a3412", color:"#fb923c",
            padding:"8px 12px", borderRadius:6, fontSize:12, marginBottom:10 }}>
            ⚠ Two reports with the same shift date are selected — the later-saved one counts for that day.
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
                      <td style={{ padding:"9px 12px", fontSize:12, fontFamily:"'DM Mono',monospace" }}>{r.shift_date}</td>
                      <td style={{ padding:"9px 12px", fontSize:11, fontFamily:"'DM Mono',monospace",
                        color:"#6b7a8d", wordBreak:"break-all" }}>{r.file_name || "—"}</td>
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
              onClick={() => exportWeeklyXLSX(rollup, rollupMeta, rules)}
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
                    {["Employee Name","Company","EIN", ...rollup.dates.map(fmtD), "Days","Total"].map((h, i) => (
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
                      {rollup.dates.map(d => {
                        const day = e.days[d];
                        return (
                          <td key={d} style={{ ...td, fontFamily:"'DM Mono',monospace", fontSize:12,
                            color: !day ? "#374151" : day.billHrs == null ? "#ef4444" : day.over13 ? "#fb923c" : "#c8d6e8" }}>
                            {!day ? "—" : day.billHrs == null ? "MISS" : decToHM(day.billHrs)}
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
                    <tr><td colSpan={5 + rollup.dates.length} style={{ padding:36, textAlign:"center", color:"#4b5e7a" }}>
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
          <button onClick={loadWeekly} style={btn(false, "#334155")}>📅 Weekly</button>
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
                <button onClick={loadWeekly} style={{ ...btn(false, "#334155"), width:"100%", marginBottom:4 }}>Weekly Rollup</button>
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
              Current rule: {rulesSummary(rules)} · lunch {rules.lunchMin} min.{extraRulesSummary(rules) && ` ${extraRulesSummary(rules)}.`} Applies to new uploads — saved reports keep the numbers they were saved with.
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
                {processing ? "Processing…" : fileName || "Tap to browse or drop CSV files"}
              </div>
              <div style={{ fontSize:11, color:"#4b5e7a", marginTop:4 }}>
                Genetec raw export (.csv) — drop several days at once to save them all to the database
              </div>
              <input ref={fileRef} type="file" accept=".csv" multiple style={{display:"none"}}
                onChange={e => { handleFiles(e.target.files); e.target.value = ""; }} />
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
