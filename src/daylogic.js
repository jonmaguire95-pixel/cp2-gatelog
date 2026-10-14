import * as XLSX from "xlsx";

// Day logic for the Gate Log: reading Genetec exports, finding each person's
// shift for a day, and the date helpers the screens share. Nothing in here
// touches the page or the database, so the same code runs in the tests.

// ── CONSTANTS ─────────────────────────────────────────────────────────────────
export const NQ_GATES_DEFAULT = new Set([
  "Helms Rd Turnstile 1","Helms Rd Turnstile 2","Helms Rd VG Staff Badge In/Out",
  "Liberty Visitors Center","PHI MMR Office","PHI Turnstile",
  "Roy Baily Turnstile 1","Roy Baily Turnstile 2","S-Curve Turnstile 1",
  "Helms Rd MMR Office","Helms Rd Worley 4 Plex","Venture Global Lake Charles Warehouse",
  // NOTE: S-Curve Turnstile 2 has significant traffic (96 events in 5/6 data)
  // and is NOT in this NQ list. Monitor and add manually via Config if needed.
]);

// ── ROUNDING RULES (configurable via Config panel, stored in database) ────────
// Defaults preserve original behavior: nearest 30 min via 15:30 (930 sec) midpoint,
// 30 min lunch deduction.
// earliestInMin: INs before this time of day (minutes from midnight) count as
//   exactly this time. 0 turns the rule off.
// lunchOverHrs: lunch comes off only when the rounded span is strictly longer
//   than this many hours. 0 means lunch comes off every day.
// inWindowHrs: an onsite IN counts only within this many hours of the
//   person's first badge anywhere that day. A later one is a Late IN. 0 turns
//   the rule off.
export const RULES_DEFAULT = { roundTo: 30, midpointSec: 930, lunchMin: 30, earliestInMin: 0, lunchOverHrs: 0, inWindowHrs: 0 };

// Night shift OUTs count until 9:00 AM the next morning. An OUT from 9:00 AM
// to noon is shown and flagged for review instead of being used.
export const NIGHT_CUT_H = 9;
export const NOTE_OUT_UNTIL_H = 12;
// An IN and an OUT read at the same gate within a minute of each other are one
// tap (someone touched both readers, or a reader caught the badge twice).
export const TAP_SEC = 60;
// Back on site 4 hours or more after an OUT, from noon on, starts a second
// shift that day. A shorter time out (lunch, a trip to the truck) is a break.
export const SPLIT_HRS = 4;
// A night shift that leaves and comes back within an hour before 9:00 AM is
// still the same night.
export const NIGHT_BREAK_HRS = 1;
const HOUR = 3600000;

// Statuses. The missing badge statuses are the gaps counted by area. The
// review statuses are flagged for a person to check before billing.
export const FLAG_STATUSES = new Set(["Missing IN", "Missing OUT", "Missing IN & OUT"]);
export const REVIEW_STATUSES = new Set(["Late IN", "Late OUT", "Split Shift", "Over 13 Hrs"]);
export const STATUS_LIST = ["Clean", "Missing OUT", "Missing IN", "Missing IN & OUT", "Late IN", "Late OUT", "Split Shift", "Over 13 Hrs"];

// ── UTILITIES ─────────────────────────────────────────────────────────────────
// Genetec writes times as "2026-10-08 17:16:12" in the csv exports, and the
// workbooks are read out as "10/08/2026 05:16:12 PM". Both are read here
// field by field as local time, so every browser (Safari and iPhone included)
// gets the same answer. Anything else falls back to the browser's own reader.
const TS_YMD = /^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3})\d*)?)?$/;
const TS_MDY = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/;

function localDate(y, mo, d, h, mi, s, ms) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59 || s > 59) return null;
  const dt = new Date(y, mo - 1, d, h, mi, s, ms);
  return dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d ? dt : null;
}

export function parseDate(str) {
  if (!str) return null;
  const s = String(str).trim();
  let m = TS_YMD.exec(s);
  if (m) return localDate(+m[1], +m[2], +m[3], +m[4], +m[5], +(m[6] || 0), m[7] ? +m[7].padEnd(3, "0") : 0);
  m = TS_MDY.exec(s);
  if (m) {
    let h = +(m[4] || 0);
    if (m[7]) {
      if (h < 1 || h > 12) return null;
      const pm = /p/i.test(m[7]);
      h = h === 12 ? (pm ? 12 : 0) : pm ? h + 12 : h;
    }
    return localDate(+m[3], +m[1], +m[2], h, +(m[5] || 0), +(m[6] || 0), 0);
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

export function toDateStr(d) {
  if (!d) return "";
  return (
    d.toLocaleDateString("en-US", { month:"2-digit", day:"2-digit", year:"numeric" }) +
    " " +
    d.toLocaleTimeString("en-US", { hour:"2-digit", minute:"2-digit", hour12:true })
  );
}

export function minutesToClock(min) {
  const h = Math.floor(min / 60), m = min % 60;
  const ap = h >= 12 ? "PM" : "AM";
  const h12 = h % 12 || 12;
  return `${h12}:${String(m).padStart(2, "0")} ${ap}`;
}

// The optional rules, summarized only when they are switched on
export function extraRulesSummary(rules) {
  const parts = [];
  if (rules.earliestInMin > 0) parts.push(`INs before ${minutesToClock(rules.earliestInMin)} count as ${minutesToClock(rules.earliestInMin)}`);
  if (rules.lunchOverHrs > 0) parts.push(`lunch only over ${rules.lunchOverHrs} hrs`);
  if (rules.inWindowHrs > 0) parts.push(`IN counted only within ${rules.inWindowHrs} hrs of first badge`);
  return parts.join("  |  ");
}

export function midpointToStr(sec) {
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return `${m}:${String(s).padStart(2, "0")}`;
}

export function rulesSummary(rules) {
  return rules.roundTo > 0
    ? `${midpointToStr(rules.midpointSec)} midpoint to nearest ${rules.roundTo} min`
    : "No rounding (exact badge times)";
}

// Round a Date per the configured rule: nearest {roundTo} min, rounding up when the
// remainder is at or past {midpointSec}. roundTo=0 disables rounding (seconds zeroed).
export function roundTime(d, rules) {
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

export function toYMD(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

export function addDays(d, n) {
  const r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

const stampOf = r => parseDate(r["Event Timestamp"] || r["EventTimestamp"] || "");

// Days with badge activity from 3:00 PM on, with a count of swipes for each
export function afternoonCounts(rows) {
  const counts = {};
  for (const r of rows) {
    const ts = stampOf(r);
    if (ts && ts.getHours() >= 15) {
      const d = toYMD(ts);
      counts[d] = (counts[d] || 0) + 1;
    }
  }
  return counts;
}

// Detect the shift date of a raw export from the data itself, never the file
// name. A pull runs from the shift day into the next morning, so the shift day
// is the last day with badge activity from 3:00 PM on that the data runs past.
// The next morning's badge INs (Monday's day shift after a Sunday, for
// example) don't decide it, and a file covering a whole week lands on its last
// full day. Files with no afternoon swipes at all fall back to the most common
// qualifying IN date.
export function detectShiftDate(rows, nqGates) {
  const pmDays = Object.keys(afternoonCounts(rows)).sort();
  if (pmDays.length) {
    const days = rowDates(rows);
    const lastDay = days[days.length - 1];
    const covered = pmDays.filter(d => d < lastDay);
    return covered.length ? covered[covered.length - 1] : pmDays[pmDays.length - 1];
  }
  const counts = {};
  let best = null, bestN = 0;
  const tally = (r) => {
    const ts = stampOf(r);
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

export function decToHM(dec) {
  if (dec == null || isNaN(dec)) return "";
  const neg = dec < 0;
  const abs = Math.abs(dec);
  const h = Math.floor(abs);
  const m = Math.round((abs - h) * 60);
  return `${neg ? "-" : ""}${h}:${String(m).padStart(2, "0")}`;
}

export function sameDay(a, b) {
  return a && b &&
    a.getFullYear() === b.getFullYear() &&
    a.getMonth()    === b.getMonth()    &&
    a.getDate()     === b.getDate();
}

export function parseCSV(text) {
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

// Genetec exports come as .csv or as Excel workbooks. Both become the same row
// objects (header name to text) so everything downstream works the same way.
// Timestamps in the workbooks are real Excel dates. They're written out in the
// CSV's own format straight from the stored value, so no seconds get lost to
// the cell's display format.
export const GATE_FILE_RE = /\.(csv|xlsx|xlsm|xls)$/i;
const pad2 = n => String(n).padStart(2, "0");

function excelDateText(v, kind, date1904) {
  const p = XLSX.SSF.parse_date_code(v, { date1904 });
  if (!p) return String(v);
  const date = `${pad2(p.m)}/${pad2(p.d)}/${p.y}`;
  const sec = Math.min(59, Math.floor(p.S));
  if (kind === "date") return date;
  if (kind === "time") return `${pad2(p.H)}:${pad2(p.M)}:${pad2(sec)}`;
  const h12 = p.H % 12 || 12;
  return `${date} ${pad2(h12)}:${pad2(p.M)}:${pad2(sec)} ${p.H < 12 ? "AM" : "PM"}`;
}

const DATE_KIND = { "Event Timestamp": "datetime", "EventTimestamp": "datetime", "Event_Date": "date", "Event_Time": "time" };

export function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: "array", cellDates: false, dense: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];
  const date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: false });
  if (aoa.length < 2) return [];
  const headers = aoa[0].map(h => String(h == null ? "" : h).trim());
  return aoa.slice(1).map(cols => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = cols[i];
      if (v == null || v === "") obj[h] = "";
      else if (typeof v === "number" && DATE_KIND[h]) obj[h] = excelDateText(v, DATE_KIND[h], date1904);
      else obj[h] = String(v).trim();
    });
    return obj;
  });
}

// A Genetec raw export has a cardholder and a timestamp on every swipe.
// Anything else (a saved Gate Log report, a roster) is turned away so it
// can't become an empty or made up day.
export function swipeCount(rows) {
  let n = 0;
  for (const r of rows) {
    if ((r["CardholderGuid"] || r["Cardholder Guid"] || r["cardholderguid"]) && stampOf(r)) n++;
  }
  return n;
}
export const NO_SWIPES = "No badge swipes found. Use the Genetec raw export.";

// Exports of the same day overlap (the .2 file is usually a later, longer pull
// of the first one). Identical swipes are kept once. The .csv and .xlsx
// exports write times differently, so the time is compared as a time.
export function mergeRows(lists) {
  const seen = new Set();
  const out = [];
  for (const rows of lists) {
    for (const r of rows) {
      const raw = r["Event Timestamp"] || r["EventTimestamp"] || "";
      const t = parseDate(raw);
      const key = [
        r["CardholderGuid"] || r["Cardholder Guid"] || "",
        t ? t.getTime() : raw,
        r["Door"] || "", r["Access Point"] || "", r["Area"] || "", r["Event"] || "",
      ].join("|");
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(r);
    }
  }
  return out;
}

// Calendar dates found in a set of rows
export function rowDates(rows) {
  const dates = new Set();
  for (const r of rows) {
    const ts = stampOf(r);
    if (ts) dates.add(toYMD(ts));
  }
  return [...dates].sort();
}

export const shortDate = ymd => `${Number(ymd.slice(5, 7))}/${Number(ymd.slice(8, 10))}`;
export const weekdayOf = ymd => ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][new Date(ymd + "T00:00:00").getDay()];
export const dayLabel = ymd => `${weekdayOf(ymd)} ${shortDate(ymd)}`;

// The weekly rollup runs Monday to Sunday by shift date. weekStart gives the
// Monday of the week a shift date falls in, weekDays the seven dates.
export const weekStart = ymd => {
  const d = new Date(ymd + "T00:00:00");
  return toYMD(addDays(d, -((d.getDay() + 6) % 7)));
};
export const weekDays = mon => [...Array(7)].map((_, i) => toYMD(addDays(new Date(mon + "T00:00:00"), i)));

// One saved report per shift date among `days`, the most recently saved, so a
// day that was saved twice is only picked once
export function newestPerDate(reports, days) {
  const best = {};
  for (const r of reports) {
    if (days && !days.includes(r.shift_date)) continue;
    const b = best[r.shift_date];
    if (!b || String(r.uploaded_at || "") > String(b.uploaded_at || "")) best[r.shift_date] = r;
  }
  return Object.values(best);
}

// ── BULK DROPS ────────────────────────────────────────────────────────────────
// Files dropped together are grouped by the shift date read from each file's
// own swipes. Any dropped file can hold swipes for a given day: the Monday
// weekly file holds the whole week (including swipes that reached Genetec
// after the daily pull), and the next day's files hold the morning after the
// daily pull stopped. Each day takes its own files plus every other file's
// swipes from noon the day before (to see who was already on a night shift)
// through noon the next morning (night shift OUTs, and later ones to flag).
// Nothing further away is used.
// read: [{ name, rows, problem }] in the order the files were dropped
export function groupDays(read, nqGates, fallbackDate) {
  const groups = {};
  const problems = [];
  for (const f of read) {
    if (f.problem) {
      problems.push({ name: f.name, rows: [], shiftDate: "", reportDate: "", note: "", status: f.problem, fileCount: 1 });
      continue;
    }
    const sd = detectShiftDate(f.rows, nqGates) || fallbackDate;
    const g = groups[sd] || (groups[sd] = { names: [], lists: [] });
    g.names.push(f.name);
    g.lists.push(f.rows);
  }
  const timed = read.filter(f => !f.problem).map(f => ({
    rows: f.rows,
    t: f.rows.map(r => { const d = stampOf(r); return d ? d.getTime() : NaN; }),
  }));
  const prepared = Object.entries(groups).map(([sd, g]) => {
    const own = g.lists.length > 1 ? mergeRows(g.lists) : g.lists[0];
    const day = new Date(sd + "T00:00:00");
    const start = addDays(day, -1).setHours(12, 0, 0, 0);
    const end = addDays(day, 1).setHours(NOTE_OUT_UNTIL_H, 0, 0, 0);
    const extra = [];
    for (const f of timed) {
      if (g.lists.includes(f.rows)) continue;
      f.rows.forEach((r, i) => { if (f.t[i] >= start && f.t[i] < end) extra.push(r); });
    }
    const rows = extra.length ? mergeRows([own, extra]) : own;
    const added = rows.length - own.length;
    const dates = rowDates(own);
    const notes = [];
    if (g.names.length > 1) notes.push(`${g.names.length} files combined, duplicate swipes removed`);
    if (added > 0) notes.push(`plus ${added} swipe${added === 1 ? "" : "s"} for this day from the other files`);
    if (dates.length > 3) notes.push(`file covers ${shortDate(dates[0])} to ${shortDate(dates[dates.length - 1])}; only the shift date is used`);
    return {
      name: g.names.join(" + "),
      rows,
      shiftDate: sd,
      reportDate: toYMD(addDays(day, 1)),
      note: notes.join(". "),
      status: "",
      fileCount: g.names.length,
    };
  });
  return [...prepared, ...problems]
    .sort((a, b) => (a.shiftDate || "9999").localeCompare(b.shiftDate || "9999"));
}

// ── MAIN ANALYSIS ─────────────────────────────────────────────────────────────
// One row per person on site for the shift date, built from their badge swipes
// in time order.
//
// A shift starts at an onsite IN on the shift date and ends at the last onsite
// OUT after it: on the shift date for a shift that starts before noon, or
// through 9:00 AM the next morning for one that starts at noon or later (a
// night shift). Going out and coming back in is a break in the same shift
// (lunch, a trip to the truck), unless the person was gone 4 hours or more
// and came back from noon on, which starts a second shift that day. Coming
// back in and never badging out again leaves the shift open: Missing OUT.
//
// An onsite IN more than {inWindowHrs} after the person's first badge of the
// day (at Helms Road, for example) is a Late IN: shown and flagged, with no
// hours counted. Two opposite reads at the same gate within a minute are one
// tap. Swipes the day before only show whether someone was still on last
// night's shift, so that night's OUTs and stray taps aren't read as a new day.
//
// No onsite IN on the shift date: an onsite OUT from 10:00 AM on, or before
// 9:00 AM the next morning, is a Missing IN. So is an onsite OUT before 10:00
// AM after an off site IN that morning (the reader tapped the wrong way).
export function analyze(rows, shiftDateStr, reportDateStr, nqGates, rules = RULES_DEFAULT) {
  const shiftDate  = new Date(shiftDateStr  + "T00:00:00");
  const reportDate = new Date(reportDateStr + "T00:00:00");
  const win = rules.inWindowHrs > 0 ? rules.inWindowHrs * HOUR : 0;

  const df = rows.map(r => ({
    guid:    r["CardholderGuid"] || r["Cardholder Guid"] || r["cardholderguid"] || "",
    name:    r["Employee Name"] || "",
    company: r["Company Name"] || "",
    ein:     r["Company EIN"] || "",
    area:    r["Area"] || "",
    access:  r["Access Point"] || "",
    ts:      stampOf(r),
    isNQ:    nqGates.has(r["Area"] || ""),
  })).filter(r => r.guid && r.ts);
  df.sort((a, b) => a.ts - b.ts);

  const people = new Map();
  for (const r of df) {
    const list = people.get(r.guid);
    if (list) list.push(r); else people.set(r.guid, [r]);
  }

  const day = dayFrame(shiftDate, reportDate);
  const prevDay = dayFrame(addDays(shiftDate, -1), shiftDate);

  const results = [];
  for (const [guid, sw] of people) {
    const res = personShift(guid, sw, day, prevDay, win, rules);
    if (res) results.push(res);
  }
  results.sort((a, b) => a.name.localeCompare(b.name));
  return results;
}

// Times that bound one shift date's search
function dayFrame(shiftDate, reportDate) {
  const at = (d, h) => new Date(d).setHours(h, 0, 0, 0);
  return {
    start: shiftDate.getTime(),
    end: addDays(shiftDate, 1).getTime(),
    ten: at(shiftDate, 10),
    repStart: reportDate.getTime(),
    cut: at(reportDate, NIGHT_CUT_H),
    note: at(reportDate, NOTE_OUT_UNTIL_H),
    repTen: at(reportDate, 10),
  };
}

// Stray reads. Reads at the same gate less than a minute apart are one tap.
// When a tap has both an IN and an OUT (someone touched both readers, or a
// reader caught the badge twice), an IN read after an OUT read is the stray
// one: the person was leaving. An IN then an OUT both stand, so a quick pass
// through a gate opens and closes on the spot. At the start of a day it isn't
// known yet whether the person is arriving or leaving: they're leaving (a
// night shift going home, and the IN reads are the stray ones) when they were
// last seen coming in on site, or going out on site within the hour, or, with
// nothing earlier in the data, when they badge out again within 3 hours.
// Otherwise they're arriving and the OUT reads are the stray ones, unless
// they're not seen again that day (in and straight back out).
function strayReads(sw, dayStart, dayEnd) {
  const stray = new Set();
  const on = sw.filter(s => !s.isNQ);
  let seenOnDay = false;   // any onsite read on the shift date before this tap
  for (let i = 0; i < on.length; ) {
    // One tap: reads at the same gate, each within a minute of the one before
    let j = i + 1;
    while (j < on.length && on[j].area === on[i].area && on[j].ts - on[j - 1].ts <= TAP_SEC * 1000) j++;
    const tap = on.slice(i, j);
    const mixed = tap.some(s => s.access === "In") && tap.some(s => s.access === "Out");
    const onDay = tap[0].ts >= dayStart && tap[0].ts < dayEnd;
    if (mixed && onDay && !seenOnDay) {
      // The first onsite activity of the day: arriving or leaving?
      const prev = lastBefore(sw, tap[0], stray, 14 * HOUR);
      let leaving;
      if (prev) leaving = !prev.isNQ && (prev.access === "In" || tap[0].ts - prev.ts <= HOUR);
      else {
        const lastT = tap[tap.length - 1].ts;
        const next = sw.find(s => s.ts > lastT && !tap.includes(s));
        leaving = !!next && next.access === "Out" && next.ts - lastT <= 3 * HOUR;
      }
      // Arriving and then nothing else all day: they came in and went right
      // back out, so both reads stand
      const lastT = tap[tap.length - 1].ts;
      const passThrough = !leaving && !sw.some(s => s.ts > lastT && s.ts < dayEnd && !tap.includes(s));
      if (!passThrough) for (const s of tap) if (s.access === (leaving ? "In" : "Out")) stray.add(s);
    } else if (mixed) {
      let seenOut = false;
      for (const s of tap) {
        if (s.access === "Out") seenOut = true;
        else if (seenOut) stray.add(s);
      }
    }
    if (onDay) seenOnDay = true;
    i = j;
  }
  return stray;
}

// The last read before `a` that isn't part of its own tap or a stray read
function lastBefore(sw, a, stray, within) {
  for (let i = sw.indexOf(a) - 1; i >= 0; i--) {
    const s = sw[i];
    if (a.ts - s.ts > within) return null;
    if (stray.has(s) || (!s.isNQ && s.area === a.area && a.ts - s.ts <= TAP_SEC * 1000)) continue;
    return s;
  }
  return null;
}

// Walk one person's swipes for a shift date and split their time on site into
// stretches. Returns the stretches, the OUTs seen with no IN before them, and
// the time of the last swipe this shift used.
function walkShift(real, f, skipUntil) {
  const stretches = [];
  const leading = [], nightExits = [], lateExits = [];
  let cur = null, inside = false, used = -Infinity, stopT = Infinity;
  const pm = s => s.start.ts.getHours() >= 12;
  for (const s of real) {
    if (s.ts < f.start || s.ts <= skipUntil) continue;
    if (s.access === "In") {
      if (s.ts >= f.end) {
        // An IN on the report date. Coming back on site within the hour on a
        // night shift is a break; any other IN there (off site at Helms Road,
        // or back on site later) means this shift is over.
        if (!s.isNQ && cur && pm(cur) && !inside && cur.end && s.ts < f.cut && s.ts - cur.end.ts < NIGHT_BREAK_HRS * HOUR) {
          cur.breaks.push({ out: cur.end, in: s }); cur.end = null; inside = true; used = +s.ts; continue;
        }
        if (!s.isNQ && inside) continue;   // already on site: another reader inside
        stopT = +s.ts;
        break;
      }
      if (s.isNQ || inside) continue;
      if (cur && cur.end && (s.ts - cur.end.ts < SPLIT_HRS * HOUR || s.ts.getHours() < 12)) {
        cur.breaks.push({ out: cur.end, in: s }); cur.end = null; inside = true; used = +s.ts; continue;
      }
      cur = { start: s, end: null, breaks: [], lateOut: null, nextMorningOut: null };
      stretches.push(cur); inside = true; used = +s.ts;
      continue;
    }
    if (s.isNQ) continue;
    // An onsite OUT
    if (!cur) {
      if (s.ts < f.end) leading.push(s);
      else if (s.ts >= f.repStart && s.ts < f.cut) { nightExits.push(s); used = +s.ts; }
      else if (s.ts >= f.cut && s.ts < f.repTen) { lateExits.push(s); used = +s.ts; }
      continue;
    }
    const limit = pm(cur) ? f.cut : f.end;
    if (s.ts >= limit) {
      if (!cur.end && s.ts >= f.repStart) {
        if (pm(cur) && s.ts < f.note && !cur.lateOut) { cur.lateOut = s; used = +s.ts; }
        if (!pm(cur) && s.ts < f.cut && !cur.nextMorningOut) cur.nextMorningOut = s;
      }
      continue;
    }
    cur.end = s; inside = false; used = +s.ts;
  }
  // A second stretch only stands as a second shift when it's a night shift
  // (it runs past midnight into the next morning). Back on site later the
  // same day is a long break in the same shift: people drive out a gate to
  // work outside the fence and come back hours later.
  for (let k = stretches.length - 1; k > 0; k--) {
    const s = stretches[k];
    const night = (s.end && s.end.ts >= f.end) || s.lateOut;
    if (night) continue;
    const p = stretches[k - 1];
    p.breaks.push({ out: p.end, in: s.start }, ...s.breaks);
    p.end = s.end; p.lateOut = s.lateOut; p.nextMorningOut = p.nextMorningOut || s.nextMorningOut;
    stretches.splice(k, 1);
  }
  return { stretches, leading, nightExits, lateExits, used, stopT };
}

const hhmm = d => { const h = d.getHours(), m = d.getMinutes(); return `${h % 12 || 12}:${String(m).padStart(2, "0")} ${h < 12 ? "AM" : "PM"}`; };

function personShift(guid, sw, f, prevF, win, rules) {
  const emp = sw[0];
  const stray = strayReads(sw, f.start, f.end);
  const real = sw.filter(s => !stray.has(s));

  // Last night's shift, if the data shows one, keeps its own OUTs and breaks
  const prev = walkShift(real.filter(s => s.ts >= prevF.start && s.ts < f.note), prevF, -Infinity);
  const lastPrev = prev.stretches[prev.stretches.length - 1];
  const skipUntil = lastPrev && lastPrev.start.ts.getHours() >= 12 && prev.used >= f.start ? prev.used : -Infinity;

  const w = walkShift(real, f, skipUntil);
  const onDay = s => s.ts >= f.start && s.ts < f.end && s.ts > skipUntil;
  const when = d => (d >= f.start && d < f.end ? "" : `${d.getMonth() + 1}/${d.getDate()} `) + hhmm(d);
  const at = s => `${when(s.ts)} at ${s.area}`;

  // Off site reference: first off site IN that day, last off site OUT
  const nqIn = real.find(s => s.isNQ && s.access === "In" && onDay(s)) || null;
  const lastSeen = [...real].reverse().find(s => onDay(s)) || null;

  const hoursFor = (rawIn, rawOut) => {
    let rndIn = roundTime(rawIn, rules);
    // Optional early start floor: an IN before the earliest counted time bills
    // from exactly that time instead of rounding (a 5:40 IN counts as 6:00)
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
    return { rndIn, rndOut, rawHrs, lessLunch, billHrs };
  };

  const row = (status, inS, outS, hrs, note, flagArea) => {
    // Last off site OUT during this shift: by midnight for a day shift, by
    // noon the next morning otherwise, and before the next shift's first IN
    const fromT = inS ? +inS.ts : f.start;
    const toT = Math.min(w.stopT, inS && inS.ts.getHours() < 12 ? f.end : f.note);
    const nqOut = [...real].reverse().find(s => s.isNQ && s.access === "Out" && s.ts >= fromT && s.ts > skipUntil && s.ts < toT) || null;
    const h = hrs || { rndIn: inS ? hoursFor(inS.ts, null).rndIn : null, rndOut: outS ? roundTime(outS.ts, rules) : null, rawHrs: null, lessLunch: null, billHrs: null };
    const over13 = h.billHrs != null && h.billHrs > 13;
    return {
      guid, name: emp.name, company: emp.company, ein: emp.ein,
      rawIn: inS ? inS.ts : null, rawOut: outS ? outS.ts : null, rndIn: h.rndIn, rndOut: h.rndOut,
      inGate: inS ? inS.area : "", outGate: outS ? outS.area : "",
      rawHrs: h.rawHrs, lessLunch: h.lessLunch, billHrs: h.billHrs,
      nqIn:  nqIn  ? { ts: nqIn.ts,  gate: nqIn.area }  : null,
      nqOut: nqOut ? { ts: nqOut.ts, gate: nqOut.area } : null,
      status: status === "Clean" && over13 ? "Over 13 Hrs" : status,
      over13,
      note: note || "",
      // Gate area where the unmatched swipe happened (see Issues by Area)
      flagArea: flagArea || "",
    };
  };

  // ── No onsite IN on the shift date ──
  if (!w.stretches.length) {
    const after10 = w.leading.filter(s => s.ts >= f.ten);
    const out = w.nightExits.length ? w.nightExits[w.nightExits.length - 1]
      : after10.length ? after10[after10.length - 1] : null;
    if (out) return row("Missing IN", null, out, null, "", out.area);
    // Badged in off site, then tapped OUT on site that morning: the IN was
    // tapped the wrong way
    const offIn = real.find(s => s.isNQ && s.access === "In" && onDay(s));
    const wrongWay = offIn ? w.leading.filter(s => s.ts > offIn.ts) : [];
    if (wrongWay.length) {
      const o = wrongWay[wrongWay.length - 1];
      return row("Missing IN", null, o, null,
        `Badged IN off site at ${offIn.area} ${hhmm(offIn.ts)}, then OUT on site ${at(o)} with no onsite IN.`, o.area);
    }
    if (w.lateExits.length) {
      const o = w.lateExits[w.lateExits.length - 1];
      return row("Missing IN", null, o, null,
        `OUT ${at(o)} next morning, after the ${NIGHT_CUT_H}:00 AM night shift cutoff.`, o.area);
    }
    return null;
  }

  // ── One or more stretches on site ──
  const S = w.stretches;
  // Window rule: each stretch's IN is measured from the first badge (any gate)
  // since the previous stretch ended, or since midnight for the first
  const anchorFor = (k) => {
    const from = k === 0 ? Math.max(f.start, skipUntil) : +S[k - 1].end.ts;
    return real.find(s => s.access === "In" && s.ts > from - (k === 0 ? 1 : 0) && s.ts >= f.start && s.ts <= S[k].start.ts) || S[k].start;
  };
  const info = S.map((s, k) => {
    const anchor = anchorFor(k);
    const late = win > 0 && s.start.ts - anchor.ts > win;
    return { s, anchor, late, hrs: s.end ? hoursFor(s.start.ts, s.end.ts) : null };
  });
  const lateNote = i => `IN on site ${at(i.s.start)} was ${((i.s.start.ts - i.anchor.ts) / HOUR).toFixed(1)} hrs after the first badge of the day, ${at(i.anchor)}.`;
  const breakNote = s => s.breaks.length
    ? s.breaks.map(b => `OUT ${at(b.out)}, back IN ${at(b.in)}`).join("; ") : "";
  const span = i => `${when(i.s.start.ts)} to ${when(i.s.end.ts)}`;

  const last = info[info.length - 1];
  const earlier = info.slice(0, -1);
  const earlierNote = earlier.length
    ? " Also on site " + earlier.map(i => i.s.end ? `${span(i)}${i.hrs && i.hrs.billHrs ? ` (${decToHM(i.hrs.billHrs)} hrs)` : ""}` : `from ${at(i.s.start)}`).join(", ") + "." : "";

  // The last stretch has no OUT
  if (!last.s.end) {
    const s = last.s;
    if (s.lateOut) {
      return row("Late OUT", s.start, s.lateOut, null,
        `Night shift OUT ${at(s.lateOut)}, after the ${NIGHT_CUT_H}:00 AM cutoff. Hours left blank for review.` + earlierNote,
        s.start.area);
    }
    if (last.late) {
      const lead = w.leading.filter(o => o.ts < s.start.ts);
      const back = lead.length && lead[lead.length - 1].area === s.start.area && lead[lead.length - 1].ts.getHours() < 12 && s.start.ts.getHours() >= 15;
      const note = (lead.length ? `OUT ${at(lead[lead.length - 1])}, then ` : "") +
        `IN ${at(s.start)} (${((s.start.ts - last.anchor.ts) / HOUR).toFixed(1)} hrs after the first badge of the day), no OUT after.` +
        (back ? " Likely tapped the wrong way both times." : "") + earlierNote;
      return row("Missing IN & OUT", null, null, null, note, lastSeen ? lastSeen.area : s.start.area);
    }
    let note = breakNote(s);
    note = note ? `${note}, no OUT after.` : "";
    if (s.nextMorningOut) note += `${note ? " " : ""}OUT ${at(s.nextMorningOut)} next morning; a day shift OUT has to be on the shift date.`;
    return row("Missing OUT", s.start, null, null, (note + earlierNote).trim(), s.start.area);
  }

  // Every stretch closed. One stretch is the shift; with more than one, the
  // longest is the shift and the others are noted.
  if (info.length === 1) {
    if (last.late) return row("Late IN", last.s.start, last.s.end, null, lateNote(last) + " Hours left blank for review.", last.s.start.area);
    return row("Clean", last.s.start, last.s.end, last.hrs, "", "");
  }
  const openEarlier = earlier.find(i => !i.s.end);
  if (openEarlier) {
    // Cannot happen: a stretch only ends when an OUT closes it
    return row("Missing OUT", openEarlier.s.start, null, null, "", openEarlier.s.start.area);
  }
  const elapsed = i => i.s.end.ts - i.s.start.ts;
  const main = info.reduce((m, i) => (elapsed(i) >= elapsed(m) ? i : m), info[0]);
  const others = info.filter(i => i !== main);
  const othersNote = others.map(i => `${span(i)}${i.hrs && i.hrs.billHrs ? ` (${decToHM(i.hrs.billHrs)} hrs)` : ""}`).join(", ");
  if (main.late) {
    return row("Late IN", main.s.start, main.s.end, null,
      lateNote(main) + ` Hours left blank for review. Also on site ${othersNote}.`, main.s.start.area);
  }
  const counted = others.filter(i => !i.late && i.hrs && i.hrs.billHrs);
  if (counted.length) {
    return row("Split Shift", main.s.start, main.s.end, main.hrs,
      `Two shifts this day. Hours shown are for ${span(main)}. Also on site ${othersNote}, not counted.`, "");
  }
  return row("Clean", main.s.start, main.s.end, main.hrs, `Also on site ${othersNote}.`, "");
}

// ── ISSUES BY AREA ────────────────────────────────────────────────────────────
export function gateFlagArea(r) {
  if (r.flagArea) return r.flagArea;
  if (r.status === "Missing OUT") return r.inGate || "";
  if (r.status === "Missing IN") return r.outGate || "";
  if (r.status === "Missing IN & OUT") return (r.nqOut && r.nqOut.gate) || (r.nqIn && r.nqIn.gate) || "";
  return "";
}

export function summarizeByArea(items, areaOf) {
  const map = {};
  for (const it of items) {
    if (!FLAG_STATUSES.has(it.status)) continue;
    const name = areaOf(it) || "Unknown";
    const row = map[name] || (map[name] = { area: name, missingIn: 0, missingOut: 0, missingBoth: 0, total: 0 });
    if (it.status === "Missing IN") row.missingIn++;
    else if (it.status === "Missing OUT") row.missingOut++;
    else row.missingBoth++;
    row.total++;
  }
  return Object.values(map).sort((a, b) => b.total - a.total || a.area.localeCompare(b.area));
}

// Counts of each status for a day's results
export function statusCounts(results) {
  const c = Object.fromEntries(STATUS_LIST.map(s => [s, 0]));
  for (const r of results) c[r.status] = (c[r.status] || 0) + 1;
  c.over13 = results.filter(r => r.over13).length;
  c.flags = results.filter(r => FLAG_STATUSES.has(r.status)).length;
  c.review = results.filter(r => REVIEW_STATUSES.has(r.status)).length;
  return c;
}
