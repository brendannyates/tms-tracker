/* TMS Tracker: Excel (web) task-pane add-in. A light augmentation over the organization's patient tracker template.
   One tab per course (LASFIR or "LASFIR 2") cloned from "New Template". Information stays where the template keeps it:
   header A1:H7 (MRN A4, # of txs D2, # of MT redos D4, MD on-site D6, status D7, provider E2, TAS E4, insurance E6, auth F4:G4,
   update protocol by G2, progress note sent G6), protocol J2:O4, freebies left N6, auth assessment N7, visit log A10:T49
   (time in T) closed by the TX SUMMARY row with its checkbox in A, home clinic S9, weekly availability V19:X24 and dates away V28:Y48.
   Benefits live on the Benefits tab, one row per MRN (see BEN_COLS). They keep their old cell names (W1 BIDF, W4 deductible,
   W5 OOP max, W6 MTR approved, W7 extensions, W8 last schedule, X12:Y16 cost by code) as keys, and applyOps sends those there.
   Tabs in the previous layout (freebies in N7, benefits in V1:Y16) are still read; Match New Template moves them over.
   PHI: MRNs are read into memory to match NextGen rows; they are never sent anywhere. Names typed in Add patient become LASFIR.
   v2 screens (Today chair board, Schedule, Patients, Chairs) follow the MHS tracker design. core.js holds the floor rules; check-in,
   session times and chair moves live on the Sessions sheet, not on the patient tabs. */
"use strict";

/* ---------- tab layout ---------- */
const LOG_FIRST = 10, LOG_LAST = 49, LOG_READ = 150;   // the template's log is A10:T49; an import can grow it (see growLogs)
/* a grown log is found from column A's PR formula: its $L$10:$L$n range stretches with every row inserted inside the log */
const prEnd = f => { const m = String(f || "").match(new RegExp(String.raw`\$L\$${LOG_FIRST}:\$L\$(\d+)`)); return m ? +m[1] : null; };   // where the PR formula's ranges stop
const logEnd = f => { const m = String(f || "").match(new RegExp(String.raw`\$L\$${LOG_FIRST}:\$L\$(\d+)`)); return m ? Math.max(LOG_LAST, +m[1]) : LOG_LAST; };
/* the TX SUMMARY row closes the log, so the row above it is the log's real end (rows added by hand below the formula's range count) */
const SUM_RE = /tx\s*summary/i, SUM_LABEL = "TX SUMMARY / TMS CONNECT / FINAL FOLLOW UP";
const endOf = (colB, f) => { const i = colB.findIndex(v => SUM_RE.test(String(v ?? ""))); return i > 0 ? LOG_FIRST + i - 1 : logEnd(f); };
async function tabEnd(ctx, ws) {
  const a10 = ws.getRange(`A${LOG_FIRST}`), b = ws.getRange(`B${LOG_FIRST}:B${LOG_FIRST + LOG_READ - 1}`); a10.load("formulas"); b.load("values"); await ctx.sync();
  return endOf(b.values.map(r => r[0]), a10.formulas[0][0]);
}
/* Benefits tab: one row per MRN. Keys are the cells the benefits block used to have on each tab. */
const BEN_SHEET = "Benefits", BEN_FIRST = 2, BEN_LAST = 1000;
const BEN_COLS = { MRN: "A", Patient: "B", W1: "C", W4: "D", W5: "E", W6: "F", W7: "G", W8: "H", X12: "I", X13: "J", X14: "K", X15: "L", X16: "M", Y12: "N", Y13: "O", Y14: "P", Y15: "Q", Y16: "R", Updated: "S" };
const BEN_HDR = ["MRN", "Patient", "BIDF estimate", "Deductible remaining", "Out-of-pocket max remaining", "MTR approved", "Extensions added (90868)", "Last schedule provided",
  "MT PR", "Daily PR", "MTR PR", "F/U in-office PR", "F/U telehealth PR", "MT before ded.", "Daily before ded.", "MTR before ded.", "F/U in-office before ded.", "F/U telehealth before ded.", "Last updated"];
const BEN_KEYS = Object.keys(BEN_COLS).filter(k => /^[WXY]\d+$/.test(k)), BEN_STAMP = "BEN_STAMP";
const isBenKey = a => a === BEN_STAMP || BEN_KEYS.includes(a);
const BLANK_ROW = { pr: null, tx: "", tech: "", date: null, time: null, pulses: [null, null, null], delivered: false, type: "", miss: false, action: "", alt: false, note: "", tele: false, freebie: "", phq: "", gad: "", measX: false, measVal: false };
const C = { pr: 0, tech: 1, tx: 2, date: 4, L: 5, R: 7, O: 9, type: 11, miss: 12, action: 14, phq: 15, gad: 16, alt: 17, note: 18, time: 19 }; // 0-based in A:T
const COST_FIRST = 12, WK_FIRST = 20, AV_FIRST = 29, AV_LAST = 48;
const BILL = [
  { k: "MT", code: "90867", label: "MT" }, { k: "Daily", code: "90868", label: "Daily" }, { k: "MTR", code: "90869", label: "MTR / MT redo" },
  { k: "FUo", code: "99214", label: "F/U in-office" }, { k: "FUt", code: "99214", label: "F/U telehealth" }
];
const TAB_RE = /^([A-Z]{1,6})(?: (\d+))?$/;
const ACTIVE = ["Pending Start", "In Progress", "Tapering"];
const FREEBIE_DAYS = 10;   // a miss's freebie decision is asked for this many days
const DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const AV_TYPE = ["Unavailable", "Available"];
const FREEBIE_TAG = { Used: "Freebie used", Waived: "Freebie waived" };
const SITES = ["Left", "Right", "Other"];
const D4_FMT = "0";                                              // # of MT redos (D4) is a count, never a date
const A4_FMT = "0";                                              // the MRN (A4) is a number, no decimals
const OUTCOMES = ["No-show", "Same-day cancel"];                 // both require a freebie classification
const STD_VISIT = ["MT", "MTR", "F/U"];
const isTxRow = r => !r.miss && ["MT", "Daily", "Taper", "MTR"].includes(r.type);   // a numbered row in C (Tx #)
const TELE_TAG = "Tele";
const SNOOZE_MS = 3 * 3600e3;
/* A tab is only read or written if it matches the current template: PATIENT: in A1, MRN: in A3, and these headers in row 8. */
const SIG = [[0, "tx #"], [2, "date"], [9, "type"], [10, "miss"], [12, "action"], [16, "notes:"]];   // offsets within C8:S8
/* Updates and Match New Template only touch trackers with a visit (column E) on or after the cutoff whose status (D7) isn't Discontinued or Complete */
const CUTOFF_DEFAULT = "2026-10-12";
const EXEMPT_TAG = "NextGen: Exempt, confirm", UNCLASS_TAG = "NextGen: no-show not classified";
/* One add-in, one workbook per device. The device comes from the workbook name unless the setting overrides it. */
const DEVICES = { MagV: { chair: "MagV1", other: /brainsway|\bBW\b/i, ours: /magv|mag\s?venture/i }, Brainsway: { chair: "Brainsway", other: /magv|mag\s?venture/i, ours: /brainsway|\bBW\b/i } };
function device() {
  const pick = state.cfg.device;
  if (pick === "MagV" || pick === "Brainsway") return pick;
  return /brainsway/i.test(state.wbName) ? "Brainsway" : "MagV";
}
const chairName = () => DEVICES[device()].chair;
/* this tracker's clinic: the Settings value, else the workbook name ("Castro MagV Patient Tracker" is Castro) */
const clinic = () => String(state.cfg.location || "").trim() || TmsCore.clinicFromName(state.wbName);
/* S9's home clinic names this clinic: its name, a short form of it (CAS), or SRL for San Rafael */
const atClinic = s => { const c = clinic().toLowerCase(), x = String(s || "").toLowerCase(); return !c || x.includes(c) || (c === "srl" && x.includes("san rafael")); };
const homeHere = h => { const x = String(h).trim().toLowerCase(), c = clinic().toLowerCase(); return x === c || (x.length >= 3 && c.startsWith(x)) || (!!c && x.includes(c)) || (x === "srl" && (!c || c === "san rafael")); };

const state = {
  cfg: { authDays: 14, protoDays: 2, recentDays: 10, device: "auto", location: "", openTab: "on", cutoff: CUTOFF_DEFAULT },
  patients: [], tabs: [], protocols: [], fix: new Map(), skipped: { layout: [], other: 0 }, old: [], stale: [], templateOk: true, templateWhy: "", pendingNew: [],
  drafts: {}, initials: "", techName: "", apSig: null, closures: new Map(), wbName: "", clinic: null, report: null,
  view: "today", ptOpen: null, ptPast: false, ptPick: null, weekOffset: 0, schedOffset: 0, dash: "day", ptTab: null, cfilter: "needs", file: null, plan: null, codeEdited: false, courseEdited: false,
  dismissed: new Set(), snoozeMem: {}, avDraft: null, lx: null, apptBy: new Map(), apptTo: null, edits: [], elogOpen: false, elogTab: "", chair: "all", sess: null, fl: null, sel: null, mctx: null, copied: new Set()
};
try { state.chair = localStorage.getItem("tms_chair") || "all"; } catch { /* storage unavailable */ }
/* Discontinued patients are left off the schedules, Attention and the roster (except its Discontinued filter) */
const onSched = p => !/^discontinued/i.test(String(p.status).trim());
/* chairs: from the saved appointment list's resource names (for example Castro MagV1 and Castro MagV2). A visit's chair is the
   resource booked that day, else the chair the patient's daily treatments are usually on. Today's view (Clinic or one chair) is
   saved per computer in state.chair; listCheck and variances count that chair only. */
const chairs = () => state.lx ? state.lx.chairs : [];
const chairOf = (p, r) => { const lx = state.lx; if (!lx || !p.mrn) return ""; return (r && lx.v[`${p.mrn}|${r.date}`]) || lx.pt[p.mrn] || ""; };
const chairPick = () => chairs().length > 1 && chairs().includes(state.chair) ? state.chair : "all";
const inChair = (p, r) => { const c = chairPick(); if (c === "all") return true; const x = chairOf(p, r); return !x || x === c; };
/* more than one patient in the same chair at once; with no chair info everyone counts as one chair, as before */
const chairClash = occ => { const m = new Map(); occ.forEach(e => { const k = chairOf(e.p, e.r); if (!m.has(k)) m.set(k, new Set()); m.get(k).add(e.p.tab); }); return [...m.values()].some(s => s.size > 1); };
/* Pop-out: the same page opened in an Office dialog window (?popout=1). A dialog cannot reach the workbook, so it asks the
   task pane (which stays open) to do every read and write, and the pane sends back a snapshot of what it read. */
const POPOUT = /[?&]popout=1\b/.test(location.search);
const pop = { dlg: null, calls: 0 };                                   // task pane side
const rpc = { id: 0, wait: new Map(), bufs: {}, snapWait: [] };       // pop-out side
const $ = id => document.getElementById(id);
const colLetter = i => { let s = ""; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };   // 0-based column index to letters (A, Z, AA)
const colIndex = l => l.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;   // letters back to a 0-based index
const esc =s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- dates and times ---------- */
const utcSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5);
const todaySerial = () => { const d = new Date(); return utcSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); };
const serialToDate = s => new Date(Date.UTC(1899, 11, 30) + s * 864e5);
const dow = s => serialToDate(s).getUTCDay();
const fmtDate = s => { if (typeof s !== "number") return ""; const d = serialToDate(s); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}`; };
const fmtDateY = s => { if (typeof s !== "number") return ""; const d = serialToDate(s); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`; };
const fmtTime = m => { if (m == null) return ""; const h = Math.floor(m / 60); return `${((h + 11) % 12) + 1}:${String(m % 60).padStart(2, "0")} ${h < 12 ? "AM" : "PM"}`; };
const weekStart = s => s - dow(s);
const inputToSerial = v => { if (!v) return null; const [y, m, d] = v.split("-").map(Number); return utcSerial(y, m, d); };
const serialToInput = s => serialToDate(s).toISOString().slice(0, 10);
const minToInput = m => m == null ? "" : `${String(Math.floor(m / 60)).padStart(2, "0")}:${String(m % 60).padStart(2, "0")}`;
const cmp = (a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0);
const isNum = v => typeof v === "number" && v > 0;
const money = v => typeof v === "number" ? `$${v.toFixed(2)}` : "";
const bizDaysUntil = (from, to) => { if (to < from) return -1; let n = 0; for (let d = from + 1; d <= to; d++) { const w = dow(d); if (w > 0 && w < 6) n++; } return n; };
function toSerial(v) {
  if (v instanceof Date) { const d = new Date(v.getTime() + 3e4); return utcSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  if (typeof v === "number") return v > 20000 ? Math.floor(v) : null;
  const s = String(v ?? "").trim(); let m;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/))) { let y = +m[3]; if (y < 100) y += 2000; return utcSerial(y, +m[1], +m[2]); }
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return utcSerial(+m[1], +m[2], +m[3]);
  const d = new Date(s); return isNaN(d) ? null : toSerial(d);
}
function toMinutes(v) {
  if (v instanceof Date) { const d = new Date(v.getTime() + 3e4); return d.getHours() * 60 + d.getMinutes() || null; }
  if (typeof v === "number") { const f = v < 1 ? v : v - Math.floor(v); return Math.round(f * 1440) || null; }
  const m = String(v ?? "").match(/(\d{1,2}):(\d{2})\s*([AaPp])?[Mm]?/);   // "9:30 A", "1:00 PM", "13:00"
  if (!m) return null;
  let h = +m[1]; const ap = (m[3] || "").toLowerCase();
  if (ap === "p" && h < 12) h += 12; if (ap === "a" && h === 12) h = 0;
  return h * 60 + +m[2];
}
const clean = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z]/g, "").toUpperCase();
const lasfir = (last, first) => clean(last).slice(0, 3) + clean(first).slice(0, 3);
const normMrn = s => String(s ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "").replace(/^0+/, "");

/* NextGen report vocabulary */
function mapSts(raw) {
  const s = String(raw ?? "").toLowerCase();
  if (/no[\s-]?show/.test(s)) return "No-show";
  if (/resched/.test(s)) return "Skip";
  if (/cancel/.test(s)) return "Cancelled";
  if (/kept|complete|checked|arrived|seen/.test(s)) return "Kept";
  return "Scheduled";                                    // Expected
}
function mapEvent(ev) {
  const s = String(ev ?? "").toLowerCase();
  if (/tele/.test(s)) return { type: "F/U", tele: true };
  if (/follow|f\/u/.test(s)) return { type: "F/U", tele: false };
  if (/redo|mtr/.test(s)) return { type: "MTR", tele: false };
  if (/\bmt\b/.test(s)) return { type: "MT", tele: false };
  if (/treatment/.test(s)) return { type: "Daily", tele: false };
  return null;
}
const chairMin = t => t === "MT" || t === "MTR" ? 60 : t === "Daily" || t === "Taper" ? 30 : 0;   // chair side
const provMin = t => t === "MT" ? 60 : t === "MTR" || t === "F/U" ? 30 : 0;                        // provider side
const billIdx = r => r.type === "MT" ? 0 : r.type === "Daily" || r.type === "Taper" ? 1 : r.type === "MTR" ? 2 : r.type === "F/U" ? (r.tele ? 4 : 3) : -1;

/* ---------- UI helpers ---------- */
let toastTimer;
/* sticky: an error that stays until closed, with Copy, so its exact text can be passed on */
function toast(msg, err, sticky) {
  if (pop.dlg && pop.calls > 0) toChild({ k: "toast", msg, err: !!err, sticky: !!sticky });   // a pop-out asked for this action: show the result there too
  const t = $("toast"); t.className = err ? "err" : ""; t.style.display = "block"; clearTimeout(toastTimer);
  if (!sticky) { t.textContent = msg; toastTimer = setTimeout(() => t.style.display = "none", err ? 9000 : 4500); return; }
  t.innerHTML = '<div class="tmsg"></div><div class="tact"><button class="sm" data-tcopy>Copy</button><button class="sm" data-tclose>Close</button></div>';
  const m = t.querySelector(".tmsg"); m.textContent = msg;
  t.querySelector("[data-tclose]").onclick = () => { t.style.display = "none"; };
  t.querySelector("[data-tcopy]").onclick = async e => {
    let ok = false; try { await navigator.clipboard.writeText(msg); ok = true; } catch { /* select it instead */ }
    if (!ok) { const r = document.createRange(); r.selectNodeContents(m); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
    e.target.textContent = ok ? "Copied" : "Press Ctrl+C";
  };
}
/* failStep names what was being done; Office errors add their code and the call that failed (for example Range.values) */
let failStep = "";
const fail = e => {
  console.error(e);
  const dbg = e && e.debugInfo, where = dbg && dbg.errorLocation ? ` at ${dbg.errorLocation}` : "";
  const code = e && e.code && e.code !== e.message ? ` [${e.code}${where}]` : where ? ` [${where.slice(4)}]` : "";
  toast(`${failStep ? failStep + ": " : ""}${(e && e.message) || String(e)}${code}`, true, true); failStep = "";
};
let modalFrom = null;
function showModal(html, wide) {
  if ($("modal").hidden) modalFrom = document.activeElement;
  const b = $("modalBody"); b.innerHTML = html; b.classList.toggle("wide", !!wide); b.setAttribute("aria-labelledby", "mtitle"); $("modal").hidden = false;
  const f = b.querySelector("input:not([type=hidden]):not(:disabled):not([hidden]),select,textarea") || b.querySelector(".btn.primary,button"); if (f) f.focus();
}
function closeModal() {
  $("modal").hidden = true; $("modalBody").innerHTML = ""; state.mctx = null;
  if (modalFrom && document.contains(modalFrom)) modalFrom.focus(); modalFrom = null;
}

/* ---------- read the workbook ---------- */
/* a missed visit's freebie decision lives in Notes ("Freebie used" or "Freebie waived"); p.fbCell (N6) counts the freebies left */
const freebieOf = note => /\bfreebie used\b/i.test(note) ? "Used" : /\bfreebie waived\b/i.test(note) ? "Waived" : "";
const withFreebie = (note, v) => [...String(note || "").split("; ").filter(x => x && !/^freebie (used|waived)$/i.test(x)), ...(v ? [FREEBIE_TAG[v]] : [])].join("; ");
const isNA = v => /^n\/?a$/i.test(String(v ?? "").trim());
/* layout 3 (current): Time in T8 and Freebie Count in J6, benefits on the Benefits tab. Layout 2: Time in T8 and the cost table
   header in V11 (freebies in N7, benefits on the tab). Anything else is an earlier layout, rebuilt from New Template. */
const layoutOf = (t8, j6, v11) => String(t8).trim().toLowerCase() !== "time" ? 0 : /^freebie count/i.test(String(j6).trim()) ? 3 : String(v11).trim().toLowerCase() === "code" ? 2 : 0;
/* vt: the log's A:D value types, when read. "Empty" there means the cell lost its formula (PR in A, Tx left in D): a formula
   that shows nothing still reads as String */
function parseLog(vals, vt) {
  const blank = x => x === "" || x == null;
  return vals.map((v, i) => {
    const note = String(v[C.note] || ""), t = (vt && vt[i]) || [];
    return {
      noA: t[0] === "Empty", noD: t[3] === "Empty",
      row: LOG_FIRST + i, pr: typeof v[C.pr] === "number" ? v[C.pr] : null, tx: v[C.tx], tech: String(v[C.tech] || "").trim(),
      date: isNum(v[C.date]) ? Math.floor(v[C.date]) : null, time: isNum(v[C.time]) ? Math.round(v[C.time] * 1440) : null,
      pulses: [v[C.L], v[C.R], v[C.O]].map(x => blank(x) ? null : x), delivered: [v[C.L], v[C.R], v[C.O]].some(x => !blank(x)),
      type: String(v[C.type] || "").trim(), miss: String(v[C.miss]).trim().toLowerCase() === "x", action: String(v[C.action] || ""),
      alt: String(v[C.alt]).trim().toLowerCase() === "x", note, tele: /\btele/i.test(note), freebie: freebieOf(note),
      phq: String(v[C.phq] ?? "").trim(), gad: String(v[C.gad] ?? "").trim(),
      measX: [v[C.phq], v[C.gad]].some(x => String(x).trim().toLowerCase() === "x"), measVal: [v[C.phq], v[C.gad]].some(x => !blank(x) && String(x).trim().toLowerCase() !== "x")
    };
  });
}
async function refresh() {
  if (POPOUT) return callPane("refresh", []);
  try {
    await firstLaunch();
    await Excel.run(async ctx => {
      ctx.workbook.load("name"); const sheets = ctx.workbook.worksheets; sheets.load("items/name"); await ctx.sync(); state.wbName = ctx.workbook.name || ""; state.sheetNames = sheets.items.map(s => s.name);
      const cand = sheets.items.filter(s => TAB_RE.test(s.name) || s.name === "New Template");
      const sg = cand.map(s => ({ name: s.name, a1: s.getRange("A1"), a3: s.getRange("A3"), h: s.getRange("C8:T8"), v11: s.getRange("V11"), j6: s.getRange("J6"), d7: s.getRange("D7"), e: s.getRange("E9:E200") }));
      sg.forEach(x => { x.a1.load("values"); x.a3.load("values"); x.h.load("values"); x.v11.load("values"); x.j6.load("values"); x.d7.load("values"); x.e.load("values"); });
      /* sheet names are already known, so the Protocols, Clinic Schedule and Benefits reads ride along with this sync (names match case-insensitively, as getItem does) */
      const sheetOf = n => sheets.items.find(s => s.name.toLowerCase() === n.toLowerCase());
      const prW = sheetOf("Protocols"), csW = sheetOf("Clinic Schedule"), bnW = sheetOf(BEN_SHEET), ssW = sheetOf(SESS_SHEET);
      const prv = prW ? prW.getRange("A2:I40") : null, cs = csW ? csW.getRange("A3:H40") : null, bn = bnW ? bnW.getRange(`A1:S${BEN_LAST}`) : null;
      const ss = ssW ? ssW.getRange(`A${SESS_FIRST}:J${SESS_LAST}`) : null, sc = ssW ? ssW.getRange(`L${CH_FIRST}:P${CH_LAST}`) : null;
      [prv, cs, bn, ss, sc].forEach(r => r && r.load("values")); await ctx.sync();
      readBen(bn ? bn.values : null); readSess(ss ? ss.values : null, sc ? sc.values : null);
      await readElog(ctx);
      const cut = inputToSerial(state.cfg.cutoff || CUTOFF_DEFAULT), eDate = v => typeof v === "number" ? (v > 20000 ? Math.floor(v) : null) : /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(String(v).trim()) ? toSerial(v) : null;
      const live = x => x.name === "New Template" || (!/^(discontinued|complete)/i.test(String(x.d7.values[0][0]).trim()) && x.e.values.some(([v]) => (eDate(v) ?? 0) >= cut));
      state.stale = sg.filter(x => !live(x)).map(x => x.name); const stale = new Set(state.stale);
      const ok = x => String(x.a1.values[0][0]).trim().toUpperCase() === "PATIENT:" && String(x.a3.values[0][0]).trim().toUpperCase() === "MRN:" && SIG.every(([i, w]) => String(x.h.values[0][i]).trim().toLowerCase() === w);
      const lay = x => ok(x) ? layoutOf(x.h.values[0][17], x.j6.values[0][0], x.v11.values[0][0]) : 0;
      const layout = new Map(sg.map(x => [x.name, lay(x)]));
      state.templateOk = layout.get("New Template") === 3;
      /* what's off about New Template, in words, for the banner and the Match New Template button */
      const nt = sg.find(x => x.name === "New Template"), shown = v => String(v ?? "").trim() === "" ? "is blank" : `reads "${String(v).trim()}"`;
      state.templateWhy = state.templateOk ? "" : !nt ? "There's no tab named New Template, so tabs can't be matched or added."
        : !ok(nt) ? `New Template's headings aren't the tracker's, so tabs can't be matched to it. It needs PATIENT: in A1, MRN: in A3, and ${SIG.map(([i, w]) => `${w[0].toUpperCase() + w.slice(1)} in ${colLetter(2 + i)}8`).join(", ")}.`
        : `New Template needs Time in T8 and Freebie Count in J6 before tabs can match it. T8 ${shown(nt.h.values[0][17])} and J6 ${shown(nt.j6.values[0][0])}.`;
      state.old = sg.filter(x => x.name !== "New Template" && live(x) && ok(x) && !layout.get(x.name)).map(x => x.name);   // a tracker tab in an earlier layout: rebuilt from New Template
      state.skipped = { layout: sg.filter(x => x.name !== "New Template" && live(x) && !ok(x)).map(x => x.name), other: 0 };   // tabs not named by LASFIR are left alone without a mention
      /* read in batches: Excel on the web caps how much one sync can return, and each log is 3,000 cells */
      const L = [];
      for (const s of sheets.items.filter(s => layout.get(s.name) >= 2)) {
        const g = (a, f) => { const r = s.getRange(a); r.load(f ? "formulas" : "values"); return r; };
        L.push({ name: s.name, layout: layout.get(s.name), hdr: g("A2:H7"), prot: g("J2:O5"), n67: g("N6:N7"), w18: g("W1:W8"), cost: g(`V11:Y${COST_FIRST + 4}`), a10: g("A10", true),
          wk: g(`W${WK_FIRST}:X${WK_FIRST + 4}`), av: g(`V${AV_FIRST}:Y${AV_LAST}`), log: g(`A${LOG_FIRST}:T${LOG_FIRST + LOG_READ - 1}`), s9: g("S9") });
        L[L.length - 1].vt = s.getRange(`A${LOG_FIRST}:D${LOG_FIRST + LOG_READ - 1}`); L[L.length - 1].vt.load("valueTypes");
        L[L.length - 1].d4 = s.getRange("D4"); L[L.length - 1].d4.load("numberFormat");
        L[L.length - 1].a4 = s.getRange("A4"); L[L.length - 1].a4.load("numberFormat");
        if (L.length % 8 === 0) await ctx.sync();
      }
      await ctx.sync();
      /* # of MT redos (D4) is a count: a date or decimal format there gets "0" (finished tabs are left alone) */
      const d4s = L.filter(l => !stale.has(l.name) && l.d4.numberFormat[0][0] !== D4_FMT);
      /* the MRN (A4) is a number with no decimals, never a date or text: on New Template and every tab whose authorization hasn't
         ended (G4 today or later) or with a visit (E) today or later. Other tabs are left as they are. */
      const t4 = todaySerial(), a4Live = l => l.name === "New Template" || (isNum(l.hdr.values[2][6]) && l.hdr.values[2][6] >= t4) || l.log.values.some(r => (eDate(r[4]) ?? 0) >= t4);
      const a4s = L.filter(l => l.a4.numberFormat[0][0] !== A4_FMT && a4Live(l));
      if (d4s.length || a4s.length) { d4s.forEach(l => { l.d4.numberFormat = [[D4_FMT]]; }); a4s.forEach(l => { l.a4.numberFormat = [[A4_FMT]]; }); await ctx.sync(); }
      state.closures = new Map();
      if (cs) cs.values.slice(1).forEach(v => { if (isNum(v[6])) state.closures.set(Math.floor(v[6]), String(v[7] || "Closed")); });
      state.patients = []; state.tabs = []; state.fix = new Map();
      for (const l of L) {
        const a10 = String(l.a10.formulas[0][0] || ""), end = endOf(l.log.values.map(r => r[1]), a10), sumA = (l.log.values[end + 1 - LOG_FIRST] || [])[0];
        /* what Match New Template would change: the layout, a PR formula that doesn't read the Benefits tab or stops short of the
           log, a summary row without its checkbox, or (New Template) the benefits block still on the tab */
        const f = { layout: l.layout === 2, pr: !/Benefits!/.test(a10) || prEnd(a10) !== end, sum: typeof sumA !== "boolean", block: l.name === "New Template" && String(l.cost.values[0][0]).trim().toLowerCase() === "code" };
        if (Object.values(f).some(Boolean) && !stale.has(l.name)) state.fix.set(l.name, f);
        if (l.name === "New Template") continue;
        const m = l.name.match(TAB_RE), pv = l.prot.values, h = l.hdr.values, H = (a, c) => h[a - 2][c];
        const mins = [0, 1, 2].map(i => pv[i][4]).filter(x => typeof x === "number");
        const m9 = String(l.s9.values[0][0] || "").match(/home\s*clinic\s*:\s*(.*)$/i), home = m9 ? m9[1].trim() : "";
        const money0 = v => typeof v === "number" ? v : null, v3 = l.layout === 3, mrn = normMrn(H(4, 0));
        /* layout 3 reads benefits from the patient's Benefits row; layout 2 from its own V1:Y16 */
        const br = v3 ? (state.ben.rows.get(mrn) || {}).v || [] : null, bc = k => br[colIndex(BEN_COLS[k])];
        const w = v3 ? ["W1", "", "", "W4", "W5", "W6", "W7", "W8"].map(k => k ? bc(k) : "") : l.w18.values.map(r => r[0]);
        const cost = v3 ? BILL.map((b, i) => ({ after: money0(bc(`X${COST_FIRST + i}`)), before: money0(bc(`Y${COST_FIRST + i}`)) }))
          : BILL.map((b, i) => ({ after: money0(l.cost.values[i + 1][2]), before: money0(l.cost.values[i + 1][3]) }));
        const n6 = l.n67.values[0][0], n7 = l.n67.values[1][0], fbv = v3 ? n6 : n7;
        state.tabs.push({ name: l.name, code: m[1], course: +(m[2] || 1) });
        const p = {
          tab: l.name, code: m[1], course: +(m[2] || 1), mrn, mrnRaw: String(H(4, 0) || "").trim(), mrnVal: H(4, 0), layout: l.layout,
          fbCell: v3 ? "N6" : "N7", assessCell: v3 ? "N7" : "N6", benUpdated: v3 && isNum(bc("Updated")) ? Math.floor(bc("Updated")) : null,
          txApproved: typeof H(2, 3) === "number" ? H(2, 3) : 36, mtRedos: typeof H(4, 3) === "number" ? H(4, 3) : 0, mdReq: String(H(6, 3) || ""), status: String(H(7, 3) || ""),
          provider: String(H(2, 4) || "").trim(), tas: String(H(4, 4) || "").trim(),
          insurance: String(H(6, 4) || ""), kaiser: /kaiser/i.test(String(H(6, 4) || "")),
          authStart: isNum(H(4, 5)) ? H(4, 5) : null, authExp: isNum(H(4, 6)) ? H(4, 6) : null,
          updateBy: isNum(H(2, 6)) ? H(2, 6) : null, progSent: isNum(H(6, 6)) ? Math.floor(H(6, 6)) : null,
          txEnd: isNum(H(7, 5)) ? H(7, 5) : null, txRem: typeof H(7, 7) === "number" ? H(7, 7) : null,
          authAssess: String((v3 ? n7 : n6) || ""), freebiesLeft: typeof fbv === "number" ? fbv : null,
          protoRows: [0, 1, 2].map(i => ({ n: i + 1, row: 2 + i, hz: pv[i][0], dur: pv[i][1], trains: pv[i][2], wait: pv[i][3], site: String(pv[i][5] || "") })),
          protoMin: mins.length ? mins.reduce((a, b) => a + b, 0) : null, protoTotal: typeof pv[3][4] === "number" && pv[3][4] > 0 ? pv[3][4] : null,   // N5
          home, external: !!home && !homeHere(home),
          ben: { bidf: money0(w[0]), dedNA: isNA(w[3]), dedRem: money0(w[3]), dedBlank: w[3] === "" || w[3] == null, oopNA: isNA(w[4]), oopRem: money0(w[4]),
            mtr: String(w[5] || ""), ext: typeof w[6] === "number" ? w[6] : 0, lastSched: isNum(w[7]) ? Math.floor(w[7]) : null },
          cost,
          canGrow: /LET\(/.test(a10),   // only a PR formula in A10 can say where a grown log ends
          rows: parseLog(l.log.values.slice(0, end - LOG_FIRST + 1), l.vt.valueTypes), end,
          week: Object.fromEntries([1, 2, 3, 4, 5].map(d => { const v = l.wk.values[d - 1] || []; return [d, { from: isNum(v[0]) ? Math.round(v[0] * 1440) : null, to: isNum(v[1]) ? Math.round(v[1] * 1440) : null }]; })),
          avail: l.av.values.filter(v => AV_TYPE.includes(v[0]) && isNum(v[1])).map(v => ({ type: v[0], from: Math.floor(v[1]), to: isNum(v[2]) ? Math.floor(v[2]) : Math.floor(v[1]), note: String(v[3] || "") }))
        };
        p.mtrAuth = p.mtRedos > 0 || /^y/i.test(p.ben.mtr);
        p.lastSched = p.ben.lastSched;
        state.patients.push(p);
      }
      state.patients.sort((a, b) => a.code.localeCompare(b.code) || a.course - b.course);
      state.patients.forEach(p => p.d = derive(p));
      state.protocols = prv ? prv.values.filter(r => r[1] !== "").map(r => ({ type: r[0], name: String(r[1]), rep: r[2], pulses: r[3], trains: r[4], iti: r[5], typed: r[7], inten: r[8] })) : [];
      state.hasAP = sheets.items.some(s => s.name === AP_SHEET);
      state.clinic = null;
      if (cs) {
        const days = {}; for (let i = 0; i < 5; i++) { const v = cs.values[i + 1], t = x => typeof x === "number" && x > 0 ? Math.round(x * 1440) : null; days[i + 1] = { open: t(v[1]), close: t(v[2]), bs: t(v[3]), be: t(v[4]) }; }
        state.clinic = { days, slot: num(cs.values[7][1], 30), closureTable: cs.values[0][6] === "Closed date" };
      }
    });
    state.report = Office.context.document.settings.get("tms_report") || null;
    if (!state.file && !state.listLoaded) {
      state.listLoaded = true;
      try { state.file = await loadList(); if (state.file) await buildPlan(true); } catch (e) { console.error(e); toast(`The ${LIST_SHEET} tab could not be read. Import the list again.`, true); }
    }
    state.pendingNew = state.pendingNew.filter(x => !state.patients.some(p => p.mrn === x.mrn));
    /* an MT redo in D4 means it was approved: W6 follows (finished trackers are left alone) */
    const mtrFix = state.patients.filter(p => p.mtRedos >= 1 && !/^y/i.test(p.ben.mtr) && !state.stale.includes(p.tab));
    if (mtrFix.length) try { await applyOps(mtrFix.map(p => ({ tab: p.tab, a: "W6", v: [["Yes"]], quiet: true })), null, { kind: "auto", what: "MTR approved set to Yes: # of MT redos (D4) has one" }); mtrFix.forEach(p => { p.ben.mtr = "Yes"; }); } catch (e) { console.error(e); }
    /* every approved treatment done (36 of 36, 15 of 15): Complete session asks to mark it Complete, and until then the course
       shows the need "Set to Complete". Nothing is written on its own, so an extension on the way isn't closed out. */
    renderBanners(); renderDash(); updateApTab(); fillAddPatientForm();
    if (pop.dlg) sendSnap();
    syncActiveSheet();
    /* once a session: visit logs edited by hand get the formatting that follows their dates (hand-colored rows stay) */
    if (!state.fmtRan) { state.fmtRan = true; formatLogs(state.patients.filter(p => !state.stale.includes(p.tab)).map(p => ({ tab: p.tab, touched: [] }))).catch(e => console.error(e)).then(sessAmPmOnce); }
  } catch (e) { fail(e); }
}
const num = (v, d) => (v === "" || v == null || isNaN(Number(v))) ? d : Number(v);

/* payment (patient responsibility) for a visit: column A's PR value (its formula applies the deductible), else the cost table */
const payOf = (p, r) => typeof r.pr === "number" ? r.pr : r.miss ? null : ((p.cost[billIdx(r)] || {}).after ?? null);
/* BIDF (benefits) info a patient still needs before visits can be priced: cost per code, the deductible (or N/A), the OOP max (or N/A) */
function bidfMissing(p) {
  const out = [], c = p.cost, need = [0, 1, 3].concat(p.mtrAuth ? [2] : []).concat(p.rows.some(r => r.type === "F/U" && r.tele) ? [4] : []);
  const noAfter = need.filter(i => c[i].after == null).map(i => BILL[i].label);
  if (noAfter.length) out.push(`cost for ${noAfter.join(", ")}`);
  if (p.ben.dedBlank) out.push("deductible remaining (or N/A)");
  else if (p.ben.dedRem > 0) { const nb = need.filter(i => c[i].before == null).map(i => BILL[i].label); if (nb.length) out.push(`before-deductible cost for ${nb.join(", ")}`); }
  if (!p.ben.oopNA && p.ben.oopRem == null) out.push("out-of-pocket max remaining (or N/A)");
  return out;
}
const happened = (r, today) => r.date != null && r.date <= today && !r.alt && (r.delivered || (r.type === "F/U" && !r.miss));
function baselineOf(p) {
  const mt = p.rows.find(r => r.type === "MT" && /baseline/i.test(r.note)); if (!mt) return null;
  const g = k => { const m = mt.note.match(new RegExp(`baseline ${k}:\\s*([^;]+)`, "i")); return m ? m[1].trim() : ""; };
  return { sleep: g("sleep"), caffeine: g("caffeine"), date: mt.date };
}
const protoName = x => { const m = state.protocols.find(q => q.rep === x.hz && q.pulses === x.dur && q.trains === x.trains && q.iti === x.wait); return m ? m.name.trim() : ""; };
const siteRows = p => p.protoRows.filter(x => isNum(x.hz) && TmsCore.siteOf(x.site));   // protocol rows with a site in O
const protoText = p => p.protoRows.filter(x => isNum(x.hz)).map(x => `${TmsCore.siteOf(x.site) || "Protocol " + x.n}: ${(TmsCore.protoMatch(x, state.protocols) || {}).name || `${x.hz} Hz, ${x.dur}/train, ${x.trains} trains, ${x.wait}s`}`).join("; ") || "not entered";
const protoLen = p => p.protoTotal || p.protoMin;   // session minutes: N5, else N2:N4 added up
const yearOf = s => serialToDate(s).getUTCFullYear();

function derive(p) {
  const today = todaySerial(), cfg = state.cfg, rows = p.rows, rec = cfg.recentDays, total = p.txApproved || 36;
  const status = p.status, active = status === "" || ACTIVE.includes(status);
  const sched = rows.filter(r => r.date != null && r.date >= today && !r.delivered && !r.miss).sort(cmp);
  const next = sched.find(r => r.type !== "F/U" && !r.alt) || sched[0];
  const tx = rows.filter(r => r.delivered && r.type !== "F/U");
  const txDone = tx.length, lastTx = tx.length ? Math.max(...tx.map(r => r.date ?? 0)) : null;
  const overdue = rows.filter(r => r.date != null && r.date < today && !r.delivered && !r.miss && !r.alt && r.type !== "F/U");
  const noTech = tx.filter(r => !r.tech && r.date != null);
  const decisions = rows.filter(r => r.miss && !r.freebie && r.date != null && today - r.date < FREEBIE_DAYS);   // older ones are past deciding
  const used = rows.filter(r => r.freebie === "Used").length, waived = rows.filter(r => r.freebie === "Waived").length;
  const left = p.freebiesLeft != null ? Math.max(0, p.freebiesLeft) : Math.max(0, 3 - used);
  const fu = rows.filter(r => r.type === "F/U");
  const fuNext = fu.filter(r => r.date != null && r.date >= today && !r.delivered && !r.miss).sort(cmp)[0];
  const fuStatus = !fu.length ? "none" : fuNext ? "scheduled" : fu.every(r => r.delivered || r.date != null && r.date < today) ? "done" : "not scheduled";
  const futureMT = rows.some(r => r.type === "MT" && r.date != null && r.date >= today && !r.delivered && !r.miss);
  const hasRecent = rows.some(r => r.date != null && r.date >= today - rec);
  const stale = (status === "In Progress" || status === "Tapering") && txDone < total && !hasRecent;
  /* cancellations on the saved appointment list (latest course only): dates still on the tracker, or nothing left booked */
  const lx = state.lx && p.mrn && !state.patients.some(q => q.mrn === p.mrn && q.course > p.course) ? state.lx.c[p.mrn] : null;
  const seen = rows.filter(r => r.date != null && r.date <= today && (r.delivered || (r.type === "F/U" && !r.miss && !r.alt))).map(r => r.date);
  const lastSeen = Math.max(lx && lx.lastKept != null ? lx.lastKept : -1, ...seen, -1), seenAt = lastSeen >= 0 ? lastSeen : null;
  const cxRows = lx ? rows.filter(r => r.date != null && !r.delivered && !r.miss && lx.cx.includes(r.date)) : [];
  const allCx = !!lx && lx.cx.length > 0 && !lx.booked;
  const quit = allCx && seenAt != null && seenAt < today - rec;   // every appointment cancelled and last seen over 10 days ago
  let suggest = null;
  if (!["Paused", "Discontinued"].includes(status)) {
    if (status !== "Complete" && txDone >= total) suggest = "Complete";   // refresh writes it (autoComplete)
    else if (status !== "Complete" && quit) suggest = "Discontinued";
    else if (status !== "Complete" && txDone >= 1 && (status === "Pending Start" || status === "")) suggest = "In Progress";
    else if (status === "" && txDone === 0 && futureMT) suggest = "Pending Start";
  }
  const cancelled = /^(discontinued|paused|complete)$/i.test(status) || suggest === "Complete" || suggest === "Discontinued" || !(cxRows.length || allCx) ? null
    : [cxRows.length ? `Cancelled in NextGen, still on the tab: ${[...new Set(cxRows.map(r => r.date))].sort((a, b) => a - b).map(fmtMon).join(", ")}` : "",
      allCx ? `Nothing booked in NextGen, last seen ${seenAt != null ? fmtMon(seenAt) : "never"}` : ""].filter(Boolean).join(". ");
  /* authorization (separate from protocol) */
  const authFlags = [];
  if (active) {
    const firstTx = rows.filter(r => r.date != null && r.type !== "F/U").map(r => r.date).sort((a, b) => a - b)[0];
    if (p.authStart != null && firstTx != null && firstTx < p.authStart) authFlags.push(`Auth starts ${fmtMon(p.authStart)}, after the first visit ${fmtMon(firstTx)}`);
    if (p.authExp == null) authFlags.push("No auth expiry entered");
    else {
      const days = p.authExp - today;
      if (days < 0) authFlags.push(`Auth expired ${fmtMon(p.authExp)}`);
      else if (days <= cfg.authDays) authFlags.push(`Auth ends ${fmtMon(p.authExp)}, in ${days} day${days === 1 ? "" : "s"}`);
      if (p.txEnd != null && p.txEnd > p.authExp) authFlags.push(`Treatments run to ${fmtMon(p.txEnd)}, past auth end`);
      if (sched.some(r => r.date > p.authExp)) authFlags.push("Visits booked past auth end");
    }
    const txBooked = rows.filter(r => r.type !== "F/U" && r.type !== "" && !r.miss && (r.delivered || r.date != null)).length;
    if (txBooked > total) authFlags.push(`${txBooked} booked or done, ${total} authorized (D2)`);
    if (!p.mtrAuth && sched.some(r => r.type === "MTR")) authFlags.push("MT redo (90869) booked, not authorized in D4 or MTR approved");
  }
  /* protocol: renewal cadence from G2 (business days) and Action = Protocol; entered at the MT */
  const protoFlags = [];
  if (active) {
    if (p.updateBy != null) { if (p.updateBy < today) protoFlags.push(`Protocol update was due ${fmtMon(p.updateBy)}`); else if (bizDaysUntil(today, p.updateBy) <= cfg.protoDays) protoFlags.push(`Protocol update due ${fmtMon(p.updateBy)}`); }
    const pa = rows.filter(r => r.action === "Protocol" && !r.delivered && r.date != null && r.date >= today).sort(cmp).find(r => bizDaysUntil(today, r.date) <= cfg.protoDays);
    if (pa) protoFlags.push(`Protocol action at the ${fmtMon(pa.date)} visit`);
  }
  /* no protocol row with a site in O: chosen at the MT, so asked for from the MT on (Confirm protocol) */
  const noProto = active && !siteRows(p).length && rows.some(r => r.type === "MT" && (r.delivered || r.date === today));
  /* patient cost so far and projected, the BIDF estimate and out-of-pocket max */
  /* TMS Follow Ups keep their own PR in column A but stay out of the course cost, the BIDF estimate and the OOPM tally */
  const accrued = rows.filter(r => happened(r, today) && r.type !== "F/U").reduce((s, r) => s + (payOf(p, r) || 0), 0);
  const futurePay = sched.filter(r => !r.alt && r.type !== "F/U").map(r => ({ r, v: payOf(p, r) || 0 }));
  const projected = accrued + futurePay.reduce((s, x) => s + x.v, 0);
  /* copays: the plan year is Last updated's year; logged copays are a floor (claims applied to the deductible count too) */
  const planYear = yearOf(p.benUpdated ?? today), owes = TmsCore.owes(p.ben);
  const tally = TmsCore.copayTally(rows, utcSerial(planYear, 1, 1), utcSerial(planYear, 12, 31));
  const toConfirm = owes ? TmsCore.copaysToConfirm(rows, r => payOf(p, r), today) : [];
  const todayRow = rows.find(r => r.date === today && !r.miss && !r.alt && r.type !== "F/U"), todayPay = todayRow && !TmsCore.copayOf(todayRow.note) ? payOf(p, todayRow) || 0 : 0;
  const oopmLikely = owes && p.ben.oopRem != null && p.ben.oopRem > 0 && tally.sum + todayPay >= p.ben.oopRem - 0.005;
  const oopmMet = p.ben.oopRem === 0 && !p.ben.oopNA ? planYear : null;
  const benFlags = [];
  if (active && owes) {
    if (p.ben.bidf != null && projected > p.ben.bidf + 0.005) benFlags.push(`Projected ${money(projected)}, over the BIDF estimate ${money(p.ben.bidf)}`);
    if (oopmLikely) benFlags.push("Out-of-pocket max likely met: set the copay to $0 at pre-check and update Insurance Maintenance in NextGen");
    else if (p.ben.oopRem != null) { let c = tally.sum; const hit = futurePay.find(x => (c += x.v) >= p.ben.oopRem); if (hit) benFlags.push(`Likely meets the out-of-pocket max around ${fmtMon(hit.r.date)}: verify the copay at pre-check`); }
  }
  const ro = p.benUpdated != null ? TmsCore.rollover({ updated: p.benUpdated, dec1: utcSerial(yearOf(p.benUpdated), 12, 1), jan1: utcSerial(yearOf(p.benUpdated) + 1, 1, 1) }, rows, today,
    r => r.date === today && !!state.sess && [...state.sess.rows.values()].some(s => s.tab === p.tab && s.date === today && s.checkin != null)) : null;
  const rollover = ro ? { ...ro, year: yearOf(p.benUpdated) + 1 } : null;
  const mdOnSite = /^y/i.test(p.mdReq);
  /* only dates the patient said they are away; weekday preferences are for scheduling, not flags */
  const awayConflicts = [];
  if (active) for (const r of sched) { const w = p.avail.find(w => w.type === "Unavailable" && w.from <= r.date && r.date <= w.to); if (w) awayConflicts.push({ r, why: `away ${fmtMon(w.from)} to ${fmtMon(w.to)}${w.note ? ` (${w.note})` : ""}` }); }
  const closedConflicts = sched.filter(r => state.closures.has(r.date)).map(r => ({ r, why: state.closures.get(r.date) }));
  const byDay = new Map(); rows.filter(r => r.date != null && r.date >= today && !r.miss && !r.alt).forEach(r => byDay.set(r.date, (byDay.get(r.date) || 0) + 1));
  const sameDay = p.kaiser ? [] : [...byDay.entries()].filter(([, n]) => n > 1).map(([d]) => d).sort((a, b) => a - b);
  const measDue = rows.filter(r => r.measX && r.date != null && r.date < today && !r.miss);
  const lastRow = rows.filter(r => r.type !== "F/U" && r.date != null).sort(cmp).pop();
  const bidf = bidfMissing(p), todayVisit = rows.some(r => r.date === today && !r.miss && !r.alt);
  return { active, sched, next, txDone, lastTx, total, overdue, noTech, decisions, used, waived, left, fu, fuNext, fuStatus, stale, suggest, cancelled, lastSeen: seenAt,
    authFlags, protoFlags, benFlags, accrued, projected, mdOnSite, sameDay, measDue, awayConflicts, closedConflicts, bidf, bidfToday: active && owes && todayVisit && bidf.length > 0,
    noProto, owes, planYear, tally, toConfirm, todayRow, oopmLikely, oopmMet, rollover,
    grad: p.txEnd ?? (lastRow ? lastRow.date : null) };
}
function availText(p) {
  const t = todaySerial(), parts = [];
  const un = p.avail.filter(w => w.type === "Unavailable" && w.to >= t).sort((a, b) => a.from - b.from)[0];
  if (un) parts.push((un.from <= t ? `Away until ${fmtDate(un.to)}` : `Away ${fmtDate(un.from)}\u2013${fmtDate(un.to)}`) + (un.note ? ` (${un.note})` : ""));
  const wkTxt = [1, 2, 3, 4, 5].map(d => { const w = p.week[d]; if (!w) return ""; const a = []; if (w.from != null) a.push(`from ${fmtTime(w.from)}`); if (w.to != null) a.push(`by ${fmtTime(w.to)}`); return a.length ? `${DAYS[d]} ${a.join(" ")}` : ""; }).filter(Boolean);
  if (wkTxt.length) parts.push("Weekly: " + wkTxt.join(", "));
  return parts.join("; ");
}
const trackerTitle = () => String(state.wbName || "").replace(/\.(xlsx|xlsm|xls)$/i, "").replace(/_/g, " ").replace(/\s{2,}/g, " ").trim() || "TMS Tracker";
/* ---------- banners: dismiss for the session, or snooze 3 hours ---------- */
function snoozed(id) {
  if (state.dismissed.has(id)) return true;
  let until = state.snoozeMem[id];
  try { until = until || JSON.parse(localStorage.getItem("tms_snooze") || "{}")[id]; } catch { /* storage unavailable */ }
  return !!until && Date.now() < until;
}
function snooze(id) {
  const until = Date.now() + SNOOZE_MS; state.snoozeMem[id] = until;
  try { const o = JSON.parse(localStorage.getItem("tms_snooze") || "{}"); o[id] = until; localStorage.setItem("tms_snooze", JSON.stringify(o)); } catch { /* memory only */ }
  renderBanners(); shareUi();
}
function dismiss(id) { state.dismissed.add(id); renderBanners(); shareUi(); }
/* full screen may not share the pane's storage, so dismissed and snoozed banners travel with the snapshot and both ways after it */
function uiSnap() {
  const sn = { ...state.snoozeMem };
  try { Object.entries(JSON.parse(localStorage.getItem("tms_snooze") || "{}")).forEach(([id, u]) => { if (!(sn[id] >= u)) sn[id] = u; }); } catch { /* memory only */ }
  return { dismissed: [...state.dismissed], snooze: sn };
}
function mergeUi(o) {
  if (!o) return;
  (o.dismissed || []).forEach(id => state.dismissed.add(id));
  Object.entries(o.snooze || {}).forEach(([id, u]) => { if (!(state.snoozeMem[id] >= u)) state.snoozeMem[id] = u; });
  try { const st = JSON.parse(localStorage.getItem("tms_snooze") || "{}"); Object.entries(state.snoozeMem).forEach(([id, u]) => { if (!(st[id] >= u)) st[id] = u; }); localStorage.setItem("tms_snooze", JSON.stringify(st)); } catch { /* memory only */ }
}
function shareUi() {
  const m = { k: "ui", ...uiSnap() };
  if (POPOUT) { try { Office.context.ui.messageParent(JSON.stringify(m)); } catch { /* pane gone */ } } else toChild(m);
}
const bctl = id => `<span class="bctl"><button class="sm" data-banner="${id}" data-bact="snooze">Snooze 3 h</button><button class="sm" data-banner="${id}" data-bact="dismiss" aria-label="Dismiss">Dismiss</button></span>`;
function renderBanners() {
  const t = todaySerial(), r = state.report, show = (id, on) => { $(id).hidden = !on || snoozed(id); };
  const stale = !r || r.asOf == null || r.asOf < t;
  $("reportMsg").textContent = !r ? "No appointment list has been uploaded yet. Run the report, then use Import."
    : r.asOf == null ? "The last appointment list has no run date. Run the report again and upload it on Import."
      : `The appointment list is from ${fmtMon(r.asOf)}${r.asOfMin != null ? " " + clock(r.asOfMin) : ""}. Run an updated report and upload it on Import.`;
  show("reportBanner", stale);
  $("wbTitle").textContent = trackerTitle(); $("devTag").textContent = `${device() === "Brainsway" ? "BrainsWay" : "MagV"} tracker`;
  const sk = state.skipped, parts = [];
  if (sk.layout.length) parts.push(`${sk.layout.length} tab${sk.layout.length > 1 ? "s" : ""} in an older layout (${sk.layout.map(esc).join(", ")})`);
  const ns = state.stale.length;
  $("cutoffNote").textContent = ns ? `${ns} tab${ns === 1 ? " is" : "s are"} left out of updates: no visit on or after this date, or Discontinued/Complete.` : "";
  $("tplNote").textContent = state.templateWhy; $("tplNote").hidden = state.templateOk;   // Settings keeps it where a dismissed banner can't hide it
  const setupUp = state.fix.size + state.old.length > 0;
  $("skipText").innerHTML = [!setupUp && state.templateWhy && esc(state.templateWhy), parts.length && `Left untouched: ${parts.join("; ")}. The add-in only reads and writes tabs that match the current template.`].filter(Boolean).join(" ");
  show("skipNote", parts.length > 0 || (!setupUp && !state.templateOk));
  const n = state.fix.size, o = state.old.length, msg = [];
  if (n) msg.push(`${n} tab${n === 1 ? " doesn't" : "s don't"} match New Template yet (${[...state.fix.keys()].map(esc).join(", ")}). Match New Template moves benefits to the ${BEN_SHEET} tab and freebies to N6, and brings the PR formula, colors by visit type, dropdowns, borders and the TX SUMMARY checkbox over every log row`);
  if (o) msg.push(`${o} tab${o === 1 ? " is" : "s are"} in the earlier layout (${state.old.map(esc).join(", ")}) and will be rebuilt from New Template. The earlier tab is deleted once its entries are copied over`);
  $("setupMsg").innerHTML = msg.join(". ") + "." + (state.templateWhy ? " " + esc(state.templateWhy) : "");
  show("setup", n + o > 0);
  $("newPtList").innerHTML = state.pendingNew.map(newPtItem).join("");
  $("newPtCount").textContent = state.pendingNew.length;
  show("newPtBanner", state.pendingNew.length > 0);
  const nm = noMrn(); $("mrnCount").textContent = nm.length;
  $("mrnList").innerHTML = nm.map(p => { const s = (state.codeMrn && state.codeMrn.get(p.code)) || ""; return `<div class="nrow"><span><b>${esc(p.tab)}</b> ${esc(p.status || "no status")}${s ? ", MRN from NextGen filled in" : ""}</span><input class="sm" data-mrnin="${esc(p.tab)}" value="${esc(s)}" placeholder="MRN" aria-label="MRN for ${esc(p.tab)}" inputmode="numeric" autocomplete="off" style="width:110px"><button class="btn ghost sm" data-act="saveMrn" data-t="${esc(p.tab)}">Save to A4</button></div>`; }).join("");
  show("mrnBanner", nm.length > 0);
  const ak = state.asks || []; $("askCount").textContent = ak.length;
  $("askList").innerHTML = ak.map(g => `<div class="nrow"><span><b>${esc(g.code || "New patient")}</b> MRN ending ${esc(String(g.raw).slice(-4))}. MT at ${esc(g.home)}, ${g.daily} daily treatments booked here${g.next != null ? `, next ${fmtMon(g.next)}` : ""}.</span><button class="btn ghost sm" data-act="newTracker" data-v="${esc(g.mrn)}">Create tracker</button><button class="btn mute sm" data-act="useTab" data-v="${esc(g.mrn)}">Use existing tab</button><button class="btn mute sm" data-act="dontAsk" data-v="${esc(g.mrn)}">Don't ask again</button></div>`).join("");
  show("askBanner", ak.length > 0);
  const gl = (state.guests || []).filter(g => !g.link); $("guestCount").textContent = gl.length;
  $("guestList").innerHTML = gl.map(g => `<div class="nrow"><span><b>${esc(g.code || "Guest")}</b> MRN ending ${esc(String(g.raw).slice(-4))}. ${g.home ? `Home: ${esc(g.home)} (${g.homeCount} chair visit${g.homeCount === 1 ? "" : "s"})` : "No MT here and under 16 treatments, home clinic not on the list"}, ${g.here} here${g.next != null ? `, next here ${fmtMon(g.next)}` : ""}.</span><input class="sm" data-guestin="${esc(g.mrn)}" placeholder="https://… link to their tab" aria-label="Link to ${esc(g.code || "the guest")}'s tab at ${esc(g.home || "their home clinic")}" autocomplete="off" style="width:220px"><button class="btn ghost sm" data-act="saveGuest" data-v="${esc(g.mrn)}">Save link</button><button class="btn mute sm" data-act="newTracker" data-v="${esc(g.mrn)}">Track here instead</button><button class="btn mute sm" data-act="useTab" data-v="${esc(g.mrn)}">Use existing tab</button></div>`).join("");
  show("guestBanner", gl.length > 0);
  fillTplPick();
  if ($("cfgLoc")) $("cfgLoc").placeholder = TmsCore.clinicFromName(state.wbName) ? `From the workbook name: ${TmsCore.clinicFromName(state.wbName)}` : "Type the clinic, for example Castro";
}
/* tabs still in progress, with a visit in the recent-visit window or later, whose MRN (A4) is blank */
const noMrn = () => { const t = todaySerial(); return state.patients.filter(p => !p.mrnRaw && (p.status === "" || ACTIVE.includes(p.status)) && !state.stale.includes(p.tab) && p.rows.some(r => r.date != null && r.date >= t - state.cfg.recentDays)); };
const newPtItem = x => `<div class="nrow"><span><b>${esc(x.code || "New patient")}</b> MRN ending ${esc(x.last4)}. ${x.n} visit${x.n > 1 ? "s" : ""} from ${fmtMon(x.first.date)} (${esc(x.first.type)}), no tracker tab.</span><button class="btn ghost sm" data-act="newTracker" data-v="${esc(x.mrn)}">Create tracker</button><button class="btn mute sm" data-act="useTab" data-v="${esc(x.mrn)}">Use existing tab</button></div>`;

/* ---------- v2 screens: Today (the chair board), Schedule, Patients and Chairs, on the design language ESK Tracker shares
   (tokens and components live in taskpane.html). The reads and writes are the same as before; core.js holds the floor rules.
   On these screens times read 24-hour (9:05) and dates "Oct 21". ---------- */
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const LONG_DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const LONG_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const clock = m => m == null ? "" : fmtTime(Math.round(m));
const fmtMon = s => { if (typeof s !== "number") return ""; const d = serialToDate(s); return `${MONTHS[d.getUTCMonth()]} ${d.getUTCDate()}`; };
const fmtDay = s => `${DAYS[dow(s)]} ${fmtMon(s)}`;
const fmtLong = s => { const d = serialToDate(s); return `${LONG_DAYS[d.getUTCDay()]}, ${LONG_MONTHS[d.getUTCMonth()]} ${d.getUTCDate()}`; };
const nowMin = () => { const d = new Date(); return d.getHours() * 60 + d.getMinutes(); };
const btn = (act, tab, val, label) => `<button class="btn ghost sm" data-act="${act}" data-t="${esc(tab)}"${val ? ` data-v="${esc(val)}"` : ""}>${esc(label)}</button>`;
const bt = (act, tab, row, val, label, cls) => `<button class="btn ${cls === "pri" ? "primary" : "ghost sm"}" data-act="${act}" data-t="${esc(tab)}" data-row="${row}"${val != null && val !== "" ? ` data-v="${esc(val)}"` : ""}>${esc(label)}</button>`;
const NO_TABS = `<p class="empty">No patient tabs found. A patient tab has "PATIENT:" in A1 and a LASFIR name such as SMIJAN or SMIJAN 2.</p>`;
const DEFAULT_LEN = 20;   // session minutes when the tab has no protocol time in N2:N4

/* ---------- Sessions sheet: check-in, start and end times and chair moves for each visit (A:J), and the chairs (L:P).
   One row per visit, keyed by date, tab and TMS or F/U (core.js sessKey). The add-in owns the sheet; rows older than
   SESS_KEEP days are dropped when it fills up. Nothing about sessions is written to the patient tabs. ---------- */
const SESS_SHEET = "Sessions", SESS_FIRST = 2, SESS_LAST = 2000, SESS_KEEP = 45, CH_FIRST = 2, CH_LAST = 21;
const SESS_HDR = ["Date", "Tab", "Visit", "Chair", "NextGen chair", "Moved in NextGen", "Moved at", "Checked in", "Started", "Ended"];
const SESS_COLS = { chair: 3, ng: 4, tick: 5, at: 6, checkin: 7, start: 8, end: 9 };   // 0-based in A:J
const CH_HDR = ["Chair", "Device", "Coil", "Last service", "Unavailable"], CH_COLS = { device: 12, coil: 13, service: 14, off: 15 };   // 0-based: the name is in L (11), these in M:P
const TIME_KEYS = ["at", "checkin", "start", "end"];
const sessTime = x => typeof x === "number" ? Math.round((x % 1) * 1440) : null;
function readSess(vals, cv) {
  const rows = new Map(), chairsM = new Map(); let next = SESS_FIRST, cnext = CH_FIRST;
  (vals || []).forEach((v, i) => {
    const tab = String(v[1] ?? "").trim(); if (!isNum(v[0]) || !tab) return;
    const row = SESS_FIRST + i, date = Math.floor(v[0]), visit = String(v[2] ?? "").trim() || "TMS"; next = Math.max(next, row + 1);
    rows.set(TmsCore.sessKey(date, tab, visit), { row, date, tab, visit, chair: String(v[3] ?? "").trim(), ng: String(v[4] ?? "").trim(),
      tick: v[5] === true || /^(true|yes)$/i.test(String(v[5]).trim()), at: sessTime(v[6]), checkin: sessTime(v[7]), start: sessTime(v[8]), end: sessTime(v[9]) });
  });
  (cv || []).forEach((v, i) => {
    const name = String(v[0] ?? "").trim(); if (!name) return;
    const row = CH_FIRST + i; cnext = Math.max(cnext, row + 1);
    chairsM.set(name, { row, device: String(v[1] ?? "").trim(), coil: String(v[2] ?? "").trim(), service: isNum(v[3]) ? Math.floor(v[3]) : null, off: v[4] === true || /^(yes|true)$/i.test(String(v[4]).trim()) });
  });
  state.sess = { rows, next, chairs: chairsM, cnext };
}
const sessTimes = ws => { ws.getRange(`G${SESS_FIRST}:J${SESS_LAST}`).numberFormat = Array(SESS_LAST - SESS_FIRST + 1).fill(Array(4).fill("h:mm AM/PM")); };
/* a Sessions sheet made before times read AM/PM gets the format once a session, on its own so a failure can't hold up a check-in */
let sessAmPm = false;
async function sessAmPmOnce() {
  if (POPOUT || sessAmPm) return;
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(SESS_SHEET); await ctx.sync();
      if (!ws.isNullObject) {
        const r = ws.getRange(`G${SESS_FIRST}:J${SESS_FIRST}`); r.load("numberFormat"); await ctx.sync();
        if (r.numberFormat[0].some(f => !/AM\/PM/i.test(f))) { sessTimes(ws); await ctx.sync(); }
      }
      sessAmPm = true;
    });
  } catch (e) { console.error(e); }
}
async function sessSheet(ctx) {
  let ws = ctx.workbook.worksheets.getItemOrNullObject(SESS_SHEET); await ctx.sync();
  if (!ws.isNullObject) return ws;
  ws = ctx.workbook.worksheets.add(SESS_SHEET);
  const h = ws.getRange("A1:J1"), c = ws.getRange("L1:P1"); h.values = [SESS_HDR]; c.values = [CH_HDR];
  [h, c].forEach(r => { r.format.font.bold = true; r.format.fill.color = "#1c2a33"; r.format.font.color = "#ffffff"; });
  ws.getRange(`A${SESS_FIRST}:A${SESS_LAST}`).numberFormat = Array(SESS_LAST - SESS_FIRST + 1).fill(["m/d/yyyy"]);
  sessTimes(ws);
  ws.getRange(`O${CH_FIRST}:O${CH_LAST}`).numberFormat = Array(CH_LAST - CH_FIRST + 1).fill(["m/d/yyyy"]);
  ws.getRange("A:P").format.columnWidth = 78; ws.getRange("K:K").format.columnWidth = 16;
  ws.getRange("A2").values = [[""]]; ws.freezePanes.freezeRows(1);
  if (!state.sess) readSess(null, null);
  return ws;
}
/* entries: { key, date, tab, visit, set: { chair, ng, tick, at, checkin, start, end } } for a visit, or { chairId, set: { device, coil, service, off } }.
   In full screen the pane does the write and the window keeps its own copy in step. */
function sessApply(entries, put) {
  const S = state.sess || (readSess(null, null), state.sess);
  for (const e of entries) {
    if (e.chairId) {
      let r = S.chairs.get(e.chairId);
      if (!r) { r = { row: S.cnext++, device: "", coil: "", service: null, off: false }; S.chairs.set(e.chairId, r); if (put) put(r.row, 11, e.chairId); }
      for (const [k, v] of Object.entries(e.set)) { r[k] = v; if (put) put(r.row, CH_COLS[k], k === "off" ? (v ? "Yes" : "") : v ?? ""); }
      continue;
    }
    let r = S.rows.get(e.key);
    if (!r) { r = { row: S.next++, date: e.date, tab: e.tab, visit: e.visit, chair: "", ng: "", tick: false, at: null, checkin: null, start: null, end: null }; S.rows.set(e.key, r); }
    if (put) { put(r.row, 0, r.date); put(r.row, 1, r.tab); put(r.row, 2, r.visit); }
    for (const [k, v] of Object.entries(e.set)) { r[k] = v; if (put) put(r.row, SESS_COLS[k], TIME_KEYS.includes(k) ? (v == null ? "" : v / 1440) : k === "tick" ? !!v : v ?? ""); }
  }
}
/* Several people can have the workbook open, so every write reads the sheet again first and adds rows after whatever is there
   now. After a failed save the copy in memory is read back from the sheet. Returns the sheet as the pane now has it. */
const sessSnap = () => state.sess && { ...state.sess, rows: [...state.sess.rows], chairs: [...state.sess.chairs] };
const setSessSnap = o => { state.sess = o ? { ...o, rows: new Map(o.rows), chairs: new Map(o.chairs) } : null; };
async function readSessNow(ctx) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(SESS_SHEET); await ctx.sync();
  if (ws.isNullObject) { readSess(null, null); return null; }
  const a = ws.getRange(`A${SESS_FIRST}:J${SESS_LAST}`), c = ws.getRange(`L${CH_FIRST}:P${CH_LAST}`); a.load("values"); c.load("values"); await ctx.sync();
  readSess(a.values, c.values); return ws;
}
let sessBusy = 0;
async function writeSess(entries) {
  if (!entries.length) return;
  sessBusy++;
  try { if (POPOUT) await callPane("writeSess", [entries]); else await writeSessPane(entries); } finally { sessBusy--; }
}
/* the pane's write; in full screen the pane then sends the table to the window in chunks (sendSess) before it answers */
async function writeSessPane(entries) {
  try {
    await Excel.run(async ctx => {
      await sessSheet(ctx); const ws = await readSessNow(ctx);
      const newChairs = new Set(entries.filter(e => e.chairId && !state.sess.chairs.has(e.chairId)).map(e => e.chairId)).size;
      if (state.sess.cnext + newChairs - 1 > CH_LAST) throw new Error(`The ${SESS_SHEET} sheet has room for ${CH_LAST - CH_FIRST + 1} chairs.`);
      if (state.sess.next + entries.length > SESS_LAST) await compactSess(ctx, ws);
      sessApply(entries, (row, col, v) => { ws.getRangeByIndexes(row - 1, col, 1, 1).values = [[v]]; });
      await ctx.sync();
    });
  } catch (e) { try { await Excel.run(readSessNow); } catch { /* keep the error that matters */ } throw e; }
}
/* the board's 30-second tick picks up other people's check-ins and moves, unless a save is still going */
async function pullSess() {
  if (sessBusy) return;
  if (POPOUT) return callPane("sessSnap", []);
  await Excel.run(readSessNow);
}
/* keeps the last SESS_KEEP days, rewritten from row 2; the copy in memory follows once the sheet has it */
async function compactSess(ctx, ws) {
  const S = state.sess, keep = [...S.rows.values()].filter(r => r.date >= todaySerial() - SESS_KEEP).sort((a, b) => a.date - b.date || a.row - b.row);
  ws.getRange(`A${SESS_FIRST}:J${SESS_LAST}`).clear("Contents");
  if (keep.length) ws.getRange(`A${SESS_FIRST}:J${SESS_FIRST + keep.length - 1}`).values = keep.map(r => [r.date, r.tab, r.visit, r.chair, r.ng, !!r.tick, ...TIME_KEYS.map(k => r[k] == null ? "" : r[k] / 1440)]);
  await ctx.sync();
  S.rows = new Map(keep.map((r, i) => [TmsCore.sessKey(r.date, r.tab, r.visit), { ...r, row: SESS_FIRST + i }])); S.next = SESS_FIRST + keep.length;
}
const sessOf = v => ({ key: v.key, date: v.date, tab: v.tab, visit: v.visit });
async function sessWrite(v, set, msg) {
  try { failStep = "Saving to the Sessions sheet"; await writeSess([{ ...sessOf(v), set }]); failStep = ""; renderDash(); if (msg) toast(msg); return true; } catch (e) { fail(e); return false; }
}

/* ---------- the floor: chairs and today's visits ---------- */
/* chairs on this tracker: the saved appointment list's resources (Castro MagV1, Castro MagV2), else the device's one chair, plus any on the Sessions sheet */
function chairList() {
  const ids = [...new Set([...(chairs().length ? chairs() : [chairName()]), ...(state.sess ? [...state.sess.chairs.keys()] : [])])];
  return ids.map(id => { const m = (state.sess && state.sess.chairs.get(id)) || {};
    return { id, label: TmsCore.chairLabel(id, clinic()), device: m.device || (device() === "Brainsway" ? "BrainsWay Deep TMS" : "MagVenture"), coil: m.coil || "", service: m.service ?? null, off: !!m.off }; });
}
const chairLbl = id => id ? TmsCore.chairLabel(id, clinic()) : "No chair";
/* every visit on a tab today, with its chair (moved, NextGen's, or none while its chair is unavailable) and its Sessions times */
function floor() {
  const t = todaySerial(), now = nowMin(), cl = chairList(), on = new Set(cl.filter(c => !c.off).map(c => c.id)), single = cl.length === 1 ? cl[0].id : "";
  const rep = state.report, listAt = rep && rep.asOf === t ? rep.asOfMin : null, visits = [];
  for (const p of state.patients.filter(onSched)) {
    for (const { r, visit } of TmsCore.visitNames(p.rows.filter(r => r.date === t && (r.type || r.delivered || r.miss)))) {
      const fu = r.type === "F/U", key = TmsCore.sessKey(t, p.tab, visit), s = (state.sess && state.sess.rows.get(key)) || null;
      const rc = fu || r.alt ? "" : chairOf(p, r) || single, ms = TmsCore.moveState(s, rc, listAt);
      const v = { key, date: t, visit, p, r, tab: p.tab, row: r.row, type: r.type || "Daily", time: r.time, tele: r.tele, alt: r.alt, chairVisit: !fu && !r.alt, s, rc,
        chair: fu || r.alt ? "" : ms.chair, moved: ms.moved, settled: ms.settled, from: ms.from, offChair: "", miss: r.miss,
        done: fu ? !!(s && s.end != null) : r.delivered, checkin: s ? s.checkin : null, start: s ? s.start : null, end: s ? s.end : null,
        len: Math.round(protoLen(p) || DEFAULT_LEN), protoLen: protoLen(p), mtDone: /\bmt %:/i.test(r.note) };
      if (v.chairVisit && v.chair && !on.has(v.chair) && ["booked", "waiting"].includes(TmsCore.phase(v))) { v.offChair = v.chair; v.chair = ""; v.moved = false; }
      visits.push(v);
    }
  }
  visits.sort((a, b) => (a.time ?? 1e4) - (b.time ?? 1e4) || a.tab.localeCompare(b.tab));
  return { t, now, chairs: cl, avail: cl.filter(c => !c.off), visits };
}
const ph = v => TmsCore.phase(v);
const findVisit = k => (state.fl || (state.fl = floor())).visits.find(v => v.key === k);
/* pulses on the tab with no NextGen visit that day, against the whole saved list (latest course only; past Expected visits count) */
const noApptRows = p => { const a = state.apptAll && p.mrn && state.apptAll.get(p.mrn); return a && !state.patients.some(q => q.mrn === p.mrn && q.course > p.course) ? TmsCore.pulsesNoAppt(p.rows, a, state.apptLo, todaySerial()) : []; };
function courseNeeds(p, fl) {
  const vs = fl.visits.filter(v => v.tab === p.tab), mt = vs.filter(v => v.type === "MT" && !v.miss);
  return TmsCore.needs(p, { today: fl.t, fmtDate: fmtMon, recentDays: state.cfg.recentDays, exemptTag: EXEMPT_TAG, unclassTag: UNCLASS_TAG,
    moved: vs.filter(v => v.moved).map(v => ({ key: v.key, chairLabel: chairLbl(v.chair) })).concat(state.sess ? [...state.sess.rows].filter(([, r]) => r.tab === p.tab && r.date < fl.t && TmsCore.moveState(r, "", null).moved)
      .map(([key, r]) => ({ key, chairLabel: chairLbl(r.chair), dayLabel: fmtMon(r.date) })) : []),
    noChair: vs.filter(v => v.chairVisit && !v.chair && ["booked", "waiting"].includes(ph(v))).map(v => ({ key: v.key, timeLabel: clock(v.time) || "today" })),
    noAppt: noApptRows(p),
    mtToday: mt.length > 0, schedGiven: p.ben.lastSched != null && mt.length > 0 && p.ben.lastSched >= fl.t });
}
const sessNo = v => typeof v.r.tx === "number" ? v.r.tx : v.p.d.txDone + (v.done ? 0 : 1);
const typeName = v => ({ Daily: "Daily", Taper: "Taper", MT: "MT (first session)", MTR: "MT redo", "F/U": "Follow-up" })[v.type] || v.type;
/* guests (and patients offered a tracker): no tab here, so their chair bookings come from the saved list. They show on Today and
   Schedule for the chair's sake; pulses, technician and measures go on their home clinic's tab, through the link. */
const guestOf = mrn => [...(state.asks || []), ...(state.guests || [])].find(g => g.mrn === mrn);
const guestVisits = day => ((state.lx && state.lx.gv) || []).filter(x => x.date === day).map(x => ({ ...x, g: guestOf(x.mrn) })).filter(x => x.g).sort((a, b) => (a.time ?? 1e4) - (b.time ?? 1e4));
const guestLink = g => g.link ? `<a class="btn ghost sm" href="${esc(g.link)}" target="_blank" rel="noopener">Open their tab</a>` : '<button class="btn ghost sm" data-act="guests">Add their link</button>';
const guestRowToday = x => `<div class="row"><div class="rowmain static"><span class="p">${esc(x.code || "Guest")}</span><span class="w">${clock(x.time) || "–"}</span><span class="muted">MRN: ${esc(String(x.g.raw).replace(/\D/g, "").slice(-4))}</span><span class="f"><span class="muted">Guest from ${esc(x.g.home)}, ${esc(chairLbl(x.chair))}: pulses go on their ${esc(x.g.home)} tab</span></span></div>${guestLink(x.g)}</div>`;
/* Today names a patient by LASFIR with the MRN's last four beside it, in grey */
const mrnTag = p => p.mrnRaw ? `MRN: ${String(p.mrnRaw).replace(/\D/g, "").slice(-4) || p.mrnRaw.slice(-4)}` : "No MRN";
const where = v => v.alt ? "Other clinic" : v.type === "F/U" ? (v.tele ? "Telehealth" : v.p.provider || "Provider") : v.chair ? chairLbl(v.chair) : "";
const runInfo = v => { const e = v.start + v.len; return { s: v.start, e, left: Math.ceil(e - state.fl.now) || 0 }; };

/* ---------- shared pieces ---------- */
const lantern = on => `<svg class="lan" viewBox="-13 -13 26 26" aria-hidden="true">${on ? '<circle r="13" fill="url(#glow)"/>' : ""}<g class="lantern${on ? " on" : ""}"><circle class="body" r="7.5"/><circle class="lens" r="4.75"/><circle class="touch" r="1.4"/></g></svg>`;
const mk = st => `<span class="k">${st === "done" ? "✓" : ""}</span>`;
function strip(p, todayHere) {
  const tot = Math.max(1, p.d.total), taper = tot >= 12 ? 6 : 0, daily = tot - taper, done = Math.min(p.d.txDone, tot);
  const cols = taper ? `repeat(${daily},1fr) 10px repeat(${taper},1fr)` : `repeat(${tot},1fr)`;
  let cells = "";
  for (let i = 1; i <= tot; i++) { if (taper && i === daily + 1) cells += "<span></span>"; cells += `<span class="c${i <= done ? " done" : i === done + 1 && todayHere ? " today" : ""}"></span>`; }
  const lab = [1, 10, 20].filter(x => x < daily).map(x => `<span style="grid-column:${x}">${x}</span>`).join("");
  return `<div class="strip" style="grid-template-columns:${cols}" role="img" aria-label="${done} of ${tot} treatments done">${cells}</div>
    <div class="stripk" style="grid-template-columns:${cols}" aria-hidden="true">${lab}${taper ? `<span style="grid-column:${daily} / span 2">${daily}</span><span style="grid-column:${daily + 2} / span ${taper}">Taper</span>` : `<span style="grid-column:${tot}">${tot}</span>`}</div>`;
}
/* PHQ-9 and GAD-7 as recorded on the log, latest four, with "Due" when today's visit still has the x */
function scores(p, dueToday) {
  const got = p.rows.filter(r => r.measVal && r.date != null).sort(cmp).slice(-4);
  const row = (lbl, k) => `<div class="phq"><span class="lbl">${lbl}</span>${got.filter(r => r[k] && !/^x$/i.test(r[k])).map(r => `<div><b>${esc(r[k])}</b><small>${fmtMon(r.date)}</small></div>`).join("")}${dueToday ? '<div class="due"><b>Due</b><small>Today</small></div>' : ""}</div>`;
  return got.length || dueToday ? row("PHQ-9", "phq") + row("GAD-7", "gad") : '<p class="body">No scores recorded yet.</p>';
}
function facts(p, short) {
  const d = p.d, b = p.ben, left = Math.max(0, d.total - d.txDone), t = todaySerial();
  const amt = (na, v) => na ? "N/A" : v != null ? money(v) : null;
  const rows = [["Protocol", protoText(p)],
    ["Authorization", `${left} of ${d.total} left${p.authExp != null ? `, ${p.authExp < t ? "expired" : "ends"} ${fmtMon(p.authExp)}` : ", no end date"}`, p.authExp == null || p.authExp < t ? "bad" : p.authExp - t <= state.cfg.authDays ? "flag" : ""],
    ["Freebies", `${d.left} left`, d.left <= 1 ? "flag" : ""]];
  if (!short) rows.push(["Insurance", p.insurance || "Not entered"],
    ["Deductible left", amt(b.dedNA, b.dedRem) ?? "Not entered", b.dedBlank ? "flag" : ""], ["Out-of-pocket left", amt(b.oopNA, b.oopRem) ?? "Not entered", !b.oopNA && b.oopRem == null ? "flag" : ""],
    ["BIDF estimate", b.bidf != null ? `${money(b.bidf)}, projected ${money(d.projected) || "$0.00"}` : "Not entered", b.bidf != null && d.projected > b.bidf + 0.005 ? "flag" : ""],
    ["Cost so far", money(d.accrued) || "$0.00"],
    ["Next visit", d.next ? `${fmtDay(d.next.date)}, ${clock(d.next.time) || "no time"}` : "None booked"],
    ["Next follow-up", d.fuNext ? `${fmtDay(d.fuNext.date)}${d.fuNext.tele ? ", telehealth" : ""}` : "Not scheduled", d.fuStatus === "not scheduled" ? "flag" : ""],
    ["Graduation", d.grad != null ? fmtMon(d.grad) : ""], ["Schedule given", b.lastSched != null ? fmtMon(b.lastSched) : "Not yet"],
    ["Provider", p.provider], ["Care navigator", p.tas], ["MD on-site", d.mdOnSite ? "Required" : "No"], ["Home clinic", p.home],
    ["MT redos", p.mtRedos ? `${p.mtRedos}${p.mtrAuth ? ", approved" : ""}` : ""], ["Auth assessment", p.authAssess], ["Progress note sent", p.progSent != null ? fmtMon(p.progSent) : ""]);
  return `<dl class="facts">${rows.filter(([, v]) => v !== "" && v != null).map(([k, v, cls]) => `<dt>${k}</dt><dd class="${cls || ""}">${esc(v)}</dd>`).join("")}</dl>`;
}

/* ---------- rail and header ---------- */
function renderRail(fl) {
  document.querySelectorAll(".rail [data-view]").forEach(b => { if (b.dataset.view === state.view || (state.view === "add" && b.dataset.view === "patients")) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current"); });
  const need = state.patients.filter(onSched).map(p => TmsCore.loud(courseNeeds(p, fl))).filter(n => n.length), cc = $("ccount");
  cc.textContent = need.length || ""; cc.className = "count" + (need.some(n => n.some(x => x.bad)) ? " bad" : "");
  cc.setAttribute("aria-label", need.length ? `${need.length} need attention` : "");
}
const head = (title, sub, withClock) => `<header class="top"><div><h1>${title}</h1><p class="sub">${sub}</p></div>${withClock ? `<div class="clock"><div class="t">${clock(state.fl.now)}</div><div class="d">${fmtLong(state.fl.t)}</div></div>` : ""}</header>`;

function renderDash() {
  const fl = state.fl = floor();
  renderRail(fl);
  if (!["today", "schedule", "patients", "chairs"].includes(state.view)) return;
  const el = $("v-" + state.view);
  if (!state.patients.length && state.view !== "chairs") { el.innerHTML = head({ today: "Today", schedule: "Schedule", patients: "Patients" }[state.view], "") + `<div class="card list">${NO_TABS}</div>`; return; }
  el.innerHTML = state.view === "today" ? viewToday(fl) : state.view === "schedule" ? viewSchedule(fl) : state.view === "patients" ? viewPatients(fl) : viewChairs(fl);
  renderPt();
  refocus();
}

/* ---------- Today ---------- */
/* Clinic, or one chair (saved per computer, like the old chair picker) */
function todayMode(fl) { const m = state.chair; return m !== "all" && fl.avail.some(c => c.id === m) ? m : "clinic"; }
function toCome(fl, shown) { return fl.visits.filter(v => ["booked", "waiting"].includes(ph(v)) && !shown.has(v.key)); }
function rowFlag(v, fl) {
  if (v.alt) return `<span class="muted">At ${esc(v.p.home || "another clinic")} today</span>`;
  if (v.chairVisit && ph(v) === "waiting") { const w = fl.now - (v.time ?? fl.now), late = TmsCore.isLate(v, fl.now);
    return `<span class="${late ? "bad" : "flag"}">${late ? `Waiting ${w} min` : `Checked in ${clock(v.checkin)}`}, ${v.chair ? `${esc(chairLbl(v.chair))} in use` : "no chair yet"}</span>`; }
  if (v.chairVisit && !v.chair) return '<span class="flag">No chair yet</span>';
  if (ph(v) === "booked" && v.time != null && v.time + 5 < fl.now && !v.tele) return `<span class="flag">${v.type === "F/U" ? "Not seen" : "Not started"}, ${fl.now - v.time} min</span>`;
  if (v.moved) return '<span class="flag">Moved here, not yet in NextGen</span>';
  const n = TmsCore.loud(courseNeeds(v.p, fl))[0];
  return n ? `<span class="${n.bad ? "bad" : "flag"}">${esc(n.t)}</span>` : `<span class="muted">${v.type === "F/U" ? "Follow-up" : `${typeName(v)}, ${sessNo(v)} of ${v.p.d.total}`}</span>`;
}
function rowAction(v) {
  if (v.alt) return "";
  if (v.chairVisit && !v.chair) return `<button class="btn ghost sm" data-act="move" data-k="${esc(v.key)}">Choose a chair</button>`;
  if (v.chairVisit && ["booked", "waiting"].includes(ph(v))) return `<button class="btn ghost sm" data-act="start" data-k="${esc(v.key)}" aria-label="Start ${esc(v.tab)}'s session">Start session</button>`;
  return "";
}
function viewToday(fl) {
  if (state.chair !== "all" && todayMode(fl) === "clinic") { state.chair = "all"; try { localStorage.setItem("tms_chair", "all"); } catch { /* this session only */ } }
  const mode = todayMode(fl), states = fl.avail.map(c => ({ c, s: TmsCore.chairState(c.id, fl.visits, fl.now) }));
  if (!state.sel || (state.sel.k && !findVisit(state.sel.k)) || (state.sel.chair && !fl.avail.some(c => c.id === state.sel.chair))) {
    const order = { late: 0, soon: 1, track: 2, open: 3 }, top = states.slice().sort((a, b) => order[a.s.state] - order[b.s.state])[0];
    const first = fl.visits.find(v => ["waiting", "booked", "running"].includes(ph(v)));
    state.sel = top && top.s.v ? { k: top.s.v.key } : first ? { k: first.key } : top ? { chair: top.c.id } : null;
  }
  const seg = `<div class="viewbar"><span class="lbl">View</span><div class="seg" role="group" aria-label="Choose a view"><button data-mode="all" aria-pressed="${mode === "clinic"}">Clinic</button>${fl.avail.length > 1 ? '<span class="sep"></span>' + fl.avail.map(c => `<button data-mode="${esc(c.id)}" aria-pressed="${mode === c.id}">${esc(c.label)}</button>`).join("") : ""}</div></div>`;
  const off = fl.chairs.filter(c => c.off).map(c => c.label), closed = state.closures.get(fl.t);
  let left, sub;
  if (mode === "clinic") {
    const shown = new Set(states.filter(x => x.s.v).map(x => x.s.v.key)), rest = toCome(fl, shown);
    const running = states.filter(x => x.s.state === "track" || x.s.state === "soon").length, late = fl.visits.filter(v => v.chairVisit && TmsCore.isLate(v, fl.now)).length;
    sub = [`${running} in session`, late ? `<strong>${late} waiting to start</strong>` : "", `${rest.length} still to come`, off.length ? `${esc(off.join(", "))} unavailable` : ""].filter(Boolean).join(", ");
    const board = states.map(({ c, s }) => {
      const running = s.state === "track" || s.state === "soon", sel = s.v ? state.sel && state.sel.k === s.v.key : state.sel && state.sel.chair === c.id;
      const bar = running ? `<span class="sbar"><i style="width:${(Math.min(1, Math.max(0, (fl.now - s.v.start) / s.v.len)) * 100).toFixed(1)}%"></i></span>` : '<span class="ph"></span>';
      const who = s.v ? `<b>${esc(s.v.tab)}</b><small>${running ? `${typeName(s.v)}, ${sessNo(s.v)} of ${s.v.p.d.total}` : `In the lobby, booked ${clock(s.v.time) || "no time"}`}</small>`
        : `<b>Open</b><small>${s.next ? `${esc(s.next.tab)} at ${clock(s.next.time) || "no time"}${s.next.time != null && s.next.time < fl.now ? ", not started" : ""}` : "Nothing else booked today"}</small>`;
      const st = TmsCore.chairStatus(s), aria = `${c.label}, ${s.v ? `${s.v.tab}, ${st.replace(/\bmin\b/, "minutes")}` : `open${s.next ? `, ${s.next.tab} at ${clock(s.next.time)}` : ""}`}`;
      return `<button class="crw" ${s.v ? `data-k="${esc(s.v.key)}"` : `data-chair="${esc(c.id)}"`} data-state="${s.state}" aria-pressed="${!!sel}" aria-label="${esc(aria)}">${lantern(running)}<span><b>${esc(c.label)}</b><small class="coil">${esc(c.coil || c.device)}</small></span><span class="who${s.v ? "" : " open"}">${who}</span>${bar}<span class="st">${st}</span></button>`;
    }).join("") || '<p class="empty">Every chair is marked unavailable. Mark one available on Chairs.</p>';
    const rows = rest.map(v => `<div class="row"><button class="rowmain" data-k="${esc(v.key)}" aria-pressed="${!!(state.sel && state.sel.k === v.key)}"><span class="p">${esc(v.tab)}</span><span class="w">${clock(v.time) || "–"}</span><span class="muted">${esc(mrnTag(v.p))}</span><span class="f">${rowFlag(v, fl)}</span></button>${rowAction(v)}</div>`).join("");
    const gs = guestVisits(fl.t).filter(x => x.time == null || x.time >= fl.now - 30), grows = gs.map(guestRowToday).join(""), nleft = rest.length + gs.length;
    left = `<div class="stack"><div class="card board" role="group" aria-label="Chairs">${board}</div>
      <div class="card list"><div class="h"><h2>Still to come</h2><span>${nleft ? `${nleft} left today` : ""}</span></div>${rows + grows || '<p class="empty">Nobody else is booked today.</p>'}</div>
      ${listCheckCard(fl.t)}${varianceCard(fl.t)}</div>`;
  } else {
    const c = fl.avail.find(x => x.id === mode), s = TmsCore.chairState(c.id, fl.visits, fl.now), vs = fl.visits.filter(v => v.chairVisit && v.chair === c.id);
    if (!state.sel || !state.sel.k || !vs.some(v => v.key === state.sel.k)) state.sel = s.v ? { k: s.v.key } : vs.find(v => ["booked", "waiting"].includes(ph(v))) ? { k: vs.find(v => ["booked", "waiting"].includes(ph(v))).key } : { chair: c.id };
    const rows = vs.map(v => {
      const p2 = ph(v), now = s.v && s.v.key === v.key, past = p2 === "done" || p2 === "missed";
      const icon = p2 === "done" ? '<span class="tick">✓</span>' : p2 === "missed" ? '<span class="xmark">✕</span>' : now ? lantern(p2 === "running") : '<span class="ring"></span>';
      const st = p2 === "done" ? "Done" : p2 === "missed" ? (OUTCOMES.find(k => v.r.note.includes(k)) || "Missed") : now ? TmsCore.chairStatus(s) : "";
      return `<button class="drow${past ? " past" : ""}" data-k="${esc(v.key)}" ${now ? `data-state="${s.state === "open" ? "track" : s.state}"` : ""} aria-pressed="${!!(state.sel && state.sel.k === v.key)}">${icon}<span class="w">${clock(v.time) || "–"}</span><span class="p">${esc(v.tab)}</span><span>${past ? "" : now ? `<span class="muted">${p2 === "running" ? `${typeName(v)}, ${sessNo(v)} of ${v.p.d.total}` : `In the lobby, booked ${clock(v.time)}`}</span>` : rowFlag(v, fl)}</span><span class="st">${st}</span></button>`;
    }).join("");
    const done = vs.filter(v => ph(v) === "done").length, todo = vs.filter(v => ["booked", "waiting"].includes(ph(v)) && !(s.v && s.v.key === v.key)).length;
    sub = `${esc(c.label)}: ${s.v ? (s.state === "late" ? `<strong>${esc(s.v.tab)} ${TmsCore.chairStatus(s).toLowerCase()}</strong>` : `${esc(s.v.tab)}, ${TmsCore.chairStatus(s).toLowerCase()}`) : "open now"}, ${done} done, ${todo} still to come`;
    left = `<div class="card day"><div class="h"><h2>${esc(c.label)} today</h2><span>${esc([c.device, c.coil].filter(Boolean).join(", "))}</span></div>${rows || '<p class="empty">Nothing booked on this chair today.</p>'}</div>`;
  }
  return head("Today", sub + (closed ? `. <strong>Clinic closed: ${esc(closed)}</strong>` : ""), true) + seg + `<div class="cols">${left}<aside class="card detail">${detailHtml(fl)}</aside></div>`;
}
/* the saved appointment list against the tabs for today (listCheck) */
function listCheckCard(day) {
  const c = listCheck(day); if (!c) return "";
  const run = state.report && state.report.asOf != null ? `List run ${fmtMon(state.report.asOf)}${state.report.asOfMin != null ? ` ${clock(state.report.asOfMin)}` : ""}` : "";
  const what = a => `${clock(a.time) || "no time"} ${esc(apptTypeName(a))}`;
  /* guests aren't a gap: they have no tab here by design, and show on Still to come and the Schedule */
  const rows = c.miss.filter(({ a }) => !a.guest).map(({ a, p, why }) => why === "new"
    ? `<div class="row"><div class="rowmain static"><span class="p">${esc(a.code || "New")}</span><span class="w">${clock(a.time)}</span><span>${esc(apptTypeName(a))}</span><span class="f flag">No tracker tab (MRN ending ${esc(a.mrn.slice(-4))})</span></div>${state.pendingNew.some(x => x.mrn === a.mrn) ? `<button class="btn ghost sm" data-act="newTracker" data-v="${esc(a.mrn)}">Create tracker</button>` : ""}</div>`
    : `<div class="row"><div class="rowmain static"><span class="p">${esc(p.tab)}</span><span class="w">${clock(a.time)}</span><span>${esc(apptTypeName(a))}</span><span class="f flag">${why === "disc" ? `${esc(p.status)} on the tracker, booked in NextGen` : "On the list, not on the tab"}</span></div>${why === "disc" ? "" : '<button class="btn ghost sm" data-act="reimport">Import again</button>'}</div>`)
    .concat(c.extra.map(({ p, r }) => `<div class="row"><div class="rowmain static"><span class="p">${esc(p.tab)}</span><span class="w">${clock(r.time)}</span><span>${esc(apptTypeName(r))}</span><span class="f flag">On the tab, not on the list. Check NextGen</span></div>${btn("opentab", p.tab, "", "Open").replace('data-act="opentab"', `data-act="opentab" data-row="${r.row}"`)}</div>`)).join("");
  return `<div class="card list"><div class="h"><h2>Appointment list check</h2><span>${esc(run)}</span></div>${rows || `<p class="body" style="margin:8px 0 10px">All ${c.n} visit${c.n === 1 ? "" : "s"} on today's list are on the tracker.</p>`}</div>`;
}
function varianceCard(day) {
  const v = variances(day, true), lines = Object.entries(v.groups).filter(([, a]) => a.length);
  if (!v.expected) return "";
  return `<div class="card list"><div class="h"><h2>Treatments today</h2><span>${v.actual} of ${v.expected} in treatment are booked</span></div>
    ${lines.length || v.grads.length ? `<dl class="facts" style="margin-bottom:12px">${lines.map(([k, a]) => `<dt>${esc(k)}</dt><dd>${a.map(([p, d]) => `${esc(p.tab)}${d ? ` <span class="muted">(${esc(d)})</span>` : ""}`).join(", ")}</dd>`).join("")}${v.grads.length ? `<dt>Graduated this week</dt><dd>${v.grads.map(p => esc(p.tab)).join(", ")}</dd>` : ""}</dl>` : '<p class="body" style="margin:8px 0 10px">Everyone in treatment is booked today.</p>'}</div>`;
}
function stepsFor(v, s, fl) {
  const p2 = ph(v), lateSt = s && s.state === "late" ? "late" : "now", co = v.p;
  /* Start session opens the treatment note (or the MT dialog); Complete session asks for pulses and the copay */
  const mtVisit = v.type === "MT" || v.type === "MTR", st = [{ st: p2 === "done" || p2 === "running" ? "done" : v.checkin != null ? lateSt : "now", t: mtVisit ? "Start session" : "Start session and treatment note", m: `Booked ${clock(v.time) || "no time"}, ${chairLbl(v.chair)}, about ${v.len} min`, w: v.start != null ? clock(v.start) : "" }];
  if (mtVisit) st.push({ st: v.mtDone || p2 === "done" ? "done" : p2 === "running" ? "now" : "next", t: v.type === "MT" ? "Mapping and motor threshold" : "Motor threshold redo", m: (v.r.note.match(/mt %:\s*([^;]+)/i) || [])[1] ? `MT ${v.r.note.match(/mt %:\s*([^;]+)/i)[1].trim()}` : "Saved to the MT row's Notes", w: "" });
  const pul = v.r.pulses.map((x, i) => x != null ? `${SITES[i]} ${x}` : "").filter(Boolean).join(", ");
  st.push({ st: p2 === "done" ? "done" : p2 === "running" && (v.mtDone || !mtVisit) ? "now" : "next", t: `Pulses${payOf(co, v.r) > 0 ? ", copay" : ""}${v.r.measX ? ", PHQ-9 and GAD-7" : ""}`, m: p2 === "done" ? `${pul}${v.r.tech ? `, ${v.r.tech}` : ""}` : p2 === "running" ? protoText(co) : "Asked when you complete the session", w: v.end != null ? clock(v.end) : "" });
  return `<ol class="steps">${st.map(x => `<li class="step ${x.st}">${mk(x.st)}<span><span class="t">${x.t}</span>${x.m ? `<span class="m">${esc(x.m)}</span>` : ""}</span><span class="w">${x.w}</span></li>`).join("")}</ol>`;
}
const chairBusy = (v, fl) => fl.visits.find(x => x !== v && x.chairVisit && x.chair === v.chair && ph(x) === "running") || null;
const openChairs = (v, fl) => fl.avail.filter(c => c.id !== v.chair && !fl.visits.some(x => x.chair === c.id && ph(x) === "running"));
function detailHtml(fl) {
  const sel = state.sel; if (!sel) return '<p class="body">Nothing is booked today.</p>';
  if (sel.chair) {
    const c = fl.avail.find(x => x.id === sel.chair); if (!c) return "";
    const s = TmsCore.chairState(c.id, fl.visits, fl.now), nx = s.next;
    return `<div class="dh"><div><div class="name">${esc(c.label)}<span class="loc">open</span></div><div class="meta">${esc([c.device, c.coil].filter(Boolean).join(", "))}</div></div></div>
      <p class="body">${nx ? `${esc(nx.tab)} is booked here at ${clock(nx.time) || "no set time"}${nx.time != null && nx.time < fl.now ? " and hasn't started" : ""}.` : "Nothing else booked on this chair today."}</p>
      ${nx ? `<button class="btn primary" data-act="start" data-k="${esc(nx.key)}">Start ${esc(nx.tab)}'s session</button><div class="section"><div class="h"><h2>${esc(nx.tab)}'s course</h2><span>${nx.p.d.txDone} of ${nx.p.d.total} done</span></div>${strip(nx.p, false)}${facts(nx.p, true)}</div>` : ""}`;
  }
  const v = findVisit(sel.k); if (!v) return "";
  const p = v.p, p2 = ph(v), c = fl.avail.find(x => x.id === v.chair), cs = c ? TmsCore.chairState(c.id, fl.visits, fl.now) : null, s = cs && cs.v && cs.v.key === v.key ? cs : null;
  const K = `data-k="${esc(v.key)}" data-t="${esc(v.tab)}" data-row="${v.row}"`;
  let lead = "", cls = "", primary = null, tl = "";
  const ghosts = [];
  if (v.alt) lead = `At ${p.home || "another clinic"} today`;
  else if (p2 === "missed") {
    const kind = OUTCOMES.find(k => v.r.note.includes(k)) || "Missed";
    lead = `${kind} at ${clock(v.time) || "today"}, ${v.r.freebie === "Used" ? `freebie used, ${p.d.left} left` : v.r.freebie === "Waived" ? "freebie waived" : "freebie not decided yet"}`; cls = "soon";
    if (!v.r.freebie) primary = ["Decide freebie", "classify"];
    ghosts.push(["Undo missed visit", "undoMiss"]);
  } else if (v.type === "F/U") {
    if (p2 === "done") { lead = `Follow-up done${v.end != null ? ` at ${clock(v.end)}` : ""}`; cls = "ready"; }
    else {
      lead = `${p.provider || "Provider"} at ${clock(v.time) || "no set time"}, ${v.tele ? "telehealth" : "in office"}`;
      if (v.r.measX) { lead += ". PHQ-9 and GAD-7 first"; primary = ["Enter PHQ-9 and GAD-7", "fumeas"]; }
      else primary = ["Mark follow-up done", "fudone"];
      ghosts.push(["Record missed", "miss"]);
    }
  } else if (p2 === "done") {
    lead = `Session ${sessNo(v)} done${v.end != null ? ` at ${clock(v.end)}` : ""}`; cls = "ready";
    if (!state.copied.has(v.key)) primary = ["Copy treatment note", "note"]; else lead += ". Note copied";
    if (v.r.delivered) ghosts.push(["Edit pulses", "editPulses"]);
    if (v.end != null) ghosts.push(["Undo complete session", "undoEnd"]);
  } else if (p2 === "running") {
    const r = runInfo(v), soon = r.left <= TmsCore.SOON_MIN, frac = Math.min(1, Math.max(0, (fl.now - r.s) / (r.e - r.s)));
    lead = soon ? (r.left > 0 ? `Ends in ${r.left} min, at ${clock(r.e)}` : r.left === 0 ? `Ends now (${clock(r.e)})` : `Session time is up (${clock(r.e)})`) : `In session, ends at ${clock(r.e)}`; cls = soon ? "soon" : "";
    tl = `<div class="tl" aria-hidden="true"><div class="bar"></div><div class="fill" style="width:${(frac * 100).toFixed(1)}%"></div><div class="mark done" style="left:0"></div><div class="mark" style="left:100%"></div><div class="now" style="left:${(frac * 100).toFixed(1)}%"></div><div class="lbl start">${clock(r.s)}</div><div class="lbl end" style="left:100%">${clock(r.e)}</div></div>`;
    const mtDue = (v.type === "MT" || v.type === "MTR") && !v.mtDone;
    primary = mtDue ? ["Record motor threshold", "mt"] : ["Complete session", "end"];
    ghosts.push(mtDue ? ["Complete session", "end"] : [v.type === "MT" || v.type === "MTR" ? "Edit motor threshold" : "Treatment note", v.type === "MT" || v.type === "MTR" ? "mt" : "note"]);
  } else {
    /* booked (or checked in on an earlier version): one Start session stamps check-in and start together */
    const busy = v.chair ? chairBusy(v, fl) : null, lateMin = v.time == null ? 0 : fl.now - v.time;
    if (!v.chair) { lead = v.offChair ? `${chairLbl(v.offChair)} is unavailable. Choose another chair` : "No chair on the appointment list. Choose one"; cls = "soon"; primary = ["Choose a chair", "move"]; }
    else {
      if (v.checkin != null) { const late = TmsCore.isLate(v, fl.now); lead = `Checked in at ${clock(v.checkin)}, booked ${clock(v.time) || "with no set time"}`; cls = late ? "late" : ""; }
      else if (v.time != null && lateMin > 5) { lead = `Not started, booked ${clock(v.time)} (${lateMin} min ago)`; cls = "soon"; } else lead = `Booked ${clock(v.time) || "with no set time"} in ${chairLbl(v.chair)}`;
      if (busy) { lead += `. ${chairLbl(v.chair)} in use by ${busy.tab} until ${clock(busy.start + busy.len)}`; primary = openChairs(v, fl).length ? ["Move to another chair", "move"] : [`Start session in ${chairLbl(v.chair)}`, "start", true]; }
      else primary = [`Start session in ${chairLbl(v.chair)}`, "start"];
      lead += `. Projected to end ${clock(Math.max(fl.now, v.time ?? fl.now, busy ? busy.start + busy.len : 0) + v.len)}`;
    }
    if (v.chair && !(primary && primary[1] === "move")) ghosts.push(["Move to another chair", "move"]);
    ghosts.push(["Record missed", "miss"]);
  }
  if (!v.alt && v.type !== "F/U" && p2 !== "missed") ghosts.push([v.type === "MT" ? "Print course schedule" : "Print schedule", "print"]);
  const moved = v.moved || (v.settled && v.chair) ? `<div class="notice${v.moved ? "" : " ok"}"><p>Moved from ${esc(chairLbl(v.from))} to ${esc(chairLbl(v.chair))}.${v.moved ? ` NextGen still has ${esc(chairLbl(v.from))}; move it there too.` : " Updated in NextGen."}</p><label class="check"><input type="checkbox" data-ngk="${esc(v.key)}"${v.moved ? "" : " checked"}> Moved in NextGen</label></div>` : "";
  const bidf = p.d.bidfToday ? `<div class="notice"><p><b>BIDF info needed.</b> Enter the ${esc(p.d.bidf.join("; "))}.</p><div class="actions" style="margin:4px 0 6px"><button class="btn ghost sm" data-act="ben" data-t="${esc(p.tab)}">Enter benefits</button></div></div>` : "";
  const priced = (p.cost[billIdx(v.r)] || {}).after != null, pay = payOf(p, v.r);
  const payline = v.alt || p2 === "missed" ? "" : `<div class="payline"><span>Payment today</span><span>${!priced && typeof v.r.pr !== "number" ? '<span class="flag">Not set</span>' : money(pay ?? 0) || "$0.00"}</span></div>`;
  const meta = v.type === "F/U" ? `Follow-up, ${v.tele ? "telehealth" : "in office"}` : `${typeName(v)}, treatment ${sessNo(v)} of ${p.d.total} · ${v.protoLen ? `protocol ${Math.round(v.protoLen)} min` : `no protocol time, ${DEFAULT_LEN} min assumed`}`;
  return `<div class="dh"><div><div class="name">${esc(v.tab)}<span class="loc">${esc(mrnTag(p))}</span></div><div class="meta">${esc(meta)}${where(v) ? ` · ${esc(where(v))}` : ""}</div></div><button class="btn mute sm" data-act="course" data-t="${esc(p.tab)}">Open</button></div>
    <p class="lead ${cls}">${esc(lead)}</p>${schedBand(v.r)}${tl}
    ${v.type === "F/U" || p2 === "missed" || v.alt ? "" : stepsFor(v, s, fl)}${moved}${bidf}
    ${primary ? `<button class="btn primary" data-act="${primary[1]}" ${K}${primary[2] ? " disabled" : ""}>${esc(primary[0])}</button>` : ""}
    <div class="actions">${ghosts.map(([l, a]) => `<button class="btn ghost sm" data-act="${a}" ${K}${a === "print" && v.type === "MT" ? ' data-v="all"' : ""}>${esc(l)}</button>`).join("")}</div>${payline}
    <div class="section"><div class="h"><h2>Course</h2><span>${p.status === "Tapering" ? "Tapering" : `${Math.max(0, p.d.total - p.d.txDone)} left`}</span></div>
      ${strip(p, v.chairVisit && ["booked", "waiting", "running"].includes(p2))}${scores(p, v.r.measX && !v.r.measVal)}${facts(p, true)}</div>`;
}

/* ---------- Schedule: a day by chair, or the week ---------- */
function viewSchedule(fl) {
  const seg = `<div class="viewbar"><span class="lbl">Show</span><div class="seg" role="group" aria-label="Day or week">${[["day", "Day by chair"], ["week", "Week"]].map(([k, l]) => `<button data-sv="${k}" aria-pressed="${state.dash === k}">${l}</button>`).join("")}</div></div>`;
  if (state.dash === "week") {
    const start = weekStart(fl.t) + 7 * state.weekOffset, nav = `<div class="nav"><button class="btn ghost sm" data-act="prev">‹ Earlier</button><button class="btn ghost sm" data-act="next">Later ›</button>${state.weekOffset ? '<button class="btn mute sm" data-act="this">This week</button>' : ""}</div>`;
    const days = [1, 2, 3, 4, 5].map(i => start + i).map(day => {
      const v = variances(day, true), lines = Object.entries(v.groups).filter(([, a]) => a.length), hrs = hoursOf(day);
      return `<div class="card wd${day === fl.t ? " today" : ""}"><div class="h"><h2>${fmtDay(day)}</h2><span>${day === fl.t ? "Today" : ""}</span></div>
        <div class="big">${v.actual} of ${v.expected} <small>${day < fl.t ? "treated" : "booked"}</small></div>${hrs && hrs.closed ? `<p class="flag" style="margin:6px 0 0">${esc(hrs.txt)}</p>` : ""}
        <ul>${lines.map(([k, a]) => `<li><span class="${k === "Missed" || k === "Away" ? "flag" : "muted"}">${esc(k)}:</span> ${a.map(([p]) => esc(p.tab)).join(", ")}</li>`).join("") || '<li class="muted">No variances</li>'}</ul>
        <ul class="vlist">${dayVisits(day)}</ul></div>`;
    }).join("");
    return head("Schedule", `Week of ${fmtMon(start)} (Sunday to Saturday). Treatments booked or done against everyone in treatment.`, true) + seg + nav + `<div class="wkgrid">${days}</div>
      <div class="actions" style="margin-top:20px"><button class="btn ghost" data-act="schedSheet">Write this week to the Chair Schedule sheet</button></div>`;
  }
  const day = fl.t + state.schedOffset, d = dow(day);
  const nav = `<div class="nav"><button class="btn ghost sm" data-act="sprev">‹ Earlier</button><b>${fmtLong(day)}</b><button class="btn ghost sm" data-act="snext">Later ›</button>${state.schedOffset ? '<button class="btn mute sm" data-act="sthis">Today</button>' : ""}</div>`;
  if (d === 0 || d === 6) return head("Schedule", "The chair schedule covers Monday to Friday.", true) + seg + nav + '<div class="card list"><p class="empty">Weekend.</p></div>';
  const ent = dayEntries(day), h = state.clinic?.days[d], times = ent.chair.flatMap(e => [e.start, e.end]).concat(ent.prov.flatMap(e => [e.start, e.end]));
  const start = Math.floor(Math.min(h?.open ?? 480, ...(times.length ? times : [480])) / 60) * 60, end = Math.ceil(Math.max(h?.close ?? 1020, ...(times.length ? times : [1020])) / 60) * 60, pxm = 1.6, height = (end - start) * pxm;
  const cl = fl.avail, single = chairList().length === 1 ? chairList()[0].id : "", isToday = day === fl.t;
  const lanes = [...cl.map(c => ({ id: c.id, name: c.label, sub: c.coil || c.device })), { id: "", name: "No chair", sub: "Not on the list" }, { id: "prov", name: "Provider", sub: "MT, MT redo and F/U" }];
  const chairFor = e => { if (isToday) { const v = fl.visits.find(x => x.tab === e.p.tab && x.row === e.r.row); if (v) return v.chair; } const c = chairOf(e.p, e.r) || single; return cl.some(x => x.id === c) ? c : ""; };
  /* visits that overlap in one lane sit side by side, not on top of each other; on a chair that's a double booking */
  const pack = list => {
    const out = list.slice().sort((a, b) => a.start - b.start || a.p.tab.localeCompare(b.p.tab));
    let grp = [], until = -1;
    const close = () => grp.forEach(e => { e.cols = grp.reduce((n, x) => Math.max(n, x.col + 1), 1); });
    for (const e of out) {
      if (e.start >= until) { close(); grp = []; }
      const cols = grp.filter(x => x.end > e.start).map(x => x.col); e.col = 0; while (cols.includes(e.col)) e.col++;
      grp.push(e); until = Math.max(until, e.end);
    }
    close(); return out;
  };
  /* guests' chair bookings here, from the saved list: they take the chair, but open their home clinic's tab */
  const gents = guestVisits(day).filter(x => x.time != null).map(x => ({ guest: x, p: { tab: x.code || "Guest" }, start: x.time, end: x.time + DEFAULT_LEN }));
  const guestBlk = e => { const x = e.guest, len = e.end - e.start, side = e.cols > 1 ? `;left:calc(4px + (100% - 8px) * ${e.col} / ${e.cols});right:auto;width:calc((100% - 8px) / ${e.cols} - 3px)` : "";
    const attrs = x.g.link ? `data-act="guestOpen" data-v="${esc(x.mrn)}"` : 'data-act="guests"';
    return `<button class="blk guest${e.cols > 1 ? " dbl" : ""}" ${attrs} style="top:${(e.start - start) * pxm}px;height:${Math.max(len * pxm - 3, 22)}px${side}" aria-label="${esc(e.p.tab)} at ${clock(e.start)}, guest from ${esc(x.g.home)}"><b>${esc(e.p.tab)}</b> ${clock(e.start)}${len >= 30 ? "<br>" : " "}Guest, ${esc(x.g.home)}${e.cols > 1 ? '<br><span class="flag">Double-booked</span>' : ""}</button>`; };
  const blocks = lane => pack(lane.id === "prov" ? ent.prov : ent.chair.filter(e => chairFor(e) === lane.id).concat(gents.filter(e => (cl.some(c => c.id === e.guest.chair) ? e.guest.chair : "") === lane.id))).map(e => {
    if (e.guest) return guestBlk(e);
    const v = isToday ? fl.visits.find(x => x.tab === e.p.tab && x.row === e.r.row) : null, p2 = v ? ph(v) : e.r.delivered ? "done" : "booked";
    const cs = v && v.chair && lane.id !== "prov" ? TmsCore.chairState(v.chair, fl.visits, fl.now) : null, live = cs && cs.v && cs.v.key === v.key;
    const k = p2 === "done" || p2 === "missed" ? "past" : live ? (cs.state === "open" ? "track" : cs.state) : "";
    const lab = p2 === "done" ? "Done" : live ? TmsCore.chairStatus(cs) : e.r.type === "MT" ? "MT, mapping" : e.r.type === "MTR" ? "MT redo" : e.r.type === "F/U" ? (e.r.tele ? "Tele F/U" : "F/U") : "Treatment";
    const len = e.end - e.start, dbl = e.cols > 1 && lane.id !== "prov" && lane.id !== "";
    const side = e.cols > 1 ? `;left:calc(4px + (100% - 8px) * ${e.col} / ${e.cols});right:auto;width:calc((100% - 8px) / ${e.cols} - 3px)` : "";
    return `<button class="blk ${k}${dbl ? " dbl" : ""}" ${v ? `data-k="${esc(v.key)}" data-go="today"` : `data-act="opentab" data-t="${esc(e.p.tab)}" data-row="${e.r.row}"`} style="top:${(e.start - start) * pxm}px;height:${Math.max(len * pxm - 3, 22)}px${side}" aria-label="${esc(e.p.tab)} at ${clock(e.start)}, ${esc(lab)}${dbl ? ", double-booked" : ""}"><b>${esc(e.p.tab)}</b> ${clock(e.start)}${len >= 30 ? "<br>" : " "}${esc(lab)}${v && v.moved ? " (moved)" : ""}${dbl ? '<br><span class="flag">Double-booked</span>' : ""}</button>`;
  }).join("");
  const used = lanes.filter(l => l.id !== "" || ent.chair.some(e => chairFor(e) === "") || gents.some(e => !cl.some(c => c.id === e.guest.chair)));
  let hrs = ""; for (let m = start + 60; m < end; m += 60) hrs += `<div class="hr" style="top:${(m - start) * pxm}px"></div>`;
  let tm = ""; for (let m = start; m <= end; m += 60) tm += `<span style="top:${(m - start) * pxm}px">${clock(m)}</span>`;
  const nowLine = isToday && fl.now >= start && fl.now <= end ? `<div class="nowline" style="top:${(fl.now - start) * pxm}px"></div>` : "";
  const grid = `<div></div>${used.map(l => `<div class="ch">${esc(l.name)}<small>${esc(l.sub)}</small></div>`).join("")}<div class="times" style="height:${height}px">${tm}</div>${used.map(l => `<div class="lane" style="height:${height}px">${hrs}${blocks(l)}${nowLine}</div>`).join("")}`;
  const hrsTxt = hoursOf(day), notes = [];
  if (hrsTxt && hrsTxt.closed) notes.push(`<p class="bad">${esc(hrsTxt.txt)}. ${ent.chair.length + ent.prov.length ? "Visits are booked on this closed date." : ""}</p>`);
  if (!state.clinic) notes.push('<p class="body">No Clinic Schedule sheet yet, so chair hours aren\'t set. <button class="btn ghost sm" data-act="mkSched">Create it</button></p>');
  if (sl0(day, ent)) notes.push(`<p class="flag">Outside chair hours: ${sl0(day, ent)}</p>`);
  if (ent.noTime.length) notes.push(`<p class="flag">No time on the tab: ${ent.noTime.map(x => esc(x.p.tab)).join(", ")}. Import the appointment list to fill times.</p>`);
  return head("Schedule", isToday ? "Select a visit to open it on Today. Hours come from the Clinic Schedule sheet." : "Select a visit to open its row on the tab.", true) + seg + nav
    + `<div class="card sched"><div class="sgrid" style="grid-template-columns:64px repeat(${used.length},minmax(110px,1fr))">${grid}</div>${notes.join("")}</div>`
    + (isToday ? "" : `<div style="margin-top:16px">${listCheckCard(day)}</div>`);
}
/* a day's visits for the Week card: time, patient, and anything but a plain daily treatment */
function dayVisits(day) {
  return state.patients.filter(onSched).flatMap(p => p.rows.filter(r => r.date === day && inChair(p, r)).map(r => ({ p, r }))).sort((a, b) => (a.r.time ?? 1e4) - (b.r.time ?? 1e4) || a.p.tab.localeCompare(b.p.tab))
    .map(({ p, r }) => { const marks = [r.type && r.type !== "Daily" ? r.type + (r.tele ? " tele" : "") : "", r.miss ? "missed" : "", r.alt ? "other clinic" : "", r.delivered ? "done" : "", p.d.awayConflicts.some(c => c.r.row === r.row) ? "away" : "", state.closures.has(day) && !r.delivered && !r.miss ? "clinic closed" : ""].filter(Boolean);
      return `<li><button class="vrow" data-act="opentab" data-t="${esc(p.tab)}" data-row="${r.row}"><span class="muted">${clock(r.time) || "–"}</span> <b>${esc(p.tab)}</b>${marks.length ? ` <span class="${/missed|away|closed/.test(marks.join()) ? "flag" : "muted"}">${esc(marks.join(", "))}</span>` : ""}</button></li>`; }).join("");
}
function sl0(day, ent) { const h = state.clinic?.days[dow(day)]; if (!h || h.open == null || h.close == null) return ""; return ent.chair.filter(e => e.start < h.open || e.end > h.close).map(e => `${esc(e.p.tab)} ${clock(e.start)}`).join(", "); }

/* ---------- Patients: one row per tab, sorted by what needs doing ---------- */
const FILTERS = [
  ["needs", "Needs attention", (p, n) => onSched(p) && TmsCore.loud(n).length > 0],
  ["active", "All active", p => p.d.active],
  ["nofu", "Follow-up not scheduled", p => p.d.active && p.d.fuStatus === "not scheduled"],
  ["nonext", "No upcoming visit", p => p.d.active && !p.d.next && p.d.txDone < p.d.total],
  ["taper", "Tapering", p => p.status === "Tapering"],
  ["paused", "Paused", p => p.status === "Paused"],
  ["complete", "Complete", p => p.status === "Complete"],
  ["disc", "Discontinued", p => !onSched(p)],
  ["copay", "Copay unconfirmed", p => p.d.toConfirm.length > 0]
];
function viewPatients(fl) {
  const N = new Map(state.patients.map(p => [p.tab, courseNeeds(p, fl)])), f = FILTERS.find(x => x[0] === state.cfilter) || FILTERS[0];
  const act = state.patients.filter(p => p.d.active).length, need = state.patients.filter(p => FILTERS[0][2](p, N.get(p.tab))).length;
  const chips = FILTERS.map(([k, l, fn]) => `<button class="chip" data-filter="${k}" aria-pressed="${k === f[0]}">${l} <span class="n">${state.patients.filter(p => fn(p, N.get(p.tab))).length}</span></button>`).join("");
  const rank = p => { const n = TmsCore.loud(N.get(p.tab)); return n.some(x => x.bad) ? 0 : n.length ? 1 : 2; };
  const list = state.patients.filter(p => f[2](p, N.get(p.tab))).sort((a, b) => rank(a) - rank(b) || a.code.localeCompare(b.code) || a.course - b.course);
  const rows = list.map(p => {
    const n = N.get(p.tab), d = p.d;
    return `<button class="crow" data-course="${esc(p.tab)}" aria-pressed="${p.tab === state.ptTab}"><span class="p">${esc(p.tab)}<small>${esc(p.status || "No status")}</small></span><span class="proto">${esc(protoText(p))}</span>
      <span class="prog"><span class="pbar"><i style="width:${Math.min(100, d.txDone / Math.max(1, d.total) * 100).toFixed(1)}%"></i></span>${d.txDone}/${d.total}</span>
      <span class="flags">${needSummary(n)}</span></button>`;
  }).join("") || `<p class="empty">Nothing here. Everyone in this group is on track.</p>`;
  return `<header class="top"><div><h1>Patients</h1><p class="sub">${act} active patient${act === 1 ? "" : "s"}, ${need} need${need === 1 ? "s" : ""} attention</p></div><button class="btn ghost" data-act="addPt">Add patient</button></header>
    <div class="chips" role="group" aria-label="Filter patients">${chips}</div>
    <div class="card courses"><div class="chead"><span>Patient</span><span>Protocol</span><span>Progress</span><span>Needs</span></div>${rows}</div>${elogHtml()}`;
}
/* the Edit log, collapsed at the bottom of Patients: what to check on each active tab, then what the add-in changed, newest first,
   each with Undo. The pick narrows both to one patient. */
const POOL_WORD = { TMS: "treatment", "F/U": "follow-up" };
function elogHtml() {
  const pick = state.elogTab, t = todaySerial(), rowBtn = (act, tab, row, label, v) => `<button class="btn ghost sm" data-act="${act}" data-t="${esc(tab)}" data-row="${row}"${v ? ` data-v="${esc(v)}"` : ""}>${esc(label)}</button>`;
  const line = (tab, text, acts) => `<div class="row"><div class="rowmain static"><span class="p">${esc(tab)}</span><span class="f flag">${text}</span></div><span class="acts">${acts}</span></div>`;
  /* "row 20", "rows 20 and 21", "rows 10–49", "12 rows, 36 to 58" */
  const rowsTxt = rs => {
    if (rs.length === 1) return `row ${rs[0]}`;
    const runs = []; rs.slice().sort((a, b) => a - b).forEach(r => { const l = runs[runs.length - 1]; if (l && r === l[1] + 1) l[1] = r; else runs.push([r, r]); });
    const parts = runs.map(([a, b]) => a === b ? String(a) : b === a + 1 ? `${a}, ${b}` : `${a}–${b}`);
    return parts.length > 4 ? `${rs.length} rows, ${runs[0][0]} to ${runs[runs.length - 1][1]}` : `rows ${andList(parts)}`;
  };
  const P = state.patients.filter(p => p.d.active && !state.stale.includes(p.tab) && (!pick || p.tab === pick)), checks = [];
  let nChecks = 0;
  for (const p of P) {
    const c = checksOf(p), tab = p.tab; if (!c.n) continue; nChecks += c.n;
    c.dup.forEach(d => { const rs = p.rows.filter(r => d.rows.includes(r.row));
      checks.push(line(tab, `${fmtMon(d.date)}: ${d.rows.length} ${POOL_WORD[d.pool]} rows (${rowsTxt(d.rows)})`, rs.filter(r => !r.delivered).map(r => rowBtn("elFix", tab, r.row, `Remove row ${r.row}`, "remove")).join("") + rowBtn("opentab", tab, d.rows[0], "Open"))); });
    if (c.tx.length) checks.push(line(tab, `Tx # off on ${plural(c.tx.length, "row")}, first on row ${c.tx[0].row}: ${esc(c.tx[0].was || "blank")}, should be ${c.tx[0].want === "" ? "blank" : c.tx[0].want}`, rowBtn("elFix", tab, c.tx[0].row, "Renumber Tx #", "tx") + rowBtn("opentab", tab, c.tx[0].row, "Open")));
    if (c.order.length) checks.push(line(tab, `Out of date order at ${rowsTxt(c.order)}`, rowBtn("elFix", tab, c.order[0], "Sort by date", "sort") + rowBtn("opentab", tab, c.order[0], "Open")));
    if (c.formulas.length) checks.push(line(tab, `${rowsTxt(c.formulas.map(x => x.row))} lost ${c.formulas.some(x => x.a) && c.formulas.some(x => x.d) ? "the PR and Tx left formulas" : c.formulas.some(x => x.a) ? "the PR formula in A" : "the Tx left formula in D"}`, rowBtn("elFix", tab, c.formulas[0].row, "Put formulas back", "formulas")));
    if (c.kaiser.length) checks.push(line(tab, `Template's Kaiser reminder on ${rowsTxt(c.kaiser)}`, rowBtn("elFix", tab, c.kaiser[0], "Remove the reminder", "kaiser")));
    if (c.over.length) checks.push(line(tab, `Treatment ${c.over[0].n} on ${fmtMon(c.over[0].date)} is past the ${p.txApproved} approved. Add an extension or cancel it in NextGen`, rowBtn("opentab", tab, c.over[0].row, "Open")));
    c.gaps.notOnTab.forEach(a => checks.push(line(tab, `On the list, not on the tab: ${fmtMon(a.date)} ${clock(a.time)} ${esc(apptTypeName(a))} (${a.status === "Kept" ? "kept" : "booked"})`, '<button class="btn ghost sm" data-act="reimport">Import again</button>')));
    c.gaps.notOnList.forEach(row => { const r = p.rows.find(x => x.row === row); checks.push(line(tab, `On the tab, not on the list: ${fmtMon(r.date)} ${esc(apptTypeName(r))} (row ${row}). Check NextGen`, rowBtn("opentab", tab, row, "Open"))); });
    (c.gaps.noAppt || []).forEach(row => { const r = p.rows.find(x => x.row === row); checks.push(line(tab, `Pulses on ${fmtMon(r.date)} (row ${row}), no visit in NextGen that day. Check NextGen`, rowBtn("editPulses", tab, row, "Edit pulses"))); });
  }
  const E = state.edits.filter(e => !pick || e.tab === pick), today = state.edits.filter(e => Math.floor(e.when) === t && e.kind !== "undo").length;
  const edits = E.slice(0, 80).map(e => `<div class="row erow${e.st && e.st !== "No undo" ? " past" : ""}"><div class="rowmain static"><span class="w">${esc(whenText(e.when))}</span><span class="p">${esc(e.tab)}</span><span class="what">${esc(e.what)}${e.by ? ` <span class="muted">${esc(e.by)}</span>` : ""}</span></div>
    <span class="acts">${e.kind === "undo" ? "" : e.st === "No undo" ? '<span class="muted">No undo</span>' : e.st ? `<span class="muted">${esc(e.st)}</span>` : `<button class="btn ghost sm" data-act="elUndo" data-v="${esc(e.id)}" aria-label="Undo ${esc(e.tab)}: ${esc(e.what)}">Undo</button>`}</span></div>`).join("");
  const tabs = [...new Set([...state.patients.filter(p => p.d.active).map(p => p.tab), ...state.edits.map(e => e.tab)])].sort();
  const sub = [`${plural(today, "edit")} today`, nChecks ? `<span class="flag">${plural(nChecks, "thing")} to check</span>` : ""].filter(Boolean).join(", ");
  return `<details class="card elog" data-elog ${state.elogOpen ? "open" : ""}><summary><span class="t">Edit log</span><span class="sub">${sub}</span></summary>
    <div class="ebar"><label>Patient <select data-elf><option value="">All patients</option>${tabs.map(x => `<option${x === pick ? " selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
    <span class="muted">Every change the add-in makes to a tab. Undo works while the cells still hold what the edit left.</span></div>
    <h3>To check</h3>${checks.join("") || `<p class="empty">Nothing to check${pick ? ` on ${esc(pick)}` : ""}. Every log matches its Tx # rule, formulas and the appointment list.</p>`}
    <h3>Changes</h3>${edits || '<p class="empty">No changes logged yet.</p>'}</details>`;
}
/* an import, sort or removed row puts the whole visit log back, so it names that first; a cell edit undoes straight away */
function elUndoAsk(id) {
  const e = state.edits.find(x => x.id === id); if (!e) return;
  if (!["import", "sort", "rows"].includes(e.kind)) return undoEdit(id);
  state.mctx = { id };
  showModal(`<h2 id="mtitle">Undo ${esc(e.tab)}'s ${e.kind === "import" ? "import" : "change"} from ${esc(whenText(e.when))}?</h2><p class="body">${esc(e.what)}.</p>
    <p class="body">The visit log goes back to how it was before${e.kind === "import" ? ": rows the import created are removed and every date, type and note it wrote is put back" : ""}. It only runs if nothing on the log changed since.</p>${mfoot(e.kind === "import" ? "Undo the import" : "Undo the change", "elUndoGo")}`);
}
function elRemoveAsk(tab, row) {
  const p = state.patients.find(x => x.tab === tab), r = p && p.rows.find(x => x.row === row); if (!r) return;
  state.mctx = { t: tab, row };
  showModal(`<h2 id="mtitle">Remove row ${row} from ${esc(tab)}?</h2><p class="body">${r.date != null ? fmtMon(r.date) : "No date"}, ${esc(r.type || "no type")}${r.miss ? ", missed" : ""}, no pulses${r.note ? `. Notes: ${esc(r.note)}` : ""}.</p>
    <p class="body">The rows below move up one. The Edit log can put it back.</p>${mfoot(`Remove row ${row}`, "elFixGo")}`);
}
/* a course row names its two most urgent needs and counts the rest; quiet labels show only when nothing needs doing */
function needSummary(n) {
  const L = TmsCore.loud(n), q = n.filter(x => x.quiet);
  if (!L.length) return q.length ? esc(q.map(x => x.t).join(", ")) : "On track";
  return L.slice(0, 2).map(x => `<span class="f${x.bad ? " bad" : ""}">${esc(x.t)}</span>`).join(", ") + (L.length > 2 ? ` <span class="muted">and ${L.length - 2} more</span>` : "");
}
/* ---------- the patient view: an overlay over Patients (in the task pane it fills the pane). Today's visit sits at the top; the
   sidebar picks one section (saved on this computer like the Today view). The list row stays selected behind it. ---------- */
const PT_SECT = [["sched", "Schedule"], ["course", "Course"], ["proto", "Protocol"], ["meas", "Measures"], ["ben", "Benefits"], ["fin", "Financial"]];
const PT_NEEDS = 4;   // needs shown before "Show all"; the next one peeks through, faded
function ptShow() {
  let v = state.ptShowMem; try { v = v || localStorage.getItem("tms_ptsect"); } catch { /* storage unavailable */ }
  return PT_SECT.some(s => s[0] === v) ? v : "sched";
}
function setPtShow(k) {
  state.ptShowMem = k; try { localStorage.setItem("tms_ptsect", k); } catch { /* this session only */ }
  renderPt();
}
let ptFrom = null;
function openPt(tab) {
  state.ptTab = tab; state.ptOpen = tab; state.ptPick = new Set(); state.ptPast = false; state.ptNeedsAll = false; ptFrom = document.activeElement; renderDash();
  const x = $("pvBody").querySelector("[data-act=ptClose]"); if (x) x.focus();
}
function closePt() {
  state.ptOpen = null; $("ptView").hidden = true; $("pvBody").innerHTML = ""; renderDash();
  const back = document.querySelector(`#v-patients [data-course="${CSS.escape(state.ptTab || "")}"]`) || ptFrom; if (back && document.contains(back)) back.focus(); ptFrom = null;
}
/* re-drawn with the dashboard (the 30 s tick, every save); scroll, focus and the ticked copays are kept */
function renderPt() {
  const box = $("ptView"), body = $("pvBody"), p = state.ptOpen && state.patients.find(x => x.tab === state.ptOpen);
  if (!p) { if (!box.hidden) { box.hidden = true; body.innerHTML = ""; state.ptOpen = null; } return; }
  const top = body.scrollTop, fk = document.activeElement && body.contains(document.activeElement) ? focusKey(document.activeElement) : null;
  const by = $("pvBy") ? $("pvBy").value : null;
  body.innerHTML = ptHtml(p, state.fl || floor()); box.hidden = false; body.scrollTop = top;
  if (by != null && $("pvBy")) $("pvBy").value = by;
  if (fk) { const el = body.querySelector(fk); if (el) el.focus(); }
}
const ptVisit = (p, fl) => { const tv = fl.visits.filter(v => v.tab === p.tab && !v.miss); return tv.find(v => v.chairVisit) || tv[0] || null; };
const chairFor = (p, r) => { const s = state.sess && [...state.sess.rows.values()].find(x => x.tab === p.tab && x.date === r.date && x.chair); return s ? s.chair : chairOf(p, r); };
/* Notes (S) hold the add-in's own entries (copay, freebie, NextGen tags, MT %, protocol renewal, a scheduled note) beside what
   staff type. Editing a note changes only the typed part; a scheduled note is one entry, "Scheduled note: ...", shown on its day. */
const SCHED_RE = /^scheduled note:\s*/i;
const ownEntry = x => !(TmsCore.copayOf(x) || SCHED_RE.test(x) || /^freebie (used|waived)$/i.test(x) || x === TELE_TAG || x === EXEMPT_TAG || x === UNCLASS_TAG || /\(NextGen\)$/.test(x) || /^mt %:/i.test(x) || /^protocol renewed/i.test(x));
const noteParts = note => String(note || "").split("; ").filter(Boolean);
const ownNote = note => noteParts(note).filter(ownEntry).join("; ");
const schedNote = note => (noteParts(note).find(x => SCHED_RE.test(x)) || "").replace(SCHED_RE, "");
const cleanNote = s => String(s || "").replace(/\s*;\s*/g, ", ").replace(/\s+/g, " ").trim();   // "; " separates entries
const withOwnNote = (note, s) => [...noteParts(note).filter(x => !ownEntry(x)), ...(cleanNote(s) ? [cleanNote(s)] : [])].join("; ");
const withSchedNote = (note, s) => [...noteParts(note).filter(x => !SCHED_RE.test(x)), ...(cleanNote(s) ? [`Scheduled note: ${cleanNote(s)}`] : [])].join("; ");
const visitNote = note => noteParts(note).filter(x => !TmsCore.copayOf(x) && !SCHED_RE.test(x)).join("; ");
/* a scheduled note on its day: the band that leads the visit */
const schedBand = r => { const s = r && schedNote(r.note); return s ? `<div class="snote"><b>Scheduled note</b><span>${esc(s)}</span></div>` : ""; };
/* color-coded markers on a visit, each with its word: MT, MTR, follow-up, auth end, final TMS visit, measures due */
function visitMarks(p, r) {
  const m = [];
  if (r.type === "MT") m.push(["mt", "MT"]); else if (r.type === "MTR") m.push(["mtr", "MTR"]); else if (r.type === "F/U") m.push(["fu", r.tele ? "Follow-up, tele" : "Follow-up"]);
  if (p.authExp != null && r.date === p.authExp) m.push(["auth", "Auth ends"]); else if (p.authExp != null && r.date > p.authExp && r.type !== "F/U") m.push(["auth", "Past auth end"]);
  if (p.d.grad != null && r.date === p.d.grad && TmsCore.isTreatment(r)) m.push(["final", "Final TMS visit"]);
  if (r.measX && !r.measVal) m.push(["meas", "Measures due"]);
  return m.map(([k, l]) => `<span class="mk mk-${k}">${esc(l)}</span>`).join("");
}
/* what the patient owes and has paid, in the place an Add extension button would take at the top (that lives in Course) */
function finLine(p) {
  const b = p.ben, d = p.d;
  const parts = d.oopmMet ? [["OOPM", `met (${d.oopmMet})`]] : !d.owes ? [["Copays", "none, BIDF $0"]]
    : [["Collected", money(d.tally.sum) || "$0.00"], ["To confirm", String(d.toConfirm.length)], ["OOPM left", benAmt(b.oopNA, b.oopRem)], ["Deductible left", benAmt(b.dedNA, b.dedRem)]];
  return `<div class="finline">${[["Cost so far", money(d.accrued) || "$0.00"], ...parts].map(([k, v]) => `<span><small>${esc(k)}</small><b>${esc(v)}</b></span>`).join("")}<button class="btn mute sm" data-ptshow="fin">Financial</button></div>`;
}
function ptHtml(p, fl) {
  const N = courseNeeds(p, fl), n = TmsCore.loud(N), first = n[0], d = p.d, v = ptVisit(p, fl), on = ptShow();
  const A = x => x.act[1] === "ext" ? 'data-ptshow="course"' : `data-act="${x.act[1]}" data-t="${esc(p.tab)}"${x.row ? ` data-row="${x.row}"` : ""}${x.key ? ` data-k="${esc(x.key)}"` : ""}${x.v ? ` data-v="${esc(x.v)}"` : ""}`;
  const label = x => x.act[1] === "ext" ? "Open Course" : x.act[0];
  const shown = new Set([first && first.act[1] !== "ext" && first.act[1]]);   // an auth lead shows the financial line, so its rows below keep one Open Course
  /* four needs in all (the lead and three); a fifth peeks through faded until Show all */
  const all = !!state.ptNeedsAll, rest = n.slice(1), keep = all ? rest : rest.slice(0, PT_NEEDS - 1), peek = all ? null : rest[PT_NEEDS - 1];
  const needRow = (x, cls) => { const dup = shown.has(x.act[1]) && !x.row && !x.key; shown.add(x.act[1]); return `<div class="row${cls || ""}"${cls ? ' aria-hidden="true"' : ""}><span class="f ${x.bad ? "bad" : "flag"} needline">${esc(x.t)}</span>${dup ? "" : `<button class="btn ghost sm" ${A(x)}${cls ? ' tabindex="-1"' : ""}>${esc(label(x))}</button>`}</div>`; };
  const also = keep.map(x => needRow(x)).join("") + (peek ? needRow(peek, " peek") : "");
  const more = rest.length > PT_NEEDS - 1 ? `<button class="btn mute sm needmore" data-act="ptNeeds" aria-expanded="${all}">${all ? "Show fewer" : `Show all ${n.length} needs`}</button>` : "";
  const quiet = N.filter(x => x.quiet).map(x => `<span class="tag">${esc(x.t)}</span>`).join("");
  const meta = [p.provider && `Dr. ${p.provider}`, `${d.txDone} of ${d.total} done`, p.status].filter(Boolean).join(" · ");
  const sect = { sched: ptSched, course: ptCourse, proto: ptProto, meas: ptMeas, ben: ptBen, fin: ptFin };
  return `<div class="pvh"><div><h2 id="pvTitle">${esc(p.tab)} <span class="mrn">MRN ${esc(p.mrnRaw || "not entered")}</span></h2><p class="meta">${esc(meta)}</p></div>
      <div class="pvr"><button class="x" data-act="ptClose" aria-label="Close">×</button></div></div>
    ${ptToday(p, v)}
    <p class="lead ${first ? (first.bad ? "late" : "soon") : ""}">${first ? esc(first.t) : d.active ? "On track" : esc(p.status || "No status")}</p>${quiet ? `<div class="tags">${quiet}</div>` : ""}
    ${first && first.act ? (first.act[1] === "ext" ? finLine(p) : `<button class="btn primary" ${A(first)}>${esc(first.act[0])}</button>`) : ""}
    ${also ? `<div class="section needs${all ? "" : " clipped"}"><div class="h"><h2>Also needs</h2><span>${n.length - 1} more</span></div>${also}${more}</div>` : ""}
    <div class="pvgrid"><nav class="pvnav" aria-label="Sections">${PT_SECT.map(([k, l]) => `<button data-ptshow="${k}" aria-current="${on === k}">${l}</button>`).join("")}</nav>
      <div class="pvs" data-sect="${on}">${sect[on](p, fl)}</div></div>`;
}
/* today's visit, first thing in the patient view: when, where, its state, the next step, and the copay on the same band */
function ptToday(p, v) {
  if (!v) return "";
  const p2 = ph(v), K = `data-k="${esc(v.key)}" data-t="${esc(p.tab)}" data-row="${v.row}"`;
  const st = v.alt ? `At ${p.home || "another clinic"}` : p2 === "missed" ? "Missed" : p2 === "done" ? `Done${v.end != null ? ` ${clock(v.end)}` : ""}` : p2 === "running" ? `In session since ${clock(v.start)}` : v.checkin != null ? `Checked in ${clock(v.checkin)}` : "Booked";
  const go = v.alt || p2 === "missed" ? "" : p2 === "running" ? `<button class="btn ghost sm" data-act="${(v.type === "MT" || v.type === "MTR") && !v.mtDone ? "mt" : "end"}" ${K}>${(v.type === "MT" || v.type === "MTR") && !v.mtDone ? "Record motor threshold" : "Complete session"}</button>`
    : p2 === "done" ? (v.r.delivered ? `<button class="btn ghost sm" data-act="ptJump" ${K}>Edit pulses</button>` : "")
    : v.type === "F/U" ? `<button class="btn ghost sm" data-act="${v.r.measX ? "fumeas" : "fudone"}" ${K}>${v.r.measX ? "Enter PHQ-9 and GAD-7" : "Mark follow-up done"}</button>`
    : v.chair ? `<button class="btn ghost sm" data-act="start" ${K}>Start session</button>` : `<button class="btn ghost sm" data-act="move" ${K}>Choose a chair</button>`;
  return `<div class="pvtoday"><span class="w">Today ${clock(v.time) || "no time"}</span><span>${esc(v.type === "F/U" ? "Follow-up" : `${typeName(v)}, ${sessNo(v)} of ${p.d.total}`)}</span><span class="muted">${esc(where(v))}</span><span class="grow">${esc(st)}</span>${go}<button class="btn mute sm" data-open="${esc(p.tab)}">On Today</button></div>${schedBand(v.r)}${ptCopayToday(p)}`;
}

/* today's copay: the payment from column A on today's row, the answer from its Notes, and who answered. Answering here also
   answers it in Complete session, and the other way round. */
function ptCopayToday(p) {
  const t = todaySerial(), r = p.rows.find(x => x.date === t && !x.miss && !x.alt && x.type !== "F/U") || p.rows.find(x => x.date === t && !x.miss && !x.alt);
  if (!r) return "";
  const c = TmsCore.copayOf(r.note), pay = payOf(p, r);
  if (!p.d.owes && !c) return `<div class="copay line"><span>Copay today</span><span class="muted">${p.d.oopmMet ? `None: ${p.d.oopmMet} OOPM met` : "None: the BIDF quote is $0"}</span></div>`;
  const by = `<label class="by">By <input id="pvBy" value="${esc(state.techName || lastTech(p))}" autocomplete="off"></label>`;
  /* answered: the marker on the same line, with Change to answer again */
  if (c && state.copayEdit !== `${p.tab}|${r.row}`) {
    const mark = c.kind === "collected" ? `<span class="got">✓ Collected ${esc(money(c.amt))}</span>` : `<span class="${c.kind === "not" ? "flag" : "muted"}">${esc(c.kind === "not" ? "Not collected" : c.kind === "oopm" ? "$0, OOPM met" : c.kind === "waived" ? "Waived" : c.line)}</span>`;
    return `<div class="copay line"><span>Copay today</span><span class="amt">${esc(money(pay ?? c.amt) || "")}</span>${mark}${c.by ? `<span class="muted">${esc(c.by)}${c.at ? `, ${esc(c.at)}` : ""}</span>` : ""}<button class="btn mute sm" data-act="copayChange" data-t="${esc(p.tab)}" data-row="${r.row}">Change</button>${by.replace("<label", '<label hidden')}</div>`;
  }
  const b = (k, l, dis) => `<button class="btn ghost sm" data-act="copayToday" data-t="${esc(p.tab)}" data-row="${r.row}" data-v="${k}" aria-pressed="${!!c && c.kind === k}"${dis ? " disabled" : ""}>${l}</button>`;
  return `<div class="copay line"><span>Copay today</span><span class="amt">${pay != null ? esc(money(pay)) : '<span class="flag">Not set</span>'}</span>${b("collected", "Collected", pay == null)}${b("not", "Not collected")}${by}</div>`;
}
/* a note on a visit: the typed part of Notes (past or today), or a scheduled note for a later visit, shown on its day */
function mPtNote(tab, row, kind) {
  const [, r] = findPR(tab, row); if (!r) return; state.mctx = { t: tab, row, kind };
  const sched = kind === "sched", cur = sched ? schedNote(r.note) : ownNote(r.note);
  showModal(`<h2 id="mtitle">${sched ? (cur ? "Edit scheduled note" : "Add scheduled note") : cur ? "Edit note" : "Add note"}: ${esc(tab)}, ${fmtDay(r.date)}</h2>
    <p class="sub">${sched ? "Shows at the top of this visit on its day, on Today and in the patient view." : "Saved to Notes on this visit's row. Copay, freebie and NextGen entries stay as they are."}</p>
    ${fld("Note", `<textarea id="mPtNote" rows="3" maxlength="400">${esc(cur)}</textarea>`)}${mfoot(cur ? "Save note" : "Add note", "doPtNote")}`);
  const ta = $("mPtNote"); if (ta) ta.focus();
}
async function savePtNote(tab, row, kind, s) {
  const [, r] = findPR(tab, row); if (!r) return;
  const note = kind === "sched" ? withSchedNote(r.note, s) : withOwnNote(r.note, s);
  if (note === r.note) return;
  try { await writeCells(tab, [[`S${row}`, note]], at(r)); patchRow(tab, row, { note }); toast(cleanNote(s) ? "Note saved." : "Note removed."); } catch (e) { fail(e); }
}
/* Edit pulses on today's band goes down to that visit in Past visits */
function ptJump(row) {
  state.ptPast = true; setPtShow("sched");
  const tr = $(`pvrow-${row}`); if (!tr) return;
  tr.scrollIntoView({ block: "center" }); tr.classList.add("hl");
  const b = tr.querySelector('[data-act="editPulses"]'); if (b) b.focus({ preventScroll: true });
}
async function copayToday(tab, row, kind) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const by = (($("pvBy") || {}).value || "").trim(); if (!by) return toast("Enter who answered the copay (By).", true);
  const note = withCopay(r.note, copayLine(kind, payOf(p, r), by));
  try { await writeCells(tab, [[`S${row}`, note]], at(r)); state.techName = by; state.copayEdit = null; patchRow(tab, row, { note }); toast("Copay saved to Notes."); } catch (e) { fail(e); }
}

/* ---------- Protocol: J2:O4 against the Protocols table; sites read F3 (left) and F4 (right) ---------- */
const pctMt = v => typeof v === "number" ? `${v <= 3 ? Math.round(v * 100) : v}% MT` : String(v || "").trim() ? `${String(v).trim()} MT` : "";
function protoLines(p) {
  return siteRows(p).map(x => {
    const m = TmsCore.protoMatch(x, state.protocols), site = TmsCore.siteOf(x.site), q = m && m.base, tot = x.dur * x.trains;
    const ch = !q ? [] : [...(m.diffs.some(dd => dd.k !== "wait") ? [`${tot.toLocaleString()} pulses (standard ${(q.pulses * q.trains).toLocaleString()})`] : []), ...(m.diffs.some(dd => dd.k === "wait") ? [`${x.wait} s wait (standard ${q.iti} s)`] : [])];
    return { x, m, site, text: [TmsCore.SITE_LABEL[site], `${x.hz} Hz`, q && pctMt(q.inten), `${tot.toLocaleString()} pulses`].filter(Boolean).join(" · "), ch };
  });
}
function ptProto(p) {
  const L = protoLines(p), d = p.d, t = todaySerial();
  const head = `<div class="h"><h2>Protocol</h2><button class="btn ghost sm" data-act="proto" data-t="${esc(p.tab)}">Edit</button></div>`;
  if (!L.length) return head + `<p class="body">Protocol not confirmed. Select the one ordered from the Protocols table and confirm it; its values fill in.</p>${protoPick(p)}`;
  const names = [...new Set(L.map(l => l.m ? l.m.name : "Custom"))].join(" + ");
  const mtRow = p.rows.filter(r => (r.type === "MT" || r.type === "MTR") && /mt %:/i.test(r.note)).sort(cmp).pop(), mt = mtRow && (mtRow.note.match(/mt %:\s*([^;]+)/i) || [])[1];
  const recheck = d.sched.find(r => r.type === "MTR"), len = protoLen(p);
  const rows = [["Treatment", `${L.length > 1 ? `Bilateral, ${L.map(l => l.site).join(" then ")}` : "Unilateral"}${len ? `, ${Math.round(len)} min` : ""}`],
    ["MT", `${mt ? mt.trim() : "Not recorded"}${recheck ? `, recheck ${fmtMon(recheck.date)}` : ""}`],
    ["Device", device() === "Brainsway" ? "BrainsWay" : "MagVenture"],
    ["Update protocol by", p.updateBy != null ? fmtMon(p.updateBy) : "Not set", d.protoFlags.length ? "flag" : ""]];
  return head + `<p class="pname">${esc(names)}</p>${L.map(l => `<p class="prow">${esc(l.text)}${l.ch.length ? `<small>${esc(l.ch.join("; "))}</small>` : ""}</p>`).join("")}
    <dl class="facts">${rows.map(([k, v, cls]) => `<dt>${k}</dt><dd class="${cls || ""}">${esc(v)}</dd>`).join("")}</dl>`;
}
/* ---------- Course: progress, the authorization and everything about it, and Add extension ---------- */
function ptCourse(p, fl) {
  const d = p.d, t = todaySerial(), tot = Math.max(1, d.total), taper = tot >= 12 ? 6 : 0, left = Math.max(0, d.total - d.txDone);
  const authCls = p.authExp == null || p.authExp < t ? "bad" : p.authExp - t <= state.cfg.authDays ? "flag" : "";
  const rows = [["Course", `${tot} treatments: ${tot - taper} daily${taper ? ` + ${taper} taper` : ""}`],
    ["Authorization", `${left} of ${d.total} left${p.authExp != null ? `, ${p.authExp < t ? "expired" : "ends"} ${fmtMon(p.authExp)}` : ", no end date"}`, authCls],
    ["Auth dates", `${p.authStart != null ? fmtMon(p.authStart) : "No start"} to ${p.authExp != null ? fmtMon(p.authExp) : "no end"}`],
    ["Extensions added", p.ben.ext ? `${p.ben.ext} daily visits` : "None"],
    ["MT redos", p.mtRedos ? `${p.mtRedos}${p.mtrAuth ? ", approved" : ", not approved"}` : p.mtrAuth ? "Approved, none used" : "None"],
    ["Auth assessment", p.authAssess || "Not entered"],
    ["Final TMS visit", d.grad != null ? fmtMon(d.grad) : "Not booked"], ["Freebies", `${d.left} left`, d.left <= 1 ? "flag" : ""]];
  const flags = d.authFlags.map(x => `<p class="${/expired|past auth|authorized \(D2\)/i.test(x) ? "flag bad" : "warnline"}">${esc(x)}</p>`).join("");
  return `<div class="h"><h2>Course</h2><span>${p.status === "Tapering" ? "Tapering" : `${left} left`}</span></div>
    <p class="sub">${d.txDone} of ${d.total} done${p.mtRedos ? `, ${p.mtRedos} MT redo${p.mtRedos > 1 ? "s" : ""}` : ""}</p>${strip(p, fl.visits.some(v => v.tab === p.tab && v.chairVisit && ["booked", "waiting", "running"].includes(ph(v))))}
    ${factList(rows)}${flags}
    <button class="btn ${d.authFlags.length ? "primary" : "ghost"} wide" data-act="ext" data-t="${esc(p.tab)}">Add extension</button>`;
}
/* ---------- Measures: PHQ-9 and GAD-7 recorded, and the visits they're due on ---------- */
function ptMeas(p, fl) {
  const t = todaySerial(), v = fl.visits.find(x => x.tab === p.tab && x.r.measX && !x.r.measVal), due = p.rows.filter(r => r.date != null && r.date >= t && r.measX && !r.measVal).sort(cmp);
  const enter = v && (v.type === "F/U" || ph(v) === "running") ? `<button class="btn ghost sm" data-act="${v.type === "F/U" ? "fumeas" : "end"}" data-k="${esc(v.key)}" data-t="${esc(p.tab)}" data-row="${v.row}">Enter scores</button>` : "";
  return `<div class="h"><h2>Measures</h2>${enter}</div>${scores(p, !!v)}
    <div class="sub2">Due</div>${due.length ? due.slice(0, 6).map(r => `<div class="srow"><span class="p">${fmtDay(r.date)}</span><span>${clock(r.time) || "no time"}</span><span></span><span>${typeof r.tx === "number" ? `#${r.tx}` : esc(apptTypeName(r))}</span><span class="nt">${visitMarks(p, r)}</span><span></span></div>`).join("") : '<p class="body">None placed on upcoming visits.</p>'}`;
}
/* no protocol yet: the dropdown, with the values greyed until one is chosen */
function protoPick(p) {
  return `<div class="g2p pick"><select id="pvProto" aria-label="Protocol"><option value="">Select the protocol</option>${state.protocols.map(q => `<option>${esc(q.name.trim())}</option>`).join("")}</select>
    <select id="pvSite" aria-label="Site">${["F3", "F4", "Other"].map(s => `<option value="${s}">${TmsCore.SITE_LABEL[s]}</option>`).join("")}</select>
    <button class="btn ghost sm" data-act="pvProtoSave" data-t="${esc(p.tab)}">Confirm protocol</button></div>
    <div class="g4 greyed" aria-hidden="true">${["Hz", "Pulses per train", "Trains", "Wait (s)"].map(l => `<label class="fld"><span>${l}</span><input disabled tabindex="-1"></label>`).join("")}</div>`;
}
/* Edit: a protocol per row from the table (its values fill in and stay editable), the site, and the next expiration date (G2,
   required). Saving also logs "Protocol renewed, next by m/d" in today's Notes. */
function mProtoEdit(tab) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return; state.mctx = { t: tab };
  const opt = v => `<option value="">(none)</option>` + state.protocols.map(q => `<option${q.name.trim() === v ? " selected" : ""}>${esc(q.name.trim())}</option>`).join("");
  const rowH = x => {
    const m = TmsCore.protoMatch(x, state.protocols), base = m && m.base ? m.base.name.trim() : "", s = TmsCore.siteOf(x.site), has = isNum(x.hz);
    return `<div class="per" data-pr="${x.n}"><select data-pk="name" aria-label="Protocol ${x.n}">${opt(base)}</select><select data-pk="site" aria-label="Site ${x.n}"><option value="">Site</option>${["F3", "F4", "Other"].map(k => `<option value="${k}"${k === s ? " selected" : ""}>${TmsCore.SITE_LABEL[k]}</option>`).join("")}</select>
      ${[["hz", "Hz"], ["dur", "Pulses per train"], ["trains", "Trains"], ["wait", "Wait (s)"]].map(([k, l]) => `<input type="number" min="0" step="any" data-pk="${k}" aria-label="${l}, row ${x.n}" value="${has && typeof x[k] === "number" ? x[k] : ""}"${has || base ? "" : " disabled"}>`).join("")}</div>`;
  };
  showModal(`<h2 id="mtitle">Protocol: ${esc(tab)}</h2><p class="sub">Choose from the Protocols table; a changed value reads Custom. One row with a site is unilateral, two are bilateral, in order. Saved to J2:O4; N calculates the minutes.</p>
    <div class="per hdr" aria-hidden="true"><span>Protocol</span><span>Site</span><span>Hz</span><span>Pulses/train</span><span>Trains</span><span>Wait (s)</span></div>${p.protoRows.map(rowH).join("")}
    ${fld("Next expiration date (update protocol by, G2)", `<input type="date" id="mProtoBy" value="${p.updateBy != null && p.updateBy > todaySerial() ? serialToInput(p.updateBy) : ""}" required>`)}
    ${mfoot("Save protocol", "doProtoEdit")}`, true);
}
function protoEditRead() {
  return [...$("modalBody").querySelectorAll(".per[data-pr]")].map(el => { const g = k => el.querySelector(`[data-pk="${k}"]`).value.trim(); return { n: +el.dataset.pr, name: g("name"), site: g("site"), hz: g("hz"), dur: g("dur"), trains: g("trains"), wait: g("wait") }; });
}
async function doProtoEdit() {
  const tab = state.mctx.t, rows = protoEditRead(), by = inputToSerial(mval("mProtoBy")), t = todaySerial();
  const used = rows.filter(r => r.site || r.hz || r.name);
  if (!used.some(r => r.site)) return merr("Choose a site for at least one protocol row.");
  const bad = used.find(r => !r.site || !["hz", "dur", "trains", "wait"].every(k => r[k] !== "" && !isNaN(+r[k]) && +r[k] >= 0) || !(+r.hz > 0));
  if (bad) return merr(`Row ${bad.n}: choose a site and fill Hz, pulses per train, trains and wait.`);
  if (by == null || by <= t) return merr("Choose the next expiration date, after today.");
  const pairs = [];
  rows.forEach(r => {
    const row = r.n + 1, u = used.includes(r);
    [["J", "hz"], ["K", "dur"], ["L", "trains"], ["M", "wait"]].forEach(([c, k]) => pairs.push([`${c}${row}`, u ? +r[k] : ""]));
    pairs.push([`O${row}`, u ? TmsCore.SITE_CELL[r.site] : ""]);
  });
  pairs.push(["G2", by]);
  const p = state.patients.find(x => x.tab === tab), lr = p && p.rows.find(r => r.date === t && !r.miss);
  if (lr) pairs.push([`S${lr.row}`, [...String(lr.note || "").split("; ").filter(x => x && !/^protocol renewed/i.test(x)), `Protocol renewed, next by ${fmtDate(by)}`].join("; ")]);
  closeModal();
  try { await writeCells(tab, pairs, lr ? at(lr) : null); toast(lr ? "Protocol saved. The renewal is noted on today's visit." : "Protocol saved. There's no visit today, so the renewal isn't in the log."); await refresh(); } catch (e) { fail(e); }
}
/* choosing a protocol in the Edit dialog fills its values (they stay editable) */
function protoEditPick(sel) {
  const el = sel.closest(".per"), q = state.protocols.find(x => x.name.trim() === sel.value); if (!el) return;
  [["hz", "rep"], ["dur", "pulses"], ["trains", "trains"], ["wait", "iti"]].forEach(([k, f]) => { const i = el.querySelector(`[data-pk="${k}"]`); if (q) { i.value = q[f]; i.disabled = false; } else i.value = ""; });
}

/* ---------- Schedule: the course with its payments, what's booked, and every past visit ---------- */
function payRow(p) {
  const tot = Math.max(1, p.d.total), taper = tot >= 12 ? 6 : 0, daily = tot - taper, cols = taper ? `repeat(${daily},1fr) 10px repeat(${taper},1fr)` : `repeat(${tot},1fr)`;
  const un = new Set(p.d.toConfirm.map(r => r.row)), done = p.rows.filter(r => TmsCore.isTreatment(r) && r.delivered).sort(cmp);
  let cells = "";
  for (let i = 1; i <= tot; i++) {
    if (taper && i === daily + 1) cells += "<span></span>";
    const r = done[i - 1], c = r && TmsCore.copayOf(r.note);
    cells += `<span class="${!r ? "" : un.has(r.row) ? "un" : c && c.kind !== "not" ? "pd" : ""}"></span>`;
  }
  const n = un.size;
  return `<div class="payrow" style="grid-template-columns:${cols}" role="img" aria-label="${n ? `${n} copay${n > 1 ? "s" : ""} to confirm` : "Copays logged"}">${cells}</div>
    <p class="hint">${!p.d.owes ? "No copays to log: " + (p.d.oopmMet ? `${p.d.oopmMet} OOPM met` : "the BIDF quote is $0") : n ? `${n} copay${n > 1 ? "s" : ""} to confirm` : "Every copay due is logged"}</p>`;
}
function ptSched(p, fl) {
  const d = p.d, t = todaySerial(), up = d.sched.slice(0, 8), av = availText(p);
  const vt = fl.visits.filter(v => v.tab === p.tab && v.chairVisit && ["booked", "waiting"].includes(ph(v)));
  /* a future visit takes a scheduled note, shown on its day; today's visit takes a plain note */
  const line = r => {
    const v = r.date === t && vt.find(x => x.row === r.row), sn = schedNote(r.note), own = visitNote(r.note), fut = r.date > t;
    const nb = `<button class="btn mute sm" data-act="ptNote" data-t="${esc(p.tab)}" data-row="${r.row}" data-v="${fut ? "sched" : "own"}">${fut ? (sn ? "Edit scheduled note" : "Add scheduled note") : own ? "Edit note" : "Add note"}</button>`;
    return `<div class="srow${sn ? " hasnote" : ""}"><span class="p">${fmtDay(r.date)}</span><span>${clock(r.time) || "no time"}</span><span>${esc(r.type === "F/U" ? (r.tele ? "Telehealth" : p.provider || "Provider") : chairLbl(chairFor(p, r)))}</span><span>${typeof r.tx === "number" ? `#${r.tx}` : esc(apptTypeName(r))}</span><span class="nt">${visitMarks(p, r)}${sn ? `<span class="snl">Note: ${esc(sn)}</span>` : ""}${own ? `<span class="muted">${esc(own)}</span>` : ""}</span><span class="ra">${nb}${v ? `<button class="btn ghost sm" data-act="move" data-k="${esc(v.key)}">Move</button>` : ""}</span></div>`;
  };
  /* the auth end on a day with no visit gets its own marker line among the visits */
  const authLine = p.authExp != null && p.authExp >= t && up.length && p.authExp < up[up.length - 1].date && !up.some(r => r.date === p.authExp) ? `<div class="srow"><span class="p">${fmtDay(p.authExp)}</span><span></span><span></span><span></span><span class="nt"><span class="mk mk-auth">Auth ends</span></span><span></span></div>` : "";
  const lines = up.map(r => ({ d: r.date, h: line(r) })).concat(authLine ? [{ d: p.authExp, h: authLine }] : []).sort((a, b) => a.d - b.d).map(x => x.h).join("");
  return `<div class="h"><h2>Schedule</h2><button class="btn pill" data-act="print" data-t="${esc(p.tab)}" data-v="14">Print schedule</button></div>
    <div class="hacts">${btn("avail", p.tab, "", "Edit availability")}${btn("statuspick", p.tab, "", "Change status")}</div>
    <div class="sub2">Next visits <span class="muted">${d.sched.length} booked</span></div>${up.length ? lines : '<p class="body">Nothing booked.</p>'}
    ${av ? `<p class="hint">Availability: ${esc(av)}</p>` : ""}
    <div class="actions"><button class="btn ghost sm" data-act="ptSchedView" data-t="${esc(p.tab)}">Open in Schedule</button><button class="btn ghost sm" data-act="goimport">Book the rest of the course</button>${p.status === "Paused" ? "" : `<button class="btn ghost sm" data-act="status" data-t="${esc(p.tab)}" data-v="Paused">Pause course</button>`}</div>
    <p class="hint">Book the rest in NextGen, then import the appointment list. The planner comes later.</p>
    <div class="h"><h2>Past visits</h2><button class="btn ghost sm" data-act="ptPast" aria-expanded="${!!state.ptPast}">${state.ptPast ? "Hide" : "Show"}</button></div>${state.ptPast ? pastVisits(p) : ""}`;
}
/* every logged visit, newest first: Tx #, chair, pulses, PR, the copay and who logged it, the note */
function pastVisits(p) {
  const t = todaySerial(), all = p.rows.filter(r => r.date != null && r.date <= t && (r.delivered || r.miss || r.alt || r.type === "F/U")).sort(cmp).reverse();
  if (!all.length) return '<p class="body">No visits logged yet.</p>';
  const cp = r => { const c = TmsCore.copayOf(r.note); return !c ? "" : c.kind === "collected" ? `${money(c.amt)}${c.by ? `, ${c.by}` : ""}` : c.kind === "not" ? "Not collected" : c.kind === "oopm" ? "$0, OOPM met" : c.kind === "waived" ? "Waived" : c.line; };
  return `<table class="tbl"><tr><th>Date</th><th class="num">Tx #</th><th>Visit</th><th>Chair</th><th>Pulses L / R / O</th><th class="num">PR</th><th>Copay</th><th class="num">PHQ-9</th><th class="num">GAD-7</th><th>Note</th><th></th></tr>${all.map(r => `<tr id="pvrow-${r.row}"><td>${fmtMon(r.date)}</td><td class="num">${typeof r.tx === "number" ? r.tx : ""}</td><td>${r.miss ? "Missed" : esc(r.type)}${r.tele ? " tele" : ""}${visitMarks(p, { ...r, measX: false })}</td><td>${esc(r.type === "F/U" || r.alt ? "" : chairLbl(chairFor(p, r)))}</td><td>${r.pulses.map(x => esc(x ?? "")).join(" / ")}</td><td class="num">${esc(money(r.pr))}</td><td>${esc(cp(r))}</td><td class="num">${esc(r.phq)}</td><td class="num">${esc(r.gad)}</td><td>${esc(visitNote(r.note))}</td><td><div class="ra"><button class="btn mute sm" data-act="ptNote" data-t="${esc(p.tab)}" data-row="${r.row}" data-v="own" aria-label="${ownNote(r.note) ? "Edit" : "Add"} note for ${fmtMon(r.date)}">${ownNote(r.note) ? "Edit note" : "Add note"}</button>${r.delivered ? `<button class="btn mute sm" data-act="editPulses" data-t="${esc(p.tab)}" data-row="${r.row}" aria-label="Edit pulses for ${fmtMon(r.date)}">Edit pulses</button>` : ""}</div></td></tr>`).join("")}</table>`;
}

/* ---------- Benefits (the Benefits row; an OOPM met is entered here as 0 remaining) and Financial (the copay tally against
   what's left, and copays to confirm) ---------- */
const benAmt = (na, v) => na ? "N/A" : v != null ? money(v) : "Not entered";
const factList = rows => `<dl class="facts">${rows.map(([k, v, cls]) => `<dt>${k}</dt><dd class="${cls || ""}">${esc(v)}</dd>`).join("")}</dl>`;
function ptBen(p) {
  const b = p.ben, d = p.d, daily = (p.cost[1] || {}).after;
  const ro = d.rollover ? `Re-verify for ${d.rollover.year}` : p.benUpdated != null ? `Current for ${yearOf(p.benUpdated)}` : "Not verified";
  const left = [["Insurance", p.insurance || "Not entered"], ["Copay per treatment visit", daily != null ? `${money(daily)} (90868 Daily)` : "Not entered"],
    ["BIDF estimate", b.bidf != null ? `${money(b.bidf)}, projected ${money(d.projected) || "$0.00"}` : "Not entered"]];
  const right = [["Deductible remaining", benAmt(b.dedNA, b.dedRem)], ["OOPM remaining", d.oopmMet ? `$0.00, ${d.oopmMet} OOPM met` : benAmt(b.oopNA, b.oopRem)],
    ["Last updated", p.benUpdated != null ? fmtDateY(p.benUpdated) : "Never"], ["Plan year", ro, d.rollover ? (d.rollover.bad ? "bad" : "flag") : ""]];
  return `<div class="h"><h2>Benefits</h2><span class="actions" style="margin:0"><button class="btn ghost sm" data-act="ben" data-t="${esc(p.tab)}">Edit benefits</button><button class="btn ghost sm" data-act="goimport">Re-run eligibility</button></span></div>
    <div class="g2c">${factList(left)}${factList(right)}</div>
    <p class="hint">OOPM met: Edit benefits and enter 0 in Out-of-pocket max remaining. Copay checks stop until the plan year ends. Re-run eligibility in the Sohar tool, then Import > Import Sohar benefits.</p>`;
}
function ptFin(p) {
  const b = p.ben, d = p.d, Y = d.planYear;
  const tally = d.oopmMet ? `<p><span class="tag">${d.oopmMet} OOPM met</span></p><p class="hint">No copay checks for this patient until the plan year ends.</p>`
    : !d.owes ? '<p class="body">The BIDF quote is $0, so there are no copays to track.</p>'
    : `${factList([["Logged collected", `${money(d.tally.sum) || "$0.00"} over ${d.tally.n} visit${d.tally.n === 1 ? "" : "s"}`], ["OOPM remaining", benAmt(b.oopNA, b.oopRem)], ["Deductible remaining", benAmt(b.dedNA, b.dedRem)], ["Cost so far", money(d.accrued) || "$0.00"]])}
      ${d.oopmLikely ? `<p class="warnline">Out-of-pocket max likely met: set the copay to $0 at pre-check and update Insurance Maintenance in NextGen.</p><p class="hint">Logged copays are a floor: claims applied to the deductible after adjudication count too. Confirm in NextGen or re-run eligibility, then enter 0 in Benefits.</p>` : ""}`;
  return `<div class="h"><h2>Financial</h2><span>${d.owes && !d.oopmMet ? `Copay tally ${Y}` : ""}</span></div>${payRow(p)}
    <div class="g2c"><div>${tally}</div><div><div class="sub2">Copays to confirm</div>${toConfirmHtml(p)}</div></div>`;
}
function toConfirmHtml(p) {
  const L = p.d.toConfirm, pick = state.ptPick || new Set();
  if (!p.d.owes) return '<p class="body">None needed.</p>';
  if (!L.length) return '<p class="body">Every copay due is logged.</p>';
  return `${L.map(r => `<label class="chk"><input type="checkbox" data-pick="${r.row}"${pick.has(r.row) ? " checked" : ""}> ${fmtMon(r.date)}${typeof r.tx === "number" ? `, #${r.tx}` : ""}: ${money(payOf(p, r))}</label>`).join("")}
    <div class="actions"><button class="btn ghost sm" data-act="copayMark" data-t="${esc(p.tab)}" data-v="sel">Mark selected paid</button><button class="btn ghost sm" data-act="copayMark" data-t="${esc(p.tab)}" data-v="all">Mark all paid</button><button class="btn ghost sm" data-act="copayMark" data-t="${esc(p.tab)}" data-v="waive">Waive</button></div>`;
}
/* "Copay $X collected (confirmed m/d by Name)" or "Copay waived by Name m/d" on each visit's Notes */
async function copayMark(tab, how) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return;
  const pick = state.ptPick || new Set(), L = p.d.toConfirm.filter(r => how === "all" || pick.has(r.row));
  if (!L.length) return toast("Tick the visits first.", true);
  const by = firstName((($("pvBy") || {}).value || "").trim() || state.techName); if (!by) return toast("Enter your name in By on today's copay line, or complete a session first.", true);
  const t = fmtDate(todaySerial()), ops = L.map(r => [r, withCopay(r.note, how === "waive" ? `Copay waived by ${by} ${t}` : `Copay ${money(payOf(p, r))} collected (confirmed ${t} by ${by})`)]);
  try {
    await applyOps(ops.map(([r, note]) => ({ tab, a: `S${r.row}`, v: [[note]] })), ops.map(([r]) => ({ tab, ...at(r) })));
    ops.forEach(([r, note]) => { r.note = note; }); p.d = derive(p); state.ptPick = new Set(); renderDash();
    toast(`${L.length} copay${L.length > 1 ? "s" : ""} ${how === "waive" ? "waived" : "marked paid"}.`);
  } catch (e) { fail(e); }
}

/* ---------- Chairs: devices, chair hours and protocols ---------- */
function viewChairs(fl) {
  const cards = fl.chairs.map(c => {
    const vs = fl.visits.filter(v => v.chairVisit && (v.chair === c.id || v.offChair === c.id)), s = c.off ? null : TmsCore.chairState(c.id, fl.visits, fl.now), run = s && (s.state === "track" || s.state === "soon");
    return `<div class="card chair${c.off ? " off" : ""}"><div class="hd"><span class="nm">${esc(c.label)}</span><span class="state" data-state="${run ? "track" : "open"}"><span class="dot"></span>${c.off ? "Unavailable" : run ? `In session, ${esc(s.v.tab)}` : "Available"}</span></div>
      <dl class="facts"><dt>On the list as</dt><dd>${esc(c.id)}</dd><dt>Device</dt><dd>${esc(c.device)}</dd><dt>Coil</dt><dd>${esc(c.coil || "Not entered")}</dd><dt>Today</dt><dd>${vs.filter(v => ph(v) === "done").length} done, ${vs.filter(v => ["booked", "waiting"].includes(ph(v))).length} still to come</dd><dt>Last service</dt><dd>${c.service != null ? fmtMon(c.service) : "Not logged"}</dd></dl>
      <div class="actions">${c.off ? `<button class="btn ghost sm" data-act="chairOn" data-v="${esc(c.id)}">Mark available</button>` : `<button class="btn ghost sm" data-act="chairOff" data-v="${esc(c.id)}">Mark unavailable</button>`}<button class="btn mute sm" data-act="chairEdit" data-v="${esc(c.id)}">Log service or coil</button></div></div>`;
  }).join("");
  const cl = state.clinic, hrs = cl ? `<table class="tbl"><tr><th>Day</th><th>Open</th><th>Close</th><th>Break</th></tr>${[1, 2, 3, 4, 5].map(i => { const h = cl.days[i]; return `<tr><td>${LONG_DAYS[i]}</td><td>${clock(h.open) || "–"}</td><td>${clock(h.close) || "–"}</td><td>${h.bs != null && h.be != null ? `${clock(h.bs)} to ${clock(h.be)}` : ""}</td></tr>`; }).join("")}</table>
      ${state.closures.size ? `<p class="body">Closed ${[...state.closures].filter(([d]) => d >= fl.t).sort((a, b) => a[0] - b[0]).slice(0, 6).map(([d, w]) => `${fmtMon(d)} (${esc(w)})`).join(", ") || "on no upcoming dates"}.</p>` : ""}
      ${cl.closureTable ? "" : '<p class="flag">Add the holiday closure table to the Clinic Schedule sheet. <button class="btn ghost sm" data-act="mkSched">Add it</button></p>'}`
    : '<p class="body">There is no Clinic Schedule sheet yet. It holds the Monday to Friday chair hours everyone sees.</p><button class="btn ghost" data-act="mkSched">Create Clinic Schedule sheet</button>';
  return head("Chairs", `${device() === "Brainsway" ? "BrainsWay" : "MagV"} chairs on this tracker, from the appointment list. A chair marked unavailable leaves the board and the schedule, and its bookings wait for another chair.`, false)
    + `<div class="cgrid">${cards}</div>
    <div class="card list" style="margin-top:24px"><div class="h"><h2>Chair hours</h2><span>From the Clinic Schedule sheet</span></div>${hrs}<p class="body">Chair time: Daily 30 min, MT and MT redo 60 min.</p></div>
    <div class="card list" style="margin-top:24px"><div class="h"><h2>Protocols</h2><span>From the Protocols sheet</span></div>${viewProtocols()}</div>`;
}

/* ---------- kept from the previous screens ---------- */
function protoEditor(p, keyBase) {
  const opts = v => `<option value="">(none)</option>` + state.protocols.map(q => `<option value="${esc(q.name)}"${q.name.trim() === v ? " selected" : ""}>${esc(q.name.trim())}</option>`).join("");
  return `<div class="pe">${p.protoRows.map(x => `<div class="g2p"><span>Protocol ${x.n}</span><select data-pname="${x.n}">${opts(protoName(x))}</select><select data-psite="${x.n}"><option value="">Site</option>${SITES.map(s => `<option${s === x.site ? " selected" : ""}>${s}</option>`).join("")}</select></div>`).join("")}
    ${bt("protoSave", p.tab, 0, keyBase, "Save protocol to J2:O4")}</div>`;
}
function variances(day, byChair) {
  const P = state.patients.filter(p => (!byChair || inChair(p)) && (["In Progress", "Tapering"].includes(p.status) || (p.status === "" && p.d.txDone > 0)));
  const ws = weekStart(day);
  const grads = state.patients.filter(p => onSched(p) && (!byChair || inChair(p)) && p.d.txDone >= p.d.total && p.d.lastTx != null && p.d.lastTx >= ws && p.d.lastTx <= day);
  const inTx = P.filter(p => p.d.txDone < p.d.total || p.rows.some(r => r.date === day && r.type !== "F/U"));
  const treated = inTx.filter(p => p.rows.some(r => r.date === day && r.type !== "F/U" && !r.miss && !r.alt));
  const groups = { "In follow-up": [], "Missed": [], "Away": [], "At another clinic": [], "Not scheduled": [] };
  for (const p of inTx.filter(p => !treated.includes(p))) {
    const rs = p.rows.filter(r => r.date === day), away = p.avail.find(w => w.type === "Unavailable" && w.from <= day && day <= w.to);
    if (rs.some(r => r.type === "F/U" && !r.miss)) groups["In follow-up"].push([p, fmtTime(rs.find(r => r.type === "F/U").time)]);
    else if (rs.some(r => r.miss)) groups["Missed"].push([p, OUTCOMES.find(k => rs.find(r => r.miss).note.includes(k)) || "Missed"]);
    else if (away) groups["Away"].push([p, `${fmtDate(away.from)}\u2013${fmtDate(away.to)}${away.note ? ` (${away.note})` : ""}`]);
    else if (rs.some(r => r.alt)) groups["At another clinic"].push([p, ""]);
    else groups["Not scheduled"].push([p, p.d.next ? `next ${fmtDate(p.d.next.date)}` : "nothing booked"]);
  }
  return { expected: inTx.length, actual: treated.length, groups, grads };
}
const lpool = t => t === "F/U" ? "F/U" : "TMS";
function listCheck(day) {
  const lx = state.lx; if (!lx || !lx.ap || day < lx.lo || day > lx.hi) return null;
  const ch = chairPick(), latest = a => state.patients.filter(p => a.mrn ? p.mrn === a.mrn : false).sort((x, y) => y.course - x.course)[0]
    || (a.code ? state.patients.filter(p => !p.mrn && p.code === a.code).sort((x, y) => y.course - x.course)[0] : null);
  const ap = lx.ap.filter(a => a.date === day && (ch === "all" || !lx.v[`${a.mrn}|${day}`] || lx.v[`${a.mrn}|${day}`] === ch));
  const miss = [], extra = [], on = r => r.date === day && !r.miss && !r.alt, groups = new Map();
  ap.forEach(a => { const p = latest(a), k = `${p ? p.tab : "#" + a.mrn}|${lpool(a.type)}`; if (!groups.has(k)) groups.set(k, { p, pool: lpool(a.type), list: [] }); groups.get(k).list.push(a); });
  for (const { p, pool, list } of groups.values()) {
    if (!p) { list.forEach(a => miss.push({ a, why: "new" })); continue; }
    if (!onSched(p)) { list.forEach(a => miss.push({ a, p, why: "disc" })); continue; }
    list.slice(p.rows.filter(r => on(r) && lpool(r.type) === pool).length).forEach(a => miss.push({ a, p, why: "tab" }));
  }
  for (const p of state.patients.filter(onSched)) for (const pool of ["TMS", "F/U"]) {
    const rs = p.rows.filter(r => on(r) && lpool(r.type) === pool && inChair(p, r)).sort((x, y) => x.delivered - y.delivered || cmp(x, y));   // open visits are named first
    const n = (groups.get(`${p.tab}|${pool}`) || { list: [] }).list.length;
    rs.slice(0, Math.max(0, rs.length - n)).forEach(r => extra.push({ p, r }));
  }
  return { n: ap.length, miss, extra };
}
function hoursOf(day) {
  if (state.closures.has(day)) return { closed: true, txt: `Closed: ${state.closures.get(day)}` };
  const h = state.clinic?.days[dow(day)];
  if (!h || h.open == null || h.close == null) return null;
  return { closed: false, txt: `${fmtTime(h.open)}–${fmtTime(h.close)}${h.bs != null && h.be != null ? `, break ${fmtTime(h.bs)}–${fmtTime(h.be)}` : ""}` };
}
function dayEntries(day) {
  const chair = [], prov = [], noTime = [];
  for (const p of state.patients.filter(onSched)) for (const r of p.rows) {
    if (r.date !== day || r.miss || r.alt) continue;
    if (r.time == null) { noTime.push({ p, r }); continue; }
    if (chairMin(r.type)) chair.push({ p, r, start: r.time, end: r.time + chairMin(r.type) });
    if (provMin(r.type)) prov.push({ p, r, start: r.time, end: r.time + provMin(r.type) });
  }
  return { chair, prov, noTime };
}
function slotsFor(day, ent) {
  const cl = state.clinic, h = cl?.days[dow(day)], step = cl?.slot || 30;
  let open = h?.open, close = h?.close;
  if (open == null || close == null) { if (!ent.chair.length) return null; open = Math.floor(Math.min(...ent.chair.map(e => e.start)) / step) * step; close = Math.max(...ent.chair.map(e => e.end)); }
  const out = []; for (let s = open; s < close; s += step) out.push({ s, e: s + step, brk: !!(h && h.bs != null && h.be != null && s >= h.bs && s < h.be) });
  return { slots: out, open, close, step, h };
}
async function mkSchedSheet() {
  if (POPOUT) return callPane("mkSchedSheet", []);
  try {
    await Excel.run(async ctx => {
      let ws = ctx.workbook.worksheets.getItemOrNullObject("Clinic Schedule"); await ctx.sync();
      if (ws.isNullObject) {
        ws = ctx.workbook.worksheets.add("Clinic Schedule");
        ws.getRange("A1").values = [[`Clinic hours template (${chairName()} chair). Enter hours for each weekday and any closure dates. Everyone sees the daily schedule on the Chair tab of the add-in.`]];
        ws.getRange("A3:E3").values = [["Day", "Open", "Close", "Break start", "Break end"]];
        ws.getRange("A4:A8").values = [["Monday"], ["Tuesday"], ["Wednesday"], ["Thursday"], ["Friday"]];
        ws.getRange("B4:E8").numberFormat = Array(5).fill(Array(4).fill("h:mm AM/PM"));
        ws.getRange("A10:B10").values = [["Slot length (minutes)", 30]];
        ws.getRange("A12").values = [["Chair time: Daily 30 min, MT 60 min, MTR 60 min. Provider time: MT 60, MTR 30, F/U 30. Only Kaiser patients may have more than one appointment in a day."]];
        ["A1", "A3:E3", "A10"].forEach(a => ws.getRange(a).format.font.bold = true);
        ws.getRange("A3:E3").format.fill.color = "#e2f1f1"; ws.getRange("B4:E8").format.fill.color = "#fff8dc"; ws.getRange("B10").format.fill.color = "#fff8dc";
        ws.getRange("A:A").format.columnWidth = 150; ws.getRange("B:E").format.columnWidth = 90;
      }
      const g3 = ws.getRange("G3"); g3.load("values"); await ctx.sync();
      if (g3.values[0][0] !== "Closed date") {
        ws.getRange("G2").values = [["Holiday and closure dates (the clinic is closed all day)"]]; ws.getRange("G2").format.font.bold = true;
        ws.getRange("G3:H3").values = [["Closed date", "Reason"]]; ws.getRange("G3:H3").format.font.bold = true; ws.getRange("G3:H3").format.fill.color = "#e2f1f1";
        ws.getRange("G4:G40").numberFormat = [["m/d/yyyy"]]; ws.getRange("G4:H40").format.fill.color = "#fff8dc";
        ws.getRange("G:G").format.columnWidth = 90; ws.getRange("H:H").format.columnWidth = 220;
      }
      ws.activate(); await ctx.sync();
    });
    toast("Clinic Schedule is ready. Enter hours and closure dates, then refresh."); await refresh();
  } catch (e) { fail(e); }
}
async function writeChairSheet(offset = state.schedOffset) {
  if (POPOUT) return callPane("writeChairSheet", [offset]);
  try {
    const mon = weekStart(todaySerial() + offset) + 1;
    const days = [0, 1, 2, 3, 4].map(i => mon + i), ents = days.map(dayEntries), sls = days.map((d, i) => slotsFor(d, ents[i]));
    const starts = sls.filter(Boolean).flatMap(s => [s.open, s.close]); if (!starts.length) return toast("Enter chair hours on the Clinic Schedule sheet first.", true);
    const step = state.clinic?.slot || 30, lo = Math.min(...starts), hi = Math.max(...starts), times = []; for (let s = lo; s < hi; s += step) times.push(s);
    await Excel.run(async ctx => {
      let ws = ctx.workbook.worksheets.getItemOrNullObject("Chair Schedule"); await ctx.sync();
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add("Chair Schedule"); else ws.getRange().clear();
      ws.getRange("A1").values = [[`${chairName()} chair, week of ${fmtDateY(mon)}. Snapshot ${fmtDateY(todaySerial())}. Re-run from the add-in to refresh.`]]; ws.getRange("A1").format.font.bold = true;
      ws.getRange("A3:F3").values = [["Time", ...days.map(d => `${DAYS[dow(d)]} ${fmtDate(d)}`)]]; ws.getRange("A3:F3").format.font.bold = true; ws.getRange("A3:F3").format.fill.color = "#e2f1f1";
      const grid = times.map(s => [fmtTime(s), ...days.map((d, i) => {
        const sl = sls[i]; if (state.closures.has(days[i])) return `CLOSED: ${state.closures.get(days[i])}`; if (!sl || s < sl.open || s >= sl.close) return "closed";
        const occ = ents[i].chair.filter(e => e.start < s + step && e.end > s);
        return occ.map(e => e.start >= s ? `${e.p.tab} ${e.r.type}` : "\u21B3").join(" / ");
      })]);
      ws.getRangeByIndexes(3, 0, grid.length, 6).values = grid;
      grid.forEach((row, ri) => row.slice(1).forEach((v, ci) => {
        const c = ws.getRangeByIndexes(3 + ri, 1 + ci, 1, 1), s = times[ri], occ = ents[ci].chair.filter(e => e.start < s + step && e.end > s);
        if (v === "closed" || v.startsWith("CLOSED")) c.format.fill.color = "#e8ecef"; else if (chairClash(occ)) c.format.fill.color = "#fbe9e7"; else if (occ.some(e => e.r.type === "MT" || e.r.type === "MTR")) c.format.fill.color = "#fdf1d8"; else if (occ.length) c.format.fill.color = "#e2f1f1";
      }));
      ws.getRange("A:A").format.columnWidth = 80; ws.getRange("B:F").format.columnWidth = 120; ws.activate(); await ctx.sync();
    });
    toast("Chair Schedule sheet updated.");
  } catch (e) { fail(e); }
}
/* Patients: the roster and week at a glance */
function dayCode(p, day) {
  const rs = p.rows.filter(r => r.date === day); if (!rs.length) return "";
  if (rs.some(r => r.miss)) return "X";
  const sp = rs.find(r => STD_VISIT.includes(r.type)); if (sp) return sp.type;
  if (rs.some(r => r.alt)) return "Alt";
  const a = rs.find(r => r.action); if (a) return a.action;
  return "\u00B7";
}
function availEditor(av, prefix = "pv") {
  const wk = [1, 2, 3, 4, 5].map(d => { const w = av.week[d] || {}; return `<div class="g3w"><span>${DAYS[d]}</span><input type="time" data-avw="${d}" data-k="from" value="${minToInput(w.from)}" aria-label="${DAYS[d]} earliest start"><input type="time" data-avw="${d}" data-k="to" value="${minToInput(w.to)}" aria-label="${DAYS[d]} latest start"></div>`; }).join("");
  const rg = av.ranges.map((w, i) => `<div class="g4r"><select data-avr="${i}" data-k="type">${AV_TYPE.map(x => `<option${x === w.type ? " selected" : ""}>${x}</option>`).join("")}</select><input type="date" data-avr="${i}" data-k="from" value="${w.from ? serialToInput(w.from) : ""}"><input type="date" data-avr="${i}" data-k="to" value="${w.to ? serialToInput(w.to) : ""}"><input data-avr="${i}" data-k="note" value="${esc(w.note)}" placeholder="Reason"><button class="sm" data-act="avDel" data-v="${i}" aria-label="Remove">\u00D7</button></div>`).join("");
  return `<div class="sub">Weekly (start-time window)</div><div class="g3w hdr"><span></span><span>Earliest start</span><span>Latest start</span></div>${wk}
    <div class="sub" style="margin-top:6px">Dates (vacation, work, other)</div>${rg || '<div class="none">None entered</div>'}<button class="sm" data-act="avAdd">Add another</button>`;
}
function hhmm(v) { if (typeof v !== "number") return ""; const h = Math.floor(v * 24 + 1e-6), m = Math.round((v * 24 - h) * 60); return `${h}:${String(m).padStart(2, "0")}`; }
function viewProtocols() {
  if (!state.protocols.length) return '<p class="body">No Protocols sheet found.</p>';
  const rows = state.protocols.map(p => {
    const calc = [p.rep, p.pulses, p.trains, p.iti].every(x => typeof x === "number" && x > 0) ? ((p.pulses / p.rep) + p.iti) * p.trains / 60 : null;
    const mm = calc == null ? "" : `${Math.floor(calc)}:${String(Math.round((calc % 1) * 60)).padStart(2, "0")}`;
    return `<tr><td>${esc(p.name)}</td><td class="num">${p.rep}</td><td class="num">${p.pulses}</td><td class="num">${p.trains}</td><td class="num">${p.iti}</td><td class="num">${p.inten ?? ""}</td><td class="num">${hhmm(p.typed)}</td><td class="num">${mm}</td></tr>`;
  }).join("");
  return `<table class="tbl"><tr><th>Protocol</th><th class="num">Hz</th><th class="num">Pulses</th><th class="num">Trains</th><th class="num">ITI</th><th class="num">% MT</th><th class="num">Time</th><th class="num">Calc</th></tr>${rows}</table>
    <p class="hint">Time is Tx Time on the Protocols sheet as entered (min:sec). Calc is ((pulses / Hz) + ITI) x trains / 60, the same formula as each tab's N2. HF 10Hz is the standard name; X is the internal copy.</p>
    <div class="actions"><button class="btn ghost sm" data-sheet="Protocols">Open Protocols sheet</button></div>`;
}

/* ---------- dialogs over the add-in (blurred backdrop). Each opens from a data-act and saves through the functions above. ---------- */
const mfoot = (label, act, extra) => `<p class="err" id="merr" role="alert"></p><div class="mfoot">${extra || ""}<button class="btn mute" data-mclose>Cancel</button><button class="btn primary" data-act="${act}">${esc(label)}</button></div>`;
const merr = m => { const e = $("merr"); if (e) e.textContent = m; };
const mval = id => { const e = $(id); return e ? String(e.value).trim() : ""; };
const radio = name => { const e = document.querySelector(`#modalBody input[name="${name}"]:checked`); return e ? e.value : ""; };
const fld = (label, inner) => `<label class="fld"><span>${label}</span>${inner}</label>`;
const isInt = (s, lo, hi) => /^\d+$/.test(s) && +s >= lo && +s <= hi;
const lastTech = p => (p.rows.filter(r => r.tech && r.date != null).sort(cmp).pop() || {}).tech || "";
const techNames = () => [...new Set(state.patients.flatMap(p => p.rows.map(r => r.tech)).filter(Boolean))].sort();
const techInput = cur => `<input id="mTech" list="mTechs" value="${esc(cur || "")}" autocomplete="off"><datalist id="mTechs">${techNames().map(n => `<option value="${esc(n)}">`).join("")}</datalist>`;

/* what the protocol calls for, per site (J:M rows with a site in O) */
function wantPulses(p) {
  const want = { Left: 0, Right: 0, Other: 0 };
  p.protoRows.filter(x => isNum(x.dur) && isNum(x.trains)).forEach(x => { want[SITES[TmsCore.pulseCol(TmsCore.siteOf(x.site) || "F3")]] += x.dur * x.trains; });
  return want;
}
function pulseFields(p, had) {
  const want = had ? Object.fromEntries(SITES.map((s, i) => [s, had[i]])) : wantPulses(p), any = SITES.some(s => want[s]);
  return `<div class="g3">${SITES.map((s, i) => !any || want[s] || i === 0 || had ? fld(`${s} pulses`, `<input type="number" min="0" id="mPul${i}" value="${want[s] || ""}">`) : `<input type="hidden" id="mPul${i}" value="">`).join("")}</div>`;
}
/* Edit pulses on a visit that has them (Today's card, or Past visits): all three sites and the technician, a cleared site included */
function mEditPulses(tab, row) {
  const [p, r] = findPR(tab, row); if (!r || !r.delivered) return; state.mctx = { t: tab, row };
  showModal(`<h2 id="mtitle">Edit pulses for ${fmtDay(r.date)}: ${esc(tab)}</h2><p class="sub">${typeof r.tx === "number" ? `Treatment ${r.tx}. ` : ""}Change what was delivered. Clearing a site removes it from the row.</p>
    ${pulseFields(p, r.pulses)}${fld("Treating technician", techInput(r.tech || state.techName))}${mfoot("Save pulses", "doEditPulses")}`);
}
async function doEditPulses() {
  const { t: tab, row } = state.mctx, [, r] = findPR(tab, row), vals = readPulses(), tech = mval("mTech"); if (!r) return closeModal();
  if (!vals.some(n => n !== "" && n > 0)) return merr("Enter the pulses delivered. To record a miss instead, undo the session first.");
  if (!tech) return merr("Enter the treating technician.");
  closeModal();
  try {
    await writeCells(tab, [[`F${row}`, vals[0] || ""], [`H${row}`, vals[1] || ""], [`J${row}`, vals[2] || ""], [`B${row}`, tech]], at(r));
    state.techName = tech; patchRow(tab, row, { tech, pulses: vals.map(x => x || null) }); toast("Pulses updated.");
  } catch (e) { fail(e); }
}
const readPulses = () => SITES.map((s, i) => { const x = mval(`mPul${i}`); return x === "" ? "" : +x; });
/* the next visit that collects measures, as the old Measures needed box chose: an MT, MT redo or F/U within 7 days, else the next visit */
function measDest(p, r) {
  const up = measUp(p, r);
  return up.find(x => STD_VISIT.includes(x.type) && x.date - r.date <= 7) || up[0] || null;
}
const measUp = (p, r) => p.rows.filter(x => x.date != null && r.date != null && x.date > Math.max(r.date, todaySerial() - 1) && !x.delivered && !x.miss).sort(cmp);
const measPick = (p, r, dest) => `<select id="mMeasTo" aria-label="Move the measures to">${measUp(p, r).slice(0, 20).map(x => `<option value="${x.row}"${x === dest ? " selected" : ""}>${fmtDay(x.date)}, ${esc(apptTypeName(x))}</option>`).join("")}</select>`;

/* "Copay $25.00 collected" (or not collected) in the visit's Notes; a later answer replaces it */
const withCopay = TmsCore.withCopayLine;
const firstName = s => String(s || "").trim().split(/[\s,]+/)[0] || "";
/* the copay line with who answered and when: "Copay $25.00 collected by Priya 11:32 AM" */
const copayLine = (kind, amt, by) => (kind === "oopm" ? "Copay $0, OOPM met" : `Copay ${money(amt)} ${kind === "not" ? "not collected" : "collected"}`) + (by ? ` by ${firstName(by)} ${clock(nowMin())}` : "");
function mEnd(v) {
  const p = v.p, nx = p.d.sched.find(r => r.row !== v.row && r.date > v.date) || p.d.sched.find(r => r.row !== v.row), dest = measDest(p, v.r), dr = draftOf(p, v.r);
  /* the copay asked here is the patient view's copay line too: an answer there is already chosen (OOPM met: not asked) */
  const cl = TmsCore.copayOf(v.r.note), pay = cl && cl.kind === "oopm" || !p.d.owes ? 0 : payOf(p, v.r);
  state.mctx = { k: v.key, pay: pay > 0 ? pay : null };
  showModal(`<h2 id="mtitle">Complete ${esc(v.tab)}'s session</h2><p class="sub">Treatment ${sessNo(v)} of ${p.d.total}. Protocol: ${esc(protoText(p))}. Pulses are filled from the protocol; change them if delivery differed.</p>
    ${pulseFields(p)}${fld("Treating technician", techInput(v.r.tech || state.techName || lastTech(p)))}
    ${pay > 0 ? `<h3>Copay today: ${esc(money(pay))}</h3><div class="opts"><label class="opt"><input type="radio" name="copay" value="collected"${cl && cl.kind === "collected" ? " checked" : ""}><span>Collected</span></label><label class="opt"><input type="radio" name="copay" value="not"${cl && cl.kind === "not" ? " checked" : ""}><span>Not collected</span></label></div>` : ""}
    <h3>Side effects</h3>
    <div class="qrow"><span>Any pain or discomfort during treatment?</span><select id="mPain"><option${dr.a.pain === "Yes" ? "" : " selected"}>No</option><option${dr.a.pain === "Yes" ? " selected" : ""}>Yes</option></select><input id="mPainD" class="sp" placeholder="What and where (needed if yes)" value="${esc(dr.d.pain || "")}"${dr.a.pain === "Yes" ? "" : " hidden"}></div>
    <div class="qrow"><span>Involuntary muscle movements?</span><select id="mMus">${["Yes (normal)", "No"].map(o => `<option${dr.a.mus === o ? " selected" : ""}>${o}</option>`).join("")}</select></div>
    ${v.r.measX ? `<h3>PHQ-9 and GAD-7 due today</h3><div class="g2">${fld("PHQ-9 (0 to 27)", '<input type="number" min="0" max="27" id="mPhq">')}${fld("GAD-7 (0 to 21)", '<input type="number" min="0" max="21" id="mGad">')}</div>
      ${dest ? `<label class="check"><input type="checkbox" id="mMeasMove"> Not collected today: move them to</label>${measPick(p, v.r, dest)}` : ""}` : ""}
    <p class="body">Next visit: ${nx ? `${fmtLong(nx.date)}, ${clock(nx.time) || "no time"}` : "none booked. Book it in NextGen."}</p>
    ${mfoot("Complete session", "doEnd")}`, true);
}
async function doEnd() {
  const v = findVisit(state.mctx.k); if (!v) return closeModal();
  const vals = readPulses(), tech = mval("mTech"), pain = mval("mPain"), painD = mval("mPainD");
  if (!vals.some(x => x !== "" && x > 0)) return merr("Enter the pulses delivered.");
  if (!tech) return merr("Enter the treating technician.");
  if (pain === "Yes" && !painD) return merr("Say what the pain or discomfort was.");
  const pay = state.mctx.pay, cp = pay ? radio("copay") : "";
  if (pay && !cp) return merr(`Say whether the ${money(pay)} copay was collected.`);
  let meas = null;
  if (v.r.measX) {
    const mv = $("mMeasMove") && $("mMeasMove").checked, P = mval("mPhq"), Q = mval("mGad");
    if (!mv && !(isInt(P, 0, 27) && isInt(Q, 0, 21))) return merr("Enter PHQ-9 (0 to 27) and GAD-7 (0 to 21), or tick Not collected today.");
    meas = mv ? { move: v.p.rows.find(x => x.row === +mval("mMeasTo")) } : { P: +P, Q: +Q };
  }
  const dr = draftOf(v.p, v.r); dr.a.pain = pain; dr.d.pain = painD; dr.a.mus = mval("mMus"); if (!dr.initials) dr.initials = state.initials || initialsOf(tech);
  closeModal();
  if (!await markPulses(v.tab, v.row, { vals, tech, note: cp ? withCopay(v.r.note, copayLine(cp, pay, tech)) : null })) return;
  if (meas && meas.move) await moveMeasures(v.tab, v.row, meas.move.row); else if (meas) await saveMeasures(v.tab, v.row, meas);
  if (!await sessWrite(v, { end: nowMin(), ...(v.start == null ? { start: v.checkin ?? nowMin() } : {}) }, `Treatment ${sessNo(v)} saved to ${v.tab}.`)) return;
  const nv = findVisit(v.key), p = state.patients.find(x => x.tab === v.tab);
  /* the last approved treatment (36 of 36): ask to mark the course Complete, then the note */
  if (p && p.d.txDone >= p.d.total && !/^(complete|discontinued)$/i.test(p.status)) return mFinish(p, nv);
  if (nv && !state.copied.has(nv.key)) mNote(nv);   // the drafted note, ready to copy into NextGen
}
function mFinish(p, nv) {
  state.mctx = { t: p.tab, k: nv ? nv.key : null };
  showModal(`<h2 id="mtitle">Treatment ${p.d.txDone} of ${p.d.total} done: ${esc(p.tab)}</h2><p class="body">Every approved treatment is done. Mark the course Complete? It drops off Active Patients. If an extension is coming, leave it as ${esc(p.status || "it is")} for now.</p>
    ${mfoot("Mark Complete", "finishYes", '<button class="btn ghost" data-act="finishNo">Not yet</button>')}`);
}
function mMt(v) {
  const first = v.type === "MT", g = k => { const m = v.r.note.match(new RegExp(`${k}:\\s*([^;]+)`, "i")); return m ? m[1].trim() : ""; };
  state.mctx = { k: v.key };
  showModal(`<h2 id="mtitle">${first ? "Mapping and motor threshold" : "Motor threshold redo"}: ${esc(v.tab)}</h2><p class="sub">Saved to the ${first ? "MT" : "MT redo"} row's Notes.</p>
    <div class="g3" data-mtbox>${fld("MT %", `<input data-mt="MT %" value="${esc(g("mt %"))}">`)}${fld("Hotspot", `<input data-mt="Hotspot" value="${esc(g("hotspot"))}">`)}${fld("Amps", `<input data-mt="Amps" value="${esc(g("amps"))}">`)}
    ${first ? fld("Baseline sleep", `<input data-mt="Baseline sleep" value="${esc(g("baseline sleep"))}" placeholder="e.g. 7 hrs">`) + fld("Baseline caffeine", `<input data-mt="Baseline caffeine" value="${esc(g("baseline caffeine"))}" placeholder="e.g. 2 cups">`) : ""}</div>
    ${first ? `<h3>Protocol chosen today</h3><p class="sub">Fills Hz, pulses, trains and wait in J2:M4; the site goes in O.</p>${protoEditor(v.p, "mt").replace(/<button[^>]*data-act="protoSave"[^>]*>[^<]*<\/button>/, "")}` : ""}
    ${mfoot("Save motor threshold", "doMt")}`, true);
}
async function doMt() {
  const v = findVisit(state.mctx.k); if (!v) return closeModal();
  const parts = [...document.querySelectorAll("#modalBody [data-mt]")].filter(i => i.value.trim()).map(i => `${i.dataset.mt}: ${i.value.trim()}`);
  const pct = (parts.find(x => /^MT %/.test(x)) || "").replace(/^MT %:\s*/, "").replace("%", "");
  if (!isInt(pct, 20, 100)) return merr("Enter the MT as a whole percent, 20 to 100.");
  const pe = document.querySelector("#modalBody .pe"), picks = pe ? [1, 2, 3].map(n => ({ name: pe.querySelector(`select[data-pname="${n}"]`).value, site: pe.querySelector(`select[data-psite="${n}"]`).value })) : null;
  if (picks && !picks.some(x => x.name)) return merr("Choose the protocol chosen today.");
  closeModal();
  if (!await saveMT(v.tab, v.row, parts)) return;
  if (picks) await saveProtocol(v.tab, picks);
}
function mMiss(v) {
  const p = v.p, left = p.d.left, after = left - 1, streak = TmsCore.missStreak(p.rows, todaySerial());
  const warn = left <= 0 ? "No freebies left." : after <= 1 ? `Using one leaves ${after} freebie${after === 1 ? "" : "s"}.` : "";
  state.mctx = { k: v.key, t: v.tab, row: v.row };
  showModal(`<h2 id="mtitle">Record missed visit: ${esc(v.tab)}, ${clock(v.time) || "today"}</h2><p class="sub">${left} freebie${left === 1 ? "" : "s"} left.${streak >= 2 ? ` This would be ${streak + 1} missed in a row.` : ""}</p>
    <div class="opts">${OUTCOMES.map((k, i) => `<label class="opt"><input type="radio" name="kind" value="${k}"${i ? "" : " checked"}><span>${k}</span></label>`).join("")}</div>
    <h3>Freebie</h3>
    <div class="opts"><label class="opt"><input type="radio" name="fb" value="Used"${left <= 0 ? " disabled" : ""}><span>Use freebie${warn ? `<small class="warnline">${warn}</small>` : ""}</span></label>
      <label class="opt"><input type="radio" name="fb" value="Waived"><span>Waived</span></label>
      <label class="opt"><input type="radio" name="fb" value="" checked><span>Decide later<small>Stays on Patients until it's decided</small></span></label></div>
    ${mfoot("Record missed visit", "doMiss")}`);
}
async function doMiss() {
  const v = findVisit(state.mctx.k) || (state.mctx.t && { tab: state.mctx.t, row: state.mctx.row }); if (!v) return closeModal();
  const kind = radio("kind"), fb = radio("fb"); closeModal();
  await recordMiss(v.tab, v.row, kind, fb);
}
/* a missed visit: x in M, the outcome (and the freebie decision, if made) in Notes, one freebie off N6 when used */
async function recordMiss(tab, row, kind, fb) {
  const [p, r] = findPR(tab, row); if (!r || !kind) return;
  const note = withFreebie((r.note ? r.note + "; " : "") + `${kind} ${fmtDate(r.date ?? todaySerial())} (marked in add-in)`, fb), left = p.d.left;
  const pairs = [[`M${row}`, "x"], [`S${row}`, note]]; if (fb === "Used") pairs.push([p.fbCell, Math.max(0, left - 1)]);
  try {
    await writeCells(tab, pairs, at(r)); if (fb === "Used") p.freebiesLeft = Math.max(0, left - 1);
    patchRow(tab, row, { miss: true, note, freebie: fb || "" });
    const n = TmsCore.missStreak(p.rows, todaySerial());
    toast(n >= 3 ? `${tab}: ${n} missed in a row. Call to reschedule is on Patients.` : fb === "Used" && left - 1 <= 1 ? `${kind} recorded. ${tab} has ${Math.max(0, left - 1)} freebie${left - 1 === 1 ? "" : "s"} left.` : `${kind} recorded for ${tab}.`);
  } catch (e) { fail(e); }
}
function freebieModal(tab, row) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const kind = OUTCOMES.find(k => r.note.includes(k)) || "Missed visit", left = p.d.left, after = left - 1;
  const warn = left <= 0 ? "This patient has no freebies left." : after === 1 ? "Using a freebie leaves 1." : after === 0 ? "Using a freebie leaves none." : "";
  const ex = r.note.includes(EXEMPT_TAG) ? "<p class=\"sub\">NextGen marked it Exempt: payer exempt, or a freebie?</p>" : r.note.includes(UNCLASS_TAG) ? '<p class="sub">Not classified in NextGen. Clean it up there too.</p>' : "";
  showModal(`<h2 id="mtitle">${esc(kind)}: ${esc(p.tab)}, ${fmtMon(r.date)}</h2><p class="sub">${left} freebie${left === 1 ? "" : "s"} left.</p>${ex}${warn ? `<p class="warnline">${esc(warn)}</p>` : ""}
    <div class="mfoot"><button class="btn mute" data-mclose>Decide later</button><button class="btn ghost" data-mf="Waived" data-t="${esc(tab)}" data-row="${row}">Waived</button><button class="btn primary" data-mf="Used" data-t="${esc(tab)}" data-row="${row}"${left <= 0 ? " disabled" : ""}>Use freebie</button></div>`);
}
function mMove(v) {
  const fl = state.fl, list = fl.avail.filter(c => c.id !== v.chair).map(c => {
    const run = fl.visits.find(x => x.chair === c.id && ph(x) === "running"), wait = fl.visits.find(x => x !== v && x.chair === c.id && ph(x) === "waiting");
    const len = chairMin(v.type) || 30, clash = fl.visits.filter(x => x !== v && x.chairVisit && x.chair === c.id && ["booked", "waiting"].includes(ph(x)) && x.time != null && v.time != null && x.time < v.time + len && x.time + (chairMin(x.type) || 30) > v.time);
    const now = run ? `In use by ${run.tab} until ${clock(run.start + run.len)}` : wait ? `${wait.tab} checked in, waiting to start` : "Open now";
    return `<label class="opt"><input type="radio" name="chair" value="${esc(c.id)}"><span><b>${esc(c.label)}</b><small>${esc(now)}${clash.length ? `. <span class="warnline">Overlaps ${clash.map(x => `${esc(x.tab)} at ${clock(x.time)}`).join(", ")}</span>` : v.time != null ? `. Free at ${clock(v.time)}` : ""}</small></span></label>`;
  }).join("");
  state.mctx = { k: v.key };
  showModal(`<h2 id="mtitle">Move ${esc(v.tab)} to another chair</h2><p class="sub">${v.chair ? `Booked ${clock(v.time) || "today"} in ${esc(chairLbl(v.chair))}` : v.offChair ? `${esc(chairLbl(v.offChair))} is unavailable` : "No chair on the appointment list"}.${v.rc ? ` NextGen has ${esc(chairLbl(v.rc))}.` : ""}</p>
    <div class="opts">${list || '<p class="body">No other chair is available.</p>'}</div>${mfoot("Move", "doMove")}`);
}
async function doMove() {
  const v = findVisit(state.mctx.k), id = radio("chair"); if (!v) return closeModal(); if (!id) return merr("Choose a chair.");
  const rep0 = state.report, ng = TmsCore.ngForMove(v.s, v.rc, rep0 && rep0.asOf === todaySerial() ? rep0.asOfMin : null);
  closeModal();
  if (!await sessWrite(v, { chair: id, ng, tick: false, at: nowMin() })) return;
  state.sel = { k: v.key };
  const w = findVisit(v.key);
  showModal(`<h2 id="mtitle">${esc(v.tab)} moved to ${esc(chairLbl(id))}</h2>
    <p class="sub">${w && w.moved ? `NextGen still has ${esc(chairLbl(ng))} at ${clock(v.time) || "this visit"}. Move the appointment there too, so the next import keeps ${esc(chairLbl(id))}.` : "Back on the chair NextGen has. Nothing to change there."}</p>
    ${w && w.moved ? `<div class="notice"><label class="check"><input type="checkbox" data-ngk="${esc(v.key)}"> Moved in NextGen</label></div>` : ""}
    <div class="mfoot"><button class="btn primary" data-mclose>Done</button></div>`);
}
async function setNgTick(k, on) {
  const s = state.sess && state.sess.rows.get(k); if (!s) return;
  await sessWrite({ key: k, date: s.date, tab: s.tab, visit: s.visit }, { tick: on }, on ? `Noted: ${s.tab} moved in NextGen.` : "Unticked. Patients will remind you.");
}
function mNote(v) {
  const p = v.p, r = v.r, dr = draftOf(p, r), bl = baselineOf(p), sel = (dk, opts, val) => `<select data-dk="${dk}">${opts.map(o => `<option${o === val ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
  state.mctx = { k: v.key };
  const qs = NOTE_QS.map(q => `<div class="qrow"><span>${esc(q.q)}</span>${sel("a." + q.k, q.opts || ["No", "Yes"], dr.a[q.k])}${q.flag ? `<input class="sp" data-dk="d.${q.k}" value="${esc(dr.d[q.k] || "")}" placeholder="Specify (needed)"${dr.a[q.k] === "Yes" ? "" : " hidden"}>` : ""}</div>`).join("");
  showModal(`<div data-draft="${esc(dr.key)}"><h2 id="mtitle">Treatment note: ${esc(v.tab)}</h2><p class="sub">Answers become sentences; the questions stay out of the note.</p>
    ${fld("Mood on arrival", sel("mood", MOODS, dr.mood))}
    ${bl && (bl.sleep || bl.caffeine) ? `<p class="body">Baseline at MT: sleep ${esc(bl.sleep || "not noted")}, caffeine ${esc(bl.caffeine || "not noted")}. Ask how they slept and what they drank, and compare.</p>` : ""}${qs}
    ${fld("Technician comments", `<textarea data-dk="comments" rows="2">${esc(dr.comments)}</textarea>`)}
    ${fld("Initials", `<input data-dk="initials" value="${esc(dr.initials)}" maxlength="4" style="width:96px">`)}
    ${fld(`Note${dr.edited ? " (edited by hand)" : ""}`, `<textarea data-dk="text" rows="10" class="ntext">${esc(dr.edited ? dr.text : noteText(p, r, dr))}</textarea>`)}
    ${(reset => v.done ? mfoot("Copy note", "doCopyNote", reset) : mfoot("Save draft", "doSaveNote", reset + '<button class="btn ghost" data-act="doCopyNote">Copy now</button>'))(dr.edited ? '<button class="btn ghost" data-act="resetNote">Rebuild from answers</button>' : "")}</div>`, true);
}
/* the draft lives in memory as it's typed; Save draft just closes it until the session is complete */
function doSaveNote() { closeModal(); toast("Draft saved. Select Complete session when treatment ends."); }
async function doCopyNote() {
  const v = findVisit(state.mctx.k); if (!v) return closeModal();
  const dr = draftOf(v.p, v.r);
  if (!dr.edited) {
    const missing = NOTE_QS.filter(q => q.flag && dr.a[q.k] === "Yes" && !(dr.d[q.k] || "").trim()).map(q => q.flag);
    if (missing.length) return merr(`Specify details for: ${missing.join(", ")}.`);
    if (!dr.initials.trim()) return merr("Enter your initials.");
  }
  const ta = document.querySelector("#modalBody .ntext"), ok = await copyText(dr.edited ? dr.text : noteText(v.p, v.r, dr), ta);
  if (!document.contains(ta)) return;   // closed while the clipboard was busy
  if (!ok) { ta.focus(); ta.select(); return merr("Copy was blocked. The note is selected: press Ctrl+C."); }
  state.initials = dr.initials.trim(); state.copied.add(v.key); closeModal(); renderDash(); toast("Note copied. Paste it into NextGen.");
}
function mFuMeas(v) {
  state.mctx = { k: v.key };
  showModal(`<h2 id="mtitle">PHQ-9 and GAD-7: ${esc(v.tab)}</h2><p class="sub">Before the follow-up${v.p.provider ? ` with ${esc(v.p.provider)}` : ""} at ${clock(v.time) || "today"}.</p>
    <div class="g2">${fld("PHQ-9 (0 to 27)", '<input type="number" min="0" max="27" id="mPhq">')}${fld("GAD-7 (0 to 21)", '<input type="number" min="0" max="21" id="mGad">')}</div>${mfoot("Save scores", "doFuMeas")}`);
}
async function doFuMeas() {
  const v = findVisit(state.mctx.k), P = mval("mPhq"), Q = mval("mGad"); if (!v) return closeModal();
  if (!(isInt(P, 0, 27) && isInt(Q, 0, 21))) return merr("Enter PHQ-9 (0 to 27) and GAD-7 (0 to 21).");
  closeModal(); await saveMeasures(v.tab, v.row, { P: +P, Q: +Q });
}
/* Patients: the needs' own dialogs, for one tab (and a log row where the need is about one visit) */
function mPastPulses(tab, row) {
  const [p, r] = findPR(tab, row); if (!r) return; state.mctx = { t: tab, row };
  showModal(`<h2 id="mtitle">Pulses for ${fmtDay(r.date)}: ${esc(tab)}</h2><p class="sub">No pulses on the row and it isn't marked missed. Enter what was delivered, or record it as missed.</p>
    ${pulseFields(p)}${fld("Treating technician", techInput(r.tech || state.techName))}${mfoot("Save pulses", "doPastPulses", '<button class="btn ghost" data-act="pastMissed">It was missed</button>')}`);
}
function mPastTech(tab, row) {
  const [, r] = findPR(tab, row); if (!r) return; state.mctx = { t: tab, row };
  showModal(`<h2 id="mtitle">Technician for ${fmtDay(r.date)}: ${esc(tab)}</h2><p class="sub">Pulses are on the row; the treating technician (B) is blank.</p>${fld("Treating technician", techInput(state.techName))}${mfoot("Save technician", "doPastTech")}`);
}
function mPastMeas(tab, row) {
  const [p, r] = findPR(tab, row); if (!r) return; const dest = measDest(p, r); state.mctx = { t: tab, row, dest: dest && dest.row };
  showModal(`<h2 id="mtitle">PHQ-9 and GAD-7 for ${fmtDay(r.date)}: ${esc(tab)}</h2><p class="sub">Still marked x on the tab.</p>
    <div class="g2">${fld("PHQ-9 (0 to 27)", '<input type="number" min="0" max="27" id="mPhq">')}${fld("GAD-7 (0 to 21)", '<input type="number" min="0" max="21" id="mGad">')}</div>
    ${dest ? `<h3>Not collected?</h3>${fld("Move them to", measPick(p, r, dest))}` : ""}
    ${mfoot("Save scores", "doPastMeas", dest ? '<button class="btn ghost" data-act="doMoveMeas">Move the measures</button>' : "")}`);
}
function mBen(tab) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return; state.mctx = { t: tab };
  const d = p.d, dedOn = p.ben.dedRem > 0, mv = v => v == null ? "" : v;
  const dedTxt = p.ben.dedNA ? "N/A" : p.ben.dedRem != null ? p.ben.dedRem : "", oopTxt = p.ben.oopNA ? "N/A" : p.ben.oopRem != null ? p.ben.oopRem : "";
  const rows = BILL.map((b, i) => `<tr><td>${b.code}</td><td>${esc(b.label)}</td><td><input type="number" min="0" step="0.01" data-cost="X${COST_FIRST + i}" value="${mv(p.cost[i].after)}" aria-label="${esc(b.label)} patient responsibility"></td><td><input type="number" min="0" step="0.01" data-cost="Y${COST_FIRST + i}" value="${mv(p.cost[i].before)}" aria-label="${esc(b.label)} before deductible"${dedOn ? "" : ' placeholder="n/a"'}></td></tr>`).join("");
  showModal(`<h2 id="mtitle">Benefits and cost: ${esc(tab)}</h2><p class="sub">Saved to ${p.layout === 3 ? `the ${BEN_SHEET} tab (this MRN's row)` : "the tab"}. Cost so far ${money(d.accrued) || "$0.00"}, projected ${money(d.projected) || "$0.00"}.</p>
    ${d.bidf.length ? `<p class="warnline">Still needed: ${esc(d.bidf.join("; "))}</p>` : ""}
    <div class="g3">${fld("Deductible remaining", `<input data-benf="W4" value="${esc(dedTxt)}" placeholder="$ or N/A">`)}${fld("Out-of-pocket max remaining", `<input data-benf="W5" value="${esc(oopTxt)}" placeholder="$ or N/A">`)}${fld("BIDF estimated course cost", `<input type="number" min="0" step="0.01" data-benf="W1" value="${mv(p.ben.bidf)}">`)}</div>
    <table class="tbl cost"><tr><th>Code</th><th>Visit</th><th>PR</th><th>Before ded.</th></tr>${rows}</table>
    <p class="hint">PR is the patient's cost per visit, after the deductible. With a deductible remaining, enter the before-deductible cost too: column A charges that until the deductible is used up, then PR. Enter N/A when there is no deductible.</p>
    ${mfoot("Save benefits", "doBen")}`, true);
}
function mExt(tab) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return; state.mctx = { t: tab };
  showModal(`<h2 id="mtitle">Add extension: ${esc(tab)}</h2><p class="sub">${p.d.total} treatments authorized (D2)${p.authExp != null ? `, ${p.authExp < todaySerial() ? "expired" : "ends"} ${fmtMon(p.authExp)}` : ""}. Adds daily visits to D2 and Extensions added on the ${p.layout === 3 ? BEN_SHEET + " tab" : "tab (W7)"}.</p>
    <div class="g2">${fld("Daily visits to add", '<input type="number" min="1" id="mExtN" value="15">')}${fld("New auth end (optional)", '<input type="date" id="mExtD">')}</div>${mfoot("Add extension", "doExt")}`);
}
function mAvail(tab, keep) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return; state.mctx = { t: tab };
  const av = keep && state.avDraft && state.avDraft.tab === tab ? state.avDraft : (state.avDraft = { tab, week: JSON.parse(JSON.stringify(p.week)), ranges: p.avail.map(w => ({ ...w })) });
  showModal(`<h2 id="mtitle">Availability: ${esc(tab)}</h2><p class="sub">${esc(availText(p) || "Nothing entered yet")}. ${p.d.awayConflicts.length ? `Booked while away: ${p.d.awayConflicts.map(c => fmtMon(c.r.date)).join(", ")}. Cancel those in NextGen, or fix the dates here.` : ""}</p>
    <div class="avail">${availEditor(av)}</div>${mfoot("Save availability", "doAvail")}`, true);
}
function mStatus(tab) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return; state.mctx = { t: tab };
  const opts = ["Pending Start", "In Progress", "Tapering", "Paused", "Complete", "Discontinued"];
  const why = [p.d.stale ? `No visit in ${state.cfg.recentDays}+ days.` : "", p.d.cancelled || "", p.d.suggest ? `Suggested: ${p.d.suggest}.` : ""].filter(Boolean).join(" ");
  showModal(`<h2 id="mtitle">Status: ${esc(tab)}</h2><p class="sub">Now ${esc(p.status || "blank")} (D7).${why ? " " + esc(why) : ""}</p>
    <div class="opts">${opts.map(o => `<label class="opt"><input type="radio" name="st" value="${o}"${o === (p.d.suggest || (p.d.stale ? "Paused" : p.status)) ? " checked" : ""}><span>${o}</span></label>`).join("")}</div>${mfoot("Save status", "doStatus")}`);
}
function mCall(tab) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return; state.mctx = { t: tab };
  showModal(`<h2 id="mtitle">Call ${esc(tab)} to reschedule</h2><p class="sub">${TmsCore.missStreak(p.rows, todaySerial())} missed in a row. The call is saved to Notes on the last missed visit.</p>
    <div class="opts">${["Rescheduled", "Left voicemail", "No answer", "Wants to pause"].map((o, i) => `<label class="opt"><input type="radio" name="out" value="${o}"${i ? "" : " checked"}><span>${o}</span></label>`).join("")}</div>
    ${fld("Note", '<input id="mCallNote" placeholder="Optional">')}${mfoot("Log call", "doCall")}`);
}
function mChairOff(id) {
  const fl = state.fl, c = fl.chairs.find(x => x.id === id); if (!c) return; state.mctx = { chair: id };
  const run = fl.visits.find(v => v.chair === id && ph(v) === "running"), off = fl.visits.filter(v => v.chair === id && ["booked", "waiting"].includes(ph(v)));
  if (run) return showModal(`<h2 id="mtitle">${esc(c.label)} is in use</h2><p class="sub">${esc(run.tab)}'s session ends at ${clock(run.start + run.len)}. End it first, then mark the chair unavailable.</p><div class="mfoot"><button class="btn primary" data-mclose>Close</button></div>`);
  showModal(`<h2 id="mtitle">Mark ${esc(c.label)} unavailable?</h2><p class="sub">${off.length ? `${off.length} booking${off.length > 1 ? "s" : ""} today will need another chair: ${off.map(v => `${esc(v.tab)} ${clock(v.time)}`).join(", ")}. They show as "No chair yet" until moved.` : "Nothing else is booked on it today."}</p>
    ${mfoot(`Mark ${c.label} unavailable`, "doChairOff")}`);
}
function mChairEdit(id) {
  const c = chairList().find(x => x.id === id); if (!c) return; state.mctx = { chair: id };
  showModal(`<h2 id="mtitle">${esc(c.label)}</h2><p class="sub">Saved to the ${SESS_SHEET} sheet.</p>
    <div class="g2">${fld("Device", `<input id="mDev" value="${esc(c.device)}">`)}${fld("Coil", `<input id="mCoil" value="${esc(c.coil)}" placeholder="e.g. Cool-B65">`)}${fld("Last service", `<input type="date" id="mSvc" value="${c.service != null ? serialToInput(c.service) : ""}">`)}</div>${mfoot("Save", "doChairEdit")}`);
}

/* ---------- floor writes: the Sessions sheet ---------- */
/* one step: check-in and start are stamped together, then the treatment note (or, on an MT or MT redo, the motor threshold) opens */
async function startSession(v) {
  const fl = state.fl, busy = chairBusy(v, fl);
  if (!v.chair) return mMove(v);
  if (busy) return toast(`${chairLbl(v.chair)} is in use by ${busy.tab}. Move ${v.tab} or wait.`, true);
  state.sel = { k: v.key };
  if (!await sessWrite(v, { start: nowMin(), ...(v.checkin == null ? { checkin: nowMin() } : {}) }, `${v.tab}'s session started in ${chairLbl(v.chair)}.`)) return;
  const nv = findVisit(v.key) || v;
  if ((nv.type === "MT" || nv.type === "MTR") && !nv.mtDone) mMt(nv); else mNote(nv);
}
/* reopening an ended session clears the pulses, technician and copay line the add-in wrote, unless the chair has had a session since */
async function undoEnd(v) {
  const fl = state.fl, c = fl.chairs.find(x => x.id === v.chair);
  if (!c || c.off) return toast(`${chairLbl(v.chair)} is unavailable, so this session can't be reopened.`, true);
  if (fl.visits.some(x => x !== v && x.chair === v.chair && x.start != null && x.start >= (v.start ?? 0))) return toast(`${chairLbl(v.chair)} has had a session since. Correct this one on the tab.`, true);
  try {
    const note = withCopay(v.r.note, "");
    await writeCells(v.tab, [[`F${v.row}`, ""], [`H${v.row}`, ""], [`J${v.row}`, ""], [`B${v.row}`, ""], ...(note !== v.r.note ? [[`S${v.row}`, note]] : [])], at(v.r));
    patchRow(v.tab, v.row, { delivered: false, pulses: [null, null, null], tech: "", note }); state.copied.delete(v.key);
    await sessWrite(v, { end: null }, "Undone. The session is running again.");
  } catch (e) { fail(e); }
}
async function logCall(tab, outcome, note) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return;
  const r = p.rows.filter(x => x.miss && x.date != null).sort(cmp).pop(); if (!r) return toast("No missed visit on the tab to note the call on.", true);
  const line = `Called ${fmtDate(todaySerial())}: ${outcome}${note ? ` (${note})` : ""}`;
  try {
    await writeCells(tab, [[`S${r.row}`, (r.note ? r.note + "; " : "") + line]], at(r)); patchRow(tab, r.row, { note: (r.note ? r.note + "; " : "") + line });
    if (outcome === "Wants to pause") await setStatus(tab, "Paused"); else toast(`Call noted on ${tab}, ${fmtMon(r.date)}.`);
  } catch (e) { fail(e); }
}

/* ---------- actions ---------- */
async function setStatus(tab, val) {
  try { await writeCells(tab, [["D7", val]]); toast(`${tab}: ${val}`); await refresh(); } catch (e) { fail(e); }
}
/* ---------- Active Patients sheet: a view-only copy of the Patients roster for this week, rebuilt whenever the data changes ---------- */
const AP_SHEET = "Active Patients", AP_SHOW = [...ACTIVE, "Paused"], AP_HEAD = 4, AP_FIRST = 7;
const AP_FILL = { MT: "#dcffff", MTR: "#e6f3d6", "F/U": "#ffe6f1", Alt: "#ffeccc" };
function activeSheetModel() {
  const t = todaySerial(), ws = weekStart(t), days = [1, 2, 3, 4, 5].map(i => ws + i);
  const P = state.patients.filter(p => p.status === "" || AP_SHOW.includes(p.status)), blank = Array(7).fill("");
  const head = ["Patient", "MRN", "Status", ...days.map(d => `${DAYS[dow(d)]} ${fmtDate(d)}${d === t ? " (today)" : ""}`), "Next F/U", "Graduation", "Auth exp", "Tx left", "MD on-site", "Schedule given", "Copays unconfirmed"];
  const vs = days.map(variances);
  const treated = ["Treated", "", "", ...vs.map(v => v.expected ? `${v.actual} of ${v.expected}` : ""), ...blank];
  const notes = ["Variances", "", "", ...vs.map((v, i) => [state.closures.has(days[i]) ? `Closed: ${state.closures.get(days[i])}` : "", ...Object.entries(v.groups).filter(([, a]) => a.length).map(([k, a]) => `${a.length} ${k.toLowerCase()}`)].filter(Boolean).join(", ")), ...blank];
  const rows = P.map(p => [p.tab, p.mrnRaw, p.status || "(blank)", ...days.map(d => dayCode(p, d)), p.d.fuNext ? p.d.fuNext.date : "", p.d.grad ?? "", p.authExp ?? "",
    p.txRem ?? Math.max(0, p.d.total - p.d.txDone), p.d.mdOnSite ? "Yes" : "", p.ben.lastSched ?? "", p.d.toConfirm.length || ""]);
  return { t, days, P, head, sum: [treated, notes], rows, sig: JSON.stringify([t, head, treated, notes, rows]) };
}
async function syncActiveSheet(force) {
  if (state.apBusy) { state.apAgain = true; return; }
  const m = activeSheetModel(), settings = Office.context.document.settings;
  if (!force && state.hasAP && (state.apSig ?? settings.get("tms_ap_sig")) === m.sig) { state.apSig = m.sig; return; }
  state.apBusy = true;
  try {
    const n = m.rows.length, last = AP_FIRST + n - 1, end = Math.max(last, AP_FIRST), legend = end + 2, col = colLetter;
    await Excel.run(async ctx => {
      let ws = ctx.workbook.worksheets.getItemOrNullObject(AP_SHEET); await ctx.sync();
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add(AP_SHEET);
      ws.protection.load("protected"); await ctx.sync();
      if (ws.protection.protected) ws.protection.unprotect();
      try { ws.autoFilter.remove(); await ctx.sync(); } catch { /* no filter */ }
      const all = ws.getRange(); all.clear(Excel.ClearApplyTo.hyperlinks); all.clear(); all.unmerge(); all.conditionalFormats.clearAll(); all.dataValidation.clear(); ws.freezePanes.unfreeze();
      all.format.rowHeight = 15; all.format.columnWidth = 60;
      all.format.font.name = "Calibri"; all.format.font.size = 10; all.format.font.color = "#1c2a33"; ws.showGridlines = false;
      const stamp = `${fmtDateY(m.t)} ${fmtTime(new Date().getHours() * 60 + new Date().getMinutes())}`;
      ws.getRange("A1").values = [[AP_SHEET]]; ws.getRange("A1").format.font.size = 14; ws.getRange("A1").format.font.bold = true;
      ws.getRange("A2").values = [[`Week of ${fmtDateY(m.days[0])}. Rebuilt by the TMS Tracker add-in ${stamp}. View only: change things on the patient tabs or in the add-in.`]];
      ws.getRange("A2").format.font.italic = true; ws.getRange("A2").format.font.color = "#5b6b75";
      const hd = ws.getRange(`A${AP_HEAD}:O${AP_HEAD}`); hd.values = [m.head];
      hd.format.fill.color = "#1c2a33"; hd.format.font.color = "#ffffff"; hd.format.font.bold = true; hd.format.wrapText = true; hd.format.verticalAlignment = "Center"; hd.format.rowHeight = 30;
      const sm = ws.getRange(`A${AP_HEAD + 1}:O${AP_HEAD + 2}`); sm.values = m.sum; sm.format.fill.color = "#f3f6f8"; sm.format.verticalAlignment = "Top";
      ws.getRange(`A${AP_HEAD + 1}:A${AP_HEAD + 2}`).format.font.bold = true;
      ws.getRange(`D${AP_HEAD + 1}:H${AP_HEAD + 1}`).format.font.bold = true;
      const vr = ws.getRange(`D${AP_HEAD + 2}:H${AP_HEAD + 2}`); vr.format.wrapText = true; vr.format.font.size = 9; vr.format.font.color = "#5b6b75";
      ws.getRange(`D${AP_HEAD}:H${end}`).format.horizontalAlignment = "Center";
      if (n) {
        const body = ws.getRange(`A${AP_FIRST}:O${last}`); body.values = m.rows;
        ["I", "J", "K", "N"].forEach(c => ws.getRange(`${c}${AP_FIRST}:${c}${last}`).numberFormat = m.rows.map(() => ["m/d/yyyy"]));
        const b = body.format.borders.getItem("InsideHorizontal"); b.style = "Continuous"; b.color = "#dde3e8";
        const bb = body.format.borders.getItem("EdgeBottom"); bb.style = "Continuous"; bb.color = "#dde3e8";
        ws.getRange(`D${AP_FIRST}:H${last}`).format.font.bold = true;
        m.rows.forEach((r, i) => {
          const row = AP_FIRST + i, p = m.P[i];
          ws.getRange(`A${row}`).hyperlink = { documentReference: `'${p.tab.replace(/'/g, "''")}'!A1`, textToDisplay: p.tab, screenTip: `Open ${p.tab}` };
          r.slice(3, 8).forEach((c, j) => {
            const cell = ws.getRange(`${col(3 + j)}${row}`);
            if (AP_FILL[c]) cell.format.fill.color = AP_FILL[c]; else if (c === "X") { cell.format.fill.color = "#fbe9e7"; cell.format.font.color = "#b3261e"; }
          });
          if (p.status === "Paused") { const c = ws.getRange(`C${row}`); c.format.fill.color = "#fdf1d8"; c.format.font.color = "#8a5a00"; }
          if (p.status === "Pending Start" || !p.status) ws.getRange(`C${row}`).format.font.italic = true;
          if (p.authExp != null) {
            const c = ws.getRange(`K${row}`), left = p.authExp - m.t;
            if (left < 0) { c.format.fill.color = "#fbe9e7"; c.format.font.color = "#b3261e"; } else if (left <= state.cfg.authDays) { c.format.fill.color = "#fdf1d8"; c.format.font.color = "#8a5a00"; }
          }
        });
      } else ws.getRange(`A${AP_FIRST}`).values = [["No patients are Pending Start, In Progress, Tapering or Paused."]];
      const ti = m.days.indexOf(m.t);
      if (ti >= 0) ["EdgeLeft", "EdgeRight"].forEach(e => { const x = ws.getRange(`${col(3 + ti)}${AP_HEAD}:${col(3 + ti)}${end}`).format.borders.getItem(e); x.style = "Continuous"; x.weight = "Medium"; x.color = "#1c2a33"; });
      ws.getRange(`A${legend}:G${legend}`).values = [["Key", "MT", "MTR", "F/U", "Alt = other clinic", "X = missed", "\u00B7 = daily treatment. Other text is the action on that visit."]];
      ws.getRange(`A${legend}`).format.font.bold = true;
      [["B", AP_FILL.MT], ["C", AP_FILL.MTR], ["D", AP_FILL["F/U"]], ["E", AP_FILL.Alt], ["F", "#fbe9e7"]].forEach(([c, f]) => { const x = ws.getRange(`${c}${legend}`); x.format.fill.color = f; x.format.horizontalAlignment = "Center"; });
      ws.getRange(`F${legend}`).format.font.color = "#b3261e";
      [72, 62, 82, 64, 64, 64, 64, 64, 66, 66, 66, 48, 58, 70, 70].forEach((w, i) => ws.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w);
      ws.getRange(`A${AP_HEAD + 2}:O${AP_HEAD + 2}`).format.autofitRows();
      ws.freezePanes.freezeAt(ws.getRange(`A1:A${AP_HEAD + 2}`));
      ws.protection.protect({ allowFormatColumns: true, allowFormatRows: true });
      await ctx.sync();
    });
    state.apSig = m.sig; state.hasAP = true; settings.set("tms_ap_sig", m.sig); settings.saveAsync(); return true;
  } catch (e) { console.error(e); if (!state.apWarned) { state.apWarned = true; toast(`The ${AP_SHEET} sheet could not be rebuilt: ${e.message || e}`, true); } return false; }
  finally { state.apBusy = false; if (state.apAgain) { state.apAgain = false; syncActiveSheet(); } }
}
async function rebuildActiveClick() {
  if (POPOUT) return callPane("rebuildActive", []);
  state.apWarned = false; if (await syncActiveSheet(true)) toast(`${AP_SHEET} sheet rebuilt.`); }
/* ---------- Guest Patients sheet: patients with chair visits here whose home clinic (most chair visits, core.js clinicRole) carries
   the tab. One row per MRN, kept when they drop off the list (On latest list: No) so a link typed there isn't lost; delete rows by
   hand. The add-in writes A:G and I; H (the link to their tab on the home clinic's tracker) and J (notes) are the clinic's. ---------- */
const GUEST_SHEET = "Guest Patients", GUEST_HEAD = 4, GUEST_FIRST = 5, GUEST_LAST = 400;
const GUEST_HDR = ["Patient", "MRN", "Home clinic", "Chair visits here", "Chair visits there", "Next visit here", "Last visit here", "Link to their tab on the home clinic's tracker", "On latest list", "Notes"];
async function guestSheet(ctx) {
  let ws = ctx.workbook.worksheets.getItemOrNullObject(GUEST_SHEET); await ctx.sync();
  if (!ws.isNullObject) return ws;
  ws = ctx.workbook.worksheets.add(GUEST_SHEET);
  ws.getRange("A1").values = [[GUEST_SHEET]]; ws.getRange("A1").format.font.size = 14; ws.getRange("A1").format.font.bold = true;
  ws.getRange("A2").values = [["Patients seen here whose home clinic carries the tab (it has most of their chair visits). Pulses, technician and measures go on their tab there: paste its link in column H."]];
  ws.getRange("A2").format.font.italic = true; ws.getRange("A2").format.font.color = "#5b6b75";
  const hd = ws.getRange(`A${GUEST_HEAD}:J${GUEST_HEAD}`); hd.values = [GUEST_HDR];
  hd.format.fill.color = "#1c2a33"; hd.format.font.color = "#ffffff"; hd.format.font.bold = true; hd.format.wrapText = true; hd.format.verticalAlignment = "Center"; hd.format.rowHeight = 30;
  [["A", 70], ["B", 80], ["C", 90], ["D", 60], ["E", 60], ["F", 70], ["G", 70], ["H", 260], ["I", 60], ["J", 200]].forEach(([c, w]) => { ws.getRange(`${c}:${c}`).format.columnWidth = w; });
  ws.freezePanes.freezeRows(GUEST_HEAD);
  return ws;
}
/* reads the sheet: MRN -> { row, link }. The link is the cell's hyperlink, else a URL typed or pasted as text. */
async function readGuests(ctx, ws) {
  const r = ws.getRange(`A${GUEST_FIRST}:J${GUEST_LAST}`); r.load("values"); await ctx.sync();
  const n = r.values.reduce((k, v, i) => String(v[1]).trim() ? i + 1 : k, 0), links = [];
  for (let i = 0; i < n; i++) { const h = ws.getRange(`H${GUEST_FIRST + i}`); h.load("hyperlink"); links.push(h); }
  if (n) await ctx.sync();
  const by = new Map();
  r.values.slice(0, n).forEach((v, i) => { const mrn = normMrn(v[1]); if (!mrn) return; const h = links[i].hyperlink, txt = String(v[7]).trim(); by.set(mrn, { row: GUEST_FIRST + i, home: String(v[2]).trim(), link: (h && h.address) || (/^https?:\/\//i.test(txt) ? txt : "") }); });
  return { by, next: GUEST_FIRST + n };
}
async function syncGuests() {
  if (POPOUT || !state.guests) return;
  const sig = JSON.stringify([todaySerial(), state.guests.map(g => [g.mrn, g.code, g.home, g.here, g.homeCount, g.next, g.last]), state.patients.map(p => p.mrn)]);
  if (state.guestSig === sig) return;
  try {
    await Excel.run(async ctx => {
      const ex = ctx.workbook.worksheets.getItemOrNullObject(GUEST_SHEET); await ctx.sync();
      if (ex.isNullObject && !state.guests.length) return;
      const ws = await guestSheet(ctx), got = await readGuests(ctx, ws), by = got.by; let next = got.next;
      const on = new Set(state.guests.map(g => g.mrn)), dt = v => v == null ? "" : v;
      for (const g of state.guests) {
        const row = by.has(g.mrn) ? by.get(g.mrn).row : next++;
        // no home clinic on the report: keep one typed in C
        ws.getRange(`A${row}:G${row}`).values = [[g.code || "", g.raw, g.home || (by.has(g.mrn) ? by.get(g.mrn).home : ""), g.here, g.homeCount, dt(g.next), dt(g.last)]];
        ws.getRange(`F${row}:G${row}`).numberFormat = [["m/d/yyyy", "m/d/yyyy"]]; ws.getRange(`B${row}`).numberFormat = [["@"]];
        ws.getRange(`I${row}`).values = [["Yes"]];
        g.link = by.has(g.mrn) ? by.get(g.mrn).link : "";
      }
      state.guestLinks = Object.fromEntries([...by].filter(([, x]) => x.link).map(([mrn, x]) => [mrn, x.link]));
      /* off the list: a tab here now, or no longer booked */
      for (const [mrn, x] of by) if (!on.has(mrn)) ws.getRange(`I${x.row}`).values = [[state.patients.some(p => p.mrn === mrn) ? "Tab here now" : "No"]];
      await ctx.sync();
    });
    state.guestSig = sig; renderBanners();
  } catch (e) { console.error(e); toast(`The ${GUEST_SHEET} sheet could not be updated: ${e.message || e}`, true); }
}
/* the link to a guest's tab on their home clinic's tracker: written to H as a hyperlink */
async function saveGuestLink(mrn, url) {
  if (POPOUT) return callPane("saveGuestLink", [mrn, url]);
  url = String(url || "").trim();
  if (!/^https:\/\/\S+$/i.test(url)) return toast("Paste the full link to their tab (it starts with https://).", true);
  const g = (state.guests || []).find(x => x.mrn === mrn); if (!g) return;
  try {
    await Excel.run(async ctx => {
      const ws = await guestSheet(ctx), { by, next } = await readGuests(ctx, ws), row = by.has(mrn) ? by.get(mrn).row : next;
      if (!by.has(mrn)) { ws.getRange(`A${row}:G${row}`).values = [[g.code || "", g.raw, g.home, g.here, g.homeCount, g.next ?? "", g.last ?? ""]]; ws.getRange(`I${row}`).values = [["Yes"]]; }
      ws.getRange(`H${row}`).hyperlink = { address: url, textToDisplay: `${g.code || "Their"} tab at ${g.home || "their home clinic"}` };
      await ctx.sync();
    });
    g.link = url; state.guestLinks = { ...(state.guestLinks || {}), [mrn]: url }; toast(`Link saved for ${g.code || "the guest"}${g.home ? ` (${g.home})` : ""}.`); renderBanners(); if (state.view === "today") renderDash();
  } catch (e) { fail(e); }
}

/* ---------- Settings > Use another tab as New Template: the chosen tab is renamed New Template and the current one deleted.
   Patient tabs and the add-in's own sheets aren't offered. Logged without undo; the workbook's version history has the old tab. ---------- */
const OWN_SHEETS = () => ["New Template", AP_SHEET, GUEST_SHEET, BEN_SHEET, SESS_SHEET, LIST_SHEET, ELOG_SHEET, "Clinic Schedule", "Chair Schedule", "Protocols"].map(sheetKey);
const tplChoices = () => (state.sheetNames || []).filter(n => !OWN_SHEETS().includes(sheetKey(n)) && !state.tabs.some(t => t.name === n) && !state.stale.includes(n) && !/ \(old\)$/.test(n));
function fillTplPick() {
  const s = $("tplPick"); if (!s || document.activeElement === s) return;
  const L = tplChoices(), cur = s.value;
  s.innerHTML = L.length ? L.map(n => `<option${n === cur ? " selected" : ""}>${esc(n)}</option>`).join("") : "<option value=\"\">No other tabs to use</option>";
  s.disabled = !L.length; $("btnTplUse").disabled = !L.length;
}
function mTplUse() {
  const name = $("tplPick").value; if (!name) return;
  const had = (state.sheetNames || []).includes("New Template"); state.mctx = { t: name };
  showModal(`<h2 id="mtitle">Make ${esc(name)} New Template?</h2><p class="body">${esc(name)} is renamed New Template${had ? " and the current New Template is <b>deleted</b>" : ""}. New tabs are copied from it, and Match New Template brings active tabs in line with it.</p>
    <p class="body">If it isn't laid out like the tracker template (PATIENT: in A1, MRN: in A3, the visit log headings in row 8), Settings says what's off and tabs can't be added until it's fixed.${had ? " The workbook's version history can bring the old New Template back." : ""}</p>
    ${mfoot(had ? "Delete the old one and use this tab" : "Use this tab", "doTplUse")}`);
}
async function useTemplateTab(name) {
  if (POPOUT) return callPane("useTemplateTab", [name]);
  await Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(name), old = ctx.workbook.worksheets.getItemOrNullObject("New Template"); await ctx.sync();
    if (ws.isNullObject) throw new Error(`There's no tab named ${name} any more. Refresh and choose again.`);
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    ws.name = "New Template"; await ctx.sync();
    await addEdits(ctx, [{ tab: "New Template", kind: "tab", what: `New Template is now the tab that was ${name}${old.isNullObject ? "" : "; the earlier New Template was deleted"}`, undo: null }]);
    await ctx.sync();
  });
}
async function doTplUse() {
  const name = state.mctx.t; closeModal();
  try { await useTemplateTab(name); toast(`${name} is now New Template.`); await refresh(); if (state.templateWhy) toast(state.templateWhy, true, true); } catch (e) { fail(e); }
}

/* ---------- first launch: a workbook's own New Template is used as it is; one without it gets the blank one shipped with the add-in
   (assets/templates). Active Patients and Guest Patients are made. Inserting sheets from a file needs ExcelApi 1.13 (Excel on the web has it). ---------- */
const TEMPLATE_FILE = () => `assets/templates/${device() === "Brainsway" ? "BrainsWay" : "MagV"} Template.xlsx`;
const TEMPLATE_SHEETS = ["New Template", "Protocols"];   // inserted when the workbook doesn't have them
const sheetKey = n => String(n).trim().replace(/\s+/g, " ").toLowerCase();
async function firstLaunch() {
  if (POPOUT || state.setupRan) return; state.setupRan = true;
  const made = [], renamed = [];
  try {
    let have = [];
    await Excel.run(async ctx => {
      ctx.workbook.load("name"); const s = ctx.workbook.worksheets; s.load("items/name"); await ctx.sync(); state.wbName = ctx.workbook.name || "";
      /* the workbook's own template under a slightly different name ("New template", "New Template ") is named New Template, so the add-in uses it */
      for (const n of TEMPLATE_SHEETS) { const near = s.items.find(x => x.name !== n && sheetKey(x.name) === sheetKey(n)); if (near && !s.items.some(x => x.name === n)) { renamed.push(`"${near.name}" to ${n}`); near.name = n; } }
      await ctx.sync(); have = s.items.map(x => sheetKey(x.name));
    });
    const want = TEMPLATE_SHEETS.filter(n => !have.includes(sheetKey(n)));
    if (want.includes("New Template")) {
      if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) throw new Error("This version of Excel can't add sheets from a file. Copy New Template in from the clinic's template workbook.");
      const res = await fetch(TEMPLATE_FILE(), { cache: "no-store" });
      if (!res.ok) throw new Error(`The blank template (${TEMPLATE_FILE()}) isn't on the add-in's host yet. Copy New Template in from the clinic's template workbook.`);
      const buf = await res.arrayBuffer(), inFile = XLSX.read(buf, { type: "array", bookSheets: true }).SheetNames, names = want.filter(n => inFile.includes(n));
      if (!names.includes("New Template")) throw new Error(`${TEMPLATE_FILE()} has no sheet named New Template.`);
      let bin = ""; const u8 = new Uint8Array(buf); for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
      await Excel.run(async ctx => { ctx.workbook.insertWorksheetsFromBase64(btoa(bin), { sheetNamesToInsert: names, positionType: Excel.WorksheetPositionType.beginning }); await ctx.sync(); });
      made.push(...names);
    }
  } catch (e) { console.error(e); toast(e.message || String(e), true, true); }
  try {
    await Excel.run(async ctx => {
      for (const n of [AP_SHEET, GUEST_SHEET]) { const ws = ctx.workbook.worksheets.getItemOrNullObject(n); await ctx.sync(); if (ws.isNullObject) { if (n === GUEST_SHEET) await guestSheet(ctx); else ctx.workbook.worksheets.add(n); made.push(n); } }
      await ctx.sync();
    });
  } catch (e) { console.error(e); }
  if (made.length || renamed.length) toast(`Set up this workbook: ${[made.length && `added ${made.join(", ")}`, renamed.length && `renamed ${renamed.join(", ")}`].filter(Boolean).join("; ")}.`);
}
/* ---------- Benefits tab: one row per MRN (BEN_COLS). Rows are found by MRN; a new patient gets the next empty row. ---------- */
function readBen(vals) {
  const rows = new Map(); let next = BEN_FIRST;
  (vals || []).slice(BEN_FIRST - 1).forEach((v, i) => {
    const raw = v[0]; if (raw === "" || raw == null) return;
    const k = normMrn(raw), row = BEN_FIRST + i; next = row + 1;
    if (k && !rows.has(k)) rows.set(k, { row, v: v.slice() });
  });
  state.ben = { rows, next, has: !!vals };
}
async function benSheet(ctx) {
  let ws = ctx.workbook.worksheets.getItemOrNullObject(BEN_SHEET); await ctx.sync();
  if (!ws.isNullObject) return ws;
  ws = ctx.workbook.worksheets.add(BEN_SHEET);
  const last = colLetter(BEN_HDR.length - 1), hd = ws.getRange(`A1:${last}1`);
  hd.values = [BEN_HDR]; hd.format.font.bold = true; hd.format.wrapText = true; hd.format.verticalAlignment = "Center";
  hd.format.fill.color = "#1c2a33"; hd.format.font.color = "#ffffff"; hd.format.rowHeight = 32;
  const money = ["C", "D", "E", ..."IJKLMNOPQR"], dates = ["H", "S"];
  money.forEach(c => ws.getRange(`${c}${BEN_FIRST}:${c}${BEN_LAST}`).numberFormat = Array(BEN_LAST - BEN_FIRST + 1).fill(["$#,##0.00"]));
  dates.forEach(c => ws.getRange(`${c}${BEN_FIRST}:${c}${BEN_LAST}`).numberFormat = Array(BEN_LAST - BEN_FIRST + 1).fill(["m/d/yyyy"]));
  ws.getRange(`A:${last}`).format.columnWidth = 78; ws.getRange("A:B").format.columnWidth = 70;
  ws.freezePanes.freezeRows(1);
  if (!state.ben) readBen(null);
  return ws;
}
/* entries: { mrn (normalized), mrnVal (as typed in A4), tab, key (W4, X12 ... or BEN_STAMP), v, quiet }. Every row written to gets
   Last updated = today, unless every entry for it is quiet (the add-in's own bookkeeping, such as MTR approved following D4). */
async function writeBen(ctx, entries) {
  if (!entries.length) return;
  const ws = await benSheet(ctx), B = state.ben, touched = new Map();
  for (const e of entries) {
    if (!e.mrn) continue;
    let r = B.rows.get(e.mrn);
    if (!r) { r = { row: B.next++, v: Array(BEN_HDR.length).fill("") }; B.rows.set(e.mrn, r); r.v[0] = e.mrnVal ?? e.mrn; ws.getRange(`A${r.row}`).values = [[r.v[0]]]; }
    /* Patient names the newest course for this MRN */
    const newest = state.patients.filter(p => p.mrn === e.mrn).sort((a, b) => b.course - a.course)[0], m = String(e.tab).match(TAB_RE);
    const tab = newest && newest.course > +((m && m[2]) || 1) ? newest.tab : e.tab;
    if (r.v[1] !== tab) { r.v[1] = tab; ws.getRange(`B${r.row}`).values = [[tab]]; }
    if (e.key !== BEN_STAMP) { const c = BEN_COLS[e.key]; r.v[colIndex(c)] = e.v; ws.getRange(`${c}${r.row}`).values = [[e.v]]; }
    touched.set(e.mrn, (touched.get(e.mrn) || false) || !e.quiet);
  }
  const t = todaySerial();
  for (const [k, loud] of touched) if (loud) { const r = B.rows.get(k); r.v[colIndex(BEN_COLS.Updated)] = t; ws.getRange(`${BEN_COLS.Updated}${r.row}`).values = [[t]]; }
}
/* a benefits key on a current-layout tab is a Benefits row write; on a previous-layout tab it is still that tab's own cell */
const benTarget = tab => state.patients.find(p => p.tab === tab && p.layout === 3);
/* every simple write goes through applyOps: { tab, a, v (2-D values) } or { tab, a, clear: true }. Benefits keys (W4, X12 ...)
   on a current-layout tab go to that patient's Benefits row; { tab, a: BEN_STAMP } only stamps Last updated. */
/* log: { what, kind } names the change in the Edit log (what defaults to the cells' meaning, see describeCells); false leaves it
   out, for writes that are part of a larger logged change (an import). An op with f: true writes a formula. */
async function applyOps(ops, guard, log) {
  if (POPOUT) return callPane("applyOps", [ops, guard, log]);
  const ben = [], cells = []; let unlogged = false;
  for (const o of ops) {
    const p = isBenKey(o.a) && benTarget(o.tab);
    if (p) ben.push({ mrn: p.mrn, mrnVal: p.mrnVal, tab: p.tab, key: o.a, v: o.clear ? "" : o.v && o.v[0][0], quiet: !!o.quiet });
    else if (o.a !== BEN_STAMP) cells.push(o);
  }
  await Excel.run(async ctx => {
    const g = (guard || []).filter(x => x.date != null).map(x => ({ ...x, r: ctx.workbook.worksheets.getItem(x.tab).getRange(`E${x.row}:L${x.row}`) }));   // E date, L type
    /* what each cell held before, so the Edit log can put it back */
    const rs = log === false ? [] : cells.map(o => { const r = ctx.workbook.worksheets.getItem(o.tab).getRange(o.a); r.load("formulas"); return { o, r }; });
    /* and which visit each visit-log row held (E date, L type), so an undo can tell the row still holds that visit */
    const ats = log === false ? [] : [...new Set(cells.map(o => logRowOf(o.a) && `${o.tab}|${logRowOf(o.a)}`).filter(Boolean))]
      .map(k => { const [tab, row] = k.split("|"), r = ctx.workbook.worksheets.getItem(tab).getRange(`E${row}:L${row}`); r.load("values"); return { tab, row: +row, r }; });
    g.forEach(x => x.r.load("values"));
    if (g.length || rs.length) await ctx.sync();
    const was = rs.map(({ o, r }) => ({ o, b: r.formulas.map(x => x.slice()) }));   // copied now, before anything is written
    const atWas = ats.map(x => {
      const a = { tab: x.tab, row: x.row, date: dayOf(x.r.values[0][0]), type: String(x.r.values[0][7] ?? "").trim() };
      cells.forEach(o => { if (o.tab !== a.tab || o.clear || !o.v) return; if (o.a === `E${a.row}`) a.date = dayOf(o.v[0][0]); if (o.a === `L${a.row}`) a.type = String(o.v[0][0] ?? "").trim(); });   // what the edit leaves there
      return a;
    });
    const moved = g.find(x => Math.floor(Number(x.r.values[0][0]) || 0) !== x.date || (x.type != null && String(x.r.values[0][7] ?? "").trim() !== String(x.type).trim()));   // both sides trimmed: "Daily " on the tab is still Daily
    if (moved) throw new Error(`${moved.tab}: the visit on row ${moved.row} has moved on the tab (an import may have re-sorted it). Nothing was saved. Select Refresh and try again.`);
    const benWas = log === false ? [] : ben.filter(e => e.key !== BEN_STAMP).map(e => ({ e, b: ((state.ben.rows.get(e.mrn) || {}).v || [])[colIndex(BEN_COLS[e.key])] ?? "" }));   // Last updated stamps aren't undone
    for (const o of cells) { const r = ctx.workbook.worksheets.getItem(o.tab).getRange(o.a); if (o.clear) r.clear("Contents"); else if (o.f) r.formulas = o.v; else r.values = o.v; }
    await writeBen(ctx, ben);
    /* the save stands even if the Edit log can't record it (for example a protected workbook structure); it says so instead */
    if (log !== false) try { await addEdits(ctx, editsFromOps(was, benWas, atWas, log || {})); } catch (e) { console.error(e); unlogged = true; }
    await ctx.sync();
  });
  if (unlogged) toast("Saved, but the Edit log couldn't record it, so it can't be undone from there.", true);
}
const dayOf = v => Math.floor(Number(v)) || null;
/* a single visit-log cell (A:T), not availability in V:Y on the same rows */
const logRowOf = a => { const m = String(a).match(/^[A-T](\d+)$/), r = m && +m[1]; return r >= LOG_FIRST && r < LOG_FIRST + LOG_READ ? r : null; };
/* ---------- Edit log: every change the add-in makes to a patient tab or the Benefits tab, one row each on the Edit Log sheet
   (hidden; it travels with the workbook, so every computer sees the same list). A Id, B when, C tab, D what, E kind, F by,
   G state ("" can be undone, "Undone ...", "No undo"), H on: what undo needs, as JSON in 30,000-character parts. An undo only
   runs when the cells still hold what the edit left; a later edit on the same cells has to be undone first.
   Not logged: the Sessions sheet (check-in, chair moves), formatting, and the Appt List tab, which each import replaces. ---------- */
const ELOG_SHEET = "Edit Log", ELOG_HDR = ["Id", "When", "Tab", "What", "Kind", "By", "State", "Undo data"], ELOG_PART = 30000, ELOG_PARTS = 20, ELOG_SHOW = 300, ELOG_KEEP = 2000;
/* what a cell means, for the log's wording: visit-log columns, then header cells */
const LOG_WORDS = { A: "PR formula", B: "technician", C: "Tx #", D: "Tx left formula", E: "date", F: "pulses", G: "pulse total", H: "pulses", I: "pulse total", J: "pulses", K: "pulse total",
  L: "visit type", M: "missed mark", N: "taper", O: "action", P: "PHQ-9", Q: "GAD-7", R: "other clinic", S: "notes", T: "time" };
const HDR_WORDS = { A4: "MRN", D2: "# of TXs", D4: "# of MT redos", D6: "MD on site", D7: "status", E2: "provider", E4: "TAS", E6: "insurance", F4: "auth start", G4: "auth expiry",
  G2: "update protocol by", G6: "progress note sent", N6: "freebies left", N7: "auth assessment", S9: "home clinic" };
const sentence = s => s ? s[0].toUpperCase() + s.slice(1) : s;
const andList = a => a.length < 2 ? a.join("") : `${a.slice(0, -1).join(", ")} and ${a[a.length - 1]}`;
const plural = (n, w, ws) => `${n} ${n === 1 ? w : ws || w + "s"}`;
/* Excel hands formulas back in its own spacing and numbers back as numbers, so cells are compared on a normal form */
const normCell = v => { const s = String(v ?? "").trim(); if (s.startsWith("=")) return s.replace(/\s+/g, "").toUpperCase(); return s !== "" && isFinite(+s) ? String(Math.round(+s * 1e6) / 1e6) : s; };
const sameGrid = (a, b) => JSON.stringify((a || []).map(r => r.map(normCell))) === JSON.stringify((b || []).map(r => r.map(normCell)));
/* "Pulses, technician and notes on Oct 27 (row 29); status" */
function describeCells(tab, addrs) {
  const p = state.patients.find(x => x.tab === tab), end = (p && p.end) || LOG_LAST, byRow = new Map(), other = [];
  for (const a of addrs) {
    if (isBenKey(a)) { other.push("benefits"); continue; }
    const m = String(a).match(/^([A-Z]+)(\d+)/); if (!m) continue;
    const col = m[1], row = +m[2];
    if (row >= LOG_FIRST && row <= end && LOG_WORDS[col]) { if (!byRow.has(row)) byRow.set(row, []); byRow.get(row).push(LOG_WORDS[col]); continue; }
    other.push(HDR_WORDS[col + row] || (row >= 2 && row <= 4 && "JKLMO".includes(col) ? "protocol" : row >= WK_FIRST - 1 && "VWXY".includes(col) ? "availability" : `${col}${row}`));
  }
  const rows = [...byRow].sort((a, b) => a[0] - b[0]).map(([row, w]) => { const r = p && p.rows.find(x => x.row === row); return `${andList([...new Set(w)])} on ${r && r.date != null ? `${fmtMon(r.date)} (row ${row})` : `row ${row}`}`; });
  return sentence([...rows, ...new Set(other)].join("; ")) || "Cells changed";
}
/* applyOps' writes as Edit log entries, one per tab */
function editsFromOps(was, benWas, atWas, log) {
  const by = new Map(), get = tab => by.get(tab) || by.set(tab, { tab, cells: [], ben: [], at: [] }).get(tab);
  atWas.forEach(a => get(a.tab).at.push({ row: a.row, date: a.date, type: a.type }));
  was.forEach(({ o, b }) => get(o.tab).cells.push({ a: o.a, b, f: o.clear ? b.map(x => x.map(() => "")) : o.v }));
  benWas.forEach(({ e, b }) => get(e.tab).ben.push({ a: e.key, b, f: e.v }));
  return [...by.values()].filter(x => x.cells.length || x.ben.length).map(x => ({ tab: x.tab, kind: log.kind || "edit", what: log.what || describeCells(x.tab, [...x.cells, ...x.ben].map(c => c.a)), undo: { cells: x.cells, ben: x.ben, at: x.at } }));
}
async function elogSheet(ctx) {
  let ws = ctx.workbook.worksheets.getItemOrNullObject(ELOG_SHEET); await ctx.sync();
  if (!ws.isNullObject) return ws;
  ws = ctx.workbook.worksheets.add(ELOG_SHEET);
  const hd = ws.getRange("A1:H1"); hd.values = [ELOG_HDR]; hd.format.font.bold = true;
  ws.visibility = Excel.SheetVisibility.hidden;
  return ws;
}
let elogSeq = 0;
const nowSerial = () => { const d = new Date(); return todaySerial() + (d.getHours() * 60 + d.getMinutes()) / 1440; };
const whenText = s => { const d = Math.floor(s), t = clock(Math.round((s - d) * 1440)); return d === todaySerial() ? t : `${fmtMon(d)}, ${t}`; };
/* entries: { tab, kind, what, undo (null: nothing to put back) }, written after the sheet's last row (read now, since another
   computer may have written since). The oldest 200 go once the log passes ELOG_KEEP. */
async function addEdits(ctx, entries) {
  if (!entries.length) return;
  const ws = await elogSheet(ctx), u = ws.getUsedRange(true); u.load("rowCount"); await ctx.sync();
  let row = Math.max(2, u.rowCount + 1); const when = nowSerial(), by = state.techName || "";
  for (const e of entries) {
    const id = `${Date.now().toString(36)}${(elogSeq++).toString(36)}`, data = e.undo ? JSON.stringify(e.undo) : "";
    const parts = []; for (let i = 0; i < data.length; i += ELOG_PART) parts.push(data.slice(i, i + ELOG_PART));
    const st = data && parts.length <= ELOG_PARTS ? "" : "No undo";
    const line = ws.getRange(`A${row}:G${row}`); line.numberFormat = [["@", "m/d/yyyy h:mm AM/PM", "@", "@", "@", "@", "@"]]; line.values = [[id, when, e.tab, e.what, e.kind, by, st]];
    if (!st) { const c = ws.getRange(`H${row}:${colLetter(7 + parts.length - 1)}${row}`); c.numberFormat = [parts.map(() => "@")]; c.values = [parts]; }
    state.edits.unshift({ id, row, when, tab: e.tab, what: e.what, kind: e.kind, by, st });
    row++;
  }
  state.edits.length = Math.min(state.edits.length, ELOG_SHOW);
  if (row - 2 > ELOG_KEEP + 200) { ws.getRange("2:201").delete(Excel.DeleteShiftDirection.up); state.edits.forEach(x => { x.row -= 200; }); }
}
async function readElog(ctx) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(ELOG_SHEET); await ctx.sync();
  if (ws.isNullObject) { state.edits = []; return; }
  const u = ws.getUsedRange(true); u.load("rowCount"); await ctx.sync();
  const last = u.rowCount, first = Math.max(2, last - ELOG_SHOW + 1);
  if (last < 2) { state.edits = []; return; }
  const r = ws.getRange(`A${first}:G${last}`); r.load("values"); await ctx.sync();
  state.edits = r.values.map((v, i) => ({ id: String(v[0]), row: first + i, when: Number(v[1]) || 0, tab: String(v[2]), what: String(v[3]), kind: String(v[4]), by: String(v[5]), st: String(v[6]) }))
    .filter(e => e.id).reverse();
}
/* each tab's log, read in the current Excel.run: where it ends, B:T formulas (F) and A:T values (V) */
async function readLogs(ctx, tabs) {
  /* six tabs per read, like refresh's batches: Excel on the web caps what one sync can return */
  if (tabs.length > 6) { const out = new Map(); for (let i = 0; i < tabs.length; i += 6) for (const [k, v] of await readLogs(ctx, tabs.slice(i, i + 6))) out.set(k, v); return out; }
  const G = tabs.map(tab => { const ws = ctx.workbook.worksheets.getItem(tab), a10 = ws.getRange(`A${LOG_FIRST}`), b = ws.getRange(`B${LOG_FIRST}:B${LOG_FIRST + LOG_READ - 1}`); a10.load("formulas"); b.load("values"); return { tab, ws, a10, b }; });
  await ctx.sync();
  G.forEach(g => { g.end = endOf(g.b.values.map(r => r[0]), g.a10.formulas[0][0]); g.r = g.ws.getRange(`B${LOG_FIRST}:T${g.end}`); g.r.load("formulas"); g.v = g.ws.getRange(`A${LOG_FIRST}:T${g.end}`); g.v.load("values"); });
  await ctx.sync();
  return new Map(G.map(g => [g.tab, { end: g.end, F: g.r.formulas, V: g.v.values }]));
}
/* a log's typed cells: B:T without the formula columns D, G, I, K */
const logHash = F => TmsCore.hashRows(F.map(r => r.map((v, j) => LOG_FORMULA_COLS.has(j + 1) ? "" : normCell(v))));
/* the log as a plan saw it (parseLog's reading), so applyPlan can tell whether a tab changed after the preview was made */
const logSig = rows => TmsCore.hashRows(rows.map(r => [r.row, r.date, r.time, r.type, r.miss, r.delivered, String(r.tx ?? "").trim(), r.note, r.alt, r.phq, r.gad, r.tech]));
async function readUndo(ctx, e) {
  const r = ctx.workbook.worksheets.getItem(ELOG_SHEET).getRange(`A${e.row}:${colLetter(7 + ELOG_PARTS - 1)}${e.row}`); r.load("values"); await ctx.sync();
  const v = r.values[0]; if (String(v[0]) !== e.id) throw new Error("The Edit log moved (another computer added to it). Select Refresh and try again.");
  return JSON.parse(v.slice(7).join("") || "null");
}
/* puts a log back: rows the edit added are deleted (ins), rows it removed come back (del), then B:T's typed cells are written as
   they were. The formula columns (D, G, I, K, and A) stay on their rows. A row comes back inside the log (never on its first row or
   the TX SUMMARY row, where the PR and header ranges wouldn't take it in), copying the row below; B:T is rewritten anyway.
   One batch: the deletes and inserts and the B:T write go to Excel together, so a failed sync leaves nothing half done. */
async function restoreLog(ctx, tab, L) {
  const ws = ctx.workbook.worksheets.getItem(tab);
  let n = L.nAfter;
  (L.ins || []).slice().sort((a, b) => b - a).forEach(row => { ws.getRange(`A${row}:T${row}`).delete(Excel.DeleteShiftDirection.up); n--; });
  for (const row of (L.del || []).slice().sort((a, b) => a - b)) {
    const at = Math.min(Math.max(row, L.first + 1), L.first + n - 1);
    ws.getRange(`A${at}:T${at}`).insert(Excel.InsertShiftDirection.down);
    ws.getRange(`A${at}:T${at}`).copyFrom(ws.getRange(`A${at + 1}:T${at + 1}`), Excel.RangeCopyType.all);
    n++;
  }
  const last = L.first + L.before.length - 1;
  [["B", "C", 0], ["E", "F", 3], ["H", "H", 6], ["J", "J", 8], ["L", "T", 10]].forEach(([c1, c2, j]) => {
    const w = colIndex(c2) - colIndex(c1) + 1;
    ws.getRange(`${c1}${L.first}:${c2}${last}`).formulas = L.before.map(r => r.slice(j, j + w));
  });
}
async function undoEdit(id) {
  if (POPOUT) return callPane("undoEdit", [id]);
  const e = state.edits.find(x => x.id === id); if (!e) return toast("That edit is no longer in the log. Select Refresh.", true);
  if (e.st) return toast(e.st === "No undo" ? "This edit can't be undone from here." : "Already undone.", true);
  let relog = false;
  try {
    failStep = "Undoing the edit";
    await Excel.run(async ctx => {
      const u = await readUndo(ctx, e); if (!u) throw new Error("This edit has nothing saved to undo.");
      const later = state.edits.slice(0, state.edits.indexOf(e)).filter(x => x.tab === e.tab && !x.st && x.kind !== "undo").length;
      const changed = `${e.tab} changed after this edit${later ? ` (${plural(later, "later edit")} on it in the log)` : ""}. Undo the later edits first, or fix the tab by hand. Nothing was changed.`;
      const ws = ctx.workbook.worksheets.getItem(e.tab), cells = (u.cells || []).map(c => { const r = ws.getRange(c.a); r.load("formulas"); return { c, r }; });
      const ats = (u.at || []).map(a => { const r = ws.getRange(`E${a.row}:L${a.row}`); r.load("values"); return { a, r }; });
      const logs = u.log ? await readLogs(ctx, [e.tab]) : (await ctx.sync(), null);
      /* the row must still hold the same visit: pulses and a technician look alike from day to day, so matching cells alone isn't enough */
      const gone = ats.find(({ a, r }) => dayOf(r.values[0][0]) !== a.date || String(r.values[0][7] ?? "").trim() !== a.type);
      if (gone) throw new Error(`${e.tab} changed after this edit: row ${gone.a.row} no longer holds the ${gone.a.date != null ? fmtMon(gone.a.date) + " " : ""}visit (a sort, import or removed row moved it). Nothing was changed; fix the cells by hand.`);
      if (cells.some(({ c, r }) => !sameGrid(r.formulas, c.f))) throw new Error(changed);
      if (u.log) { const now = logs.get(e.tab); if (now.F.length !== u.log.nAfter || logHash(now.F) !== u.log.hash) throw new Error(changed); }
      const bp = benTarget(e.tab), benNow = (u.ben || []).map(c => ({ c, v: bp ? ((state.ben.rows.get(bp.mrn) || {}).v || [])[colIndex(BEN_COLS[c.a])] ?? "" : "" }));
      if (benNow.some(x => normCell(x.v) !== normCell(x.c.f))) throw new Error(changed);
      if (u.log) { await restoreLog(ctx, e.tab, u.log); relog = true; }
      cells.forEach(({ c, r }) => { r.formulas = c.b; });
      if (benNow.length && bp) await writeBen(ctx, benNow.map(({ c }) => ({ mrn: bp.mrn, mrnVal: bp.mrnVal, tab: bp.tab, key: c.a, v: c.b, quiet: true })));
      const st = `Undone ${fmtDate(todaySerial())} ${clock(Math.round((nowSerial() % 1) * 1440))}`;
      ctx.workbook.worksheets.getItem(ELOG_SHEET).getRange(`G${e.row}`).values = [[st]]; e.st = st;
      await addEdits(ctx, [{ tab: e.tab, kind: "undo", what: `Undid: ${e.what}`, undo: null }]);
      await ctx.sync();
    });
    failStep = "";
    await refresh();
    if (relog) { const p = state.patients.find(x => x.tab === e.tab); if (p) await formatLogs([{ tab: p.tab, touched: p.rows.map(r => r.row) }]); }
    toast(`Undone: ${e.what}`);
  } catch (err) { if (/changed after this edit|no longer|moved|nothing saved/i.test(err.message)) { failStep = ""; toast(err.message, true, true); } else fail(err); }
}

/* ---------- import bookkeeping: applyPlan checks each tab against its preview first, and logs one entry per tab when done ---------- */
const imp = { busy: 0, snaps: new Map() };
/* list: { tab, sig, cells } per changed tab. Returns { id }, or { stale: [tabs] } when a tab no longer matches the preview (nothing
   is held then). One import at a time through this pane, whichever window asked. */
async function importStart(list) {
  if (POPOUT) return callPane("importStart", [list]);
  if (imp.busy && Date.now() - imp.busy < 10 * 60e3) throw new Error("An import is already being applied. Wait for it to finish, then check the preview again.");
  imp.busy = Date.now();
  try {
    const stale = [], snap = new Map();
    await Excel.run(async ctx => {
      const logs = await readLogs(ctx, list.map(x => x.tab));
      const cells = list.map(x => x.cells.map(a => { const r = ctx.workbook.worksheets.getItem(x.tab).getRange(a); r.load("formulas"); return { a, r }; }));
      await ctx.sync();
      list.forEach((x, i) => {
        const L = logs.get(x.tab);
        if (logSig(parseLog(L.V.slice(0, LOG_READ))) !== x.sig) stale.push(x.tab);   // refresh reads LOG_READ rows at most
        snap.set(x.tab, { before: L.F, cells: cells[i].map(({ a, r }) => ({ a, b: r.formulas })) });
      });
    });
    if (stale.length) { imp.busy = 0; return { stale }; }
    const id = String(Date.now()); imp.snaps.set(id, snap); return { id, stale: [] };
  } catch (e) { imp.busy = 0; throw e; }
}
/* done: { tab, what, ins } per tab; stopped names the step that failed, if one did. Each tab's entry can put its log back, unless
   the row count says rows were added or lost some other way (then it's logged without an undo). */
async function importEnd(id, done, stopped) {
  if (POPOUT) return callPane("importEnd", [id, done, stopped]);
  const snap = imp.snaps.get(id); imp.snaps.delete(id); imp.busy = 0; if (!snap || !done.length) return;
  await Excel.run(async ctx => {
    const logs = await readLogs(ctx, done.map(x => x.tab));
    const cells = done.map(x => snap.get(x.tab).cells.map(c => { const r = ctx.workbook.worksheets.getItem(x.tab).getRange(c.a); r.load("formulas"); return { c, r }; }));
    await ctx.sync();
    await addEdits(ctx, done.map((x, i) => {
      const s = snap.get(x.tab), L = logs.get(x.tab), what = `${stopped ? `Import stopped at "${stopped}"` : "Import"}: ${x.what}`;
      if (L.F.length - s.before.length !== x.ins.length) return { tab: x.tab, kind: "import", what, undo: null };
      return { tab: x.tab, kind: "import", what, undo: { log: { first: LOG_FIRST, before: s.before, ins: x.ins, nAfter: L.F.length, hash: logHash(L.F) }, cells: cells[i].map(({ c, r }) => ({ a: c.a, b: c.b, f: r.formulas })) } };
    }));
    await ctx.sync();
  });
}
/* "Created 2 new rows, set 5 dates, marked 1 missed, re-sorted by date" */
function importWhat(p) {
  const w = [p.added && `created ${plural(p.added, "new row")}`, p.set && `set ${plural(p.set, "date")}`, p.types && `changed ${plural(p.types, "visit type")}`,
    p.cleared && `cleared ${plural(p.cleared, "date")}`, p.miss && `marked ${p.miss} missed`, p.unmissed && `undid ${plural(p.unmissed, "miss", "misses")} NextGen corrected`,
    p.resorted && "re-sorted by date", p.renum && !p.resorted && "renumbered Tx #", p.e2 && `provider set to ${p.e2}`].filter(Boolean);
  return sentence(w.join(", ")) || (p.ops.length ? `Updated ${plural(p.ops.length, "cell")} (times, notes or marks)` : "Checked against the list");
}

/* ---------- the Edit log's checks: what's off on each active tab, from its own log and from the saved appointment list ---------- */
function checksOf(p) {
  const c = TmsCore.logChecks(p.rows, { pr: p.canGrow, approved: p.txApproved }), appts = state.apptBy.get(p.tab);
  c.gaps = appts ? TmsCore.apptGaps(p.rows, appts, { from: todaySerial(), to: state.apptTo ?? todaySerial() }) : { notOnTab: [], notOnList: [] };
  c.gaps.noAppt = noApptRows(p).map(r => r.row);
  c.n += c.gaps.notOnTab.length + c.gaps.notOnList.length + c.gaps.noAppt.length;
  return c;
}
/* the log in date order, as an import leaves it (undated treated or missed rows stay after the visit above them) */
function sortOrder(rows) {
  let lk = -Infinity, lt = 0;
  return rows.map((r, i) => { if (r.date != null) { lk = r.date; lt = r.time ?? 0; } return r.date != null || r.delivered || r.miss ? { i, k: lk, t: lt } : { i, k: Infinity, t: 0 }; })
    .sort((a, b) => a.k - b.k || a.t - b.t || a.i - b.i).map(x => x.i);
}
/* a logged change to the whole log (sort, remove a row): fn(ctx, ws) makes it; the entry can put the log back */
async function logChange(tab, what, kind, fn, del) {
  await Excel.run(async ctx => {
    const before = (await readLogs(ctx, [tab])).get(tab), ws = ctx.workbook.worksheets.getItem(tab);
    await fn(ctx, ws); await ctx.sync();
    const after = (await readLogs(ctx, [tab])).get(tab);
    try { await addEdits(ctx, [{ tab, kind, what, undo: { log: { first: LOG_FIRST, before: before.F, ins: [], del: del || [], nAfter: after.F.length, hash: logHash(after.F) } } }]); await ctx.sync(); }
    catch (e) { console.error(e); toast("Done, but the Edit log couldn't record it, so it can't be undone from there.", true); }
  });
}
/* the checks' Fix buttons. kind: tx (renumber C), formulas (A and D back), kaiser (drop the reminder), sort, remove (a row) */
async function fixCheck(tab, kind, row) {
  if (POPOUT) return callPane("fixCheck", [tab, kind, row]);
  const p = state.patients.find(x => x.tab === tab); if (!p) return;
  const c = checksOf(p);
  try {
    /* the rows as this pane read them: a guard on each dated row stops the write if the tab was re-sorted since */
    const g = rows => rows.map(n => p.rows.find(r => r.row === n)).filter(r => r && r.date != null).map(r => ({ tab, ...at(r) }));
    if (kind === "tx") await applyOps(c.tx.map(x => ({ tab, a: `C${x.row}`, v: [[x.want]] })), g(c.tx.map(x => x.row)), { kind: "txnum", what: `Renumbered Tx # on ${plural(c.tx.length, "row")}: the MT is 1, each treatment after it one more` });
    else if (kind === "formulas") await applyOps(c.formulas.flatMap(x => [x.a && { tab, a: `A${x.row}`, v: [[prF(x.row, p.end)]], f: true }, x.d && { tab, a: `D${x.row}`, v: [[ROW_F.D(x.row)]], f: true }].filter(Boolean)), g(c.formulas.map(x => x.row)),
      { kind: "formulas", what: `Put back the formulas on ${plural(c.formulas.length, "row")} (${andList([c.formulas.some(x => x.a) && "PR in A", c.formulas.some(x => x.d) && "Tx left in D"].filter(Boolean))})` });
    else if (kind === "kaiser") await applyOps(c.kaiser.map(r => ({ tab, a: `S${r}`, v: [[TmsCore.dropKaiser(p.rows.find(x => x.row === r).note)]] })), g(c.kaiser), { kind: "notes", what: `Removed the template's Kaiser reminder from ${plural(c.kaiser.length, "row")}` });
    else if (kind === "sort") {
      const order = sortOrder(p.rows), tx = TmsCore.txNumbers(order.map(i => p.rows[i]));
      await logChange(tab, "Sorted the visit log by date and renumbered Tx #", "sort", () => arrangeLogs([{ tab, order, tx, dates: p.rows.map(r => r.date) }]));
    } else if (kind === "remove") {
      const r = p.rows.find(x => x.row === row); if (!r || r.delivered) return toast("Only a row without pulses can be removed.", true);
      await logChange(tab, `Removed row ${row} (${r.date != null ? fmtMon(r.date) + " " : ""}${r.type || "no type"}${r.miss ? ", missed" : ""}), a second row for that day`, "rows", async (ctx, ws) => {
        const e = ws.getRange(`E${row}:L${row}`); e.load("values"); await ctx.sync();
        if ((Math.floor(Number(e.values[0][0])) || null) !== r.date || String(e.values[0][7] ?? "").trim() !== r.type) throw new Error(`${tab}: row ${row} has changed. Nothing was removed. Select Refresh and try again.`);
        ws.getRange(`A${row}:T${row}`).delete(Excel.DeleteShiftDirection.up);
      }, [row]);
    }
    toast("Fixed. The Edit log at the bottom of Patients can undo it.");
    await refresh();
  } catch (e) { fail(e); }
}
/* new visit-log rows, in A:T only so the benefits block beside the log stays put. Each goes inside the log (planTab never puts one
   past the last row), so Excel stretches the PR formulas, header counts and rules over it. It takes everything from a row of the
   same visit type (fill, borders, validation, conditional formats, height) and keeps that row's formulas, with the typed cells blank. */
async function growLogs(list) {
  if (!list.length) return;
  if (POPOUT) return callPane("growLogs", [list]);
  await Excel.run(async ctx => {
    for (const { tab, ins } of list) {
      const ws = ctx.workbook.worksheets.getItem(tab);
      ins.forEach(({ row }) => ws.getRange(`A${row}:T${row}`).insert(Excel.InsertShiftDirection.down));   // ascending, so each lands on its final row
      const src = ins.map(({ src }) => { const r = ws.getRange(`A${src}:T${src}`); r.load("formulasR1C1"); r.format.load("rowHeight"); return r; });
      await ctx.sync();
      ins.forEach(({ row }, i) => {
        const r = ws.getRange(`A${row}:T${row}`);
        r.copyFrom(src[i], Excel.RangeCopyType.all);
        /* copyFrom brought the formulas; blank the typed cells (writing the formulas back in R1C1 fails on the PR formula) */
        src[i].formulasR1C1[0].forEach((f, c) => { if (!(typeof f === "string" && f.startsWith("="))) r.getCell(0, c).clear("Contents"); });
        r.format.rowHeight = src[i].format.rowHeight;
      });
      await ctx.sync();
      /* the colors by visit type and the dropdowns are laid down again over the grown log, so the new rows get them too */
      const end = await tabEnd(ctx, ws); applyRules(ws, end, end + 1); await ctx.sync();
    }
  });
}
/* after the import's writes: puts each log row's typed cells (B:T) on the row order planTab chose, so E runs in date order,
   and writes the Tx # in C. The template's formulas (D, G, I, K) stay on their own rows, so D recalculates from the new C. */
const LOG_FORMULA_COLS = new Set([3, 6, 8, 10]);   // D, G, I, K, 0-based in A:T
async function arrangeLogs(list) {
  if (!list.length) return;
  if (POPOUT) return callPane("arrangeLogs", [list]);
  await Excel.run(async ctx => {
    const got = list.map(({ tab, order }) => { const r = ctx.workbook.worksheets.getItem(tab).getRange(`B${LOG_FIRST}:T${LOG_FIRST + order.length - 1}`); r.load("formulas,values"); return r; });
    await ctx.sync();
    /* the order is a move of rows, so it is only right on the rows it was planned from: every E must hold the date the plan wrote
       there. Anything else (the log changed while the import ran) and nothing moves, on any tab. */
    list.forEach(({ tab, dates }, n) => {
      if (!dates) return;
      const i = dates.findIndex((d, k) => dayOf(got[n].values[k][3]) !== (d ?? null));   // values: an E that is a formula (=E12+1) still reads as its date
      if (i >= 0) throw new Error(`${tab}: row ${LOG_FIRST + i} changed while the import ran, so the visit log was not re-sorted. Undo this import from the Edit log at the bottom of Patients, then import again.`);
    });
    list.forEach(({ tab, order, tx }, n) => {
      const F = got[n].formulas, ws = ctx.workbook.worksheets.getItem(tab);
      order.forEach((s, i) => {
        const out = F[i].map((v, j) => LOG_FORMULA_COLS.has(j + 1) ? v : j + 1 === C.tx ? tx[i] : F[s][j]);
        if (out.some((v, j) => v !== F[i][j])) ws.getRange(`B${LOG_FIRST + i}:T${LOG_FIRST + i}`).formulas = [out];
      });
    });
    await ctx.sync();
  });
}
/* Visit-log formatting follows the dates. A row with a date in E takes the first log row's formatting (A10:T10: no gray, bordered,
   date and time formats); a row without one takes New Template's blank row for that position (gray; the last two rows map to its
   last two). Rows an import changed always follow; any other row only when it has no highlight of its own (every cell's fill is
   row 10's or the blank row's), so a hand-colored row is left alone. Cell properties and number formats only: the template's
   conditional formats (the MT, MTR, F/U and missed highlights) and validation stay as they are. Returns the rows reformatted. */
const FMT_Q = { format: { fill: { color: true }, font: { bold: true, color: true, italic: true, name: true, size: true }, borders: { color: true, style: true, weight: true }, horizontalAlignment: true } };
/* the log grid, header row 9 through the log's last row in A:R, is drawn in regular (thin) lines: no bold outline, whatever the template has */
const GRID_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"];
const thinGrid = (ws, end) => GRID_EDGES.forEach(e => { const b = ws.getRange(`A9:R${end}`).format.borders.getItem(e); b.style = "Continuous"; b.weight = "Thin"; });
/* Excel reports a setting it can't give one value for (mixed fonts in a cell, a pattern fill) as null, and rejects null when it's written back */
const dropBlank = o => {
  if (o === null || o === undefined || o === "") return undefined;
  if (typeof o !== "object") return o;
  const r = {}; for (const [k, v] of Object.entries(o)) { const x = dropBlank(v); if (x !== undefined) r[k] = x; }
  return Object.keys(r).length ? r : undefined;
};
const fmtCell = c => {
  const f = c.format || {}, b = f.borders || {};
  return dropBlank({ format: { fill: { color: (f.fill || {}).color }, font: f.font, horizontalAlignment: f.horizontalAlignment,
    borders: Object.fromEntries(["top", "bottom", "left", "right"].filter(k => b[k]).map(k => [k, b[k].style === "None" ? { style: "None" } : b[k]])) } }) || {};
};
const fillOf = c => String((((c || {}).format || {}).fill || {}).color || "").toUpperCase();
const ampm = f => /AM\/PM/i.test(f) ? f : "h:mm AM/PM";
async function formatLogs(list) {
  if (!list.length) return 0;
  if (POPOUT) return callPane("formatLogs", [list]);
  let n = 0; const skipped = [];
  await Excel.run(async ctx => {
    const tpl = ctx.workbook.worksheets.getItemOrNullObject("New Template"); await ctx.sync();
    const bl = tpl.isNullObject ? null : tpl.getRange(`A${LOG_FIRST + 1}:T${LOG_LAST}`), bp = bl && bl.getCellProperties(FMT_Q);
    if (bl) bl.load("numberFormat");
    const G = list.map(({ tab, touched }) => {
      const ws = ctx.workbook.worksheets.getItem(tab), a10 = ws.getRange(`A${LOG_FIRST}`), b = ws.getRange(`B${LOG_FIRST}:B${LOG_FIRST + LOG_READ - 1}`);
      a10.load("formulas"); b.load("values"); return { tab, ws, a10, b, touched: new Set(touched || []) };
    });
    await ctx.sync();
    /* three tabs per read: cell properties are bulky, and Excel on the web caps how much one sync can return */
    for (let k = 0; k < G.length; k += 3) {
      const part = G.slice(k, k + 3);
      part.forEach(g => { g.end = endOf(g.b.values.map(r => r[0]), g.a10.formulas[0][0]); g.rg = g.ws.getRange(`A${LOG_FIRST}:T${g.end}`); g.rg.load("values,numberFormat"); g.cp = g.rg.getCellProperties(FMT_Q); });
      await ctx.sync();
      const jobs = [];
      for (const g of part) {
        const P = g.cp.value, V = g.rg.values, NF = g.rg.numberFormat, first = P[0].map(fmtCell);
        /* the visit time in T reads 2:30 PM, whatever the template's row 10 says, on hand-colored rows too (only T changes there) */
        const nf0 = NF[0].map((f, j) => j === C.time ? ampm(f) : f), tOnly = i => { if (NF[i][C.time] !== ampm(NF[i][C.time])) jobs.push({ g, r: LOG_FIRST + i, t: true }); };
        tOnly(0);
        for (let i = 1; i < V.length; i++) {
          const r = LOG_FIRST + i, dated = isNum(V[i][C.date]); if (!dated && !bp) continue;
          const ti = (r >= g.end - 1 ? LOG_LAST - (g.end - r) : Math.min(r, LOG_LAST - 2)) - LOG_FIRST - 1;   // New Template row for this position, 0 = row 11
          const want = dated ? first : bp.value[ti].map(fmtCell), wantNf = dated ? nf0 : bl.numberFormat[ti];
          const plain = P[i].every((c, j) => fillOf(c) === fillOf(P[0][j]) || (bp && fillOf(c) === fillOf(bp.value[ti][j])));
          if (!g.touched.has(r) && !plain) { if (dated) tOnly(i); continue; }
          if (P[i].every((c, j) => JSON.stringify(fmtCell(c)) === JSON.stringify(want[j])) && NF[i].every((f, j) => f === wantNf[j])) continue;
          jobs.push({ g, r, want, wantNf });
        }
      }
      const put = j => { if (j.t) { j.g.ws.getRange(`T${j.r}`).numberFormat = [["h:mm AM/PM"]]; return; } const row = j.g.ws.getRange(`A${j.r}:T${j.r}`); row.setCellProperties([j.want]); row.numberFormat = [j.wantNf]; };
      /* one bad row shouldn't stop the rest: if the batch fails, rows go one at a time and the ones Excel refuses are skipped */
      jobs.forEach(put);
      try { await ctx.sync(); n += jobs.length; }
      catch (e) { for (const j of jobs) { put(j); try { await ctx.sync(); n++; } catch (e2) { console.error(e2); skipped.push(`${j.g.tab} row ${j.r}`); } } }
      part.forEach(g => thinGrid(g.ws, g.end));
      try { await ctx.sync(); } catch (e) { console.error(e); }
    }
  });
  if (skipped.length) toast(`Couldn't reformat ${skipped.length} visit-log row${skipped.length === 1 ? "" : "s"} (${skipped.slice(0, 5).join(", ")}${skipped.length > 5 ? ", ..." : ""}). Everything else was applied.`, true, true);
  return n;
}
async function formatLogsClick() {
  try {
    const n = await formatLogs(state.patients.filter(p => !state.stale.includes(p.tab)).map(p => ({ tab: p.tab, touched: [] })));
    toast(n ? `Reformatted ${n} visit-log row${n === 1 ? "" : "s"}. Rows with their own highlight were left as they are.` : "Visit-log formatting already follows the dates.");
  } catch (e) { fail(e); }
}
/* guard: the log row and the date it should still hold, checked before anything is written */
const writeCells = (tab, pairs, guard) => applyOps(pairs.map(([a, v]) => ({ tab, a, v: [[v]] })), guard ? [].concat(guard).map(x => ({ tab, ...x })) : null);
const at = r => ({ row: r.row, date: r.date, type: r.type });
function patchRow(tab, row, f) { const p = state.patients.find(x => x.tab === tab), r = p && p.rows.find(x => x.row === row); if (!r) return; Object.assign(r, f); p.d = derive(p); renderDash(); }
const findPR = (tab, row) => { const p = state.patients.find(x => x.tab === tab); return [p, p && p.rows.find(x => x.row === row)]; };

/* missed visits are recorded from the Record missed dialog (recordMiss); a freebie left undecided is classified later (freebieModal) */
/* the freebies left (N6; N7 on a previous-layout tab): using one takes one off, undoing it gives it back */
async function setFreebie(tab, row, v) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const note = withFreebie(r.note, v), n7 = Math.max(0, p.d.left + (r.freebie === "Used" ? 1 : 0) - (v === "Used" ? 1 : 0));
  try {
    await writeCells(tab, [[`S${row}`, note], [p.fbCell, n7]], at(r)); closeModal(); p.freebiesLeft = n7; patchRow(tab, row, { freebie: v, note });
    const left = p.d.left;
    toast(v === "Used" && left <= 1 ? (left === 1 ? `Freebie used. ${tab} has 1 left.` : `Freebie used. ${tab} has none left.`) : v === "Used" ? "Freebie used." : "Marked waived.", v === "Used" && left === 0);
  } catch (e) { fail(e); }
}
async function undoMiss(tab, row) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const note = withFreebie(r.note.split("; ").filter(x => !/\(marked in add-in\)$/.test(x)).join("; "), "");
  const pairs = [[`M${row}`, ""], [`S${row}`, note]]; if (r.freebie === "Used") pairs.push([p.fbCell, p.d.left + 1]);
  try { await writeCells(tab, pairs, at(r)); if (r.freebie === "Used") p.freebiesLeft = p.d.left + 1; patchRow(tab, row, { miss: false, freebie: "", note }); toast("Missed visit undone."); } catch (e) { fail(e); }
}
/* pulses delivered = treatment complete (F, H, J). Treating technician goes in B; a new Notes (S), e.g. with the copay, when given.
   Returns whether it saved. */
async function markPulses(tab, row, { vals, tech, note }) {
  const [, r] = findPR(tab, row); if (!r) return false;
  if (!vals.some(v => v !== "" && v > 0)) { toast("Enter the pulses delivered.", true); return false; }
  if (!tech) { toast("Enter the treating technician.", true); return false; }
  const pairs = [["F", vals[0]], ["H", vals[1]], ["J", vals[2]]].filter(([, v]) => v !== "" && v > 0).map(([c, v]) => [`${c}${row}`, v]);
  pairs.push([`B${row}`, tech]); state.techName = tech;
  if (note != null) pairs.push([`S${row}`, note]);
  try { await writeCells(tab, pairs, at(r)); patchRow(tab, row, { delivered: true, tech, pulses: [vals[0] || null, vals[1] || null, vals[2] || null], ...(note != null ? { note } : {}) }); return true; } catch (e) { fail(e); return false; }
}
/* PHQ-9 in P, GAD-7 in Q */
async function saveMeasures(tab, row, vals) {
  const [, r] = findPR(tab, row); if (!r) return false;
  if (vals.P != null && (vals.P < 0 || vals.P > 27)) { toast("PHQ-9 must be 0 to 27.", true); return false; }
  if (vals.Q != null && (vals.Q < 0 || vals.Q > 21)) { toast("GAD-7 must be 0 to 21.", true); return false; }
  const pairs = Object.entries(vals).filter(([, n]) => n != null).map(([c, n]) => [`${c}${row}`, n]); if (!pairs.length) { toast("Enter a score first.", true); return false; }
  try {
    await writeCells(tab, pairs, at(r));
    const phq = vals.P != null ? String(vals.P) : r.phq, gad = vals.Q != null ? String(vals.Q) : r.gad;
    patchRow(tab, row, { phq, gad, measX: phq.toLowerCase() === "x" || gad.toLowerCase() === "x", measVal: true }); toast("Measures saved."); return true;
  } catch (e) { fail(e); return false; }
}
/* not collected: the x's move to another visit (destRow) and the original row's Notes say so */
async function moveMeasures(tab, row, destRow) {
  const [p, r] = findPR(tab, row); if (!r) return false;
  const dest = p.rows.find(x => x.row === destRow); if (!dest) return false;
  const note = (r.note ? r.note + "; " : "") + `Measures not collected ${fmtDate(r.date)}; moved to ${fmtDate(dest.date)}`;
  const pairs = [[`S${row}`, note]];
  if (r.phq.toLowerCase() === "x") pairs.push([`P${row}`, ""]); if (r.gad.toLowerCase() === "x") pairs.push([`Q${row}`, ""]);
  if (!dest.phq) pairs.push([`P${dest.row}`, "x"]); if (!dest.gad) pairs.push([`Q${dest.row}`, "x"]);
  try { await writeCells(tab, pairs, [at(r), at(dest)]); toast(`Measures moved to ${fmtMon(dest.date)}.`); await refresh(); return true; } catch (e) { fail(e); return false; }
}
/* MT details ("MT %: 55", "Hotspot: ...") replace the earlier ones in the MT row's Notes */
async function saveMT(tab, row, parts) {
  const [, r] = findPR(tab, row); if (!r) return false;
  if (!parts.length) { toast("Enter at least one MT detail.", true); return false; }
  const keep = r.note.split("; ").filter(x => !/^(hotspot|mt %|amps|baseline sleep|baseline caffeine):/i.test(x));
  const note = [...keep.filter(Boolean), ...parts].join("; ");
  try { await writeCells(tab, [[`S${row}`, note]], at(r)); patchRow(tab, row, { note }); toast("MT details saved to Notes."); return true; } catch (e) { fail(e); return false; }
}
/* protocol rows 2-4: name from the Protocols sheet fills J:M (N calculates), site label in O. picks: [{ name, site }] for rows 1-3 */
async function saveProtocol(tab, picks) {
  const pairs = [];
  picks.forEach(({ name, site }, i) => {
    const row = i + 2, q = state.protocols.find(x => x.name === name);
    if (q) pairs.push([`J${row}`, q.rep], [`K${row}`, q.pulses], [`L${row}`, q.trains], [`M${row}`, q.iti]);
    else if (!name) pairs.push([`J${row}`, ""], [`K${row}`, ""], [`L${row}`, ""], [`M${row}`, ""]);
    pairs.push([`O${row}`, q || site ? site : ""]);
  });
  try { await writeCells(tab, pairs); toast("Protocol saved."); await refresh(); return true; } catch (e) { fail(e); return false; }
}
/* benefits: W1 BIDF estimate, W4 deductible and W5 OOP max (an amount or N/A), cost by code in X (PR) and Y (before deductible) */
const benVal = v => { const s = String(v ?? "").trim().replace(/[$,]/g, ""); return s === "" ? "" : isNA(s) ? "N/A" : isNaN(Number(s)) ? null : Number(s); };
async function saveBenefits(tab, data) {
  const pairs = [];
  for (const [a, raw] of Object.entries(data)) {
    const v = benVal(raw); if (v === null) return toast(`${a}: enter an amount${/^W[45]$/.test(a) ? " or N/A" : ""}.`, true);
    if (v === "N/A" && !/^W[45]$/.test(a)) return toast(`${a}: enter an amount.`, true);
    pairs.push([a, v]);
  }
  try { await writeCells(tab, pairs); toast("Benefits saved."); await refresh(); } catch (e) { fail(e); }
}
/* an extension adds daily visits: # of TXs (D2) and Extensions added (W7) go up; the auth expiry moves if given */
async function addExtension(tab, { n, exp }) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return;
  if (!(n > 0)) return toast("Enter the number of visits.", true);
  const pairs = [["D2", (p.txApproved || 36) + n], ["W7", (p.ben.ext || 0) + n]];
  if (exp) pairs.push(["G4", exp]);
  try { await writeCells(tab, pairs); toast(`Added ${n} daily visits. # of TXs is now ${(p.txApproved || 36) + n}${exp ? `, auth expires ${fmtDate(exp)}` : ""}.`); await refresh(); } catch (e) { fail(e); }
}
function readAvailEditor(root, av) {
  root.querySelectorAll("[data-avw]").forEach(i => { const d = +i.dataset.avw, k = i.dataset.k; av.week[d] = av.week[d] || {}; av.week[d][k] = i.value ? (([h, m]) => +h * 60 + +m)(i.value.split(":")) : null; });
  root.querySelectorAll("[data-avr]").forEach(i => { const w = av.ranges[+i.dataset.avr]; if (!w) return; const k = i.dataset.k; w[k] = k === "from" || k === "to" ? inputToSerial(i.value) : i.value; });
}
/* weekly start window in W20:X24 (Monday to Friday), dates in V29:Y48 */
async function saveAvailability(tab, av) {
  const wk = [1, 2, 3, 4, 5].map(d => { const w = av.week[d] || {}; return [w.from != null ? w.from / 1440 : "", w.to != null ? w.to / 1440 : ""]; });
  const rg = av.ranges.filter(w => w.from).map(w => [w.type || "Unavailable", w.from, w.to || w.from, w.note || ""]);
  if (rg.length > AV_LAST - AV_FIRST + 1) return toast("Too many date ranges for the tab.", true);
  const ops = [{ tab, a: `W${WK_FIRST}:X${WK_FIRST + 4}`, v: wk }, { tab, a: `V${AV_FIRST}:Y${AV_LAST}`, clear: true }];
  if (rg.length) ops.push({ tab, a: `V${AV_FIRST}:Y${AV_FIRST + rg.length - 1}`, v: rg });
  try { await applyOps(ops); state.avDraft = null; toast("Availability saved."); await refresh(); } catch (e) { fail(e); }
}

/* ---------- treatment note drafter (Daily treatments only) ---------- */
const NOTE_QS = [
  { k: "mus", q: "Were there any involuntary muscle movements?", opts: ["Yes (normal)", "No"], def: "Yes (normal)" },
  { k: "med", q: "Are there any changes in medication?", flag: "change in medication" },
  { k: "pain", q: "Was there any pain or discomfort during treatment?", flag: "pain or discomfort during treatment" },
  { k: "meal", q: "Any change in meals?", flag: "change in meals" },
  { k: "sleep", q: "Any change in sleep patterns?", flag: "change in sleep patterns" },
  { k: "caf", q: "Any changes in caffeine consumption?", flag: "change in caffeine consumption" },
  { k: "alc", q: "Any alcohol consumption since last treatment?", flag: "alcohol consumption since last treatment" },
  { k: "drug", q: "Any recreational drug use since last treatment?", flag: "recreational drug use since last treatment" }
];
const MOODS = ["good", "neutral", "reserved", "stressed"];
const initialsOf = name => { const s = String(name || "").trim(); if (/^[A-Za-z]{2,3}$/.test(s)) return s.toUpperCase(); return s.split(/[\s,.]+/).filter(Boolean).map(w => w[0].toUpperCase()).join("").slice(0, 3); };
function draftOf(p, r) {
  const key = `${p.tab}|${r.row}`;
  return state.drafts[key] ||= { key, open: false, mood: "good", a: Object.fromEntries(NOTE_QS.map(q => [q.k, q.def || "No"])), d: {}, comments: "", initials: state.initials || initialsOf(r.tech), edited: false, text: "" };
}
function noteText(p, r, dr) {
  const yes = q => q.flag && dr.a[q.k] === "Yes", det = k => (dr.d[k] || "").trim() || "[details needed]";
  const screen = NOTE_QS.filter(q => yes(q) && q.k !== "pain");
  const out = [`Comments: Pt. arrived in ${dr.mood} mood, ` + (screen.length ? `Pt reports ${screen.map(q => `${q.flag}: ${det(q.k)}`).join("; ")}.` : "Pt reports no changes to daily routine screening questions.")];
  if (dr.comments.trim()) out.push("", dr.comments.trim());
  /* the yes/no observations feed these sentences; the questions themselves stay out of the note */
  const obs = [
    /^yes/i.test(dr.a.mus) ? "Involuntary muscle movements observed, within normal limits." : "No involuntary muscle movements observed.",
    dr.a.pain === "Yes" ? `Patient reported pain or discomfort during treatment: ${det("pain")}.` : "Patient denied discomfort."];
  out.push("", obs.join(" ") + " Pt successfully completed Tx w/out incident.");
  const n = typeof r.tx === "number" ? r.tx : p.d.txDone + 1;
  out.push("", `${n}/${p.d.total} completed`, "", `-${dr.initials.trim() || "[initials]"}`);
  return out.join("\n");
}
async function copyText(text, ta) {
  try { await navigator.clipboard.writeText(text); return true; } catch { /* fall back below */ }
  try { ta.focus(); ta.select(); if (document.execCommand("copy")) return true; } catch { /* manual copy */ }
  return false;
}
/* Rebuild from answers: drops the hand edits and redraws the note dialog */
function resetNote() {
  const v = findVisit(state.mctx && state.mctx.k); if (!v) return;
  const dr = draftOf(v.p, v.r); dr.edited = false; dr.text = ""; mNote(v);
}
/* note dialog inputs: answers update the draft and the note text; a Yes shows its "Specify" box */
function draftInput(el) {
  const det = el.closest("[data-draft]"); if (!det) return false;
  const dr = state.drafts[det.dataset.draft]; if (!dr) return false;
  const dk = el.dataset.dk, v = el.value;
  if (dk === "text") { dr.text = v; dr.edited = true; return true; }
  if (dk.startsWith("a.")) { dr.a[dk.slice(2)] = v; const d = det.querySelector(`[data-dk="d.${dk.slice(2)}"]`); if (d) d.hidden = v !== "Yes"; }
  else if (dk.startsWith("d.")) dr.d[dk.slice(2)] = v; else dr[dk] = v;
  if (!dr.edited) { const [tab, row] = dr.key.split("|"), p = state.patients.find(x => x.tab === tab), r = p.rows.find(x => x.row === +row); det.querySelector(".ntext").value = noteText(p, r, dr); }
  return true;
}

/* ---------- measures placeholders: the rule is TmsCore.planMeasures (core.js) ---------- */
const planMeasures = p => TmsCore.planMeasures(p.rows, p.txApproved || 36, todaySerial());
/* log: false when an import's own Edit log entry covers it */
async function placeMeasures(patients, totalsDone, log) {
  if (!totalsDone) await fixTotals(patients.map(p => p.tab), log);
  const plans = patients.map(p => ({ p, ops: planMeasures(p) })).filter(x => x.ops.length);
  if (!plans.length) return 0;
  const rowsOf = x => [...new Set(x.ops.map(([a]) => +a.slice(1)))].map(n => x.p.rows.find(r => r.row === n)).filter(Boolean);
  await applyOps(plans.flatMap(x => x.ops.map(([a, v]) => ({ tab: x.p.tab, a, v: [[v]] }))), plans.flatMap(x => rowsOf(x).map(r => ({ tab: x.p.tab, ...at(r) }))), log === false ? false : { kind: "measures", what: "Measures (PHQ-9, GAD-7) re-planned" });
  return plans.length;
}
/* the running pulse totals (G, I, K: this row's pulses plus the total above, as on New Template). A row without the formula, such
   as an F/U row typed in by hand, breaks the chain and every total below it restarts from 0, so any empty G, I or K in the log gets
   it back. Cells with a value or a formula of their own are left alone. Returns the number of cells filled. */
const TOTALS = [["F", "G"], ["H", "I"], ["J", "K"]];
async function fixTotals(tabs, log) {
  if (!tabs.length) return 0;
  if (POPOUT) return callPane("fixTotals", [tabs, log]);
  const ops = [];
  await Excel.run(async ctx => {
    for (const tab of tabs) {
      const ws = ctx.workbook.worksheets.getItem(tab), end = await tabEnd(ctx, ws), r = ws.getRange(`G${LOG_FIRST}:K${end}`);
      r.load("formulas"); await ctx.sync();
      r.formulas.forEach((v, i) => TOTALS.forEach(([p, t]) => {
        const c = t.charCodeAt(0) - 71, row = LOG_FIRST + i;
        if (v[c] === "" || v[c] == null) ops.push({ tab, a: `${t}${row}`, v: [[`=${p}${row} + INDEX(${t}:${t}, ROW()-1)`]], f: true });
      }));
    }
  });
  if (ops.length) await applyOps(ops, null, log === false ? false : { kind: "formulas", what: `Put back ${plural(ops.length, "pulse total")} (G, I, K)` });
  return ops.length;
}
async function placeMeasuresClick() {
  try {
    const P = state.patients.filter(p => p.d.active), t = await fixTotals(P.map(p => p.tab)), n = await placeMeasures(P, true);
    const parts = [n ? `measures replanned on ${n} tab${n > 1 ? "s" : ""}` : "", t ? `${t} pulse total${t > 1 ? "s" : ""} put back` : ""].filter(Boolean).join(", ");
    toast(parts ? parts[0].toUpperCase() + parts.slice(1) + "." : "Measures and pulse totals are already in place."); await refresh();
  } catch (e) { fail(e); }
}
/* ---------- Match New Template: New Template owns the layout. Each active tracker gets its header formats, Freebie Count in N6
   (auth assessment in N7), the PR formula in column A, its benefits moved to the Benefits tab, a TX SUMMARY row with its
   checkbox, and New Template's conditional formats and dropdowns stretched over the whole log. Tabs from an earlier layout
   are rebuilt as fresh copies of New Template. ---------- */
/* PR (patient responsibility) per visit, from the patient's Benefits row (matched on the MRN in A4). With a deductible remaining
   (D), visits charge the before-deductible cost (N:R) in date order until it is used up; the visit that crosses it pays what is
   left plus PR on the rest; then PR (I:M). Missed visits are not charged. No Benefits row: blank. */
const BR = c => `Benefits!$${c}$${BEN_FIRST}:$${c}$${BEN_LAST}`, BRX = (a, b) => `Benefits!$${a}$${BEN_FIRST}:$${b}$${BEN_LAST}`;
const prF = (r, last = LOG_LAST) => `=IF(OR($L${r}="",$M${r}="x",$A$4=""),"",LET(bi,IFERROR(MATCH($A$4,${BR("A")},0),MATCH($A$4&"",${BR("A")}&"",0)),IF(ISNA(bi),"",`
  + `LET(L,$L$${LOG_FIRST}:$L$${last},M,$M$${LOG_FIRST}:$M$${last},E,$E$${LOG_FIRST}:$E$${last},tl,ISNUMBER(SEARCH("tele",$S$${LOG_FIRST}:$S$${last})),`
  + `c,LAMBDA(col,(L="MT")*N(INDEX(col,1))+((L="Daily")+(L="Taper"))*N(INDEX(col,2))+(L="MTR")*N(INDEX(col,3))+(L="F/U")*NOT(tl)*N(INDEX(col,4))+(L="F/U")*tl*N(INDEX(col,5))),`
  + `aft,c(INDEX(${BRX("I", "M")},bi,0)),bef,c(INDEX(${BRX("N", "R")},bi,0)),i,ROW()-${LOG_FIRST - 1},a,INDEX(aft,i),b,INDEX(bef,i),d,INDEX(${BR("D")},bi),`
  + `IF(OR(NOT(ISNUMBER(d)),d<=0,b<=0),a,LET(key,IF(E="",1E6,E)*100+ROW(E),rem,MAX(0,d-SUMPRODUCT(bef*(M<>"x")*(key<INDEX(key,i)))),IF(rem<=0,a,IF(b<=rem,b,rem+a*(b-rem)/b))))))))`;
/* a PR cell is ours to replace when it is blank or still holds a cost-table lookup; a typed amount is left alone */
const isPrFormula = f => f === "" || f == null || (typeof f === "string" && f.startsWith("=") && /\$X\$12|\$AE\$18|LET\(/.test(f));
/* New Template's visit-log formulas, for log rows that lost theirs (rows inserted by hand) */
const ROW_F = { D: r => `=IF(C${r}="", "", $D$2-C${r})`, G: r => `=F${r} + INDEX(G:G, ROW()-1)`, I: r => `=H${r} + INDEX(I:I, ROW()-1)`, K: r => `=J${r} + INDEX(K:K, ROW()-1)` };
/* New Template's conditional formats, highest priority first, with the log's ranges stretched to its last row (end) and the
   summary row (s) turning green when its checkbox in A is ticked. Colors are the template's own (white text = #FFFFFF). */
const CF_W = "#FFFFFF", CF_K = "#000000";
function cfSpec(end, s) {
  const T = `T${LOG_FIRST}:T${end}`, A9 = `A9:S${end}`, A10 = `A${LOG_FIRST}:S${end}`;
  const daily = { fill: "#D9D9D9", font: CF_K }, cap = { fill: "#5497C6", font: CF_W }, proto = { fill: "#5497C6", font: CF_K }, meas = { fill: "#FFF2CC", font: CF_K };
  const mt = { fill: "#66FFFF", font: CF_K, bold: true }, alt = { fill: "#E79507", font: CF_W }, miss = { fill: "#7030A0", font: CF_W };
  const fu = { fill: "#FF99CC", font: CF_K, bold: true }, mtr = { fill: "#92D050", font: CF_W, bold: true };
  return [
    { a: `A${s}:T${s}`, f: `=$A$${s}=TRUE`, fill: "#C6EFCE", font: "#006100" },
    { a: T, f: '=$O10="Daily"', ...daily }, { a: T, f: '=$O10="Give Cap"', ...cap }, { a: T, f: '=$O10="Protocol"', ...proto }, { a: T, f: '=$O10="Measures"', ...meas },
    { a: T, f: '=$L10="MT"', ...mt }, { a: T, f: '=$R10="x"', ...alt }, { a: T, f: '=$M10="x"', ...miss }, { a: T, f: '=$L10="F/U"', ...fu }, { a: T, f: '=$L10="MTR"', ...mtr },
    { a: A9, f: '=$O9="Daily"', ...daily },
    { a: "E9", f: "=LEN(TRIM(E9))>0", fill: "#5B9BD5", font: CF_K },
    { a: "F7", f: '=AND(F7<>"",F7>$G$4)', fill: "#FFC7CE", font: "#9C0006" },
    { a: A10, f: '=$O10="Give Cap"', ...cap }, { a: A9, f: '=$O9="Protocol"', ...proto }, { a: A9, f: '=$O9="Measures"', ...meas },
    { a: A10, f: '=$L10="MT"', ...mt }, { a: `A9:A${end}`, f: '=$L9="MTR"', fill: "#92D050", font: CF_W },
    { a: A9, f: '=$R9="x"', ...alt }, { a: A10, f: '=$M10="x"', ...miss },
    { a: `E${LOG_FIRST}:E${end}`, cell: { formula1: "=$F$4", formula2: "=$G$4", operator: "NotBetween" }, fill: "#FF0000", font: CF_K },
    { a: A10, f: '=$L10="F/U"', ...fu }, { a: A10, f: '=$L10="MTR"', ...mtr }
  ];
}
const X_ONLY = { title: "Error", message: 'Mark Cell by using "x". Otherwise, leave blank' };
const dvSpec = end => [
  ["D6", "Yes,No"], ["F6", "Yes,No"], ["D7", "Pending Start,In Progress,Tapering,Paused,Discontinued,Complete"], ["O2:O4", SITES.join(",")],
  [`L${LOG_FIRST}`, "Daily,MT,MTR,F/U"], [`L${LOG_FIRST + 1}:L${end}`, "Daily,MT,MTR,F/U,Taper"], [`M${LOG_FIRST}:N${end}`, "x", X_ONLY],
  [`O${LOG_FIRST}:O${end}`, "Measures,Give Cap,Protocol"], [`V${AV_FIRST}:V${AV_LAST}`, AV_TYPE.join(",")]
];
/* replaces every conditional format on the tab with cfSpec, and the dropdowns in A1:T(summary) and the dates-away types with dvSpec */
function applyRules(ws, end, s) {
  ws.getRange().conditionalFormats.clearAll();
  const fmt = (f, x) => { f.fill.color = x.fill; f.font.color = x.font; if (x.bold) { f.font.bold = true; f.font.italic = false; } };
  /* add() puts each new rule at the top, so the lowest priority goes in first. Stop If True stays off: as on New Template,
     a rule that only colors lets a lower one add its bold (a missed or red-dated F/U row is still bold). */
  for (const x of cfSpec(end, s).reverse()) {
    const cfs = ws.getRange(x.a).conditionalFormats;
    const cf = cfs.add(x.cell ? Excel.ConditionalFormatType.cellValue : Excel.ConditionalFormatType.custom);
    if (x.cell) { cf.cellValue.rule = x.cell; fmt(cf.cellValue.format, x); } else { cf.custom.rule.formula = x.f; fmt(cf.custom.format, x); }
    cf.stopIfTrue = false;
  }
  ws.getRange(`A1:T${s}`).dataValidation.clear(); ws.getRange(`V${AV_FIRST}:V${AV_LAST}`).dataValidation.clear();
  for (const [a, source, err] of dvSpec(end)) {
    const dv = ws.getRange(a).dataValidation; dv.rule = { list: { inCellDropDown: true, source } };
    if (err) dv.errorAlert = { showAlert: true, style: "Stop", title: err.title, message: err.message };
  }
}
/* the benefits a previous-layout tab keeps in V1:Y16, as Benefits row entries (quiet: moving them is not an update). Every tracker
   gets its row, so the Benefits tab lists every patient. */
function benEntries(p, get) {
  const out = [{ mrn: p.mrn, mrnVal: p.mrnVal, tab: p.tab, key: BEN_STAMP, quiet: true }];   // the row exists even with nothing entered yet
  for (const k of BEN_KEYS) { const v = get(k); if (v !== "" && v != null) out.push({ mrn: p.mrn, mrnVal: p.mrnVal, tab: p.tab, key: k, v, quiet: true }); }
  return out;
}
const TPL_SUM = LOG_LAST + 1;   // New Template's TX SUMMARY row
async function matchTab(ctx, name) {
  const isTpl = name === "New Template", ws = ctx.workbook.worksheets.getItem(name), tpl = ctx.workbook.worksheets.getItem("New Template");
  const end = await tabEnd(ctx, ws), s = end + 1;
  const hd = ws.getRange("A1:Y17"), lg = ws.getRange(`A${LOG_FIRST}:T${s}`), tl = tpl.getRange("J6:J7"); hd.load("values"); lg.load("values,formulas"); tl.load("values");
  await ctx.sync();
  const V = hd.values, at = a => { const m = a.match(/^([A-Z]+)(\d+)$/); return V[+m[2] - 1][colIndex(m[1])]; };
  const layout = layoutOf(at("T8"), at("J6"), at("V11")), p = state.patients.find(x => x.tab === name);
  /* 1. the benefits block moves to the Benefits tab (its values are the live ones: the tab has not been moved over yet) */
  if (!isTpl && layout === 2 && p) await writeBen(ctx, benEntries(p, k => at(k)));
  if (layout === 2 || isTpl) ws.getRange("V1:AA17").clear(Excel.ClearApplyTo.all);
  /* 2. header: New Template's formats; freebies left move to N6 and the auth assessment to N7 */
  if (!isTpl) {
    ws.getRange("A1:T9").copyFrom(tpl.getRange("A1:T9"), Excel.RangeCopyType.formats);
    if (layout === 2) ws.getRange("N6:N7").values = [[typeof at("N7") === "number" ? at("N7") : 3], [at("N6") ?? ""]];
    ws.getRange("J6:J7").values = tl.values;
  }
  ws.getRange("D4").numberFormat = [[D4_FMT]]; ws.getRange("A4").numberFormat = [[A4_FMT]];
  /* 3. visit log: the PR formula over the whole log, and the template's formulas on rows that lost them */
  const F = lg.formulas;
  ws.getRange(`A${LOG_FIRST}:A${end}`).formulas = F.slice(0, end - LOG_FIRST + 1).map((v, i) => [isPrFormula(v[0]) ? prF(LOG_FIRST + i, end) : v[0]]);
  for (let i = 1; i <= end - LOG_FIRST; i++) for (const [c, fx] of Object.entries(ROW_F)) if (F[i][colIndex(c)] === "") ws.getRange(`${c}${LOG_FIRST + i}`).formulas = [[fx(LOG_FIRST + i)]];
  /* 4. the TX SUMMARY row: New Template's formats (merge, checkbox), FALSE in A until it's ticked */
  if (!isTpl) ws.getRange(`A${s}:T${s}`).copyFrom(tpl.getRange(`A${TPL_SUM}:T${TPL_SUM}`), Excel.RangeCopyType.formats);
  const sv = lg.values[s - LOG_FIRST];
  if (typeof sv[0] !== "boolean") ws.getRange(`A${s}`).values = [[false]];
  if (!SUM_RE.test(String(sv[1] ?? ""))) ws.getRange(`B${s}`).values = [[SUM_LABEL]];
  /* 5. colors by visit type and the dropdowns, over every log row, in a thin grid */
  applyRules(ws, end, s); thinGrid(ws, end);
  await ctx.sync();
  return { name, end };
}
/* copy an earlier-layout tab's entries into a fresh copy of New Template. Every copied cell is read back from the new tab;
   only when all of them match is the earlier tab deleted. Otherwise it stays as "NAME (old)". Benefits go to the Benefits tab. */
async function rebuildTab(ctx, name) {
  const old = ctx.workbook.worksheets.getItem(name), tpl = ctx.workbook.worksheets.getItem("New Template");
  const src = old.getRange("A1:AN49"); src.load("formulas"); old.load("position"); await ctx.sync();
  const F = src.formulas, at = a => { const m = a.match(/^([A-Z]+)(\d+)$/), c = colIndex(m[1]); return F[+m[2] - 1][c]; };
  const isK = v => !(typeof v === "string" && v.startsWith("="));   // a typed entry, not a formula
  const txt = a => String(at(a) ?? "").trim();
  const pos = old.position; old.name = `${name} (old)`.slice(0, 31);
  const ws = tpl.copy(Excel.WorksheetPositionType.end); ws.name = name; ws.position = pos;
  const wrote = new Map(), put = (a, v) => { ws.getRange(a).values = [[v]]; wrote.set(a, v); }, ben = {};
  ["A2", "A4", "A6", "D2", "D4", "D6", "D7", "E2", "E4", "E6", "F2", "F4", "F6", "G2", "G4", "G6", "E9", "P9", "Q9", "S9"].forEach(a => { const v = at(a); if (v !== "" && isK(v)) put(a, v); });
  ws.getRange("D4").numberFormat = [[D4_FMT]]; ws.getRange("A4").numberFormat = [[A4_FMT]];
  for (let r = 2; r <= 4; r++) { ["J", "K", "L", "M"].forEach(c => { const v = at(c + r); if (isK(v)) put(c + r, v); }); if (txt("O1") === "Site" && SITES.includes(txt("O" + r))) put("O" + r, txt("O" + r)); }
  const assess = at("M7"); if (assess !== "" && isK(assess) && !/:$/.test(String(assess))) put("N7", assess);
  /* visit log: typed entries only, so the template's formulas (A, D, G, I, K) stay */
  const pilot = txt("T8").toLowerCase() === "freebie", timeCol = pilot ? "U" : txt("T8").toLowerCase() === "time" ? "T" : null;
  let used = 0;
  for (let r = LOG_FIRST; r <= LOG_LAST; r++) {
    ["B", "C", "E", "F", "H", "J", "L", "M", "N", "O", "P", "Q", "R"].forEach(c => { const v = at(c + r); if (isK(v)) put(c + r, v); });
    let note = isK(at("S" + r)) ? String(at("S" + r) ?? "") : "";
    if (pilot && FREEBIE_TAG[txt("T" + r)]) { note = withFreebie(note, txt("T" + r)); if (txt("T" + r) === "Used") used++; }
    if (note !== "") put("S" + r, note);   // a blank keeps the template's prompt (MT row, final Kaiser reminder)
    if (timeCol) { const v = at(timeCol + r); if (isK(v)) put("T" + r, v); }
  }
  /* benefits, costs and availability from the pilot add-in's columns, when present */
  if (txt("Y7") === "Plan type") [["Z9", "W4"], ["Z10", "W5"], ["Z11", "W6"], ["Z12", "W7"], ["Z13", "W8"], ["Z15", "W1"]].forEach(([s, d]) => { const v = at(s); if (v !== "" && isK(v)) ben[d] = v; });
  if (txt("Y16") === "BILLING AND AUTH BY CODE") for (let i = 0; i < 5; i++) { const v = at(`AE${18 + i}`); if (typeof v === "number") ben[`X${COST_FIRST + i}`] = v; }
  if (txt("AD25") === "WEEKLY AVAILABILITY") for (let i = 0; i < 5; i++) ["AF", "AG"].forEach((c, j) => { const v = at(`${c}${27 + i}`); if (typeof v === "number") put(`${j ? "X" : "W"}${WK_FIRST + i}`, v); });
  if (txt("Y26") === "Type") {
    const rg = []; for (let r = 27; r <= 48; r++) if (AV_TYPE.includes(txt("Y" + r))) rg.push(["Y", "Z", "AA", "AB"].map(c => at(c + r)));
    rg.slice(0, AV_LAST - AV_FIRST + 1).forEach((v, i) => ["V", "W", "X", "Y"].forEach((c, j) => put(`${c}${AV_FIRST + i}`, v[j])));
  }
  const fb = old.getRange("U4"); fb.load("values"); await ctx.sync();
  put("N6", txt("T1") === "Freebies allowed" && typeof fb.values[0][0] === "number" ? fb.values[0][0] : Math.max(0, 3 - used));
  const mrnVal = at("A4"), mrn = normMrn(mrnVal);
  if (mrn) await writeBen(ctx, benEntries({ mrn, mrnVal, tab: name }, k => ben[k]));
  await ctx.sync();
  /* check the transfer, then delete the earlier tab */
  const back = ws.getRange("A1:Y49"); back.load("values"); await ctx.sync();
  const same = (a, b) => (a === "" || a == null) ? (b === "" || b == null) : typeof a === "number" ? typeof b === "number" && Math.abs(a - b) < 1e-9 : String(a) === String(b);
  const now = a => { const m = a.match(/^([A-Z]+)(\d+)$/); return back.values[+m[2] - 1][colIndex(m[1])]; };
  const bad = [...wrote].filter(([a, v]) => !same(v, now(a))).map(([a, v]) => ({ a, old: v, now: now(a) }));   // for the side-by-side check
  if (!bad.length) { old.delete(); await ctx.sync(); }
  return { name, old: `${name} (old)`.slice(0, 31), deleted: !bad.length, bad };
}
/* rebuilding deletes the earlier tabs, so it asks first. force: every active tracker, not only the ones that look off */
function setupClick(force) {
  if (!state.templateOk) return toast(state.templateWhy, true, true);
  if (!state.old.length) return setupColumns(force === true);
  const n = state.old.length;
  showModal(`<h2 id="mtitle">Rebuild ${n} tab${n > 1 ? "s" : ""} from New Template?</h2><p class="sub">${state.old.map(esc).join(", ")} will be copied onto New Template. Each earlier tab is deleted once every copied entry checks out on the new tab. If anything doesn't match, the earlier tab is kept as "(old)".</p>
    <p class="sub">To get a deleted tab back, use the workbook's version history (File > Info > Previous versions).</p>
    <div class="mfoot"><button class="btn mute" data-mclose>Cancel</button><button class="btn primary" data-act="setupGo">Rebuild and delete</button></div>`);
}
async function setupColumns(force) {
  if (POPOUT) return callPane("setupColumns", [force]);
  try {
    if (!state.templateOk) return toast(state.templateWhy, true, true);
    const stale = new Set(state.stale), courseOf = n => (state.patients.find(p => p.tab === n) || { course: 0 }).course;
    /* oldest course first, so the newest course of an MRN writes its Benefits row last */
    const names = (force ? state.patients.filter(p => !stale.has(p.tab)).map(p => p.tab) : [...state.fix.keys()].filter(n => n !== "New Template")).sort((a, b) => courseOf(a) - courseOf(b));
    const res = [], matched = [];
    await Excel.run(async ctx => {
      await benSheet(ctx);   // first: a PR formula written before the Benefits tab exists would turn into #REF!
      if (force || state.fix.has("New Template")) await matchTab(ctx, "New Template");
      for (const n of state.old) res.push(await rebuildTab(ctx, n));
      for (const n of names) matched.push((await matchTab(ctx, n)).name);
      await addEdits(ctx, [...state.old.map(n => ({ tab: n, kind: "tab", what: "Rebuilt from New Template (earlier layout)", undo: null })), ...matched.map(n => ({ tab: n, kind: "tab", what: "Matched New Template (formats, PR formula, dropdowns)", undo: null }))]);
      /* the Benefits tab lists every tracker: finished ones in the previous layout keep their own cells, and get a copy here */
      const rest = state.patients.filter(p => p.layout === 2 && stale.has(p.tab) && p.mrn && !state.ben.rows.has(p.mrn)).sort((a, b) => b.course - a.course)
        .filter((p, i, a) => a.findIndex(q => q.mrn === p.mrn) === i);
      for (const p of rest) {
        const b = p.ben, cost = Object.fromEntries(BILL.flatMap((x, i) => [[`X${COST_FIRST + i}`, p.cost[i].after], [`Y${COST_FIRST + i}`, p.cost[i].before]]));
        const vals = { W1: b.bidf, W4: b.dedNA ? "N/A" : b.dedRem, W5: b.oopNA ? "N/A" : b.oopRem, W6: b.mtr, W7: b.ext || "", W8: b.lastSched, ...cost };
        await writeBen(ctx, benEntries(p, k => vals[k]));
      }
      await ctx.sync();
    });
    await refresh();
    /* the log rows follow New Template's formatting (dated rows like row 10, undated ones like the template's blank rows), and measures are placed */
    const done = state.patients.filter(p => matched.includes(p.tab) || res.some(x => x.name === p.tab));
    await formatLogs(done.map(p => ({ tab: p.tab, touched: p.rows.map(r => r.row) })));
    await placeMeasures(done.filter(p => p.d.active));
    const kept = res.filter(x => !x.deleted), rebuilt = res.length - kept.length;
    const parts = [matched.length ? `${matched.length} tab${matched.length === 1 ? "" : "s"} now match New Template` : "", rebuilt ? `${rebuilt} earlier tab${rebuilt === 1 ? "" : "s"} rebuilt and deleted` : ""].filter(Boolean);
    toast(kept.length
      ? `${parts.join(", ") || "Tabs updated"}. ${kept.length} tab${kept.length === 1 ? "" : "s"} to check before the old one is deleted.`
      : `${parts.join(", ") || "Tabs already match New Template"}.`, kept.length > 0);
    await refresh();
    if (kept.length) { state.verify = kept; showVerify(); }
  } catch (e) { fail(e); }
}
/* a rebuilt tab whose cells didn't all read back the same: each field, old tab and new tab side by side, for someone to confirm
   before the old tab is deleted. Keeping it leaves "NAME (old)". One tab at a time, from state.verify. */
function showVerify() {
  const x = (state.verify || [])[0]; if (!x) return;
  const fmt = (v, date) => v === "" || v == null ? "(blank)" : date && typeof v === "number" ? fmtDateY(v) : String(v);
  const rows = x.bad.map((b, i) => { const l = TmsCore.cellLabel(b.a); return `<tr><td>${esc(l.label)} <span class="sub">${esc(b.a)}</span></td><td>${esc(fmt(b.old, l.date))}</td><td>${esc(fmt(b.now, l.date))}</td><td><label class="opt"><input type="checkbox" data-vok="${i}"><span>Matches</span></label></td></tr>`; }).join("");
  state.mctx = { verify: x };
  showModal(`<h2 id="mtitle">Check ${esc(x.name)} before the old tab is deleted</h2><p class="sub">${x.bad.length} field${x.bad.length === 1 ? "" : "s"} didn't read back the same on the new tab. Compare each one; when they all match, delete ${esc(x.old)}. Otherwise keep it and fix the new tab by hand.</p>
    <table class="tbl"><tr><th>Field</th><th>Old tab (${esc(x.old)})</th><th>New tab (${esc(x.name)})</th><th></th></tr>${rows}</table>
    ${mfoot("Delete the old tab", "verifyGo", `<button class="btn ghost" data-act="verifyKeep">Keep the old tab</button>`)}`, true);
}
function verifyNext() { state.verify = (state.verify || []).slice(1); closeModal(); if (state.verify.length) showVerify(); }
async function verifyGo() {
  const x = (state.mctx || {}).verify; if (!x) return;
  if ([...document.querySelectorAll("[data-vok]")].some(c => !c.checked)) return merr("Tick Matches on every field first, or keep the old tab.");
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(x.old); await ctx.sync();
      if (!ws.isNullObject) ws.delete();
      await addEdits(ctx, [{ tab: x.name, kind: "tab", what: `Deleted ${x.old} after ${x.bad.map(b => b.a).join(", ")} were checked by hand`, undo: null }]);
      await ctx.sync();
    });
    toast(`Deleted ${x.old}.`); verifyNext(); await refresh();
  } catch (e) { fail(e); }
}
function verifyKeep() { const x = (state.mctx || {}).verify; if (x) toast(`Kept ${x.old}. Fix ${x.name} by hand, then delete the old tab.`); verifyNext(); }
/* Use existing tab: a patient offered a tracker (New patients, Create a tracker here?, a guest) goes onto a tab already in the workbook
   instead of a new one. Only tabs with no MRN, or this one, are offered, the same LASFIR first. The MRN goes in A4 (and the LASFIR in A2
   when it's blank); a tab in an earlier layout is rebuilt from New Template with its data moved over, and checked as Match New Template does. */
function useTabAsk(mrn) {
  if (POPOUT) return toast("Use this from the task pane.", true);
  const y = state.pendingNew.find(z => z.mrn === mrn) || [...(state.asks || []), ...(state.guests || [])].find(z => z.mrn === mrn); if (!y) return;
  if (!state.templateOk) return toast(state.templateWhy, true, true);
  const cur = state.tabs.map(t => { const p = state.patients.find(q => q.tab === t.name); return { name: t.name, code: t.code, mrn: p ? p.mrn : "", old: false }; });
  const tabs = [...cur, ...state.old.map(n => ({ name: n, code: n.split(" ")[0], mrn: "", old: true }))].filter(t => !t.mrn || t.mrn === mrn)
    .sort((a, b) => (b.code === y.code) - (a.code === y.code) || a.name.localeCompare(b.name));
  if (!tabs.length) return toast("Every tracker tab already has an MRN. Create a tracker instead.", true);
  state.mctx = { use: y };
  showModal(`<h2 id="mtitle">Use an existing tab for ${esc(y.code || "this patient")}</h2><p class="sub">MRN ending ${esc(String(y.raw).slice(-4))} goes in A4 of the tab you pick. A tab in an earlier layout is rebuilt from New Template with its data moved over; any field that doesn't copy cleanly is shown side by side for you to check before the old tab is deleted.</p>
    <div class="qrow"><span>Tab</span><select id="useTabSel" aria-label="Tab">${tabs.map(t => `<option value="${esc(t.name)}">${esc(t.name)}${t.mrn ? " (this MRN)" : " (no MRN)"}${t.old ? ", earlier layout" : ""}</option>`).join("")}</select></div>
    ${mfoot("Use this tab", "useTabGo")}`);
}
async function useTabGo() {
  const y = (state.mctx || {}).use, name = ($("useTabSel") || {}).value; if (!y || !name) return;
  const rebuild = state.old.includes(name), match = !rebuild && state.fix.has(name);
  let res = null;
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(name), hd = ws.getRange("A2:A4"); hd.load("values"); await ctx.sync();
      const had = normMrn(hd.values[2][0]);
      if (had && had !== y.mrn) throw new Error(`${name} already has another MRN in A4.`);
      ws.getRange("A4").values = [[y.raw]]; ws.getRange("A4").numberFormat = [[A4_FMT]];
      if (y.code && String(hd.values[0][0] ?? "").trim() === "") ws.getRange("A2").values = [[y.code]];
      await ctx.sync();
      if (rebuild) { await benSheet(ctx); res = await rebuildTab(ctx, name); } else if (match) await matchTab(ctx, name);
      await addEdits(ctx, [{ tab: name, kind: "tab", what: `Assigned MRN ending ${String(y.raw).slice(-4)} from the appointment list${rebuild ? "; rebuilt from New Template (earlier layout)" : match ? "; matched New Template" : ""}`, undo: null }]);
      await ctx.sync();
    });
    closeModal();
    state.pendingNew = state.pendingNew.filter(x => x.mrn !== y.mrn);
    await refresh();
    toast(res && !res.deleted ? `${y.code || "The patient"} is on ${name}. Check the fields that didn't copy over.` : `${y.code || "The patient"} is on ${name}.`);
    if (res && !res.deleted) { state.verify = [res]; showVerify(); }
    else if (state.file) { show("imp"); await buildPlan(); }
  } catch (e) { fail(e); }
}

/* ---------- add patient ---------- */
function apCode() { return ($("apCode").value || lasfir($("apLast").value, $("apFirst").value)).toUpperCase().replace(/[^A-Z]/g, ""); }
function updateApTab() {
  const code = apCode(), mx = Math.max(0, ...state.tabs.filter(t => t.code === code).map(t => t.course));
  if (!state.courseEdited) $("apCourse").value = mx + 1;
  const c = +$("apCourse").value || 1;
  $("apTab").textContent = code ? `Tab name: ${c > 1 ? `${code} ${c}` : code}${mx ? ` (existing courses: ${mx})` : ""}` : "";
}
function renderApRanges() {
  $("apRanges").innerHTML = state.apRanges.map((w, i) => `<div class="g4r"><input type="date" data-apr="${i}" data-k="from" value="${w.from || ""}" aria-label="Away from"><input type="date" data-apr="${i}" data-k="to" value="${w.to || ""}" aria-label="Away to"><input data-apr="${i}" data-k="note" value="${esc(w.note || "")}" placeholder="Reason"><button class="sm" data-apact="del" data-v="${i}" aria-label="Remove">×</button></div>`).join("") || '<div class="none">None</div>';
}
function fillAddPatientForm() {
  if (!$("apBill").children.length) {
    $("apBill").innerHTML = BILL.map((b, i) => `<div class="row3"><span>${b.code}<br>${esc(b.label)}</span><input id="apX${i}" type="number" min="0" step="0.01" aria-label="${esc(b.label)} PR"><input id="apY${i}" type="number" min="0" step="0.01" aria-label="${esc(b.label)} before deductible"></div>`).join("");
    $("apWeek").innerHTML = [1, 2, 3, 4, 5].map(d => `<div class="row3"><span>${DAYS[d]}</span><input id="apWf${d}" type="time" aria-label="${DAYS[d]} earliest start"><input id="apWt${d}" type="time" aria-label="${DAYS[d]} latest start"></div>`).join("");
    state.apRanges = state.apRanges || [{}]; renderApRanges();
  }
}
const AP_FIELDS = ["apLast", "apFirst", "apCode", "apCourse", "apMrn", "apDob", "apHome", "apMd", "apIns", "apAuthStart", "apAuthExp", "apAssess", "apTxs", "apRedos", "apMtr", "apDed", "apOop", "apBidf"];
function collectAddPatient() {
  document.querySelectorAll("[data-apr]").forEach(i => { state.apRanges[+i.dataset.apr][i.dataset.k] = i.value; });
  const d = Object.fromEntries(AP_FIELDS.map(id => [id, $(id).value.trim()]));
  d.code = apCode(); d.cost = BILL.map((b, i) => [$("apX" + i).value, $("apY" + i).value]);
  d.week = [1, 2, 3, 4, 5].map(k => [$("apWf" + k).value, $("apWt" + k).value]); d.ranges = state.apRanges.filter(w => w.from).map(w => ({ ...w }));
  return d;
}
function resetAddPatient() {
  AP_FIELDS.forEach(i => $(i).value = ""); BILL.forEach((b, i) => { $("apX" + i).value = ""; $("apY" + i).value = ""; });
  [1, 2, 3, 4, 5].forEach(k => { $("apWf" + k).value = ""; $("apWt" + k).value = ""; });
  $("apHome").value = clinic() || "SRL"; $("apTxs").value = 36; $("apRedos").value = 0; $("apCourse").value = 1; $("apNote").textContent = ""; state.apRanges = [{}]; renderApRanges();
  state.codeEdited = state.courseEdited = false; updateApTab();
}
/* copies New Template and fills the header, benefits, cost by code and availability in the template's own cells */
async function addPatient(d) {
  if (POPOUT) { const ok = await callPane("addPatient", [d]); if (ok) { resetAddPatient(); if (state.file) { show("imp"); await buildPlan(); } else openCourse(apTabName(d)); } return ok; }
  const code = d.code, course = +d.apCourse || 1, name = course > 1 ? `${code} ${course}` : code;
  if (!code) return toast("Enter a name or a LASFIR code.", true);
  if (!state.templateOk) return toast(state.templateWhy, true, true);
  const ded = benVal(d.apDed), oop = benVal(d.apOop), bidf = benVal(d.apBidf);
  if (ded === null || oop === null) return toast("Deductible and out-of-pocket max: enter an amount or N/A.", true);
  if (bidf === null || bidf === "N/A") return toast("BIDF estimate: enter an amount.", true);
  const tf = s => { const m = String(s || "").match(/^(\d{2}):(\d{2})$/); return m ? (+m[1] * 60 + +m[2]) / 1440 : ""; };
  try {
    await Excel.run(async ctx => {
      const wsAll = ctx.workbook.worksheets;
      const ex = wsAll.getItemOrNullObject(name), tpl = wsAll.getItemOrNullObject("New Template"); await ctx.sync();
      if (!ex.isNullObject) throw new Error(`A tab named ${name} already exists.`);
      if (tpl.isNullObject) throw new Error("There is no New Template tab to copy.");
      const ws = tpl.copy(Excel.WorksheetPositionType.end); ws.name = name; thinGrid(ws, LOG_LAST);
      const set = (a, val) => { if (val !== "" && val != null) ws.getRange(a).values = [[val]]; };
      set("A2", code); set("A4", d.apMrn); set("D2", +d.apTxs || 36); set("D4", d.apRedos === "" ? "" : +d.apRedos); set("D6", d.apMd); set("E6", d.apIns); set("N7", d.apAssess);
      set("A6", inputToSerial(d.apDob)); set("F4", inputToSerial(d.apAuthStart)); set("G4", inputToSerial(d.apAuthExp));
      ws.getRange("D4").numberFormat = [[D4_FMT]]; ws.getRange("A4").numberFormat = [[A4_FMT]]; ws.getRange("D7").values = [["Pending Start"]];
      if (d.apHome) ws.getRange("S9").values = [[`Home Clinic: ${d.apHome}`]];
      /* benefits and cost by code go on the patient's Benefits row */
      const mrn = normMrn(d.apMrn), bv = { W1: bidf, W4: ded, W5: oop, W6: d.apMtr, W7: 0 };
      d.cost.forEach(([x, y], i) => { bv[`X${COST_FIRST + i}`] = x === "" ? "" : +x; bv[`Y${COST_FIRST + i}`] = y === "" ? "" : +y; });
      if (mrn) await writeBen(ctx, Object.entries(bv).filter(([, v]) => v !== "" && v != null).map(([key, v]) => ({ mrn, mrnVal: d.apMrn, tab: name, key, v })));
      d.week.forEach(([f, t], i) => { if (f || t) ws.getRange(`W${WK_FIRST + i}:X${WK_FIRST + i}`).values = [[tf(f), tf(t)]]; });
      const rg = d.ranges.map(w => ["Unavailable", inputToSerial(w.from), inputToSerial(w.to) || inputToSerial(w.from), w.note || ""]).slice(0, AV_LAST - AV_FIRST + 1);
      if (rg.length) ws.getRange(`V${AV_FIRST}:Y${AV_FIRST + rg.length - 1}`).values = rg;
      await addEdits(ctx, [{ tab: name, kind: "tab", what: "Created the tab from New Template", undo: null }]);
      ws.activate(); await ctx.sync();
    });
    state.pendingNew = state.pendingNew.filter(x => x.mrn !== normMrn(d.apMrn));
    if (!POPOUT) resetAddPatient();
    await refresh();
    const np = state.patients.find(x => x.tab === name), meas = !!np && planMeasures(np).length > 0; if (meas) await placeMeasures([np]);
    toast(`Created ${name}.`); if (!state.file) openCourse(name);
    // buildPlan() refreshes on its own, so refresh here only when measures were written and no preview rebuild follows
    if (state.file && state.file.rows.some(r => normMrn(r[state.file.map.mMrn]) === normMrn(d.apMrn))) { show("imp"); await buildPlan(); toast(`Created ${name}. The preview now includes its dates.`); }
    else if (meas) await refresh();
    return true;
  } catch (e) { fail(e); }
}
const apTabName = d => (+d.apCourse || 1) > 1 ? `${d.code} ${+d.apCourse}` : d.code;
/* Don't ask again: kept in the workbook's settings, so nobody on this tracker is asked about that patient again; they become a guest */
const noAskList = () => { if (!POPOUT) try { state.noAsk = Office.context.document.settings.get("tms_noask") || state.noAsk || []; } catch { /* settings unavailable */ } return state.noAsk || []; };
function dontAsk(mrn) {
  const a = (state.asks || []).find(x => x.mrn === mrn); if (!a) return;
  state.noAsk = [...new Set([...noAskList(), mrn])]; saveSetting("tms_noask", state.noAsk);
  state.asks = state.asks.filter(x => x.mrn !== mrn); state.guests = [...(state.guests || []), a];
  toast(`${a.code || "That patient"} won't be asked about again. They're on the ${GUEST_SHEET} sheet: add the link to their ${a.home} tab.`);
  renderBanners(); syncGuests(); renderDash();
}
function startTracker(mrn) {
  const x = state.pendingNew.find(y => y.mrn === mrn), g = !x && [...(state.asks || []), ...(state.guests || [])].find(y => y.mrn === mrn); if (!x && !g) return;
  ["apLast", "apFirst", "apDob", "apIns", "apAuthStart", "apAuthExp", "apAssess"].forEach(i => $(i).value = "");
  const y = x || g; $("apMrn").value = y.raw; $("apCode").value = y.code; state.codeEdited = !!y.code; state.courseEdited = false;
  $("apNote").textContent = x ? `From NextGen: ${x.n} visit${x.n > 1 ? "s" : ""} today or later, first ${fmtDateY(x.first.date)} (${x.first.type}). Enter coverage and billing, then create the tab. The import preview reruns afterward to fill its dates.`
    : `${g.home} carries this course (the MT is there, or most chair visits), with ${g.daily} daily visit${g.daily === 1 ? "" : "s"} booked here. Create a tab here only if pulses will be entered here. The import preview reruns afterward to fill its dates.`;
  updateApTab(); show("add");
}

/* ---------- NextGen import: column layout is fixed by the "1 Appt List - ALL TMS" report, so it is detected, not mapped ---------- */
/* a list of patterns is tried in order, so "Provider" wins over "Res Name" when a report has both */
const MAP_RE = { mMrn: /^mrn$/i, mDate: /appt\s*date|appointment\s*date/i, mTime: [/^time$/i, /^beg(in)?\s*(tm|time)$/i], mSts: /^sts$|^status$/i, mEvent: /^event$|appointment type/i,
  mProv: [/^provider$/i, /^res(ource)?\s*name$/i], mChg: /charged/i, mMod: /^mod\s*dt|modified\s*date/i, mLoc: [/^location$/i, /^loc(ation)?\s*name$/i] };
const findCol = (headers, re) => { for (const x of [].concat(re)) { const i = headers.findIndex(h => x.test(h)); if (i >= 0) return i; } return -1; };
async function onFile(e) {
  // cleared at once so picking the same file again (a re-export, or after an error) still fires change
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  try {
    failStep = "Reading the file";
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    state.file = parseReport(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" }), f.name);
    failStep = "Building the preview"; await buildPlan(); failStep = "";
  } catch (err) {
    // a file that fails must not leave the last preview up, where Apply would write the earlier list
    if (state.file && !state.file.saved) state.file = null;
    state.plan = null; $("planBox").hidden = true; renderSavedList(); fail(err);
  }
}
function parseReport(all, name, saved) {
  const hi = all.findIndex(r => r.filter(c => String(c).trim() !== "").length >= 5);
  if (hi < 0) throw new Error("Could not find the report's header row.");
  let asOf = null, asOfMin = null;
  for (const r of all.slice(0, hi)) for (const c of r) { const m = String(c).match(/(\d{1,2}\/\d{1,2}\/\d{2,4})\s+(\d{1,2}:\d{2}\s*[AaPp][Mm])/); if (m) { asOf = toSerial(m[1]); asOfMin = toMinutes(m[2]); } }
  const headers = all[hi].map(h => String(h).trim());
  const map = Object.fromEntries(Object.entries(MAP_RE).map(([k, re]) => [k, findCol(headers, re)]));
  map.mPat = headers.findIndex(h => /other id/i.test(h)); if (map.mPat < 0) map.mPat = headers.findIndex(h => /pat\s*name|^patient$/i.test(h));
  const missing = [["mMrn", "MRN"], ["mDate", "Appt Date"], ["mSts", "Sts"], ["mEvent", "Event"]].filter(([k]) => map[k] < 0).map(([, n]) => n);
  if (missing.length) throw new Error(`This doesn't look like the 1 Appt List - ALL TMS report. Missing: ${missing.join(", ")}.`);
  return { headers, rows: all.slice(hi + 1).filter(r => r.some(c => String(c).trim() !== "")), asOf, asOfMin, map, name, saved: !!saved };
}

/* ---------- Sohar benefits import (benefits.js does the matching; this shows the plan and applies ticked rows) ---------- */
async function onBenFile(e) {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", raw: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: "" });
    const parsed = SoharBenefits.parseSoharBenefits(rows);
    if (parsed.error) throw new Error(parsed.error);
    state.benPlan = SoharBenefits.planBenefits(parsed.records, state.patients);
    state.benPlan.file = f.name;
    state.benPlan.otherService = parsed.otherService || 0;
    renderBenPlan();
  } catch (err) { fail(err); }
}
function renderBenPlan() {
  const plan = state.benPlan, box = $("benPlan"); if (!plan) { box.hidden = true; return; }
  const show = v => v === "" || v == null ? "blank" : typeof v === "number" ? money(v) : esc(v);
  const n = plan.rows.reduce((s, r) => s + r.changes.filter(c => c.on && c.kind !== "same").length, 0), st = benStamps(plan).length;
  const rows = plan.rows.map((r, i) => {
    const x = r.rec, info = [x.payer, x.plan, x.copay !== "" ? `copay $${x.copay}` : "", x.coins ? `coinsurance ${x.coins}` : "", x.flag].filter(Boolean).map(esc).join(" · ");
    const ch = r.changes.filter(c => c.kind !== "same").map((c, j) => `<label class="chk"><input type="checkbox" data-benrow="${i}" data-benc="${r.changes.indexOf(c)}" ${c.on ? "checked" : ""}> ${esc(c.label)} (${benTarget(r.tab) && isBenKey(c.cell) ? BEN_SHEET + " tab" : c.cell}): ${show(c.from)} &rarr; <b>${show(c.to)}</b>${c.kind === "change" ? ' <span class="chip amb">differs</span>' : ""}</label>`).join("");
    const same = r.changes.filter(c => c.kind === "same").map(c => esc(c.label)).join(", ");
    return `<div class="item${r.note && !r.changes.length ? " amb" : ""}"><b>${esc(r.tab)}</b> <span class="sub">MRN ${esc(r.mrnRaw)}</span><div class="sub">${info}</div>${ch}${same ? `<div class="sub">Already matches: ${same}</div>` : ""}${r.note ? `<div class="sub">${esc(r.note)}</div>` : ""}</div>`;
  }).join("");
  const un = plan.unmatched.length ? `<div class="hint">${plan.unmatched.length} patient${plan.unmatched.length > 1 ? "s" : ""} in the file ha${plan.unmatched.length > 1 ? "ve" : "s"} no tracker tab here (MRN ending ${plan.unmatched.map(r => esc(String(r.mrn).slice(-4))).join(", ")}).</div>` : "";
  const other = plan.otherService ? ` ${plan.otherService} patient${plan.otherService === 1 ? "" : "s"} booked only for ESK left out.` : "";
  box.innerHTML = `<div class="hint">${esc(plan.file)}: ${plan.rows.length} tracker tab${plan.rows.length === 1 ? "" : "s"} matched.${other}</div>${rows || '<div class="hint">No tabs matched by MRN.</div>'}${un}
    <div class="actions"><button id="benApply" ${n || st ? "" : "disabled"}>${n ? `Apply ${n} change${n === 1 ? "" : "s"}` : "Mark as checked today"}</button><button id="benClear">Clear</button></div>`;
  box.hidden = false;
  box.querySelectorAll("[data-benrow]").forEach(cb => cb.onchange = () => { plan.rows[+cb.dataset.benrow].changes[+cb.dataset.benc].on = cb.checked; renderBenPlan(); });
  $("benApply").onclick = applyBenPlan;
  $("benClear").onclick = () => { state.benPlan = null; renderBenPlan(); };
}
/* every patient the Sohar file verified (amounts in hand, even when they already match) gets Last updated = today on the Benefits tab */
const benStamps = plan => plan.rows.filter(r => r.changes.length && benTarget(r.tab)).map(r => ({ tab: r.tab, a: BEN_STAMP }));
async function applyBenPlan() {
  const ops = SoharBenefits.benefitOps(state.benPlan), st = benStamps(state.benPlan);
  if (!ops.length && !st.length) return;
  try {
    await applyOps([...ops, ...st]);
    const tabs = new Set(ops.map(o => o.tab)).size;
    toast(`Benefits imported: ${ops.length} cell${ops.length === 1 ? "" : "s"} on ${tabs} tab${tabs === 1 ? "" : "s"}${st.length ? `. Last updated set to today for ${st.length} patient${st.length === 1 ? "" : "s"}` : ""}.`);
    state.benPlan = null; renderBenPlan(); await refresh();
  } catch (e) { fail(e); }
}

/* The last applied ALL TMS list lives on the Appt List tab, so it survives closing the pane: new-patient banners and
   new tabs keep their dates without uploading it again. The run date is stored as text so parseReport reads it back. */
const LIST_SHEET = "Appt List", LIST_HEAD = 6;
const toCell = v => {
  if (!(v instanceof Date)) return v;
  const s = (v.getFullYear() < 1900 ? 0 : utcSerial(v.getFullYear(), v.getMonth() + 1, v.getDate())) + Math.round((v.getHours() * 3600 + v.getMinutes() * 60 + v.getSeconds()) / 60) / 1440;
  return s < 0 ? 0 : s;
};
function saveList(f) {
  const cols = f.headers.length, rows = f.rows.map(r => Array.from({ length: cols }, (_, i) => toCell(r[i] ?? "")));
  const dateCol = f.headers.map((_, i) => f.rows.some(r => r[i] instanceof Date));
  const meta = [["Run", f.asOf == null ? "unknown" : `${fmtDateY(f.asOf)} ${f.asOfMin != null ? fmtTime(f.asOfMin) : ""}`.trim()], ["File", f.name || ""], ["Imported", fmtDateY(todaySerial())],
    ["Replaced each time an ALL TMS list is applied. The add-in reads it back, so leave it as is.", ""]];
  const x = { headers: f.headers, rows, dateCol, meta };   // dates are serials by now, so this survives the trip from a pop-out
  return POPOUT ? callPane("writeList", [x]) : writeList(x);
}
async function writeList({ headers, rows, dateCol, meta }) {
  const cols = headers.length;
  const last = colLetter(cols - 1);
  await Excel.run(async ctx => {
    let ws = ctx.workbook.worksheets.getItemOrNullObject(LIST_SHEET); await ctx.sync();
    if (ws.isNullObject) ws = ctx.workbook.worksheets.add(LIST_SHEET); else ws.getRange().clear();
    ws.getRange("A1:B4").numberFormat = meta.map(() => ["@", "@"]); ws.getRange("A1:B4").values = meta; ws.getRange("A1:A3").format.font.bold = true;
    const hd = ws.getRange(`A${LIST_HEAD}:${last}${LIST_HEAD}`); hd.values = [headers]; hd.format.font.bold = true;
    if (rows.length) {
      const body = ws.getRange(`A${LIST_HEAD + 1}:${last}${LIST_HEAD + rows.length}`);
      body.numberFormat = rows.map(r => r.map((v, i) => !dateCol[i] ? "@" : v < 1 ? "h:mm AM/PM" : v % 1 ? "m/d/yyyy h:mm AM/PM" : "m/d/yyyy"));
      body.values = rows;
    }
    ws.getRange(`A${LIST_HEAD}:${last}${LIST_HEAD + rows.length}`).format.autofitColumns();
    await ctx.sync();
  });
  if (!POPOUT && (!state.file || state.file.saved)) { state.file = null; state.listLoaded = false; }   // a pop-out applied it: the pane reads it back on its next refresh
}
async function loadList() {
  let vals = null;
  await Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(LIST_SHEET); await ctx.sync(); if (ws.isNullObject) return;
    const u = ws.getUsedRange(true); u.load("rowIndex,columnIndex,rowCount,columnCount"); await ctx.sync(); vals = [];
    /* 500 rows per read: a long report can pass the cap on what one sync returns in Excel on the web */
    const c1 = colLetter(u.columnIndex), c2 = colLetter(u.columnIndex + u.columnCount - 1);
    for (let i = 0; i < u.rowCount; i += 500) {
      const r = ws.getRange(`${c1}${u.rowIndex + i + 1}:${c2}${u.rowIndex + Math.min(i + 500, u.rowCount)}`); r.load("values"); await ctx.sync(); vals.push(...r.values);
    }
  });
  if (!vals || vals.length <= LIST_HEAD) return null;
  const name = String((vals.find(r => String(r[0]).trim() === "File") || [])[1] || LIST_SHEET);
  return parseReport(vals, name, true);
}
function planTab(p, appts, from, to, o) {
  const R = p.rows.map(r => ({ ...r, _o: r }));
  const st = { cleared: 0, miss: 0, none: 0, unrecorded: 0, exempt: 0, unclass: 0, types: 0, sameCancel: 0, early: 0, earlyNew: 0, noMod: 0, elsewhere: 0, unmissed: 0, dup: 0, treated: 0 };
  const tag = (r, t) => { if (!r.note.split("; ").includes(t)) r.note = r.note ? `${r.note}; ${t}` : t; };
  const untag = (r, t) => { r.note = r.note.split("; ").filter(x => x !== t).join("; "); };
  const kindOf = a => a.status === "No-show" ? "No-show" : (a.status === "Cancelled" && a.mod != null && a.mod >= a.date) ? "Same-day cancel" : null;
  // 0. a miss an earlier list recorded that NextGen has since corrected (that day's visit is now Expected or Kept, with no no-show or
  //    same-day cancel) is undone, freebie included, so the visit isn't on the tab twice. Misses marked in the add-in stay.
  for (const r of R) {
    if (!r.miss || r.delivered || r.date == null) continue;
    const tags = r.note.split("; ").filter(x => x === `No-show ${fmtDate(r.date)} (NextGen)` || x === `Same-day cancel ${fmtDate(r.date)} (NextGen)`);
    const day = appts.filter(a => a.date === r.date && lpool(a.type) === lpool(r.type));
    if (!tags.length || day.some(kindOf) || !day.some(a => a.status === "Scheduled" || a.status === "Kept")) continue;
    r.miss = false; r.freebie = ""; [...tags, EXEMPT_TAG, UNCLASS_TAG].forEach(t => untag(r, t)); st.unmissed++;
  }
  // 1. no-shows and same-day cancels first (any date): Miss rows that need a freebie classification
  for (const a of appts.filter(a => a.status === "No-show" || a.status === "Cancelled")) {
    const kind = kindOf(a);
    if (a.status === "Cancelled" && !kind) {
      st.early++; if (o.lastAsOf != null && a.mod != null && a.mod >= o.lastAsOf) st.earlyNew++; if (a.mod == null) st.noMod++;
      // cancelled ahead of time: clear the date (any date) unless the patient is still booked or seen that day; the row is free for another visit
      const r0 = R.find(x => x.date === a.date && !x.delivered && (o.ps || !x.alt) && !x.miss);
      if (r0 && !o.keptPast.has(`${o.mrn}|${a.date}`) && !appts.some(b => b.date === a.date && (b.status === "Scheduled" || b.status === "Kept"))) {
        r0.date = null; r0.time = null; r0.freebie = ""; untag(r0, TELE_TAG); untag(r0, `Cancelled ${fmtDate(a.date)} (NextGen)`); if (o.ps) r0.alt = false; st.cleared++;
      }
      continue;
    }
    // a day already treated (pulses in) for this kind of visit: the no-show is a second booking NextGen kept, not a missed visit,
    // so it never takes a free row and dates it as a miss beside the treatment
    if (R.some(x => x.date === a.date && x.delivered && lpool(x.type) === lpool(a.type))) { st.treated++; continue; }
    // that day's row for the same kind of visit (a follow-up no-show never lands on the treatment row), else one with no type
    const day = R.filter(x => x.date === a.date && !x.delivered && (o.ps || !x.alt));
    let r = day.find(x => x.type !== "" && lpool(x.type) === lpool(a.type)) || day.find(x => x.type === "");
    if (!r) { r = R.find(x => !x.delivered && !x.alt && !x.miss && x.date == null && lpool(x.type) === lpool(a.type)); if (r) { r.date = a.date; r.time = a.time; } }
    if (!r) { st.unrecorded++; continue; }
    if (!r.miss) { r.miss = true; st.miss++; if (kind === "Same-day cancel") st.sameCancel++; }
    if (r.type === "") { r.type = a.type; st.types++; }
    r._p = true; tag(r, `${kind} ${fmtDate(a.date)} (NextGen)`);
    if (!r.freebie) {
      if (a.charged === "Waived") r.freebie = "Used";
      else if (a.charged === "Exempt") { tag(r, EXEMPT_TAG); st.exempt++; }
      else { tag(r, UNCLASS_TAG); st.unclass++; }
    }
  }
  // 2. today and later: Expected (and today's Kept) fill undelivered rows in row order, by exact type for external patients.
  //    A past Kept visit since the course's first dated row gets a row too when the tab has none that day.
  const start = Math.min(...p.rows.filter(r => r.date != null).map(r => r.date));
  // the same visit listed twice (for example on the provider's schedule and on the chair) counts once
  const evKey = new Set(), evAll = appts.filter(a => a.status === "Scheduled" && a.date >= from || a.status === "Kept" && a.date >= Math.min(from, start)).sort(cmp);
  const events = evAll.filter(a => { const k = `${a.date}|${a.type}|${a.time}`; if (evKey.has(k)) return false; evKey.add(k); return true; });
  st.dup = evAll.length - events.length;
  const strict = !!p.external || !!o.ps, norm = t => (t === "" || t === "Taper") ? "Daily" : t;
  const poolOf = r => strict ? norm(r.type) : (r.type === "F/U" ? "fu" : "tx"), evPool = e => strict ? e.type : (e.type === "F/U" ? "fu" : "tx");
  const eligible = r => !r._p && !r.delivered && (o.ps || !r.alt) && !r.miss;
  const inWin = r => eligible(r) && r.date != null && r.date >= from && r.date <= to;
  const over = r => eligible(r) && (r.date == null || r.date > to);
  // a miss NextGen hasn't recorded (step 1 matched no no-show or same-day cancel to it) still answers that day's booking at the same
  // time, so the visit isn't given a second row and everything after it pushed down. A booking at another time that day is another visit.
  const sameTime = (r, e) => r.time == null || e.time == null || r.time === e.time;
  const ownMiss = r => r.miss && !r._p && !r.delivered && r.date != null;
  // one whose L was emptied first takes back the type of that day's booking no open row holds, a treatment ahead of a follow-up
  for (const r of R.filter(r => ownMiss(r) && r.type === "")) {
    const held = e => R.some(x => x !== r && !x.miss && !x.delivered && x.date === e.date && poolOf(x) === evPool(e));
    const e = events.filter(e => e.date === r.date && sameTime(r, e) && !held(e)).sort((a, b) => (a.type === "F/U") - (b.type === "F/U"))[0];
    if (e) { r.type = e.type; st.types++; }
  }
  const fill = rows => {
    const f = { cleared: 0, types: 0, elsewhere: 0, past: 0 }, left = [];
    for (const pool of strict ? ["Daily", "MT", "MTR", "F/U"] : ["fu", "tx"]) {
      const free = rows.filter(r => poolOf(r) === pool && (inWin(r) || over(r)));
      const done = rows.filter(r => poolOf(r) === pool && r.delivered && r.date != null && r.date >= from);   // a visit already treated (pulses in) answers its own event, so it never takes a second row that day
      const missed = rows.filter(r => ownMiss(r) && r.type !== "" && poolOf(r) === pool);
      for (const e of events.filter(e => evPool(e) === pool)) {
        const k = done.findIndex(r => r.date === e.date); if (k >= 0) { done.splice(k, 1); continue; }
        const m = missed.findIndex(r => r.date === e.date && sameTime(r, e)); if (m >= 0) { missed.splice(m, 1); continue; }
        if (e.date < from && rows.some(r => poolOf(r) === pool && r.date === e.date)) continue;
        const r = free.shift(); if (!r) { left.push(e); continue; }
        if (e.date < from) f.past++;
        r.date = e.date; r.time = e.time; r.miss = false; r.freebie = ""; r.note = TmsCore.dropKaiser(r.note);
        if (!strict && pool === "tx" && !(r.type === "Taper" && e.type === "Daily") && r.type !== e.type) { r.type = e.type; f.types++; }
        if (e.type === "F/U") { if (e.tele) tag(r, TELE_TAG); else untag(r, TELE_TAG); }
        if (o.ps) { r.alt = !o.ours(e.loc); if (r.alt) f.elsewhere++; }
      }
      free.filter(r => r.date != null && r.date <= to).forEach(r => { r.date = null; r.time = null; r.miss = false; r.freebie = ""; untag(r, TELE_TAG); if (o.ps) r.alt = false; f.cleared++; });
    }
    return { f, left };
  };
  let rows = R.map(r => ({ ...r })), res = fill(rows);
  // more booked visits than free rows (each no-show or MT redo uses one up): add rows to the log, each after the visits it follows
  // and never past the last row, so they land inside the log and Excel stretches every reference to it
  const ins = [];
  if (res.left.length && o.grow) {
    const G = R.map((r, i) => ({ ...r, _d: rows[i].date }));
    for (const e of res.left) {
      let at = 0; G.forEach((r, i) => { if (r.delivered || (r._d != null && r._d <= e.date) || poolOf(r) === evPool(e)) at = i + 1; });
      at = Math.min(Math.max(at, 1), G.length - 1);
      G.splice(at, 0, { ...BLANK_ROW, type: e.type, pulses: [null, null, null], _o: null, _d: e.date });
    }
    if (G.length <= LOG_READ) { rows = G; res = fill(rows); }
  }
  for (const [k, v] of Object.entries(res.f)) st[k] = (st[k] || 0) + v;   // steps 0-1 counted some of these already
  st.none = res.left.length;
  // 3. every visit dated this week (from Sunday) or later has NextGen's time, including the first import on a tab typed by hand:
  //    a row the steps above left without one (earlier this week, already treated, missed, or past the list) takes the time of
  //    that day's booking of the same kind, each booking once
  const wk = weekStart(o.today ?? from), used = new Set(rows.filter(r => r.date != null && r.time != null).map(r => `${r.date}|${poolOf(r)}|${r.time}`));
  st.timed = 0;
  for (const r of rows) {
    if (r.date == null || r.date < wk || r.time != null) continue;
    const e = appts.filter(a => a.date === r.date && a.time != null && (a.status !== "Cancelled" || kindOf(a)) && evPool(a) === poolOf(r) && !used.has(`${a.date}|${evPool(a)}|${a.time}`)).sort(cmp)[0];
    if (e) { r.time = e.time; used.add(`${e.date}|${evPool(e)}|${e.time}`); st.timed++; }
  }
  rows.forEach((r, i) => { r.row = LOG_FIRST + i; });
  rows.forEach((r, i) => {
    if (r._o) return;
    const same = rows.filter(x => x._o && x.type === r.type).sort((x, y) => Math.abs(x.row - r.row) - Math.abs(y.row - r.row))[0];
    const src = same || rows.slice(0, i).reverse().find(x => x._o) || rows.find(x => x._o);
    ins.push({ row: r.row, src: src.row });
  });
  const ops = []; let fbDelta = 0;
  rows.forEach(r => {
    const b = r._o || BLANK_ROW;
    if (r.freebie !== b.freebie) { r.note = withFreebie(r.note, r.freebie); fbDelta += (r.freebie === "Used" ? 1 : 0) - (b.freebie === "Used" ? 1 : 0); }
    if (r.date !== b.date) ops.push([`E${r.row}`, r.date ?? ""]);
    if (r.time !== b.time) ops.push([`T${r.row}`, r.time == null ? "" : r.time / 1440]);
    if (r.type !== b.type) ops.push([`L${r.row}`, r.type]);
    if (r.alt !== b.alt) ops.push([`R${r.row}`, r.alt ? "x" : ""]);
    if (r.miss !== b.miss) ops.push([`M${r.row}`, r.miss ? "x" : ""]);
    if (r.note !== b.note) ops.push([`S${r.row}`, r.note]);
  });
  if (fbDelta) ops.push([p.fbCell, Math.max(0, p.d.left - fbDelta)]);   // NextGen "Waived" = a freebie used
  // the log reads top to bottom by date (undated open rows last; an undated treated or missed row stays after the visit above it),
  // and C numbers the treatments (MT, Daily, Taper, MTR) that weren't missed, so D (= D2 - C) is what's left
  let lk = -Infinity, lt = 0;
  const order = rows.map((r, i) => { if (r.date != null) { lk = r.date; lt = r.time ?? 0; } return r.date != null || r.delivered || r.miss ? { i, k: lk, t: lt } : { i, k: Infinity, t: 0 }; })
    .sort((a, b) => a.k - b.k || a.t - b.t || a.i - b.i).map(x => x.i);
  const txNo = TmsCore.txNumbers(order.map(i => rows[i]));   // the MT is 1, each treatment after it one more
  const txWas = r => r._o && isNum(r._o.tx) ? r._o.tx : r._o && String(r._o.tx ?? "").trim() !== "" ? String(r._o.tx) : "";
  const resorted = order.some((s, i) => s !== i), renum = rows.filter((r, i) => txWas(r) !== txNo[i]).length;
  // dates: what each row's E holds once the ops are written, checked before anything moves
  const arrange = resorted || renum ? { order, tx: txNo, dates: rows.map(r => r.date) } : null;
  // significant disruptions re-flow the measures: an MTR or F/U moved, or 2+ missed visits within a week
  const was = r => r._o || BLANK_ROW;
  const moved = rows.some(r => (r.type === "MTR" || r.type === "F/U") && was(r).date != null && was(r).date !== r.date) || rows.some(r => r.type === "MTR" && was(r).type !== "MTR");
  const missDates = rows.filter(r => r.miss && r.date != null).map(r => r.date).sort((a, b) => a - b), newMiss = new Set(rows.filter(r => r.miss && !was(r).miss).map(r => r.date));
  let multi = false; for (let i = 1; i < missDates.length; i++) if (missDates[i] - missDates[i - 1] <= 6 && (newMiss.has(missDates[i]) || newMiss.has(missDates[i - 1]))) multi = true;
  const dates = rows.filter(r => r.date !== was(r).date && r.date != null).map(r => r.date);
  // rows whose formatting must follow (see formatLogs): new rows, rows whose date changed, and rows the re-sort fills with another visit
  const touched = [...new Set([...ins.map(x => x.row), ...ops.filter(([a]) => /^E\d+$/.test(a)).map(([a]) => +a.slice(1)), ...(arrange ? order.map((s, i) => s !== i ? LOG_FIRST + i : 0).filter(Boolean) : [])])];
  return { name: p.tab, ops, ins, arrange, resorted, renum, touched, added: ins.length, set: dates.length, ...st, changed: ops.length > 0 || ins.length > 0 || !!arrange, reflow: moved || multi || ins.length > 0 || resorted };
}
async function buildPlan(quiet) {
  if (!state.file) return;
  const ps = !state.file.saved && document.querySelector('input[name=rtype]:checked').value === "ps", from = todaySerial(), m = state.file.map, dev = DEVICES[device()];
  if (!quiet) await refresh();
  const cell = (r, k) => m[k] >= 0 ? r[m[k]] : "";
  // device is decided per patient from where their daily treatments are booked, so an MT on a provider schedule follows its patient
  const tally = new Map();
  for (const r of state.file.rows) {
    const mrn = normMrn(cell(r, "mMrn")); if (!mrn) continue;
    const ev = String(cell(r, "mEvent") || ""), pv = String(cell(r, "mProv") || ""), t = tally.get(mrn) || { o: 0, u: 0, evOther: false }, me = mapEvent(ev);
    if (me && me.type === "Daily") { if (dev.other.test(pv)) t.o++; else if (dev.ours.test(pv)) t.u++; }
    if (dev.other.test(ev)) t.evOther = true; tally.set(mrn, t);
  }
  const isOther = mrn => { const t = tally.get(mrn); if (!t) return false; if (t.o || t.u) return t.o > t.u; return t.evOther && !state.patients.some(p => p.mrn === mrn); };
  /* which clinic carries each patient (core.js clinicRole). A chair is ours when its resource or location names this tracker's
     clinic. No chair visit here: left out. Most chair visits at another clinic: a guest, listed on the Guest Patients sheet with a
     link to their tab there, unless someone made a tab here anyway. A report whose chairs never name the clinic is not split. */
  const roleIn = new Map(), lasfir = r => { const c = String(cell(r, "mPat") ?? "").trim().toUpperCase(); return /^[A-Z]{6}$/.test(c) ? c : ""; };
  const isHere = (pv, loc) => loc ? atClinic(loc) : dev.ours.test(pv) ? atClinic(pv) : null;   // a provider's schedule with no location: unknown
  const whereOf = (pv, loc) => loc || pv.replace(/\s*(magv|mag\s?venture|brainsway|bw)\s*\d*$/i, "").trim();
  const split = !!clinic() && state.file.rows.some(r => { const pv = String(cell(r, "mProv") || ""), loc = String(cell(r, "mLoc") || ""); return dev.ours.test(pv) && atClinic(loc || pv); });
  // every report gets the no-tab check (core.js noTabRole); only one that names the clinic is split by clinicRole
  for (const r of state.file.rows) {
    const mrn = normMrn(cell(r, "mMrn")), pv = String(cell(r, "mProv") || "").trim(), loc = String(cell(r, "mLoc") || "").trim();
    if (!mrn || isOther(mrn) || dev.other.test(pv)) continue;
    const sts = mapSts(cell(r, "mSts")), ev = mapEvent(cell(r, "mEvent")), date = toSerial(cell(r, "mDate"));
    if (!ev || date == null || sts === "Skip" || sts === "Cancelled") continue;
    const x = roleIn.get(mrn) || { visits: [], raw: String(cell(r, "mMrn")).trim(), code: "" }; roleIn.set(mrn, x);
    x.visits.push({ chair: dev.ours.test(pv), here: split ? isHere(pv, loc) : null, where: whereOf(pv, loc), type: ev.type, date, sts, status: sts }); x.code = x.code || lasfir(r);
  }
  const hasTab = (mrn, code) => state.patients.some(p => p.mrn === mrn || (!p.mrn && code && p.code === code));
  /* "ask" (MT elsewhere, 16+ dailies here) is offered a tracker here until someone says Don't ask again; then it's a guest. Until a
     tab exists both are tracked at home: no tab, no import, their bookings on Today and Schedule marked as guests. No tab, no MT
     here and fewer than 16 treatments here is a guest too, its home clinic's link asked for on the Guest Patients banner. */
  const noAsk = new Set(noAskList());
  const roles = new Map([...roleIn].map(([mrn, x]) => { const g = hasTab(mrn, x.code) ? { role: "home" } : TmsCore.noTabRole(split ? TmsCore.clinicRole(x.visits) : { role: "home" }, x.visits, from); return [mrn, g.role === "ask" && noAsk.has(mrn) ? { ...g, role: "guest" } : g]; }));
  const roleOf = mrn => (roles.get(mrn) || { role: "home" }).role, awayRole = r => r === "guest" || r === "ask";
  const t0 = todaySerial(), hereDates = x => x.visits.filter(v => v.chair && v.here !== false).map(v => v.date);
  if (!ps) {
    const list = role => [...roles].filter(([, g]) => g.role === role).map(([mrn, g]) => {
      const x = roleIn.get(mrn), ds = hereDates(x), fut = ds.filter(d => d >= t0), past = ds.filter(d => d < t0);
      return { mrn, raw: x.raw, code: x.code, home: g.home, here: g.here, daily: g.daily, homeCount: g.homeCount, next: fut.length ? Math.min(...fut) : null, last: past.length ? Math.max(...past) : null, link: (state.guestLinks || {})[mrn] || "" };
    }).sort((a, b) => (a.next ?? 1e9) - (b.next ?? 1e9));
    state.guests = list("guest"); state.asks = list("ask");
    state.clinicSplit = { on: split, clinic: clinic(), none: [...roles.values()].filter(g => g.role === "none").length };
  }
  // from the ALL TMS list only: each visit's chair (the resource booked), and per patient the cancellations, whether anything is still booked, the last kept visit
  if (!ps) {
    const v = {}, pt = {}, c = {}, names = new Set(), cnt = {}, ap = [], apKey = new Set(), gv = [];
    for (const r of state.file.rows) {
      const mrn = normMrn(cell(r, "mMrn")), pv = String(cell(r, "mProv") || "").trim(), role = roleOf(mrn), ourChair = dev.ours.test(pv) && (!split || isHere(pv, String(cell(r, "mLoc") || "").trim()) !== false);
      if (!mrn || isOther(mrn) || dev.other.test(pv) || role === "none" || (awayRole(role) && !ourChair)) continue;
      const date = toSerial(cell(r, "mDate")), sts = mapSts(cell(r, "mSts")), ev = mapEvent(cell(r, "mEvent"));
      if (date == null || sts === "Skip" || !ev) continue;
      /* a guest's chair bookings here: on the Day view's list check, and on Today and Schedule as guests (state.lx.gv). No cancellations,
         no chair counts, since there's no tab here. */
      if (awayRole(role)) {
        if (sts !== "Cancelled" && date >= from - 31) {
          let time = m.mTime >= 0 ? toMinutes(cell(r, "mTime")) : null; if (time == null) time = toMinutes(cell(r, "mDate"));
          const k = `${mrn}|${date}|${ev.type}|${time}`;
          if (!apKey.has(k)) { apKey.add(k); ap.push({ mrn, code: lasfir(r), date, time, type: ev.type, tele: ev.tele, guest: true }); gv.push({ mrn, code: lasfir(r), date, time, type: ev.type, chair: pv, status: sts }); }
        }
        continue;
      }
      const x = c[mrn] || (c[mrn] = { cx: [], on: [], booked: false, lastKept: null });
      if (sts === "Cancelled") x.cx.push(date);
      else if (sts === "Scheduled" || sts === "Kept") {
        /* every visit on the list for the Day view's check (the last month on; the same visit listed twice counts once) */
        let time = m.mTime >= 0 ? toMinutes(cell(r, "mTime")) : null; if (time == null) time = toMinutes(cell(r, "mDate"));
        const k = `${mrn}|${date}|${ev.type}|${time}`, code = String(cell(r, "mPat") ?? "").trim().toUpperCase();
        if (date >= from - 31 && !apKey.has(k)) { apKey.add(k); ap.push({ mrn, code: /^[A-Z]{6}$/.test(code) ? code : "", date, time, type: ev.type, tele: ev.tele }); }
        x.on.push(date);
        if (sts === "Scheduled" && date >= from) x.booked = true;
        if (sts === "Kept" && date <= from) x.lastKept = Math.max(x.lastKept ?? 0, date);
        if (ourChair) { v[`${mrn}|${date}`] = pv; names.add(pv); const k = cnt[mrn] || (cnt[mrn] = {}); k[pv] = (k[pv] || 0) + 1; }
      }
    }
    for (const x of Object.values(c)) { x.cx = [...new Set(x.cx)].filter(d => !x.on.includes(d)); delete x.on; }   // rebooked the same day: not a cancellation
    for (const [mrn, k] of Object.entries(cnt)) pt[mrn] = Object.entries(k).sort((a, b) => b[1] - a[1])[0][0];
    const ds = ap.map(a => a.date);
    state.lx = { v, pt, c, chairs: [...names].sort(), ap, gv, lo: ds.length ? Math.min(...ds) : null, hi: ds.length ? Math.max(...ds) : null };
    state.patients.forEach(p => p.d = derive(p));
  }
  const by = new Map(), codes = new Map(), rawMrn = new Map(), clin = new Map(), keptPast = new Set(); let maxDate = from, ignored = 0;
  for (const r of state.file.rows) {
    const mrn = normMrn(cell(r, "mMrn")), pv = String(cell(r, "mProv") || "").trim();
    if (!mrn) continue;
    if (isOther(mrn) || dev.other.test(pv) || roleOf(mrn) === "none") { ignored++; continue; }
    if (awayRole(roleOf(mrn))) continue;   // tracked at their home clinic: the Guest Patients sheet, or the ask banner
    const date = toSerial(cell(r, "mDate")), sts = mapSts(cell(r, "mSts")), ev = mapEvent(cell(r, "mEvent"));
    if (date == null || sts === "Skip" || !ev) continue;
    const past = date < from;
    if (sts === "Kept" && past) keptPast.add(`${mrn}|${date}`);   // guards a cancel on the same day, and gets a row if the tab has none that day
    else if (sts !== "No-show" && sts !== "Cancelled" && past) continue;   // past visits never update, except no-shows, cancels and kept visits with no row
    let time = m.mTime >= 0 ? toMinutes(cell(r, "mTime")) : null; if (time == null) time = toMinutes(cell(r, "mDate"));
    if (sts === "Scheduled" || sts === "Kept") maxDate = Math.max(maxDate, date);
    const chg = String(cell(r, "mChg") ?? "").trim(), charged = /waive/i.test(chg) ? "Waived" : /exempt/i.test(chg) ? "Exempt" : "";
    rawMrn.set(mrn, String(cell(r, "mMrn")).trim());
    if (!by.has(mrn)) by.set(mrn, []);
    by.get(mrn).push({ date, time, status: sts, type: ev.type, tele: ev.tele, charged, mod: m.mMod >= 0 ? toSerial(cell(r, "mMod")) : null, loc: String(cell(r, "mLoc") || "").trim(), here: split && isHere(pv, String(cell(r, "mLoc") || "").trim()) === false ? false : null });   // false: booked at another clinic
    if (pv && !past && !/san rafael|magv|brainsway|mag\s?venture/i.test(pv)) { const c = clin.get(mrn) || {}, ln = TmsCore.providerLast(pv); if (ln) { c[ln] = (c[ln] || 0) + 1; clin.set(mrn, c); } }
    const c = String(cell(r, "mPat") ?? "").trim().toUpperCase(); if (/^[A-Z]{6}$/.test(c)) codes.set(mrn, c);
  }
  if (!ps) state.codeMrn = new Map([...codes].map(([mrn, c]) => [c, rawMrn.get(mrn) || mrn]));   // LASFIR -> MRN, to suggest a blank A4
  const latest = new Map(), latestCode = new Map();
  for (const p of state.patients) {
    if (p.mrn) { const c = latest.get(p.mrn); if (!c || p.course > c.course) latest.set(p.mrn, p); }
    const cc = latestCode.get(p.code); if (!cc || p.course > cc.course) latestCode.set(p.code, p);
  }
  const o = { lastAsOf: state.report && state.report.asOf != null ? state.report.asOf : null, today: from, ps, keptPast, ours: loc => !loc || atClinic(loc) };
  const plans = [], disc = [];
  if (!ps) { state.apptBy = new Map(); state.apptTo = maxDate; const ds = state.file.rows.map(r => toSerial(cell(r, "mDate"))).filter(x => x != null); state.apptLo = ds.length ? Math.min(...ds) : null;   // apptLo: the list's first day
    /* every booking per MRN, past Expected ones included, for pulses with no NextGen visit (noApptRows) */
    state.apptAll = new Map();
    for (const r of state.file.rows) {
      const mrn = normMrn(cell(r, "mMrn")), ev = mapEvent(cell(r, "mEvent")), date = toSerial(cell(r, "mDate")), sts = mapSts(cell(r, "mSts"));
      if (!mrn || !ev || date == null || sts === "Skip" || isOther(mrn) || dev.other.test(String(cell(r, "mProv") || ""))) continue;
      if (!state.apptAll.has(mrn)) state.apptAll.set(mrn, []); state.apptAll.get(mrn).push({ date, type: ev.type, status: sts });
    }
  }   // each tab's visits on the ALL TMS list, for the Edit log's checks
  for (const [mrn, appts] of by) {
    let p = latest.get(mrn); const code = codes.get(mrn);
    if (!p && code) { const q = latestCode.get(code); if (q && !q.mrn) p = q; }
    if (!p) continue;
    // a Discontinued tab is left as it is; the preview and the Day view's list check name any visits NextGen still has booked
    if (!onSched(p)) { if (appts.some(a => a.status === "Scheduled" && a.date >= from)) disc.push(p.tab); continue; }
    const pl = planTab(p, appts, from, maxDate, { ...o, mrn, grow: p.canGrow });
    pl.sig = logSig(p.rows); pl.fb = p.fbCell; if (!ps) state.apptBy.set(p.tab, appts);   // sig: the log this plan was made from (applyPlan checks it)
    // the report is the source of truth for the provider, so E2 follows it (last name only) even when something was typed there
    const c = clin.get(mrn), top = c && Object.entries(c).sort((a, b) => b[1] - a[1])[0][0]; if (top && top !== p.provider) pl.e2 = top;
    if (pl.e2) pl.changed = true;
    plans.push(pl);
  }
  const noTab = [...by.entries()].filter(([k]) => !latest.has(k) && !(codes.get(k) && latestCode.has(codes.get(k)) && !latestCode.get(codes.get(k)).mrn));
  // offered a tracker only with an MT booked here or 16+ treatments here (core.js needsTracker); F/U alone don't need one
  state.pendingNew = noTab.filter(([, a]) => TmsCore.needsTracker(a, from)).map(([k, a]) => { const fut = a.filter(x => x.status !== "No-show" && x.status !== "Cancelled" && x.date >= from).sort(cmp); return { mrn: k, raw: rawMrn.get(k) || k, last4: k.slice(-4), code: codes.get(k) || "", n: fut.length, first: fut[0] }; }).filter(x => x.n > 0);
  state.plan = { ps, plans, from, to: maxDate, ignored, disc };
  if (!quiet) renderPlan(); renderBanners(); renderSavedList();
  if (!ps) syncGuests();
}
const listRun = f => f.asOf == null ? "no run date" : `run ${fmtDateY(f.asOf)}${f.asOfMin != null ? " " + fmtTime(f.asOfMin) : ""}`;
function renderSavedList() {
  const f = state.file && state.file.saved ? state.file : null; $("savedList").hidden = !f;
  if (f) $("savedListMsg").textContent = `The ${LIST_SHEET} tab holds the last applied list, ${listRun(f)}, ${f.rows.length} rows.`;
}
function renderPlan() {
  const { ps, plans, from, to } = state.plan, ch = plans.filter(p => p.changed), tot = k => ch.reduce((n, p) => n + (p[k] || 0), 0);
  let h = `<h2>Preview${ps ? ": patient-specific report" : ""}</h2><div class="sub">${state.file.saved ? `Saved list (${esc(listRun(state.file))})` : esc(state.file.name || "")}. Today through ${fmtDateY(to)}. Past visits only change for no-shows, same-day cancels and kept visits with no row.</div>`;
  h += ch.length ? `<table><thead><tr><th>Tab</th><th class="num">Dates</th><th class="num">Types</th><th class="num">Cleared</th><th class="num">Missed</th></tr></thead><tbody>${ch.map(p => `<tr><td>${esc(p.name)}</td><td class="num">${p.set}</td><td class="num">${p.types}</td><td class="num">${p.cleared}</td><td class="num">${p.miss}</td></tr>`).join("")}
    <tr><th>Total</th><th class="num">${tot("set")}</th><th class="num">${tot("types")}</th><th class="num">${tot("cleared")}</th><th class="num">${tot("miss")}</th></tr></tbody></table>` : `<div class="none">The tracker already matches this report.</div>`;
  const grown = ch.filter(p => p.added), full = plans.filter(p => p.none);
  if (grown.length) h += `<div class="sub" style="margin-top:6px">Rows added to the visit log: ${grown.map(p => `${esc(p.name)} ${p.added}`).join(", ")}. They take the formatting of a row of the same visit type.</div>`;
  const past = ch.filter(p => p.past), sorted = ch.filter(p => p.resorted), renum = ch.filter(p => p.renum && !p.resorted);
  if (past.length) h += `<div class="sub" style="margin-top:6px">Past kept visits with no row on the tab, added with their date (enter the pulses): ${past.map(p => `${esc(p.name)} ${p.past}`).join(", ")}.</div>`;
  if (sorted.length) h += `<div class="sub" style="margin-top:6px">Visit log re-sorted by date, with Tx # renumbered: ${sorted.map(p => esc(p.name)).join(", ")}.</div>`;
  if (renum.length) h += `<div class="sub" style="margin-top:6px">Tx # renumbered: ${renum.map(p => esc(p.name)).join(", ")}.</div>`;
  const nf = full.reduce((n, p) => n + p.none, 0);
  if (nf) h += `<div class="item red" style="margin-top:6px">No row for ${nf} booked visit${nf === 1 ? "" : "s"} on ${full.map(p => esc(p.name)).join(", ")}: the PR formula in A10 is missing. Run Match New Template on the banner, then import again.</div>`;
  const um = ch.filter(p => p.unmissed), dup = plans.reduce((n, p) => n + (p.dup || 0), 0), disc = state.plan.disc || [];
  if (um.length) h += `<div class="sub" style="margin-top:6px">Missed visits undone, since NextGen now shows them booked or kept: ${um.map(p => `${esc(p.name)} ${p.unmissed}`).join(", ")}.</div>`;
  if (dup) h += `<div class="sub" style="margin-top:6px">${dup} visit${dup === 1 ? " is" : "s are"} listed twice on the report (same day, type and time) and ${dup === 1 ? "gets" : "get"} one row.</div>`;
  if (disc.length) h += `<div class="item amb" style="margin-top:6px">Discontinued on the tracker but still booked in NextGen, so not updated: ${disc.map(esc).join(", ")}. Cancel the visits in NextGen, or start a new course tab.</div>`;
  const pv = ch.filter(p => p.e2).length;
  if (pv) h += `<div class="sub" style="margin-top:6px">Provider (E2) set from the report on ${pv} tab${pv > 1 ? "s" : ""}.</div>`;
  const cs = state.clinicSplit, ng = (state.guests || []).length, na = (state.asks || []).length;
  if (cs && !ps) h += cs.on ? `<div class="sub" style="margin-top:6px">${esc(cs.clinic)} chairs only: ${cs.none} patient${cs.none === 1 ? "" : "s"} with no chair visit at ${esc(cs.clinic)} left out${ng ? `, and ${ng} guest${ng === 1 ? "" : "s"} whose home clinic carries the tab listed on the ${GUEST_SHEET} sheet` : ""}${na ? `. ${na} with the MT elsewhere but ${TmsCore.ASK_MIN}+ dailies here: see Create a tracker here? at the top` : ""}.</div>`
    : `<div class="item amb" style="margin-top:6px">${cs.clinic ? `No chair on this report names ${esc(cs.clinic)}` : "The clinic isn't known (set it in Settings)"}, so every patient on it was kept${ng ? ` (${ng} with no tab, no MT and under ${TmsCore.ASK_MIN} treatments listed as guest${ng === 1 ? "" : "s"} on the ${GUEST_SHEET} sheet)` : ""}. Check the clinic in Settings.</div>`;
  if (state.pendingNew.length) h += `<div class="sub" style="margin-top:6px">${state.pendingNew.length} patient${state.pendingNew.length > 1 ? "s" : ""} in the report ${state.pendingNew.length > 1 ? "have" : "has"} no tracker yet. See the top of the dashboard.</div>`;
  h += `<div class="actions"><button class="pri" data-act="apply">Apply</button></div>`;
  $("planBody").innerHTML = h; $("planBox").hidden = false;
}
/* An import plan addresses rows by number and moves them (growLogs, arrangeLogs), so it is only right on the logs it was made from.
   Applying it twice, or after the tab changed (an edit, or the same list applied on another computer), added the rows again and
   moved visits onto the wrong rows. So: one Apply at a time, every tab is checked against its preview before anything is written
   (a changed tab rebuilds the preview instead), and each tab's import is one Edit log entry that can put the log back. */
let applying = false;
async function applyPlan() {
  const plan = state.plan; if (applying || !plan) return;
  const ch = plan.plans.filter(p => p.changed), reflow = new Set(plan.plans.filter(p => p.reflow).map(p => p.name));
  const busy = on => { const b = document.querySelector("#planBody [data-act=apply]"); if (b) { b.disabled = on; b.textContent = on ? "Applying" : "Apply"; } };
  applying = true; busy(true);
  let held = null, stopped = "";
  const done = ch.map(p => ({ tab: p.name, what: importWhat(p), ins: p.ins.map(x => x.row) }));
  try {
    failStep = "Checking the tabs against the preview";
    held = await importStart(ch.map(p => ({ tab: p.name, sig: p.sig, cells: ["E2", "D7", p.fb || "N6"] })));
    if (held.stale.length) {
      const tabs = held.stale; held = null; failStep = "";
      await buildPlan();   // refreshes, then shows the preview made from the tabs as they are now
      toast(`${andList(tabs)} changed after the preview was made (an edit, or this list applied on another computer). Nothing was written. The preview is rebuilt: check it, then select Apply again.`, true, true);
      return;
    }
    failStep = "Adding visit-log rows";
    await growLogs(ch.filter(p => p.ins.length).map(p => ({ tab: p.name, ins: p.ins })));   // rows first: the ops address the grown log
    failStep = "Writing visits";
    await applyOps(ch.flatMap(p => [...p.ops, ...(p.e2 ? [["E2", p.e2]] : [])].map(([a, v]) => ({ tab: p.name, a, v: [[v]] }))), null, false);
    failStep = "Putting the log in date order";
    await arrangeLogs(ch.filter(p => p.arrange).map(p => ({ tab: p.name, ...p.arrange })));   // it moves the rows the ops just wrote
    failStep = "Formatting the visit logs";
    await formatLogs(ch.map(p => ({ tab: p.name, touched: p.touched })));                       // last: formatting follows the dates where they ended up
    const f = state.file;
    if (!plan.ps && !f.saved) { const rep = { asOf: f.asOf, asOfMin: f.asOfMin, importedOn: todaySerial() }; state.report = rep; await saveSetting("tms_report", rep); failStep = `Saving the ${LIST_SHEET} tab`; await saveList(f); }
    state.plan = null; $("planBox").hidden = true; $("file").value = ""; state.file = plan.ps ? null : { ...f, saved: true };
    await refresh();
    const sug = state.patients.filter(p => p.d.suggest);
    failStep = "Updating statuses and measures";
    const sugIn = sug.filter(p => done.some(d => d.tab === p.tab)), sugOut = sug.filter(p => !sugIn.includes(p));
    if (sugIn.length) await applyOps(sugIn.map(p => ({ tab: p.tab, a: "D7", v: [[p.d.suggest]] })), null, false);   // in that tab's import entry
    for (const p of sugOut) await applyOps([{ tab: p.tab, a: "D7", v: [[p.d.suggest]] }], null, { kind: "auto", what: `Status set to ${p.d.suggest} after the import` });
    const nx = reflow.size ? await placeMeasures(state.patients.filter(p => reflow.has(p.tab)), false, false) : 0;
    sug.forEach(p => { const d = done.find(x => x.tab === p.tab); if (d) d.what += `, status set to ${p.d.suggest}`; });
    if (nx) done.forEach(d => { if (reflow.has(d.tab)) d.what += ", measures re-planned"; });
    failStep = "Saving the Edit log"; const id = held.id; held = null; await importEnd(id, done, "");
    toast(`Updated ${ch.length} tab${ch.length === 1 ? "" : "s"}${sug.length ? `, ${sug.length} status${sug.length > 1 ? "es" : ""}` : ""}${nx ? `, measures re-planned for ${reflow.size}` : ""}. Each tab's import is in the Edit log on Patients.`);
    await refresh();
    failStep = ""; show("today");
  } catch (e) { stopped = failStep; fail(e); }
  finally {
    /* a failed import is still logged, so whatever it wrote can be put back from the Edit log */
    if (held && held.id) await importEnd(held.id, done, stopped || "Import").then(() => refresh()).catch(e => console.error(e));
    applying = false; busy(false);
  }
}

/* ---------- patient schedule: opens in a dialog window (schedule.html) to print or save as PDF ---------- */
const pad2 = n => String(n).padStart(2, "0");
const apptTypeName = r => r.type === "MT" ? (device() === "Brainsway" ? "MT BW" : "MT MagV") : r.type === "MTR" ? "MT Redo" : r.type === "F/U" ? (r.tele ? "Telepsych TMS F/U" : "TMS Follow Up") : "TMS Treatment";
function scheduleRows(p, from, to) {
  const t = todaySerial(), loc = clinic(), lo = Math.max(from ?? t, t), hi = to ?? Infinity;
  return p.rows.filter(r => r.date != null && r.date >= lo && r.date <= hi && !r.delivered && !r.miss).sort(cmp)
    .map(r => [p.mrnRaw, r.alt ? (p.home || "Other clinic") : loc, r.date, r.time == null ? "" : r.time / 1440, chairMin(r.type) && !provMin(r.type) ? `${loc} ${chairName()}` : p.provider, apptTypeName(r), "Expected"]);
}
function asOfText() {
  const rep = state.report, d = new Date();
  return rep && rep.asOf != null ? `${fmtDateY(rep.asOf).replace(/^(\d)\//, "0$1/").replace(/\/(\d)\//, "/0$1/")}  ${fmtTime(rep.asOfMin ?? 0)}` : `${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}/${d.getFullYear()}  ${fmtTime(d.getHours() * 60 + d.getMinutes())}`;
}
function printModal(tab, def = "14") {
  const p = state.patients.find(x => x.tab === tab); if (!p) return;
  const t = todaySerial(), lastVisit = Math.max(t, ...p.rows.filter(r => r.date != null && !r.delivered && !r.miss).map(r => r.date));
  const quick = [["14", "Next 14 days", t + 14], ["30", "Next 30 days", t + 30], ["all", "Rest of course", lastVisit]], sel = quick.find(q => q[0] === def) || quick[0];
  showModal(`<h2 id="mtitle">Schedule for ${esc(p.tab)}</h2><p class="sub">Choose the visits to include. The schedule opens over the add-in, to print or save as a PDF.</p>
    <div class="chips" role="group" aria-label="Range" style="margin:14px 0 0">${quick.map(q => `<button class="chip" data-prq="${q[2]}" aria-pressed="${q === sel}">${q[1]}</button>`).join("")}</div>
    <div class="g2">${fld("From", `<input type="date" id="prFrom" value="${serialToInput(t)}">`)}${fld("To", `<input type="date" id="prTo" value="${serialToInput(sel[2])}">`)}</div>
    <p class="body" id="prCount"></p>
    <div class="mfoot"><button class="btn mute" data-mclose>Cancel</button><button class="btn primary" data-act="doPrint" data-t="${esc(tab)}">Open schedule</button></div>`);
  prCount(tab);
}
function prRange() { return [inputToSerial($("prFrom").value), inputToSerial($("prTo").value)]; }
function prCount(tab) { const p = state.patients.find(x => x.tab === tab), [f, to] = prRange(), n = p && f && to ? scheduleRows(p, f, to).length : 0; $("prCount").textContent = `${n} visit${n === 1 ? "" : "s"} in this range.`; }
async function logSchedule(p) {
  const t = todaySerial(), r = p.rows.find(x => x.date === t) || p.d.sched[0], tagText = `Schedule provided ${fmtDate(t)}`;
  const pairs = [["W8", t]];
  if (r) pairs.push([`S${r.row}`, [...r.note.split("; ").filter(x => x && !/^Schedule provided /.test(x)), tagText].join("; ")]);
  await writeCells(p.tab, pairs);
}
function openSchedule(tab, from, to) {
  const p = state.patients.find(x => x.tab === tab); if (!p) return;
  const body = scheduleRows(p, from, to); if (!body.length) return toast("No scheduled visits in that date range.", true);
  const d = serialToDate(todaySerial());
  const obj = { fname: `${p.code} schedule ${d.getUTCMonth() + 1}.${d.getUTCDate()}.pdf`, asOf: `As of: ${asOfText()}`,
    rows: body.map(r => [r[0], r[1], fmtDateY(r[2]), r[3] === "" ? "" : fmtTime(Math.round(r[3] * 1440)), r[4], r[5], r[6]]) };
  /* never a separate window: the schedule floats over the add-in on a blurred backdrop */
  openSchedOverlay(p, obj);
}

let schedFor = null;
function openSchedOverlay(p, obj) {
  schedFor = { p, logged: false };
  const f = $("schedFrame"), go = () => { try { f.contentWindow.render(obj); } catch (e) { fail(e); } };
  $("schedOverlay").hidden = false;
  if (f.dataset.ready) go(); else f.onload = () => { f.dataset.ready = "1"; go(); };
  if (!f.getAttribute("src")) f.src = "schedule.html?embed=1";
}
function closeSchedOverlay() { $("schedOverlay").hidden = true; schedFor = null; }
async function onSchedMessage(m) {
  if (m === "close") return closeSchedOverlay();
  if ((m === "printed" || m === "saved") && schedFor && !schedFor.logged) { schedFor.logged = true; try { await logSchedule(schedFor.p); await refresh(); } catch (e) { fail(e); } }
}

/* ---------- wiring ---------- */
const VIEWS = ["today", "schedule", "patients", "chairs", "imp", "settings", "add"];
function show(view) {
  if (view !== "patients" && state.ptOpen) { state.ptOpen = null; $("ptView").hidden = true; $("pvBody").innerHTML = ""; }
  state.view = view;
  VIEWS.forEach(v => $("v-" + v).hidden = v !== view);
  $("page").scrollTop = 0;
  renderDash();
}
/* ---------- pop-out window: full screen or a separate window, driven by the task pane ---------- */
const CHUNK = 20000;
function toChild(m) { if (pop.dlg) try { pop.dlg.messageChild(JSON.stringify(m)); } catch { /* window closed */ } }
function sendSnap() {
  const s = JSON.stringify({ cfg: state.cfg, patients: state.patients.map(({ d, ...p }) => p), tabs: state.tabs, protocols: state.protocols, fix: [...state.fix], old: state.old, stale: state.stale,
    skipped: state.skipped, templateOk: state.templateOk, templateWhy: state.templateWhy, pendingNew: state.pendingNew, guests: state.guests || null, asks: state.asks || null, noAsk: state.noAsk || null, sheetNames: state.sheetNames || [], codeMrn: state.codeMrn ? [...state.codeMrn] : null, closures: [...state.closures], wbName: state.wbName, clinic: state.clinic, report: state.report, lx: state.lx,
    edits: state.edits, apptBy: [...state.apptBy], apptTo: state.apptTo, apptLo: state.apptLo, apptAll: state.apptAll ? [...state.apptAll] : null, sess: sessSnap(), ui: uiSnap() });
  sendParts("snap", s);
}
/* dialog messages are size-limited, so anything big goes in CHUNK-sized parts the window puts back together */
function sendParts(k, s) { const n = Math.ceil(s.length / CHUNK) || 1, seq = Date.now() + Math.random(); for (let i = 0; i < n; i++) toChild({ k, seq, i, n, part: s.slice(i * CHUNK, (i + 1) * CHUNK) }); }
const sendSess = () => sendParts("sess", JSON.stringify(sessSnap()));
function applySnap(o) {
  Object.assign(state, { cfg: o.cfg, patients: o.patients, tabs: o.tabs, protocols: o.protocols, fix: new Map(o.fix), old: o.old, stale: o.stale || [], skipped: o.skipped, templateOk: o.templateOk, templateWhy: o.templateWhy || "",
    pendingNew: o.pendingNew, guests: o.guests || null, asks: o.asks || null, noAsk: o.noAsk || null, sheetNames: o.sheetNames || [], codeMrn: o.codeMrn ? new Map(o.codeMrn) : null, closures: new Map(o.closures), wbName: o.wbName, clinic: o.clinic, report: o.report, lx: o.lx || null,
    edits: o.edits || [], apptBy: new Map(o.apptBy || []), apptTo: o.apptTo ?? null, apptLo: o.apptLo ?? null, apptAll: o.apptAll ? new Map(o.apptAll) : null, sess: null }); setSessSnap(o.sess); mergeUi(o.ui);
  state.patients.forEach(p => p.d = derive(p));
  syncCfgInputs(); renderBanners(); renderDash(); updateApTab(); fillAddPatientForm();
  rpc.snapWait.splice(0).forEach(f => f());
}
const CALL_MS = 180000;
function callPane(fn, args) {
  return new Promise((res, rej) => {
    const id = ++rpc.id, gone = setTimeout(() => { if (rpc.wait.delete(id)) rej(new Error("Full screen lost touch with the task pane. Keep the pane open, then check the tab before trying again.")); }, CALL_MS);
    rpc.wait.set(id, { res: v => { clearTimeout(gone); res(v); }, rej: e => { clearTimeout(gone); rej(e); } });
    try { Office.context.ui.messageParent(JSON.stringify({ k: "call", id, fn, args })); } catch (e) { clearTimeout(gone); rpc.wait.delete(id); rej(e); }
  });
}
/* what a pop-out may ask the pane to do: reads and writes need the workbook, settings need the document */
const PANE_FNS = {
  refresh: () => refresh(), applyOps: (ops, g, log) => applyOps(ops, g, log), importStart: l => importStart(l), importEnd: (id, d, s) => importEnd(id, d, s), undoEdit: id => undoEdit(id), fixCheck: (t, k, r) => fixCheck(t, k, r), growLogs: x => growLogs(x), arrangeLogs: x => arrangeLogs(x), formatLogs: x => formatLogs(x), writeList: x => writeList(x), rebuildActive: () => rebuildActiveClick(), mkSchedSheet: () => mkSchedSheet(),
  writeChairSheet: o => writeChairSheet(o), setupColumns: force => setupColumns(force), addPatient: d => addPatient(d), openTabName: (t, r) => openTabName(t, r),
  saveSetting: (k, v) => saveSetting(k, v), fixTotals: (t, log) => fixTotals(t, log), saveGuestLink: (m, u) => saveGuestLink(m, u), useTemplateTab: n => useTemplateTab(n),
  writeSess: async e => { try { await writeSessPane(e); } finally { sendSess(); } return null; },
  sessSnap: async () => { await Excel.run(readSessNow); sendSess(); return null; }
};
async function onPopMessage(arg) {
  let m; try { m = JSON.parse(arg.message); } catch { return; }
  if (m.k === "ready") return sendSnap();
  if (m.k === "close") { closePopout(); return; }
  if (m.k === "ui") { mergeUi(m); renderBanners(); return; }
  if (window.DevUpdates && DevUpdates.receive(m)) return;
  if (m.k !== "call" || !PANE_FNS[m.fn]) return;
  pop.calls++; let ok = true, err = "", result = null;
  try { result = await PANE_FNS[m.fn](...(m.args || [])); } catch (e) { ok = false; err = e.message || String(e); }
  pop.calls--;
  toChild({ k: "ack", id: m.id, ok, err, result: result === undefined ? null : result });
}
/* Full screen floats over the Excel page (an in-page dialog, displayInIframe), never a separate browser window */
function openPopout() {
  if (pop.dlg) return toast("Full screen is already open.");
  if (!Office.context.requirements.isSetSupported("DialogApi", "1.2")) return toast("This version of Excel can't open full screen.", true);
  const url = new URL(location.href); url.search = "?popout=1&full=1"; url.hash = "";
  Office.context.ui.displayDialogAsync(window.DevUpdates ? DevUpdates.dialogUrl(url.href) : url.href, { height: 95, width: 95, displayInIframe: true }, res => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) return toast(res.error.code === 12007 ? "Close the other add-in window first." : `Full screen couldn't open (${res.error.message}). Try again.`, true);
    pop.dlg = res.value; $("popNote").hidden = false; if (window.DevUpdates) DevUpdates.setSend(toChild);
    pop.dlg.addEventHandler(Office.EventType.DialogMessageReceived, onPopMessage);
    pop.dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { pop.dlg = null; $("popNote").hidden = true; if (window.DevUpdates) DevUpdates.setSend(null); });
  });
}
function closePopout() { if (pop.dlg) { try { pop.dlg.close(); } catch { /* already closed */ } pop.dlg = null; } $("popNote").hidden = true; if (window.DevUpdates) DevUpdates.setSend(null); }
/* the window side: snapshot and Sessions parts, answers to its calls, and toasts from the pane */
function onPaneMessage(a) {
    let m; try { m = JSON.parse(a.message); } catch { return; }
    if (m.k === "snap" || m.k === "sess") {
      const b = rpc.bufs[m.k] && rpc.bufs[m.k].seq === m.seq ? rpc.bufs[m.k] : (rpc.bufs[m.k] = { seq: m.seq, parts: Array(m.n) });
      b.parts[m.i] = m.part;
      if (b.parts.filter(x => x != null).length === m.n) { const o = JSON.parse(b.parts.join("")); delete rpc.bufs[m.k]; if (m.k === "snap") applySnap(o); else { setSessSnap(o); renderDash(); } }
    } else if (m.k === "ack") {
      const w = rpc.wait.get(m.id); if (!w) return; rpc.wait.delete(m.id);
      if (m.ok) w.res(m.result); else w.rej(new Error(m.err || "The task pane could not do that."));
    } else if (m.k === "toast") toast(m.msg, m.err, m.sticky);
    else if (m.k === "ui") { mergeUi(m); renderBanners(); }
    else if (window.DevUpdates) DevUpdates.receive(m);
}
function initPopout() {
  document.body.classList.add("popout");
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onPaneMessage, () => Office.context.ui.messageParent(JSON.stringify({ k: "ready" })));
}
/* Full screen: the tracker floats inside the Excel page over the workbook; the menu sits on frosted glass over a blurred page (body.fs).
   In full screen the same button closes it. */
function fullScreen() {
  if (POPOUT) { try { Office.context.ui.messageParent(JSON.stringify({ k: "close" })); } catch { /* pane gone */ } return; }
  openPopout();
}
const FULL = POPOUT;
function syncFs() {
  document.body.classList.toggle("fs", FULL);
  $("btnFull").textContent = FULL ? "Exit full screen" : "Full screen";
  $("conn").lastChild.textContent = FULL ? "Saving through the task pane" : "Connected to the workbook";
}
/* settings live in the document, which only the task pane can save */
function saveSetting(k, v) {
  if (POPOUT) return callPane("saveSetting", [k, v]);
  const s = Office.context.document.settings; s.set(k, v); s.saveAsync();
  if (k === "tms_cfg") { state.cfg = v; state.patients.forEach(p => p.d = derive(p)); renderBanners(); renderDash(); syncCfgInputs(); }
  if (k === "tms_report") state.report = v;
}
async function openTabName(tab, row) {
  if (POPOUT) return callPane("openTabName", [tab, row]);
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(tab); ws.activate();
      if (row) ws.getRange(`A${row}:T${row}`).select();
      await ctx.sync();
    });
  } catch (e) { fail(e); }
}
/* a course on Patients; it switches to All active when the current filter hides it */
function openCourse(tab) {
  state.ptTab = tab; const p = state.patients.find(x => x.tab === tab), fl = state.fl || floor(), f = FILTERS.find(x => x[0] === state.cfilter) || FILTERS[0];
  if (p && !f[2](p, courseNeeds(p, fl))) state.cfilter = onSched(p) ? (p.d.active ? "active" : FILTERS.find(x => x[2](p, []))?.[0] || "active") : "disc";
  show("patients"); if (p) openPt(tab);
}
/* appearance is per computer (localStorage), like the Today view; the pane and full screen share it */
const themeOf = () => document.documentElement.dataset.theme || "light";
function syncTheme() {
  const t = themeOf(), dark = t === "dark" || (t === "system" && matchMedia("(prefers-color-scheme: dark)").matches), b = $("btnTheme"), c = $("cfgTheme");
  if (b) b.setAttribute("aria-pressed", String(dark));
  if (c) c.value = t;
}
function setTheme(t) {
  document.documentElement.dataset.theme = t;
  try { localStorage.setItem("tms_theme", t); } catch { /* this session only */ }
  syncTheme();
}
const CFG_IDS = { cfgLoc: "location", cfgDevice: "device", cfgAuth: "authDays", cfgProto: "protoDays", cfgRecent: "recentDays", cfgOpenTab: "openTab", cfgCutoff: "cutoff" };
function syncCfgInputs() { Object.entries(CFG_IDS).forEach(([id, k]) => { if ($(id) && document.activeElement !== $(id)) $(id).value = state.cfg[k]; }); }
const benData = el => Object.fromEntries([...el.closest("details").querySelectorAll("[data-benf],[data-cost]")].map(i => [i.dataset.benf || i.dataset.cost, i.value]));
/* re-renders replace the markup, so focus goes back to the same control (by its data attributes), else to the card's next step */
const FKEYS = ["view", "mode", "k", "chair", "course", "filter", "sv", "open", "act", "t", "row", "v"];
let focusWant = null;
function focusKey(a) {
  if (!a || a === document.body || !a.dataset || a.closest("#modal")) return null;
  return FKEYS.filter(x => a.dataset[x] != null).map(x => `[data-${x.replace(/[A-Z]/g, c => "-" + c.toLowerCase())}="${CSS.escape(a.dataset[x])}"]`).join("") || null;
}
function refocus() {
  const a = document.activeElement;
  if (!focusWant || !$("modal").hidden || (a && a !== document.body && document.contains(a))) return;
  const el = (focusWant && document.querySelector(`#app ${focusWant}`)) || document.querySelector(`#v-${state.view} .detail .btn.primary`);
  if (el && el.offsetParent) { el.focus(); focusWant = null; }
}
const skipWeekend = (k, step) => { do { state[k] += step; } while ([0, 6].includes(dow(todaySerial() + state[k]))); renderDash(); };
const benOk = data => { for (const [a, raw] of Object.entries(data)) { const v = benVal(raw); if (v === null || (v === "N/A" && !/^W[45]$/.test(a))) return false; } return true; };

Office.onReady(info => {
  if (!POPOUT && info.host !== Office.HostType.Excel) return;
  if (!POPOUT) { const saved = Office.context.document.settings.get("tms_cfg"); if (saved) Object.assign(state.cfg, saved); }
  syncCfgInputs();
  Object.entries(CFG_IDS).forEach(([id, k]) => { $(id).onchange = () => {
    const v = $(id).value; state.cfg[k] = k === "cutoff" ? (v || CUTOFF_DEFAULT) : ["location", "device", "openTab"].includes(k) ? (v.trim() || { location: "", device: "auto", openTab: "on" }[k]) : (+v || { authDays: 14, protoDays: 2, recentDays: 10 }[k]);
    saveSetting("tms_cfg", { ...state.cfg }); state.patients.forEach(p => p.d = derive(p)); renderBanners(); renderDash(); if (k === "cutoff") refresh(); }; });
  $("apBack").onclick = () => show("patients");
  $("btnFull").onclick = fullScreen;
  $("popClose").onclick = closePopout;
  syncFs();
  window.addEventListener("message", e => { if (e.origin === location.origin && e.data && e.data.sched) onSchedMessage(e.data.sched); });
  /* the clock moves the board: every 30 seconds, unless someone is typing or a dialog is open */
  setInterval(() => {
    const a = document.activeElement;
    if (document.hidden || !$("modal").hidden || (a && a.matches && a.matches("input,select,textarea")) || !["today", "schedule"].includes(state.view)) return;
    focusWant = focusKey(a); pullSess().catch(e => console.error(e)).then(() => { renderDash(); refocus(); });
  }, 30000);
  document.addEventListener("click", e => {
    const ctl = e.target.closest("button,input,select"); if (ctl) focusWant = focusKey(ctl);
    onClick(e);
    if (ctl) refocus();
  });
  function onClick(e) {
    const el = e.target;
    const bn = el.closest("[data-banner]"); if (bn) { if (bn.dataset.bact === "snooze") snooze(bn.dataset.banner); else dismiss(bn.dataset.banner); return; }
    if (el.closest("[data-mclose]") || el.id === "modal") return closeModal();
    const mf = el.closest("[data-mf]"); if (mf) return setFreebie(mf.dataset.t, +mf.dataset.row, mf.dataset.mf);
    const pq = el.closest("[data-prq]"); if (pq) { $("prFrom").value = serialToInput(todaySerial()); $("prTo").value = serialToInput(+pq.dataset.prq); document.querySelectorAll("[data-prq]").forEach(b => b.setAttribute("aria-pressed", String(b === pq))); return prCount($("modalBody").querySelector("[data-act=doPrint]").dataset.t); }
    const ap = el.closest("[data-apact]"); if (ap) { document.querySelectorAll("[data-apr]").forEach(i => { state.apRanges[+i.dataset.apr][i.dataset.k] = i.value; }); if (ap.dataset.apact === "add") state.apRanges.push({}); else state.apRanges.splice(+ap.dataset.v, 1); renderApRanges(); return; }
    const x = el.closest("[data-act],[data-view],[data-mode],button[data-k],[data-chair],[data-course],[data-filter],[data-sv],[data-open],[data-sheet],[data-pt],[data-ptshow]"); if (!x) return;
    const d = x.dataset;
    if (d.act) return act(d.act, x, e);
    if (d.view) { if (d.view === "patients" && state.view === "add") resetAddPatient(); return show(d.view); }
    if (d.mode) { state.chair = d.mode; state.sel = null; try { localStorage.setItem("tms_chair", d.mode); } catch { /* this session only */ } return renderDash(); }
    /* selecting a visit also opens its tab at that row, unless Settings turns it off */
    if (d.k) { state.sel = { k: d.k }; const v = findVisit(d.k); if (v && state.chair !== "all" && v.chair !== state.chair) state.chair = "all"; if (v && state.cfg.openTab !== "off") openTabName(v.tab, v.row); if (d.go) return show("today"); return renderDash(); }
    if (d.chair) { state.sel = { chair: d.chair }; return renderDash(); }
    if (d.course) return openPt(d.course);
    if (d.ptshow) return setPtShow(d.ptshow);
    if (d.filter) { state.cfilter = d.filter; return renderDash(); }
    if (d.sv) { state.dash = d.sv; return renderDash(); }
    if (d.open) { const v = (state.fl || floor()).visits.find(y => y.tab === d.open && ["running", "waiting", "booked"].includes(ph(y))) || state.fl.visits.find(y => y.tab === d.open); if (v) { state.sel = { k: v.key }; state.chair = "all"; } return show("today"); }
    if (d.sheet) return openTabName(d.sheet, null);
    if (d.pt) return openCourse(d.pt);
  }
  function act(a, el, e) {
    const t = el.dataset.t, v = el.dataset.v, row = +el.dataset.row || null, k = el.dataset.k, V = k ? findVisit(k) : null, m = state.mctx || {};
    const needV = fn => () => { if (!V) return toast("That visit isn't on today's list anymore. Refresh to see the latest.", true); return fn(V); };
    const acts = {
      /* import, banners, Match New Template, Clinic Schedule */
      /* the Edit log on Patients: undo (whole-log changes ask first), and the checks' fixes (removing a row asks first) */
      elUndo: () => elUndoAsk(v), elUndoGo: () => { const id = m.id; closeModal(); return undoEdit(id); },
      elFix: () => v === "remove" ? elRemoveAsk(t, row) : fixCheck(t, v, row), elFixGo: () => { const { t: tt, row: rr } = m; closeModal(); return fixCheck(tt, "remove", rr); },
      apply: () => applyPlan(), previewSaved: () => buildPlan(), setupGo: () => { closeModal(); setupColumns(); }, newTracker: () => startTracker(v),
      dontAsk: () => dontAsk(v), doTplUse, verifyGo, verifyKeep, useTab: () => useTabAsk(v), useTabGo,
      guestOpen: () => { const g = guestOf(v); if (g && g.link) window.open(g.link, "_blank", "noopener"); },
      saveGuest: () => saveGuestLink(v, (document.querySelector(`[data-guestin="${CSS.escape(v)}"]`) || {}).value),
      guests: () => { state.dismissed.delete("guestBanner"); delete state.snoozeMem.guestBanner; renderBanners(); const b = $("guestBanner"); b.hidden = false; b.scrollIntoView({ block: "start" }); },
      saveMrn: () => { const x = String((document.querySelector(`[data-mrnin="${CSS.escape(t)}"]`) || {}).value || "").trim();
        if (!/^\d[\d-]{2,}$/.test(x)) return toast("Enter the MRN from NextGen (digits).", true);
        const dup = state.patients.find(p => p.tab !== t && p.mrn === normMrn(x) && p.code !== (state.patients.find(q => q.tab === t) || {}).code);
        if (dup) return toast(`That MRN is already on ${dup.tab}. Check NextGen.`, true);
        return writeCells(t, [["A4", x]]).then(() => { toast(`${t}: MRN saved to A4.`); return refresh(); }).catch(fail); },
      reimport: () => { show("imp"); buildPlan(); }, goimport: () => show("imp"),
      mkSched: () => mkSchedSheet(), schedSheet: () => writeChairSheet(state.dash === "week" ? 7 * state.weekOffset : state.schedOffset), addPt: () => { resetAddPatient(); show("add"); },
      /* today's visits */
      start: needV(startSession), end: needV(mEnd), mt: needV(mMt), miss: needV(mMiss), move: needV(mMove), note: needV(mNote), fumeas: needV(mFuMeas),
      fudone: needV(x => sessWrite(x, { end: nowMin(), ...(x.checkin == null ? { checkin: nowMin() } : {}) }, `${x.tab}'s follow-up marked done.`)), undoEnd: needV(undoEnd), ngdone: () => setNgTick(k, true),
      classify: () => freebieModal(t, row), freebie: () => freebieModal(t, row), undoMiss: () => undoMiss(t, row), print: () => printModal(t, v || "14"), opentab: () => openCourse(t), course: () => openCourse(t),
      /* Patients */
      status: () => setStatus(t, v), statuspick: () => mStatus(t), call: () => mCall(t), ben: () => mBen(t), ext: () => mExt(t), proto: () => mProtoEdit(t), avail: () => mAvail(t),
      ptClose: () => closePt(), ptPast: () => { state.ptPast = !state.ptPast; renderPt(); }, ptNeeds: () => { state.ptNeedsAll = !state.ptNeedsAll; renderPt(); },
      ptJump: () => ptJump(row), ptNote: () => mPtNote(t, row, v), doPtNote: () => { const s = mval("mPtNote"); closeModal(); savePtNote(m.t, m.row, m.kind, s); },
      ptSchedView: () => { const p = state.patients.find(y => y.tab === t), nx = p && p.d.next; closePt(); state.dash = "day"; show("schedule"); if (nx) toast(`${t}'s next visit: ${fmtDay(nx.date)}, ${clock(nx.time) || "no time"}.`); },
      copayToday: () => copayToday(t, row, v), copayMark: () => copayMark(t, v),
      pvProtoSave: () => { const name = ($("pvProto") || {}).value, site = ($("pvSite") || {}).value; if (!name) return toast("Select the protocol to confirm.", true); saveProtocol(t, [{ name, site: TmsCore.SITE_CELL[site] || "Left" }]); },
      pastpulses: () => mPastPulses(t, row), editPulses: () => mEditPulses(t, row), doEditPulses, copayChange: () => { state.copayEdit = `${t}|${row}`; renderPt(); },
      finishYes: () => { const { t: tab, k: key } = m; closeModal(); setStatus(tab, "Complete").then(() => { const nv = key && findVisit(key); if (nv && !state.copied.has(nv.key)) mNote(nv); }); },
      finishNo: () => { const nv = m.k && findVisit(m.k); closeModal(); if (nv && !state.copied.has(nv.key)) mNote(nv); }, pasttech: () => mPastTech(t, row), pastmeas: () => mPastMeas(t, row),
      /* Chairs */
      chairOff: () => mChairOff(v), chairEdit: () => mChairEdit(v),
      chairOn: async () => { try { await writeSess([{ chairId: v, set: { off: false } }]); renderDash(); toast(`${chairLbl(v)} is available again.`); } catch (err) { fail(err); } },
      /* Schedule */
      prev: () => { state.weekOffset--; renderDash(); }, next: () => { state.weekOffset++; renderDash(); }, this: () => { state.weekOffset = 0; renderDash(); },
      sprev: () => skipWeekend("schedOffset", -1), snext: () => skipWeekend("schedOffset", 1), sthis: () => { state.schedOffset = 0; renderDash(); },
      /* dialog saves */
      doEnd, doProtoEdit, doMt, doMiss, doMove, doCopyNote, doSaveNote, resetNote, doFuMeas,
      doPrint: () => { const [f, to] = prRange(); if (!f || !to || to < f) return merr("Pick a valid date range."); closeModal(); openSchedule(t, f, to); },
      doPastPulses: () => { const vals = readPulses(), tech = mval("mTech"); if (!vals.some(n => n !== "" && n > 0)) return merr("Enter the pulses delivered."); if (!tech) return merr("Enter the treating technician."); closeModal(); markPulses(m.t, m.row, { vals, tech }).then(ok => ok && toast("Pulses saved.")); },
      pastMissed: () => { const [p, r] = findPR(m.t, m.row); if (!r) return closeModal(); mMiss({ p, r, tab: m.t, row: m.row, time: r.time, key: null }); },
      doPastTech: async () => { const tech = mval("mTech"); if (!tech) return merr("Enter the treating technician."); closeModal(); try { const [, r0] = findPR(m.t, m.row); await writeCells(m.t, [[`B${m.row}`, tech]], r0 && at(r0)); state.techName = tech; patchRow(m.t, m.row, { tech }); toast("Technician saved."); } catch (err) { fail(err); } },
      doPastMeas: () => { const P = mval("mPhq"), Q = mval("mGad"); if (!(isInt(P, 0, 27) && isInt(Q, 0, 21))) return merr("Enter PHQ-9 (0 to 27) and GAD-7 (0 to 21)."); closeModal(); saveMeasures(m.t, m.row, { P: +P, Q: +Q }); },
      doMoveMeas: () => { const to = +mval("mMeasTo") || m.dest; closeModal(); moveMeasures(m.t, m.row, to); },
      doBen: () => { const data = Object.fromEntries([...$("modalBody").querySelectorAll("[data-benf],[data-cost]")].map(i => [i.dataset.benf || i.dataset.cost, i.value])); if (!benOk(data)) return merr("Use amounts like 25 or 25.00. Only the deductible and out-of-pocket max can be N/A."); closeModal(); saveBenefits(m.t, data); },
      doExt: () => { const n = mval("mExtN"); if (!isInt(n, 1, 99)) return merr("Add 1 or more visits."); const exp = inputToSerial(mval("mExtD")); if (exp != null && exp <= todaySerial()) return merr("Choose an end date after today, or leave it blank."); closeModal(); addExtension(m.t, { n: +n, exp }); },
      doAvail: () => { const av = state.avDraft; if (!av || av.tab !== m.t) return closeModal(); readAvailEditor($("modalBody"), av); closeModal(); saveAvailability(m.t, av); },
      avAdd: () => { readAvailEditor($("modalBody"), state.avDraft); state.avDraft.ranges.push({ type: "Unavailable", from: null, to: null, note: "" }); mAvail(m.t, true); },
      avDel: () => { readAvailEditor($("modalBody"), state.avDraft); state.avDraft.ranges.splice(+v, 1); mAvail(m.t, true); },
      doStatus: () => { const st = radio("st"); if (!st) return merr("Choose a status."); closeModal(); setStatus(m.t, st); },
      doCall: () => { const out = radio("out"), note = mval("mCallNote"); closeModal(); logCall(m.t, out, note); },
      doChairOff: async () => { const id = m.chair; closeModal(); try { await writeSess([{ chairId: id, set: { off: true } }]); renderDash(); toast(`${chairLbl(id)} marked unavailable.`); } catch (err) { fail(err); } },
      doChairEdit: async () => { const id = m.chair, svc = inputToSerial(mval("mSvc")), set = { device: mval("mDev"), coil: mval("mCoil"), service: svc }; if (svc != null && svc > todaySerial()) return merr("Choose today or an earlier date."); closeModal(); try { await writeSess([{ chairId: id, set }]); renderDash(); toast(`${chairLbl(id)} saved.`); } catch (err) { fail(err); } }
    };
    if (acts[a]) { e.preventDefault(); acts[a](); }
  }
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && !$("schedOverlay").hidden) return closeSchedOverlay();
    if (e.key === "Escape" && !$("modal").hidden) return closeModal();
    if (e.key === "Escape" && !$("ptView").hidden) return closePt();
    if (e.key === "Tab" && $("modal").hidden && !$("ptView").hidden) {   // and inside the patient view
      const f = [...$("pvBody").querySelectorAll("button,input,select,textarea")].filter(x => !x.disabled && x.type !== "hidden" && x.offsetParent);
      if (f.length && e.shiftKey && document.activeElement === f[0]) { e.preventDefault(); f[f.length - 1].focus(); }
      else if (f.length && !e.shiftKey && document.activeElement === f[f.length - 1]) { e.preventDefault(); f[0].focus(); }
    }
    if (e.key === "Tab" && !$("modal").hidden) {   // keep focus inside the dialog
      const f = [...$("modalBody").querySelectorAll("button,input,select,textarea")].filter(x => !x.disabled && !x.hidden && x.type !== "hidden" && x.offsetParent);
      if (!f.length) return;
      if (e.shiftKey && document.activeElement === f[0]) { e.preventDefault(); f[f.length - 1].focus(); }
      else if (!e.shiftKey && document.activeElement === f[f.length - 1]) { e.preventDefault(); f[0].focus(); }
    }
  });
  document.addEventListener("change", e => {
    const el = e.target;
    if (el.dataset.ngk) { focusWant = el.closest("#modal") ? null : focusKey(el); return setNgTick(el.dataset.ngk, el.checked).then(refocus); }
    if (el.id === "mPain") { $("mPainD").hidden = el.value !== "Yes"; return; }
    if (el.dataset.pick) { const n = +el.dataset.pick; state.ptPick = state.ptPick || new Set(); if (el.checked) state.ptPick.add(n); else state.ptPick.delete(n); return; }
    if (el.dataset.pk === "name") return protoEditPick(el);
    if (el.matches("[data-elf]")) { state.elogTab = el.value; return renderDash(); }
    if (el.dataset.dk) draftInput(el);
  });
  document.addEventListener("input", e => { if (e.target.dataset.dk && e.target.tagName !== "SELECT") draftInput(e.target); if (e.target.id === "prFrom" || e.target.id === "prTo") { document.querySelectorAll("[data-prq]").forEach(b => b.setAttribute("aria-pressed", "false")); prCount($("modalBody").querySelector("[data-act=doPrint]").dataset.t); } });
  $("refresh").onclick = () => refresh(); $("btnApRebuild").onclick = () => rebuildActiveClick(); $("btnSetup").onclick = () => setupClick(false); $("btnMeas").onclick = placeMeasuresClick;
  $("btnTheme").onclick = () => setTheme($("btnTheme").getAttribute("aria-pressed") === "true" ? "light" : "dark");
  $("cfgTheme").onchange = e => setTheme(e.target.value);
  window.addEventListener("storage", e => { if (e.key === "tms_theme") { document.documentElement.dataset.theme = e.newValue || "light"; syncTheme(); } });
  matchMedia("(prefers-color-scheme: dark)").addEventListener("change", syncTheme);
  document.addEventListener("toggle", e => { if (e.target.matches && e.target.matches("[data-elog]")) state.elogOpen = e.target.open; }, true);   // the Edit log stays open across redraws
  syncTheme();
  $("btnMatch").onclick = () => setupClick(true); $("btnTplUse").onclick = mTplUse;
  ["file", "benFile"].forEach(i => { const el = $(i); el.ondragenter = () => el.classList.add("over"); el.ondragleave = el.ondrop = () => el.classList.remove("over"); });
  $("file").onchange = onFile; $("benFile").onchange = onBenFile; document.querySelectorAll("input[name=rtype]").forEach(r => r.onchange = () => state.file && !state.file.saved && buildPlan());
  ["apLast", "apFirst"].forEach(i => $(i).oninput = () => { if (!state.codeEdited) $("apCode").value = lasfir($("apLast").value, $("apFirst").value); updateApTab(); });
  $("apCode").oninput = () => { state.codeEdited = true; updateApTab(); };
  $("apCourse").oninput = () => { state.courseEdited = true; updateApTab(); };
  $("btnAdd").onclick = () => addPatient(collectAddPatient());
  $("btnFmt").onclick = formatLogsClick;
  fillAddPatientForm(); show("today");
  /* Dev Update notices; full screen tells the pane what was dismissed (the pane tells full screen once it opens it) */
  if (window.DevUpdates) DevUpdates.init({ app: "tms", send: POPOUT ? m => Office.context.ui.messageParent(JSON.stringify(m)) : null });
  if (POPOUT) initPopout(); else refresh();
});
