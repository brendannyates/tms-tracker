/* TMS Tracker: Excel (web) task-pane add-in. A light augmentation over the organization's patient tracker template.
   One tab per course (LASFIR or "LASFIR 2") cloned from "New Template". Information stays where the template keeps it:
   header A1:H7 (MRN A4, # of txs D2, # of MT redos D4, MD on-site D6, status D7, provider E2, TAS E4, insurance E6, auth F4:G4,
   update protocol by G2, progress note sent G6), protocol J2:O4, freebies left N6, auth assessment N7, visit log A10:T49
   (time in T) closed by the TX SUMMARY row with its checkbox in A, home clinic S9, weekly availability V19:X24 and dates away V28:Y48.
   Benefits live on the Benefits tab, one row per MRN (see BEN_COLS). They keep their old cell names (W1 BIDF, W4 deductible,
   W5 OOP max, W6 MTR approved, W7 extensions, W8 last schedule, X12:Y16 cost by code) as keys, and applyOps sends those there.
   Tabs in the previous layout (freebies in N7, benefits in V1:Y16) are still read; Match New Template moves them over.
   PHI: MRNs are read into memory to match NextGen rows; they are never sent anywhere. Names typed in Add patient become LASFIR. */
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
const DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const AV_TYPE = ["Unavailable", "Available"];
const FREEBIE_TAG = { Used: "Freebie used", Waived: "Freebie waived" };
const SITES = ["Left", "Right", "Other"];
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

const state = {
  cfg: { authDays: 14, protoDays: 2, recentDays: 10, device: "auto", location: "San Rafael", openTab: "on", cutoff: CUTOFF_DEFAULT },
  patients: [], tabs: [], protocols: [], fix: new Map(), skipped: { layout: [], other: 0 }, old: [], stale: [], templateOk: true, pendingNew: [],
  drafts: {}, initials: "", techName: "", apSig: null, closures: new Map(), wbName: "", clinic: null, report: null,
  view: "sch", clinicTab: "chair", ptMode: "list", ptFilter: "all", wideAtt: null, weekOffset: 0, schedOffset: 0, dayOffset: 0, dash: "today", ptTab: null, ptAll: false, file: null, plan: null, codeEdited: false, courseEdited: false,
  open: new Set(), dismissed: new Set(), snoozeMem: {}, avDraft: null, lx: null, chair: "all"
};
try { state.chair = localStorage.getItem("tms_chair") || "all"; } catch { /* storage unavailable */ }
/* Discontinued patients are left off the schedules, Attention and the roster (except its Discontinued filter) */
const onSched = p => !/^discontinued/i.test(String(p.status).trim());
/* chairs: from the saved appointment list's resource names (for example Castro MagV1 and Castro MagV2). A visit's chair is the
   resource booked that day, else the chair the patient's daily treatments are usually on. Each person picks a chair on their own device. */
const chairs = () => state.lx ? state.lx.chairs : [];
const chairOf = (p, r) => { const lx = state.lx; if (!lx || !p.mrn) return ""; return (r && lx.v[`${p.mrn}|${r.date}`]) || lx.pt[p.mrn] || ""; };
const chairPick = () => chairs().length > 1 && chairs().includes(state.chair) ? state.chair : "all";
const inChair = (p, r) => { const c = chairPick(); if (c === "all") return true; const x = chairOf(p, r); return !x || x === c; };
function chairBar() {
  const cs = chairs(); if (cs.length < 2) return "";
  const cur = chairPick(), opts = [["all", cs.length === 2 ? "Both" : "All chairs"], ...cs.map(c => [c, c])];
  return `<div class="fchips" role="group" aria-label="Chair" style="margin-bottom:8px">${opts.map(([k, l]) => `<button class="sm${k === cur ? " on" : ""}" data-act="chair" data-v="${esc(k)}" aria-pressed="${k === cur}">${esc(l)}</button>`).join("")}</div>`;
}
/* more than one patient in the same chair at once; with no chair info everyone counts as one chair, as before */
const chairClash = occ => { const m = new Map(); occ.forEach(e => { const k = chairOf(e.p, e.r); if (!m.has(k)) m.set(k, new Set()); m.get(k).add(e.p.tab); }); return [...m.values()].some(s => s.size > 1); };
/* Pop-out: the same page opened in an Office dialog window (?popout=1). A dialog cannot reach the workbook, so it asks the
   task pane (which stays open) to do every read and write, and the pane sends back a snapshot of what it read. */
const POPOUT = /[?&]popout=1\b/.test(location.search);
const pop = { dlg: null, calls: 0 };                                   // task pane side
const rpc = { id: 0, wait: new Map(), parts: [], snapWait: [] };       // pop-out side
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
function toast(msg, err) {
  if (pop.dlg && pop.calls > 0) toChild({ k: "toast", msg, err: !!err });   // a pop-out asked for this action: show the result there too
  const t = $("toast"); t.textContent = msg; t.className = err ? "err" : ""; t.style.display = "block"; clearTimeout(toastTimer); toastTimer = setTimeout(() => t.style.display = "none", err ? 9000 : 4500); }
const fail = e => { console.error(e); toast(e.message || String(e), true); };
function showModal(html) { $("modalBody").innerHTML = html; $("modal").hidden = false; const f = $("modalBody").querySelector("button"); if (f) f.focus(); }
function closeModal() { $("modal").hidden = true; $("modalBody").innerHTML = ""; }

/* ---------- read the workbook ---------- */
/* a missed visit's freebie decision lives in Notes ("Freebie used" or "Freebie waived"); p.fbCell (N6) counts the freebies left */
const freebieOf = note => /\bfreebie used\b/i.test(note) ? "Used" : /\bfreebie waived\b/i.test(note) ? "Waived" : "";
const withFreebie = (note, v) => [...String(note || "").split("; ").filter(x => x && !/^freebie (used|waived)$/i.test(x)), ...(v ? [FREEBIE_TAG[v]] : [])].join("; ");
const isNA = v => /^n\/?a$/i.test(String(v ?? "").trim());
/* layout 3 (current): Time in T8 and Freebie Count in J6, benefits on the Benefits tab. Layout 2: Time in T8 and the cost table
   header in V11 (freebies in N7, benefits on the tab). Anything else is an earlier layout, rebuilt from New Template. */
const layoutOf = (t8, j6, v11) => String(t8).trim().toLowerCase() !== "time" ? 0 : /^freebie count/i.test(String(j6).trim()) ? 3 : String(v11).trim().toLowerCase() === "code" ? 2 : 0;
function parseLog(vals) {
  const blank = x => x === "" || x == null;
  return vals.map((v, i) => {
    const note = String(v[C.note] || "");
    return {
      row: LOG_FIRST + i, pr: typeof v[C.pr] === "number" ? v[C.pr] : null, tx: v[C.tx], tech: String(v[C.tech] || "").trim(),
      date: isNum(v[C.date]) ? Math.floor(v[C.date]) : null, time: isNum(v[C.time]) ? Math.round(v[C.time] * 1440) : null,
      pulses: [v[C.L], v[C.R], v[C.O]].map(x => blank(x) ? null : x), delivered: [v[C.L], v[C.R], v[C.O]].some(x => !blank(x)),
      type: String(v[C.type] || ""), miss: String(v[C.miss]).trim().toLowerCase() === "x", action: String(v[C.action] || ""),
      alt: String(v[C.alt]).trim().toLowerCase() === "x", note, tele: /\btele/i.test(note), freebie: freebieOf(note),
      phq: String(v[C.phq] ?? "").trim(), gad: String(v[C.gad] ?? "").trim(),
      measX: [v[C.phq], v[C.gad]].some(x => String(x).trim().toLowerCase() === "x"), measVal: [v[C.phq], v[C.gad]].some(x => !blank(x) && String(x).trim().toLowerCase() !== "x")
    };
  });
}
async function refresh() {
  if (POPOUT) return callPane("refresh", []);
  try {
    await Excel.run(async ctx => {
      ctx.workbook.load("name"); const sheets = ctx.workbook.worksheets; sheets.load("items/name"); await ctx.sync(); state.wbName = ctx.workbook.name || "";
      const cand = sheets.items.filter(s => TAB_RE.test(s.name) || s.name === "New Template");
      const sg = cand.map(s => ({ name: s.name, a1: s.getRange("A1"), a3: s.getRange("A3"), h: s.getRange("C8:T8"), v11: s.getRange("V11"), j6: s.getRange("J6"), d7: s.getRange("D7"), e: s.getRange("E9:E200") }));
      sg.forEach(x => { x.a1.load("values"); x.a3.load("values"); x.h.load("values"); x.v11.load("values"); x.j6.load("values"); x.d7.load("values"); x.e.load("values"); });
      /* sheet names are already known, so the Protocols, Clinic Schedule and Benefits reads ride along with this sync (names match case-insensitively, as getItem does) */
      const sheetOf = n => sheets.items.find(s => s.name.toLowerCase() === n.toLowerCase());
      const prW = sheetOf("Protocols"), csW = sheetOf("Clinic Schedule"), bnW = sheetOf(BEN_SHEET);
      const prv = prW ? prW.getRange("A2:I40") : null, cs = csW ? csW.getRange("A3:H40") : null, bn = bnW ? bnW.getRange(`A1:S${BEN_LAST}`) : null;
      [prv, cs, bn].forEach(r => r && r.load("values")); await ctx.sync();
      readBen(bn ? bn.values : null);
      const cut = inputToSerial(state.cfg.cutoff || CUTOFF_DEFAULT), eDate = v => typeof v === "number" ? (v > 20000 ? Math.floor(v) : null) : /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(String(v).trim()) ? toSerial(v) : null;
      const live = x => x.name === "New Template" || (!/^(discontinued|complete)/i.test(String(x.d7.values[0][0]).trim()) && x.e.values.some(([v]) => (eDate(v) ?? 0) >= cut));
      state.stale = sg.filter(x => !live(x)).map(x => x.name); const stale = new Set(state.stale);
      const ok = x => String(x.a1.values[0][0]).trim().toUpperCase() === "PATIENT:" && String(x.a3.values[0][0]).trim().toUpperCase() === "MRN:" && SIG.every(([i, w]) => String(x.h.values[0][i]).trim().toLowerCase() === w);
      const lay = x => ok(x) ? layoutOf(x.h.values[0][17], x.j6.values[0][0], x.v11.values[0][0]) : 0;
      const layout = new Map(sg.map(x => [x.name, lay(x)]));
      state.templateOk = layout.get("New Template") === 3;
      state.old = sg.filter(x => x.name !== "New Template" && live(x) && ok(x) && !layout.get(x.name)).map(x => x.name);   // a tracker tab in an earlier layout: rebuilt from New Template
      state.skipped = { layout: sg.filter(x => x.name !== "New Template" && live(x) && !ok(x)).map(x => x.name), other: 0 };   // tabs not named by LASFIR are left alone without a mention
      const L = sheets.items.filter(s => layout.get(s.name) >= 2).map(s => {
        const g = (a, f) => { const r = s.getRange(a); r.load(f ? "formulas" : "values"); return r; };
        return { name: s.name, layout: layout.get(s.name), hdr: g("A2:H7"), prot: g("J2:O4"), n67: g("N6:N7"), w18: g("W1:W8"), cost: g(`V11:Y${COST_FIRST + 4}`), a10: g("A10", true),
          wk: g(`W${WK_FIRST}:X${WK_FIRST + 4}`), av: g(`V${AV_FIRST}:Y${AV_LAST}`), log: g(`A${LOG_FIRST}:T${LOG_FIRST + LOG_READ - 1}`), s9: g("S9") };
      });
      await ctx.sync();
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
          protoMin: mins.length ? mins.reduce((a, b) => a + b, 0) : null,
          home, external: !!home && !/^(srl|san rafael)$/i.test(home),
          ben: { bidf: money0(w[0]), dedNA: isNA(w[3]), dedRem: money0(w[3]), dedBlank: w[3] === "" || w[3] == null, oopNA: isNA(w[4]), oopRem: money0(w[4]),
            mtr: String(w[5] || ""), ext: typeof w[6] === "number" ? w[6] : 0, lastSched: isNum(w[7]) ? Math.floor(w[7]) : null },
          cost,
          canGrow: /LET\(/.test(a10),   // only a PR formula in A10 can say where a grown log ends
          rows: parseLog(l.log.values.slice(0, end - LOG_FIRST + 1)),
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
    if (mtrFix.length) try { await applyOps(mtrFix.map(p => ({ tab: p.tab, a: "W6", v: [["Yes"]], quiet: true }))); mtrFix.forEach(p => { p.ben.mtr = "Yes"; }); } catch (e) { console.error(e); }
    renderBanners(); renderDash(); updateApTab(); fillAddPatientForm();
    if (pop.dlg) sendSnap();
    syncActiveSheet();
    /* once a session: visit logs edited by hand get the formatting that follows their dates (hand-colored rows stay) */
    if (!state.fmtRan) { state.fmtRan = true; formatLogs(state.patients.filter(p => !state.stale.includes(p.tab)).map(p => ({ tab: p.tab, touched: [] }))).catch(e => console.error(e)); }
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
const protoText = p => p.protoRows.filter(x => isNum(x.hz)).map(x => `${x.site || "Protocol " + x.n}: ${protoName(x) || `${x.hz} Hz, ${x.dur}/train, ${x.trains} trains, ${x.wait}s`}`).join("; ") || "not entered";

function derive(p) {
  const today = todaySerial(), cfg = state.cfg, rows = p.rows, rec = cfg.recentDays, total = p.txApproved || 36;
  const status = p.status, active = status === "" || ACTIVE.includes(status);
  const sched = rows.filter(r => r.date != null && r.date >= today && !r.delivered && !r.miss).sort(cmp);
  const next = sched.find(r => r.type !== "F/U" && !r.alt) || sched[0];
  const tx = rows.filter(r => r.delivered && r.type !== "F/U");
  const txDone = tx.length, lastTx = tx.length ? Math.max(...tx.map(r => r.date ?? 0)) : null;
  const overdue = rows.filter(r => r.date != null && r.date < today && !r.delivered && !r.miss && !r.alt && r.type !== "F/U");
  const noTech = tx.filter(r => !r.tech && r.date != null);
  const decisions = rows.filter(r => r.miss && !r.freebie);
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
    if (status !== "Complete" && txDone >= total && (lastTx == null || lastTx < today - rec)) suggest = "Complete";
    else if (status !== "Complete" && quit) suggest = "Discontinued";
    else if (status !== "Complete" && txDone >= 1 && (status === "Pending Start" || status === "")) suggest = "In Progress";
    else if (status === "" && txDone === 0 && futureMT) suggest = "Pending Start";
  }
  const cancelled = /^(discontinued|paused|complete)$/i.test(status) || suggest === "Complete" || suggest === "Discontinued" || !(cxRows.length || allCx) ? null
    : [cxRows.length ? `Cancelled in NextGen, still on the tracker: ${[...new Set(cxRows.map(r => r.date))].sort((a, b) => a - b).map(fmtDate).join(", ")}` : "",
      allCx ? `Nothing booked in NextGen, last seen ${seenAt != null ? fmtDate(seenAt) : "never"}` : ""].filter(Boolean).join(". ");
  /* authorization (separate from protocol) */
  const authFlags = [];
  if (active) {
    const firstTx = rows.filter(r => r.date != null && r.type !== "F/U").map(r => r.date).sort((a, b) => a - b)[0];
    if (p.authStart != null && firstTx != null && firstTx < p.authStart) authFlags.push(`Auth starts ${fmtDate(p.authStart)}, after the first TMS visit ${fmtDate(firstTx)}`);
    if (p.authExp == null) authFlags.push("No auth expiry entered");
    else {
      const days = p.authExp - today;
      if (days < 0) authFlags.push(`Auth expired ${fmtDate(p.authExp)}`);
      else if (days <= cfg.authDays) authFlags.push(`Auth expires in ${days} d (${fmtDate(p.authExp)})`);
      if (p.txEnd != null && p.txEnd > p.authExp) authFlags.push(`Treatments run to ${fmtDate(p.txEnd)}, past auth expiry`);
      if (sched.some(r => r.date > p.authExp)) authFlags.push("Visits booked past auth expiry");
    }
    const txBooked = rows.filter(r => r.type !== "F/U" && r.type !== "" && !r.miss && (r.delivered || r.date != null)).length;
    if (txBooked > total) authFlags.push(`${txBooked} treatments booked or done, ${total} authorized (D2)`);
    if (!p.mtrAuth && sched.some(r => r.type === "MTR")) authFlags.push("MT redo (90869) booked, but D4 and MTR approved don't show it authorized");
  }
  /* protocol: renewal cadence from G2 (business days) and Action = Protocol; entered at the MT */
  const protoFlags = [];
  if (active) {
    if (p.updateBy != null) { if (p.updateBy < today) protoFlags.push(`Protocol update was due ${fmtDate(p.updateBy)}`); else if (bizDaysUntil(today, p.updateBy) <= cfg.protoDays) protoFlags.push(`Protocol update by ${fmtDate(p.updateBy)}`); }
    const pa = rows.filter(r => r.action === "Protocol" && !r.delivered && r.date != null && r.date >= today).sort(cmp).find(r => bizDaysUntil(today, r.date) <= cfg.protoDays);
    if (pa) protoFlags.push(`Protocol action at the ${fmtDate(pa.date)} visit`);
    if (!p.protoRows.some(x => isNum(x.hz)) && rows.some(r => r.type === "MT" && r.delivered)) protoFlags.push("Protocol not entered in J2:M4 since the MT");
  }
  /* patient cost so far and projected, the BIDF estimate and out-of-pocket max */
  const accrued = rows.filter(r => happened(r, today)).reduce((s, r) => s + (payOf(p, r) || 0), 0);
  const futurePay = sched.filter(r => !r.alt).map(r => ({ r, v: payOf(p, r) || 0 }));
  const projected = accrued + futurePay.reduce((s, x) => s + x.v, 0);
  const benFlags = [];
  if (active) {
    if (p.ben.bidf != null && projected > p.ben.bidf + 0.005) benFlags.push(`Projected patient cost ${money(projected)} exceeds the BIDF estimate ${money(p.ben.bidf)}`);
    if (p.ben.oopRem != null) {
      if (accrued >= p.ben.oopRem) benFlags.push("Out-of-pocket max likely met. Set the copay to $0 at pre-check and update Insurance Maintenance in NextGen");
      else { let c = accrued; const hit = futurePay.find(x => (c += x.v) >= p.ben.oopRem); if (hit) benFlags.push(`Likely meets the out-of-pocket max around ${fmtDate(hit.r.date)}. Verify the copay at pre-check and update Insurance Maintenance in NextGen`); }
    }
  }
  const mdOnSite = /^y/i.test(p.mdReq);
  /* only dates the patient said they are away; weekday preferences are for scheduling, not flags */
  const awayConflicts = [];
  if (active) for (const r of sched) { const w = p.avail.find(w => w.type === "Unavailable" && w.from <= r.date && r.date <= w.to); if (w) awayConflicts.push({ r, why: `away ${fmtDate(w.from)}\u2013${fmtDate(w.to)}${w.note ? ` (${w.note})` : ""}` }); }
  const closedConflicts = sched.filter(r => state.closures.has(r.date)).map(r => ({ r, why: state.closures.get(r.date) }));
  const byDay = new Map(); rows.filter(r => r.date != null && r.date >= today && !r.miss && !r.alt).forEach(r => byDay.set(r.date, (byDay.get(r.date) || 0) + 1));
  const sameDay = p.kaiser ? [] : [...byDay.entries()].filter(([, n]) => n > 1).map(([d]) => d).sort((a, b) => a - b);
  const measDue = rows.filter(r => r.measX && r.date != null && r.date < today && !r.miss);
  const lastRow = rows.filter(r => r.type !== "F/U" && r.date != null).sort(cmp).pop();
  const bidf = bidfMissing(p), todayVisit = rows.some(r => r.date === today && !r.miss && !r.alt);
  return { active, sched, next, txDone, lastTx, total, overdue, noTech, decisions, used, waived, left, fu, fuNext, fuStatus, stale, suggest, cancelled, lastSeen: seenAt,
    authFlags, protoFlags, benFlags, accrued, projected, mdOnSite, sameDay, measDue, awayConflicts, closedConflicts, bidf, bidfToday: active && todayVisit && bidf.length > 0,
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
  renderBanners();
}
const bctl = id => `<span class="bctl"><button class="sm" data-banner="${id}" data-bact="snooze">Snooze 3 h</button><button class="sm" data-banner="${id}" data-bact="dismiss" aria-label="Dismiss">Dismiss</button></span>`;
function renderBanners() {
  const t = todaySerial(), r = state.report, show = (id, on) => { $(id).hidden = !on || snoozed(id); };
  const stale = !r || r.asOf == null || r.asOf < t;
  $("reportMsg").textContent = !r ? "No appointment list has been uploaded yet. Run the report and use Import list at the top right."
    : r.asOf == null ? "The last appointment list has no run date. Run the report again and upload it."
      : `The appointment list is from ${fmtDateY(r.asOf)}${r.asOfMin != null ? " " + fmtTime(r.asOfMin) : ""}. Run an updated report and use Import list at the top right.`;
  show("reportBanner", stale);
  $("wbTitle").textContent = trackerTitle(); $("devTag").textContent = `${device()} tracker`;
  const sk = state.skipped, parts = [];
  if (sk.layout.length) parts.push(`${sk.layout.length} tab${sk.layout.length > 1 ? "s" : ""} in an older layout (${sk.layout.map(esc).join(", ")})`);
  if (!state.templateOk) parts.push("the New Template tab does not match the current layout (Time in T8, Freebie Count in J6)");
  const ns = state.stale.length;
  $("cutoffNote").textContent = ns ? `${ns} tab${ns === 1 ? " is" : "s are"} left out of updates: no visit on or after this date, or Discontinued/Complete.` : "";
  $("skipText").innerHTML = parts.length ? `Left untouched: ${parts.join("; ")}. The add-in only reads and writes tabs that match the current template.` : "";
  show("skipNote", parts.length > 0);
  const n = state.fix.size, o = state.old.length, msg = [];
  if (n) msg.push(`${n} tab${n === 1 ? " doesn't" : "s don't"} match New Template yet (${[...state.fix.keys()].map(esc).join(", ")}). Match New Template moves benefits to the ${BEN_SHEET} tab and freebies to N6, and brings the PR formula, colors by visit type, dropdowns, borders and the TX SUMMARY checkbox over every log row`);
  if (o) msg.push(`${o} tab${o === 1 ? " is" : "s are"} in the earlier layout (${state.old.map(esc).join(", ")}) and will be rebuilt from New Template. The earlier tab is deleted once its entries are copied over`);
  $("setupMsg").innerHTML = msg.join(". ") + ".";
  $("btnSetup").disabled = o > 0 && !state.templateOk;
  show("setup", n + o > 0);
  $("newPtList").innerHTML = state.pendingNew.map(newPtItem).join("");
  $("newPtCount").textContent = state.pendingNew.length;
  show("newPtBanner", state.pendingNew.length > 0);
}
const newPtItem = x => `<div class="item red"><b>${esc(x.code || "New patient")}</b> MRN ending ${esc(x.last4)}<div class="sub">${x.n} visit${x.n > 1 ? "s" : ""} today or later, first ${fmtDate(x.first.date)} ${esc(x.first.type)}. No tracker tab exists.</div><button class="sm" data-act="newTracker" data-v="${esc(x.mrn)}">Create tracker</button></div>`;

/* ---------- dashboard ---------- */
const item = (p, main, sub, cls, row, extra) =>
  `<div class="item ${cls || ""}" tabindex="0" data-tab="${esc(p.tab)}"${row ? ` data-row="${row}"` : ""}><b>${esc(p.tab)}</b> ${main || ""}${sub ? `<div class="sub">${esc(sub)}</div>` : ""}${extra || ""}</div>`;
const group = (title, items, empty) =>
  `<div class="group"><h2>${esc(title)}<span class="n">${items.length}</span></h2>${items.length ? items.join("") : `<div class="none">${esc(empty)}</div>`}</div>`;
const btn = (act, tab, val, label) => `<button class="sm" data-act="${act}" data-t="${esc(tab)}"${val ? ` data-v="${esc(val)}"` : ""}>${esc(label)}</button>`;
const bt = (act, tab, row, val, label, cls) => `<button class="sm${cls ? " " + cls : ""}" data-act="${act}" data-t="${esc(tab)}" data-row="${row}"${val != null && val !== "" ? ` data-v="${esc(val)}"` : ""}>${esc(label)}</button>`;
const ptLink = (p, row) => `<b class="lnk" data-pt="${esc(p.tab)}"${row ? ` data-sched="${row}"` : ""} tabindex="0">${esc(p.tab)}</b>`;

const NO_TABS = `<div class="none">No patient tabs found. A patient tab has "PATIENT:" in A1 and a LASFIR name such as SMIJAN or SMIJAN 2.</div>`;
function renderDash() {
  const v = state.view;
  document.querySelectorAll("#mainTabs button").forEach(b => b.classList.toggle("on", b.dataset.view === v || (v === "add" && b.dataset.view === "pts")));
  if (v === "sch") {
    document.querySelectorAll("#schToggle button").forEach(b => { const on = b.dataset.sv === state.dash; b.classList.toggle("on", on); b.setAttribute("aria-pressed", String(on)); });
    if (!state.patients.length) { $("schBody").innerHTML = NO_TABS; $("attBody").innerHTML = '<div class="none">All clear</div>'; $("attCount").textContent = "0"; return; }
    $("schBody").innerHTML = state.dash === "week" ? viewWeek() : viewToday();
    const att = viewAttention(); $("attBody").innerHTML = att.html; $("attCount").textContent = att.count;
  } else if (v === "pts") {
    $("ptsBody").innerHTML = !state.patients.length ? NO_TABS : state.ptMode === "pt" ? viewPatient() : viewPatients();
  } else if (v === "clinic") {
    document.querySelectorAll("#clinicTabs button").forEach(b => { const on = b.dataset.cv === state.clinicTab; b.classList.toggle("on", on); b.setAttribute("aria-pressed", String(on)); });
    $("clinicBody").innerHTML = state.clinicTab === "proto" ? viewProtocols() : viewSched();
  }
}
function syncAtt() { document.body.style.setProperty("--chrome-h", `${$("chrome").offsetHeight}px`); const wide = document.body.clientWidth >= 640; document.body.classList.toggle("wide", wide); document.body.classList.toggle("attc", wide && !!state.attCollapsed); if (wide !== state.wideAtt) { state.wideAtt = wide; $("attWrap").open = wide; } }

function attentionHtml() {
  const P = state.patients.filter(onSched), A = P.filter(p => p.d.active);
  const exempt = r => r.note.includes(EXEMPT_TAG);
  const dec = P.flatMap(p => p.d.decisions.map(r => ({ p, r })));
  return gi(`No appointment in ${state.cfg.recentDays}+ days: pause or discontinue?`, P.filter(p => p.d.stale).map(p => item(p, "", `${p.d.txDone}/${p.d.total} treatments, last ${p.d.lastTx ? fmtDate(p.d.lastTx) : "none"}`, "red", null, btn("status", p.tab, "Paused", "Paused") + btn("status", p.tab, "Discontinued", "Discontinued"))), "Nobody is stale")
    + gi("Status update suggested", P.filter(p => p.d.suggest).map(p => item(p, `<span class="chip">${esc(p.status || "blank")} \u2192 ${esc(p.d.suggest)}</span>`, p.d.suggest === "Discontinued" ? `Every NextGen appointment cancelled, last seen ${fmtDate(p.d.lastSeen)}` : "", "amb", null, btn("status", p.tab, p.d.suggest, "Apply"))), "Statuses are current")
    + gi("Cancelled in NextGen: update status?", P.filter(p => p.d.cancelled).map(p => item(p, "", p.d.cancelled, "amb", null, ["Paused", "Discontinued", "Complete"].map(s => btn("status", p.tab, s, s)).join(""))), "None")
    + gi("BIDF info needed today", A.filter(p => p.d.bidfToday).map(p => item(p, "", `Scheduled today. Enter ${p.d.bidf.join("; ")}`, "red", null, `<button class="sm" data-pt="${esc(p.tab)}" data-sec="pb">Enter benefits</button>`)), "None")
    + gi("Auth flags", A.filter(p => p.d.authFlags.length).map(p => item(p, "", p.d.authFlags.join("; "), "red")), "No auth issues")
    + gi("Protocol", A.filter(p => p.d.protoFlags.length).map(p => item(p, "", p.d.protoFlags.join("; "), "amb")), "None due")
    + gi("Benefits and patient cost", A.filter(p => p.d.benFlags.length).map(p => item(p, "", p.d.benFlags.join("; "), "amb")), "Nothing to check")
    + gi("Booked while the patient is away", A.filter(p => p.d.awayConflicts.length).map(p => item(p, "", p.d.awayConflicts.map(c => `${fmtDate(c.r.date)}: ${c.why}`).join(" | "), "red", p.d.awayConflicts[0].r.row)), "None")
    + gi("Booked on a clinic closure date", P.filter(p => p.d.closedConflicts.length).map(p => item(p, "", p.d.closedConflicts.map(c => `${fmtDate(c.r.date)} (${c.why})`).join(", "), "red", p.d.closedConflicts[0].r.row)), "None")
    + gi("Same-day appointments (not Kaiser)", A.filter(p => p.d.sameDay.length).map(p => item(p, "", `More than one appointment on ${p.d.sameDay.map(fmtDate).join(", ")}`, "red")), "None")
    + gi("Freebie decision needed", dec.filter(({ r }) => !exempt(r)).map(({ p, r }) => item(p, `<span class="chip amb">Missed</span> ${fmtDateY(r.date)}`, r.note.includes(UNCLASS_TAG) ? "No-show not classified in NextGen. Clean up there, then classify" : "Classify the freebie", "amb", r.row)), "Every missed visit is classified")
    + gi("Exempt in NextGen: confirm", dec.filter(({ r }) => exempt(r)).map(({ p, r }) => item(p, `<span class="chip">Exempt</span> ${fmtDateY(r.date)}`, "Payor exempt, or a freebie? Use freebie or Waived", "amb", r.row)), "None")
    + gi("Pulses or technician missing", P.flatMap(p => [...p.d.overdue.map(r => item(p, `${fmtDate(r.date)}`, "No pulses recorded and not marked missed", "amb", r.row)), ...p.d.noTech.map(r => item(p, `${fmtDate(r.date)}`, "Pulses entered, treating technician blank", "amb", r.row))]), "Nothing outstanding")
    + gi("Measures not collected", P.filter(p => p.d.measDue.length).map(p => item(p, "", `${p.d.measDue.length} past visit${p.d.measDue.length > 1 ? "s" : ""} still marked x, oldest ${fmtDate(p.d.measDue[0].date)}. Enter the scores or move them`, "amb", p.d.measDue[0].row)), "None");
}

const gi = (title, items) => items.length ? group(title, items, "") : "";
function viewAttention() { const html = attentionHtml(); return { html: html || '<div class="none">All clear</div>', count: (html.match(/class="item /g) || []).length }; }

/* ---------- Today: one card per visit, colored by event type like the tab's row highlight ---------- */
const cardCls = r => r.alt ? "c-alt" : r.type === "MT" ? "c-mt" : r.type === "MTR" ? "c-mtr" : r.type === "F/U" ? "c-fu" : "c-daily";
function completeBlock(p, r) {
  const pul = ["L", "R", "O"].map((s, i) => r.pulses[i] != null ? `${s} ${r.pulses[i]}` : "").filter(Boolean).join(", ");
  if (r.delivered) return `<div class="done"><span class="chip ok">Delivered</span> <span class="sub">${esc(pul)}${r.tech ? ` \u00B7 ${esc(r.tech)}` : ""}</span></div>`;
  const key = `cmp|${p.tab}|${r.row}`, rows = p.protoRows.filter(x => isNum(x.dur) && isNum(x.trains)), want = { Left: 0, Right: 0, Other: 0 };
  rows.forEach(x => { want[SITES.includes(x.site) ? x.site : "Left"] += x.dur * x.trains; });
  return `<details class="nt" data-open="${esc(key)}"${state.open.has(key) ? " open" : ""}><summary>Complete treatment</summary>
    <div class="sub">${rows.length ? "Protocol calls for: " + rows.map(x => `${esc(x.site || "Protocol " + x.n)} ${x.dur * x.trains} (${x.trains} trains x ${x.dur})`).join("; ") : "No protocol in J2:M4 yet. Enter the pulses delivered."}</div>
    <div class="g3">${SITES.map((s, i) => `<label>${s}<input type="number" min="0" data-pul="${i}" value="${want[s] || ""}"></label>`).join("")}</div>
    <label>Treating technician<input data-tech value="${esc(r.tech || state.techName || "")}"></label>
    ${bt("pulses", p.tab, r.row, "", "Save treatment", "pri")}</details>`;
}
function cardMenu(p, r) {
  const items = [];
  if (r.miss) items.push(bt("undoMiss", p.tab, r.row, "", "Undo missed visit"));
  else if (!r.delivered && !r.alt) OUTCOMES.forEach(k => items.push(bt("missKind", p.tab, r.row, k, `Record ${k.toLowerCase()}`)));
  items.push(bt("print", p.tab, r.row, "", "Print schedule"), `<button class="sm" data-pt="${esc(p.tab)}">Open patient</button>`);
  return `<details class="kebab"><summary title="More actions" aria-label="More actions for ${esc(p.tab)}">\u22EF</summary><div class="menu">${items.join("")}</div></details>`;
}
function measBlock(p, r) {
  const key = `meas|${p.tab}|${r.row}`, open = state.open.has(key);
  const up = p.rows.filter(x => x.date != null && r.date != null && x.date > r.date && !x.delivered && !x.miss).sort(cmp);
  const ms = up.find(x => STD_VISIT.includes(x.type) && x.date - r.date <= 7), def = ms || up[0];
  let h = `<div class="meas"><button class="mbtn" data-act="measToggle" data-t="${esc(p.tab)}" data-row="${r.row}" aria-expanded="${open}">Measures needed</button>`;
  if (open) {
    h += `<div class="row2" style="margin-top:6px"><input type="number" min="0" max="27" data-m="P" placeholder="PHQ-9"><input type="number" min="0" max="21" data-m="Q" placeholder="GAD-7"></div>${bt("saveMeas", p.tab, r.row, "", "Save measures", "pri")}`;
    if (up.length) h += `<div class="sub" style="margin-top:6px">Not collected today? Move to</div><select data-mvto>${up.slice(0, 20).map(x => `<option value="${x.row}"${x === def ? " selected" : ""}>${DAYS[dow(x.date)]} ${fmtDate(x.date)} ${esc(apptTypeName(x))}</option>`).join("")}</select>${ms ? `<div class="sub">${esc(apptTypeName(ms))} on ${fmtDate(ms.date)} collects measures at check-in.</div>` : ""}${bt("moveMeas", p.tab, r.row, "", "Move measures")}`;
  }
  return h + `</div>`;
}
function protoEditor(p, keyBase) {
  const opts = v => `<option value="">(none)</option>` + state.protocols.map(q => `<option value="${esc(q.name)}"${q.name.trim() === v ? " selected" : ""}>${esc(q.name.trim())}</option>`).join("");
  return `<div class="pe">${p.protoRows.map(x => `<div class="g2p"><span>Protocol ${x.n}</span><select data-pname="${x.n}">${opts(protoName(x))}</select><select data-psite="${x.n}"><option value="">Site</option>${SITES.map(s => `<option${s === x.site ? " selected" : ""}>${s}</option>`).join("")}</select></div>`).join("")}
    ${bt("protoSave", p.tab, 0, keyBase, "Save protocol to J2:O4")}</div>`;
}
function mtBlock(p, r) {
  const key = `mt|${p.tab}|${r.row}`, g = k => { const m = r.note.match(new RegExp(`${k}:\\s*([^;]+)`, "i")); return m ? m[1].trim() : ""; };
  const noAvail = ![1, 2, 3, 4, 5].some(d => p.week[d] && (p.week[d].avail || p.week[d].from != null)) && !p.avail.length;
  return `<details class="nt" data-open="${esc(key)}"${state.open.has(key) ? " open" : ""}><summary>MT details (saved to Notes)</summary>
    <div class="g2"><label>Hotspot<input data-mt="Hotspot" value="${esc(g("hotspot"))}"></label><label>MT %<input data-mt="MT %" value="${esc(g("mt %"))}"></label>
    <label>Amps<input data-mt="Amps" value="${esc(g("amps"))}"></label><label>Baseline sleep<input data-mt="Baseline sleep" value="${esc(g("baseline sleep"))}" placeholder="e.g. 9 hrs"></label>
    <label>Baseline caffeine<input data-mt="Baseline caffeine" value="${esc(g("baseline caffeine"))}" placeholder="e.g. 2 cups"></label></div>
    ${bt("mtSave", p.tab, r.row, "", "Save MT details")}
    <div class="sub" style="margin-top:8px">Protocol chosen today (fills Hz, Dur, Trains, Wait; O = site)</div>${protoEditor(p, r.row)}
    ${noAvail ? `<div class="sub" style="margin-top:6px">Availability not entered yet. <b class="lnk" data-pt="${esc(p.tab)}">Open the patient</b> to add it.</div>` : ""}</details>`;
}
function noteForm(p, r) {
  const dr = draftOf(p, r), bl = baselineOf(p), sel = (dk, opts, v) => `<select data-dk="${dk}">${opts.map(o => `<option${o === v ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
  const qs = NOTE_QS.map(q => `<div class="nq"><span>${esc(q.q)}</span>${sel("a." + q.k, q.opts || ["No", "Yes"], dr.a[q.k])}${q.flag && dr.a[q.k] === "Yes" ? `<input data-dk="d.${q.k}" value="${esc(dr.d[q.k] || "")}" placeholder="Specify (required)">` : ""}</div>`).join("");
  const ref = bl && (bl.sleep || bl.caffeine) ? `<div class="bl">Baseline at MT: sleep ${esc(bl.sleep || "not noted")}; caffeine ${esc(bl.caffeine || "not noted")}. Ask how they slept and what they drank, and compare.</div>` : "";
  return `<details class="nt" data-draft="${esc(dr.key)}"${dr.open ? " open" : ""}><summary>Draft treatment note</summary>
    <label>Mood on arrival</label>${sel("mood", MOODS, dr.mood)}
    <label>Observations</label>${ref}${qs}
    <label>Technician comments</label><textarea data-dk="comments" rows="2">${esc(dr.comments)}</textarea>
    <label>Initials</label><input data-dk="initials" value="${esc(dr.initials)}" maxlength="4" style="width:80px">
    <label>Note${dr.edited ? " (edited by hand)" : ""}</label><textarea data-dk="text" rows="12" class="ntext">${esc(dr.edited ? dr.text : noteText(p, r, dr))}</textarea>
    <div>${bt("copyNote", p.tab, r.row, "", "Copy note")}${dr.edited ? bt("resetNote", p.tab, r.row, "", "Rebuild from answers") : ""}</div></details>`;
}
function visitCard(p, r) {
  const treat = r.type !== "F/U", pay = payOf(p, r), kind = OUTCOMES.find(k => r.note.includes(k)), [tm, ap] = (fmtTime(r.time) || "\u2013 ").split(" ");
  let h = `<div class="card ${cardCls(r)}"><div class="chead"><div class="ctime">${tm}<span>${ap || ""}</span></div>
    <div class="cmain"><div class="cname">${ptLink(p, r.row)} <span class="chip">${esc(apptTypeName(r))}</span>${r.alt ? ' <span class="chip amb">Other clinic</span>' : ""}${chairs().length > 1 && chairPick() === "all" && chairOf(p, r) ? ` <span class="chip">${esc(chairOf(p, r))}</span>` : ""}</div><div class="sub">MRN ${esc(p.mrnRaw || "not entered")}</div></div>${cardMenu(p, r)}</div>`;
  if (r.miss) return h + `<div class="cbody"><div class="mline"><span class="chip amb">${esc(kind || "Missed")}</span> ${r.freebie ? `<span class="chip">${r.freebie === "Used" ? "Freebie used" : "Waived"}</span>` : bt("classify", p.tab, r.row, "", "Classify freebie", "pri")}</div></div></div>`;
  h += `<div class="cbody">`;
  if (p.d.bidfToday) h += `<div class="item red bidf"><b>BIDF info needed</b><div class="sub">Enter ${esc(p.d.bidf.join("; "))}</div><button class="sm" data-pt="${esc(p.tab)}" data-sec="pb">Enter benefits</button></div>`;
  const priced = (p.cost[billIdx(r)] || {}).after != null;
  if (!r.alt) h += `<div class="pay">${!priced ? `<span class="sub">Payment not set</span>` : pay > 0 ? `<span>Payment ${money(pay)}</span>` : `<span class="sub">Payment $0.00</span>`}</div>`;
  /* every MT gets the course schedule printout; the card says whether one has been given since */
  if (r.type === "MT") { const ls = p.ben.lastSched, given = ls != null && ls >= r.date; h += `<div class="mline">${bt("print", p.tab, r.row, "all", "Print course schedule", given ? "" : "pri")}<span class="sub">${given ? `Schedule given ${fmtDate(ls)}` : ls != null ? `Last schedule given ${fmtDate(ls)}, before this MT` : "No schedule given yet"}</span></div>`; }
  if (treat && !r.alt) h += completeBlock(p, r);
  if (r.measX) h += measBlock(p, r);
  if (r.type === "MT") h += mtBlock(p, r);
  if (r.type === "Daily" || r.type === "Taper") h += noteForm(p, r);
  return h + `</div></div>`;
}
/* schedule variances: why today's treatment count differs from the number of patients in treatment */
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
function varianceHtml(day, open) {
  const v = variances(day, true), lines = Object.entries(v.groups).filter(([, a]) => a.length);
  const gradTxt = v.grads.length ? `<div class="vg"><b>Graduated this week</b> ${v.grads.map(p => `${ptLink(p)} <span class="sub">${fmtDate(p.d.lastTx)}</span>`).join(", ")}</div>` : "";
  return `<details class="group var" data-open="var|${day}"${open || state.open.has("var|" + day) ? " open" : ""}><summary class="gs">Schedule variances <span class="sub">${v.actual} of ${v.expected} patients in treatment are being treated</span></summary>
    ${lines.length || gradTxt ? lines.map(([k, a]) => `<div class="vg"><b>${esc(k)}</b> <span class="n">${a.length}</span><div>${a.map(([p, d]) => `${ptLink(p)}${d ? ` <span class="sub">${esc(d)}</span>` : ""}`).join(", ")}</div></div>`).join("") + gradTxt : '<div class="none">No variances</div>'}</details>`;
}
/* the saved appointment list against the tabs for one day: every visit Expected or Kept in NextGen should be on a tab that day,
   and every visit on a tab should be on the list. Counted per patient as treatments (MT, MTR, Daily) and follow-ups. */
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
function listCheckHtml(day) {
  const c = listCheck(day); if (!c) return "";
  const run = state.report && state.report.asOf != null ? ` (run ${fmtDateY(state.report.asOf)}${state.report.asOfMin != null ? " " + fmtTime(state.report.asOfMin) : ""})` : "";
  if (!c.miss.length && !c.extra.length) return c.n ? `<div class="sub" style="margin:-4px 0 8px">All ${c.n} visit${c.n === 1 ? "" : "s"} on the appointment list${run} are on the tracker.</div>` : "";
  const what = a => `${fmtTime(a.time) || "no time"} ${esc(apptTypeName(a))}`;
  const items = c.miss.map(({ a, p, why }) => why === "new"
    ? `<div class="item red"><b>${esc(a.code || "New patient")}</b> MRN ending ${esc(a.mrn.slice(-4))} ${what(a)}<div class="sub">On the appointment list, no tracker tab.</div>${state.pendingNew.some(x => x.mrn === a.mrn) ? `<button class="sm" data-act="newTracker" data-v="${esc(a.mrn)}">Create tracker</button>` : ""}</div>`
    : why === "disc" ? item(p, what(a), `${p.status} on the tracker but booked in NextGen. Cancel it in NextGen, or start a new course tab.`, "red")
      : item(p, what(a), "On the appointment list, not on the tab.", "red", null, `<button class="sm" data-act="reimport">Import the list again</button>`))
    .concat(c.extra.map(({ p, r }) => item(p, `${fmtTime(r.time) || "no time"} ${esc(apptTypeName(r))}`, `On the tab, not on the appointment list${run}. Cancelled or booked since? Check NextGen.`, "amb", r.row)));
  return group("Appointment list check", items, "");
}
/* a day's hours: that weekday's hours on Clinic Schedule, unless the date is listed in G as closed */
function hoursOf(day) {
  if (state.closures.has(day)) return { closed: true, txt: `Closed: ${state.closures.get(day)}` };
  const h = state.clinic?.days[dow(day)];
  if (!h || h.open == null || h.close == null) return null;
  return { closed: false, txt: `${fmtTime(h.open)}–${fmtTime(h.close)}${h.bs != null && h.be != null ? `, break ${fmtTime(h.bs)}–${fmtTime(h.be)}` : ""}` };
}
function viewToday() {
  const t = todaySerial() + state.dayOffset, P = state.patients.filter(onSched), lbl = state.dayOffset ? "this day" : "today", hrs = hoursOf(t);
  const todays = P.flatMap(p => p.rows.filter(r => r.date === t && inChair(p, r)).map(r => ({ p, r }))).sort((a, b) => (a.r.time ?? 1e4) - (b.r.time ?? 1e4) || a.p.tab.localeCompare(b.p.tab));   // untimed last
  const tx = todays.filter(({ r }) => r.type !== "F/U"), fu = todays.filter(({ r }) => r.type === "F/U");
  let h = chairBar() + `<div class="wk"><button data-act="dprev">‹ Prev</button><b>${DAYS[dow(t)]} ${fmtDateY(t)}</b><button data-act="dnext">Next ›</button></div>`;
  h += `<div class="sub" style="margin:-4px 0 10px">${tx.length} treatment${tx.length === 1 ? "" : "s"}, ${fu.length} follow-up${fu.length === 1 ? "" : "s"}${hrs && !hrs.closed ? `. Open ${esc(hrs.txt)}` : ""}${state.dayOffset ? ' <button class="sm" data-act="dthis">Back to today</button>' : ""}</div>`;
  if (hrs?.closed) h += `<div class="item red"><b>Clinic closed</b> ${esc(state.closures.get(t))}</div>`;
  h += listCheckHtml(t) + varianceHtml(t, false);
  h += `<div class="group"><h2>Treatments ${lbl}<span class="n">${tx.length}</span></h2>${tx.length ? tx.map(({ p, r }) => visitCard(p, r)).join("") : `<div class="none">No treatments scheduled ${lbl}</div>`}</div>`;
  h += `<div class="group"><h2>Follow-ups ${lbl}<span class="n">${fu.length}</span></h2>${fu.length ? fu.map(({ p, r }) => visitCard(p, r)).join("") : `<div class="none">No follow-ups ${lbl}</div>`}</div>`;
  return h;
}

function viewWeek() {
  const t = todaySerial(), start = weekStart(t) + 7 * state.weekOffset, P = state.patients.filter(onSched);
  let html = chairBar() + `<div class="wk"><button data-act="prev">\u2039 Prev</button><b>${fmtDate(start)} \u2013 ${fmtDate(start + 6)}</b><button data-act="next">Next \u203A</button></div>`;
  if (state.weekOffset) html += `<div style="margin:-4px 0 10px"><button data-act="this">Back to this week</button></div>`;
  for (let i = 1; i < 6; i++) {
    const day = start + i, v = variances(day, true);
    const rows = P.flatMap(p => p.rows.filter(r => r.date === day && inChair(p, r)).map(r => ({ p, r }))).sort((x, y) => (x.r.time ?? 1e4) - (y.r.time ?? 1e4) || x.p.tab.localeCompare(y.p.tab));
    const note = Object.entries(v.groups).filter(([, a]) => a.length).map(([k, a]) => `${a.length} ${k.toLowerCase()}`).join(", ");
    const hrs = hoursOf(day);
    html += `<div class="day${day === t ? " today" : ""}"><h2><span>${DAYS[i]} ${fmtDate(day)}</span><span class="n">${v.actual} of ${v.expected} treated</span></h2>${hrs ? `<div class="${hrs.closed ? "item red" : "sub"}" style="margin-bottom:4px">${esc(hrs.closed ? hrs.txt : `Open ${hrs.txt}`)}</div>` : ""}${note ? `<div class="sub" style="margin-bottom:4px">Variances: ${esc(note)}</div>` : ""}`;
    html += rows.length ? rows.map(({ p, r }) => {
      const chips = [r.type && r.type !== "Daily" ? `<span class="chip ${cardCls(r)}">${esc(r.type)}${r.tele ? " tele" : ""}</span>` : "", r.action ? `<span class="chip">${esc(r.action)}</span>` : "", r.measX ? `<span class="chip">Measures</span>` : "",
        r.delivered ? `<span class="chip">Done</span>` : "", r.miss ? `<span class="chip amb">Missed</span>` : "", r.alt ? `<span class="chip amb">Other clinic</span>` : "",
        p.d.awayConflicts.some(c => c.r.row === r.row) ? `<span class="chip red">patient away</span>` : "", state.closures.has(day) && !r.delivered && !r.miss ? `<span class="chip red">clinic closed</span>` : ""].filter(Boolean).join(" ");
      return item(p, `${fmtTime(r.time)} ${chips}`, "", "", r.row);
    }).join("") : `<div class="none">No visits</div>`;
    html += `</div>`;
  }
  return html;
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
function viewSched() {
  if (!state.clinic) return `<div class="none">The Clinic Schedule sheet does not exist yet. It holds the Monday to Friday chair hours everyone sees.</div><div class="actions"><button class="pri" data-act="mkSched">Create Clinic Schedule sheet</button></div>`;
  const day = todaySerial() + state.schedOffset, all = dayEntries(day), keep = e => inChair(e.p, e.r), ent = { chair: all.chair.filter(keep), prov: all.prov.filter(keep), noTime: all.noTime.filter(keep) };
  let html = chairBar() + `<div class="wk"><button data-act="sprev">\u2039 Prev</button><b>${DAYS[dow(day)]} ${fmtDateY(day)}</b><button data-act="snext">Next \u203A</button></div>`;
  if (state.schedOffset) html += `<div style="margin:-4px 0 10px"><button data-act="sthis">Back to today</button></div>`;
  const d = dow(day); if (d === 0 || d === 6) return html + `<div class="none">Weekend. The chair schedule covers Monday to Friday.</div>`;
  if (!state.clinic.closureTable) html += `<div class="item amb">Add the holiday closure table to the Clinic Schedule sheet. <button class="sm" data-act="mkSched">Add it</button></div>`;
  if (state.closures.has(day)) html += `<div class="item red"><b>Clinic closed</b> ${esc(state.closures.get(day))}</div>${ent.chair.length + ent.prov.length + ent.noTime.length ? `<div class="item red">Visits are booked on this closed date: ${[...ent.chair, ...ent.prov, ...ent.noTime].map(e => esc(e.p.tab)).filter((v, i, a) => a.indexOf(v) === i).join(", ")}</div>` : ""}`;
  const sl = slotsFor(day, ent);
  if (!sl) html += `<div class="none">No chair hours set for this day. Enter them on the Clinic Schedule sheet.</div>`;
  else {
    if (!sl.h || sl.h.open == null) html += `<div class="hint">No hours entered for this day on the Clinic Schedule sheet. Showing the booked range.</div>`;
    html += sl.slots.map(s => {
      const occ = ent.chair.filter(e => e.start < s.e && e.end > s.s), clash = chairClash(occ);
      const txt = occ.map(e => `<b>${esc(e.p.tab)}</b> ${e.start >= s.s ? `<span class="chip${e.r.type === "Daily" || e.r.type === "Taper" ? "" : " amb"}">${esc(e.r.type)}</span>` : "\u21B3"}`).join(" &nbsp; ");
      return `<div class="slot${clash ? " bad" : s.brk ? " closed" : ""}"><span class="t">${fmtTime(s.s)}</span><span>${s.brk && !occ.length ? "Break" : txt}${clash ? ' <span class="chip red">chair conflict</span>' : ""}</span></div>`;
    }).join("");
    const out = ent.chair.filter(e => e.start < sl.open || e.end > sl.close);
    if (out.length) html += `<div class="item red" style="margin-top:8px">Outside chair hours: ${out.map(e => `${esc(e.p.tab)} ${fmtTime(e.start)}`).join(", ")}</div>`;
  }
  html += group(`${chairName()} chair. Provider (MT 60 min, MTR and F/U 30 min)`, ent.prov.sort((a, b) => a.start - b.start).map(e => item(e.p, `${fmtTime(e.start)}\u2013${fmtTime(e.end)} <span class="chip">${esc(e.r.type)}${e.r.tele ? " tele" : ""}</span>`, "", "", e.r.row)), "No provider visits");
  if (ent.noTime.length) html += group("No time recorded", ent.noTime.map(({ p, r }) => item(p, esc(r.type), "Import the appointment list to fill times", "amb", r.row)), "");
  return html + `<div class="actions"><button data-act="schedSheet">Write this week to the Chair Schedule sheet</button></div>`;
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
const PT_FILTERS = { all: ["All", onSched], active: ["Active", p => p.d.active], nofu: ["Follow-up not scheduled", p => p.d.active && p.d.fuStatus === "not scheduled"], nonext: ["No upcoming visit", p => p.d.active && !p.d.next && p.d.txDone < p.d.total], disc: ["Discontinued", p => !onSched(p)] };
function viewPatients() {
  const ws = weekStart(todaySerial()) + 7 * state.weekOffset, days = [1, 2, 3, 4, 5].map(i => ws + i), f = PT_FILTERS[state.ptFilter] || PT_FILTERS.all;
  const bar = `<div class="wk"><div class="fchips">${Object.entries(PT_FILTERS).map(([k, [lab, fn]]) => `<button class="sm${k === state.ptFilter ? " on" : ""}" data-act="ptFilter" data-v="${k}" aria-pressed="${k === state.ptFilter}">${esc(lab)} <span class="n">${state.patients.filter(fn).length}</span></button>`).join("")}</div><button class="pri" data-act="addPt">Add patient</button></div>`;
  const rows = state.patients.filter(f[1]).map(p => `<tr data-pt="${esc(p.tab)}" tabindex="0"><td><b>${esc(p.tab)}</b><div class="sub">${esc(p.mrnRaw)}</div></td><td>${esc(p.status)}</td>
    ${days.map(d => { const c = dayCode(p, d); return `<td class="dc ${c === "X" ? "x" : c === "MT" ? "c-mt" : c === "MTR" ? "c-mtr" : c === "F/U" ? "c-fu" : c === "Alt" ? "c-alt" : ""}">${esc(c)}</td>`; }).join("")}
    <td>${p.d.fuNext ? fmtDate(p.d.fuNext.date) : ""}</td><td>${fmtDate(p.d.grad)}</td><td>${p.d.mdOnSite ? "Yes" : ""}</td><td>${fmtDate(p.ben.lastSched)}</td></tr>`).join("");
  return `${bar}<div class="wk"><button data-act="prev">\u2039</button><b>Week of ${fmtDate(ws + 1)}</b><button data-act="next">\u203A</button></div>
    <div class="tscroll"><table class="pts"><thead><tr><th>Patient / MRN</th><th>Status</th>${days.map(d => `<th>${DAYS[dow(d)]}<br>${fmtDate(d)}</th>`).join("")}<th>Next F/U</th><th>Graduation</th><th>MD on-site</th><th>Schedule given</th></tr></thead><tbody>${rows}</tbody></table></div>
    <div class="hint">Day cells show MT, MTR, F/U, an action, Alt (other clinic) or X (missed). A dot is a daily treatment. Select a patient to open their view.</div>`;
}
/* Patient view */
function viewPatient() {
  const p = state.patients.find(x => x.tab === state.ptTab); if (!p) { state.ptMode = "list"; return viewPatients(); }
  const d = p.d, t = todaySerial(), fact = (k, v) => v === "" || v == null ? "" : `<div class="fct"><span>${esc(k)}</span><b>${esc(v)}</b></div>`;
  const up = d.sched.slice(0, 12), all = p.rows.filter(r => r.date != null).sort(cmp);
  const meas = p.rows.filter(r => r.measVal && r.date != null).sort(cmp);
  const av = state.avDraft && state.avDraft.tab === p.tab ? state.avDraft : (state.avDraft = { tab: p.tab, week: JSON.parse(JSON.stringify(p.week)), ranges: p.avail.map(w => ({ ...w })) });
  const dedOn = p.ben.dedRem > 0, mv = v => v == null ? "" : v;
  const billRows = BILL.map((b, i) => `<tr><td>${b.code}</td><td>${esc(b.label)}</td><td class="num"><input type="number" min="0" step="0.01" data-cost="X${COST_FIRST + i}" value="${mv(p.cost[i].after)}" aria-label="${esc(b.label)} patient responsibility"></td><td class="num"><input type="number" min="0" step="0.01" data-cost="Y${COST_FIRST + i}" value="${mv(p.cost[i].before)}" aria-label="${esc(b.label)} before deductible"${dedOn ? "" : ` placeholder="n/a"`}></td></tr>`).join("");
  const dedTxt = p.ben.dedNA ? "N/A" : p.ben.dedRem != null ? p.ben.dedRem : "", oopTxt = p.ben.oopNA ? "N/A" : p.ben.oopRem != null ? p.ben.oopRem : "";
  return `<div class="wk"><button data-act="ptBack">\u2039 Patients</button><span class="cact">${btn("print", p.tab, "", "Print or save schedule")}<button class="sm" data-tab="${esc(p.tab)}">Open tab</button></span></div>
    <h2 class="pth">${esc(p.tab)} <span class="sub">${esc(p.status)}</span></h2>
    <div class="facts">${fact("MRN", p.mrnRaw)}${fact("Insurance", p.insurance)}${fact("MD on-site", d.mdOnSite ? "Required" : "No")}${fact("Provider", p.provider)}${fact("Care navigator (TAS)", p.tas)}
      ${fact("Authorization", p.authStart || p.authExp ? `${fmtDate(p.authStart)} to ${fmtDate(p.authExp)}` : "")}${fact("Treatments", `${d.txDone} of ${d.total}${p.mtRedos ? `, ${p.mtRedos} MT redo${p.mtRedos > 1 ? "s" : ""}` : ""}`)}${fact("Graduation", fmtDateY(d.grad))}
      ${fact("Next follow-up", d.fuNext ? `${fmtDateY(d.fuNext.date)} ${fmtTime(d.fuNext.time)}${d.fuNext.tele ? " tele" : ""}` : "Not scheduled")}${fact("Freebies left", d.left)}${fact("Home clinic", p.home)}${fact("Schedule last given", fmtDateY(p.lastSched))}
      ${fact("Auth assessment", p.authAssess)}${fact("Progress note sent", fmtDateY(p.progSent))}</div>
    ${d.bidf.length ? `<div class="item ${d.bidfToday ? "red" : "amb"}"><b>BIDF info needed</b> ${esc(d.bidf.join("; "))}</div>` : ""}
    ${[...d.authFlags, ...d.protoFlags, ...d.benFlags].map(f => `<div class="item amb">${esc(f)}</div>`).join("")}
    <details class="group" data-open="pp|${esc(p.tab)}"${state.open.has("pp|" + p.tab) ? " open" : ""}><summary class="gs">Protocol <span class="sub">${esc(protoText(p))}</span></summary>${protoEditor(p, "pv")}</details>
    <details class="group ben" data-open="pb|${esc(p.tab)}"${state.open.has("pb|" + p.tab) ? " open" : ""}><summary class="gs">Benefits and cost <span class="sub">projected ${money(d.projected) || "$0.00"}</span></summary>
      <div class="facts">${fact("Patient cost so far", money(d.accrued) || "$0.00")}${fact("Projected course cost", money(d.projected) || "$0.00")}
        ${fact("Out-of-pocket max", p.ben.oopRem != null ? `${money(p.ben.oopRem)} remaining (${Math.min(100, Math.round(100 * d.projected / Math.max(p.ben.oopRem, 0.01)))}% projected)` : "")}</div>
      <div class="g3"><label>Deductible remaining<input data-benf="W4" value="${esc(dedTxt)}" placeholder="$ or N/A"></label><label>Out-of-pocket max remaining<input data-benf="W5" value="${esc(oopTxt)}" placeholder="$ or N/A"></label><label>BIDF estimated course cost<input type="number" min="0" step="0.01" data-benf="W1" value="${mv(p.ben.bidf)}"></label></div>
      <table class="cost"><thead><tr><th>Code</th><th>Visit</th><th class="num">PR</th><th class="num">Before ded.</th></tr></thead><tbody>${billRows}</tbody></table>
      <div class="hint">PR is the patient's cost per visit (after the deductible). With a deductible remaining, enter the before-deductible cost too: column A charges that until the deductible is used up, then PR. Enter N/A when there is no deductible.</div>
      ${btn("benSave", p.tab, "", "Save benefits")}
      <div class="g2" style="margin-top:10px"><label>Extension (daily visits)<input type="number" min="1" data-extn value="15"></label><label>New auth expiry<input type="date" data-extd></label></div>
      <div class="hint">Adds to # of TXs (D2, now ${d.total}) and Extensions added on the ${p.layout === 3 ? BEN_SHEET + " tab" : "tab (W7)"}.</div>${btn("extend", p.tab, "", "Add extension")}</details>
    <details class="group" data-open="pm|${esc(p.tab)}"${state.open.has("pm|" + p.tab) ? " open" : ""}><summary class="gs">Measures over time <span class="sub">${meas.length} recorded</span></summary>
      ${meas.length ? `<table><thead><tr><th>Date</th><th>Visit</th><th class="num">PHQ-9</th><th class="num">GAD-7</th></tr></thead><tbody>${meas.map(r => `<tr><td>${fmtDate(r.date)}</td><td>${esc(r.type)}</td><td class="num">${esc(r.phq)}</td><td class="num">${esc(r.gad)}</td></tr>`).join("")}</tbody></table>` : '<div class="none">No scores yet</div>'}</details>
    <div class="group"><h2>Upcoming appointments<span class="n">${d.sched.length}</span></h2>${up.length ? up.map(r => item(p, `${DAYS[dow(r.date)]} ${fmtDate(r.date)} ${fmtTime(r.time)} <span class="chip ${cardCls(r)}">${esc(apptTypeName(r))}</span>`, "", "", r.row)).join("") : '<div class="none">Nothing booked</div>'}
      <button class="sm" data-act="ptAll">${state.ptAll ? "Hide all visits" : "Show all visits"}</button>
      ${state.ptAll ? `<table style="margin-top:6px"><thead><tr><th>Date</th><th>Tx #</th><th>Type</th><th>Pulses L/R/O</th><th>Status</th></tr></thead><tbody>${all.map(r => `<tr><td>${fmtDate(r.date)}</td><td>${typeof r.tx === "number" ? r.tx : ""}</td><td>${esc(r.type)}</td><td>${r.pulses.map(x => x ?? "").join(" / ")}</td><td>${r.miss ? "Missed" : r.delivered ? "Done" : r.date < t ? "" : "Scheduled"}</td></tr>`).join("")}</tbody></table>` : ""}</div>
    <details class="group" data-open="pa|${esc(p.tab)}"${state.open.has("pa|" + p.tab) ? " open" : ""}><summary class="gs">Availability <span class="sub">${esc(availText(p) || "not entered")}</span></summary>${availEditor(av)}${btn("availSave", p.tab, "", "Save availability")}</details>`;
}
function availEditor(av, prefix = "pv") {
  const wk = [1, 2, 3, 4, 5].map(d => { const w = av.week[d] || {}; return `<div class="g3w"><span>${DAYS[d]}</span><input type="time" data-avw="${d}" data-k="from" value="${minToInput(w.from)}" aria-label="${DAYS[d]} earliest start"><input type="time" data-avw="${d}" data-k="to" value="${minToInput(w.to)}" aria-label="${DAYS[d]} latest start"></div>`; }).join("");
  const rg = av.ranges.map((w, i) => `<div class="g4r"><select data-avr="${i}" data-k="type">${AV_TYPE.map(x => `<option${x === w.type ? " selected" : ""}>${x}</option>`).join("")}</select><input type="date" data-avr="${i}" data-k="from" value="${w.from ? serialToInput(w.from) : ""}"><input type="date" data-avr="${i}" data-k="to" value="${w.to ? serialToInput(w.to) : ""}"><input data-avr="${i}" data-k="note" value="${esc(w.note)}" placeholder="Reason"><button class="sm" data-act="avDel" data-v="${i}" aria-label="Remove">\u00D7</button></div>`).join("");
  return `<div class="sub">Weekly (start-time window)</div><div class="g3w hdr"><span></span><span>Earliest start</span><span>Latest start</span></div>${wk}
    <div class="sub" style="margin-top:6px">Dates (vacation, work, other)</div>${rg || '<div class="none">None entered</div>'}<button class="sm" data-act="avAdd">Add another</button>`;
}
function hhmm(v) { if (typeof v !== "number") return ""; const h = Math.floor(v * 24 + 1e-6), m = Math.round((v * 24 - h) * 60); return `${h}:${String(m).padStart(2, "0")}`; }
function viewProtocols() {
  if (!state.protocols.length) return `<div class="none">No Protocols sheet found.</div>`;
  const rows = state.protocols.map(p => {
    const calc = [p.rep, p.pulses, p.trains, p.iti].every(x => typeof x === "number" && x > 0) ? ((p.pulses / p.rep) + p.iti) * p.trains / 60 : null;
    const mm = calc == null ? "" : `${Math.floor(calc)}:${String(Math.round((calc % 1) * 60)).padStart(2, "0")}`;
    return `<tr><td>${esc(p.name)}</td><td class="num">${p.rep}</td><td class="num">${p.pulses}</td><td class="num">${p.trains}</td><td class="num">${p.iti}</td><td class="num">${p.inten ?? ""}</td><td class="num">${hhmm(p.typed)}</td><td class="num">${mm}</td></tr>`;
  }).join("");
  return `<table><thead><tr><th>Protocol</th><th class="num">Hz</th><th class="num">Pulses</th><th class="num">Trains</th><th class="num">ITI</th><th class="num">% MT</th><th class="num">Time</th><th class="num">Calc</th></tr></thead><tbody>${rows}</tbody></table>
    <div class="hint">Time = Tx Time on the Protocols sheet as entered (min:sec). Calc = ((pulses / Hz) + ITI) x trains / 60, the same formula as each tab's N2. HF 10Hz is the standard name; X is the internal copy.</div>
    <div class="actions"><button data-sheet="Protocols">Open Protocols sheet</button></div>`;
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
  const P = state.patients.filter(p => p.status === "" || AP_SHOW.includes(p.status)), blank = Array(6).fill("");
  const head = ["Patient", "MRN", "Status", ...days.map(d => `${DAYS[dow(d)]} ${fmtDate(d)}${d === t ? " (today)" : ""}`), "Next F/U", "Graduation", "Auth exp", "Tx left", "MD on-site", "Schedule given"];
  const vs = days.map(variances);
  const treated = ["Treated", "", "", ...vs.map(v => v.expected ? `${v.actual} of ${v.expected}` : ""), ...blank];
  const notes = ["Variances", "", "", ...vs.map((v, i) => [state.closures.has(days[i]) ? `Closed: ${state.closures.get(days[i])}` : "", ...Object.entries(v.groups).filter(([, a]) => a.length).map(([k, a]) => `${a.length} ${k.toLowerCase()}`)].filter(Boolean).join(", ")), ...blank];
  const rows = P.map(p => [p.tab, p.mrnRaw, p.status || "(blank)", ...days.map(d => dayCode(p, d)), p.d.fuNext ? p.d.fuNext.date : "", p.d.grad ?? "", p.authExp ?? "",
    p.txRem ?? Math.max(0, p.d.total - p.d.txDone), p.d.mdOnSite ? "Yes" : "", p.ben.lastSched ?? ""]);
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
      const hd = ws.getRange(`A${AP_HEAD}:N${AP_HEAD}`); hd.values = [m.head];
      hd.format.fill.color = "#1c2a33"; hd.format.font.color = "#ffffff"; hd.format.font.bold = true; hd.format.wrapText = true; hd.format.verticalAlignment = "Center"; hd.format.rowHeight = 30;
      const sm = ws.getRange(`A${AP_HEAD + 1}:N${AP_HEAD + 2}`); sm.values = m.sum; sm.format.fill.color = "#f3f6f8"; sm.format.verticalAlignment = "Top";
      ws.getRange(`A${AP_HEAD + 1}:A${AP_HEAD + 2}`).format.font.bold = true;
      ws.getRange(`D${AP_HEAD + 1}:H${AP_HEAD + 1}`).format.font.bold = true;
      const vr = ws.getRange(`D${AP_HEAD + 2}:H${AP_HEAD + 2}`); vr.format.wrapText = true; vr.format.font.size = 9; vr.format.font.color = "#5b6b75";
      ws.getRange(`D${AP_HEAD}:H${end}`).format.horizontalAlignment = "Center";
      if (n) {
        const body = ws.getRange(`A${AP_FIRST}:N${last}`); body.values = m.rows;
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
      [72, 62, 82, 64, 64, 64, 64, 64, 66, 66, 66, 48, 58, 70].forEach((w, i) => ws.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w);
      ws.getRange(`A${AP_HEAD + 2}:N${AP_HEAD + 2}`).format.autofitRows();
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
async function applyOps(ops) {
  if (POPOUT) return callPane("applyOps", [ops]);
  const ben = [], cells = [];
  for (const o of ops) {
    const p = isBenKey(o.a) && benTarget(o.tab);
    if (p) ben.push({ mrn: p.mrn, mrnVal: p.mrnVal, tab: p.tab, key: o.a, v: o.clear ? "" : o.v && o.v[0][0], quiet: !!o.quiet });
    else if (o.a !== BEN_STAMP) cells.push(o);
  }
  await Excel.run(async ctx => {
    for (const o of cells) { const r = ctx.workbook.worksheets.getItem(o.tab).getRange(o.a); if (o.clear) r.clear("Contents"); else r.values = o.v; }
    await writeBen(ctx, ben);
    await ctx.sync();
  });
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
        r.formulasR1C1 = [src[i].formulasR1C1[0].map(f => typeof f === "string" && f.startsWith("=") ? f : "")];
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
    const got = list.map(({ tab, order }) => { const r = ctx.workbook.worksheets.getItem(tab).getRange(`B${LOG_FIRST}:T${LOG_FIRST + order.length - 1}`); r.load("formulas"); return r; });
    await ctx.sync();
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
const fmtCell = c => {
  const f = c.format || {}, b = f.borders || {};
  return { format: { fill: { color: (f.fill || {}).color }, font: f.font, horizontalAlignment: f.horizontalAlignment,
    borders: Object.fromEntries(["top", "bottom", "left", "right"].filter(k => b[k]).map(k => [k, b[k].style === "None" ? { style: "None" } : b[k]])) } };
};
const fillOf = c => String((((c || {}).format || {}).fill || {}).color || "").toUpperCase();
async function formatLogs(list) {
  if (!list.length) return 0;
  if (POPOUT) return callPane("formatLogs", [list]);
  let n = 0;
  await Excel.run(async ctx => {
    const tpl = ctx.workbook.worksheets.getItemOrNullObject("New Template"); await ctx.sync();
    const bl = tpl.isNullObject ? null : tpl.getRange(`A${LOG_FIRST + 1}:T${LOG_LAST}`), bp = bl && bl.getCellProperties(FMT_Q);
    if (bl) bl.load("numberFormat");
    const G = list.map(({ tab, touched }) => {
      const ws = ctx.workbook.worksheets.getItem(tab), a10 = ws.getRange(`A${LOG_FIRST}`), b = ws.getRange(`B${LOG_FIRST}:B${LOG_FIRST + LOG_READ - 1}`);
      a10.load("formulas"); b.load("values"); return { ws, a10, b, touched: new Set(touched || []) };
    });
    await ctx.sync();
    G.forEach(g => { g.end = endOf(g.b.values.map(r => r[0]), g.a10.formulas[0][0]); g.rg = g.ws.getRange(`A${LOG_FIRST}:T${g.end}`); g.rg.load("values,numberFormat"); g.cp = g.rg.getCellProperties(FMT_Q); });
    await ctx.sync();
    for (const g of G) {
      const P = g.cp.value, V = g.rg.values, NF = g.rg.numberFormat, first = P[0].map(fmtCell);
      for (let i = 1; i < V.length; i++) {
        const r = LOG_FIRST + i, dated = isNum(V[i][C.date]); if (!dated && !bp) continue;
        const ti = (r >= g.end - 1 ? LOG_LAST - (g.end - r) : Math.min(r, LOG_LAST - 2)) - LOG_FIRST - 1;   // New Template row for this position, 0 = row 11
        const want = dated ? first : bp.value[ti].map(fmtCell), wantNf = dated ? NF[0] : bl.numberFormat[ti];
        const plain = P[i].every((c, j) => fillOf(c) === fillOf(P[0][j]) || (bp && fillOf(c) === fillOf(bp.value[ti][j])));
        if (!g.touched.has(r) && !plain) continue;
        if (P[i].every((c, j) => JSON.stringify(fmtCell(c)) === JSON.stringify(want[j])) && NF[i].every((f, j) => f === wantNf[j])) continue;
        const row = g.ws.getRange(`A${r}:T${r}`); row.setCellProperties([want]); row.numberFormat = [wantNf]; n++;
      }
    }
    await ctx.sync();
  });
  return n;
}
async function formatLogsClick() {
  try {
    const n = await formatLogs(state.patients.filter(p => !state.stale.includes(p.tab)).map(p => ({ tab: p.tab, touched: [] })));
    toast(n ? `Reformatted ${n} visit-log row${n === 1 ? "" : "s"}. Rows with their own highlight were left as they are.` : "Visit-log formatting already follows the dates.");
  } catch (e) { fail(e); }
}
const writeCells = (tab, pairs) => applyOps(pairs.map(([a, v]) => ({ tab, a, v: [[v]] })));
function patchRow(tab, row, f) { const p = state.patients.find(x => x.tab === tab), r = p && p.rows.find(x => x.row === row); if (!r) return; Object.assign(r, f); p.d = derive(p); renderDash(); }
const findPR = (tab, row) => { const p = state.patients.find(x => x.tab === tab); return [p, p && p.rows.find(x => x.row === row)]; };

/* missed visits: record, then classify the freebie in a pop-up */
async function markMiss(tab, row, kind) {
  const [p, r] = findPR(tab, row); if (!r || !kind) return;
  const note = (r.note ? r.note + "; " : "") + `${kind} ${fmtDate(r.date ?? todaySerial())} (marked in add-in)`;
  try { await writeCells(tab, [[`M${row}`, "x"], [`S${row}`, note]]); patchRow(tab, row, { miss: true, note }); freebieModal(tab, row); } catch (e) { fail(e); }
}
function freebieModal(tab, row) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const kind = OUTCOMES.find(k => r.note.includes(k)) || "Missed visit", left = p.d.left, after = left - 1;
  const warn = left <= 0 ? "This patient has no freebies left." : after === 1 ? "Using a freebie leaves this patient with 1 freebie remaining." : after === 0 ? "Using a freebie leaves this patient with zero remaining freebies." : "";
  showModal(`<h3>${esc(kind)}: ${esc(p.tab)}, ${fmtDateY(r.date)}</h3><p class="sub">Classify this missed visit.</p>${warn ? `<div class="item amb">${esc(warn)}</div>` : ""}
    <div class="actions"><button class="pri" data-mf="Used" data-t="${esc(tab)}" data-row="${row}"${left <= 0 ? " disabled" : ""}>Use freebie</button><button data-mf="Waived" data-t="${esc(tab)}" data-row="${row}">Waived</button><button data-mclose>Decide later</button></div>`);
}
/* the freebies left (N6; N7 on a previous-layout tab): using one takes one off, undoing it gives it back */
async function setFreebie(tab, row, v) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const note = withFreebie(r.note, v), n7 = Math.max(0, p.d.left + (r.freebie === "Used" ? 1 : 0) - (v === "Used" ? 1 : 0));
  try {
    await writeCells(tab, [[`S${row}`, note], [p.fbCell, n7]]); closeModal(); p.freebiesLeft = n7; patchRow(tab, row, { freebie: v, note });
    const left = p.d.left;
    toast(v === "Used" && left <= 1 ? (left === 1 ? "This patient has 1 freebie remaining after marking today's visit." : "This patient has zero remaining freebies.") : v === "Used" ? "Freebie used." : "Marked waived.", v === "Used" && left === 0);
  } catch (e) { fail(e); }
}
async function undoMiss(tab, row) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const note = withFreebie(r.note.split("; ").filter(x => !/\(marked in add-in\)$/.test(x)).join("; "), "");
  const pairs = [[`M${row}`, ""], [`S${row}`, note]]; if (r.freebie === "Used") pairs.push([p.fbCell, p.d.left + 1]);
  try { await writeCells(tab, pairs); if (r.freebie === "Used") p.freebiesLeft = p.d.left + 1; patchRow(tab, row, { miss: false, freebie: "", note }); toast("Undone."); } catch (e) { fail(e); }
}
/* pulses delivered = treatment complete (F, H, J). Treating technician goes in B. */
async function markPulses(tab, row, el) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const box = el.closest("details"), vals = [...box.querySelectorAll("input[data-pul]")].map(i => i.value === "" ? "" : +i.value), tech = box.querySelector("input[data-tech]").value.trim();
  if (!vals.some(v => v !== "" && v > 0)) return toast("Enter the pulses delivered.", true);
  if (!tech) return toast("Enter the treating technician.", true);
  const pairs = [["F", vals[0]], ["H", vals[1]], ["J", vals[2]]].filter(([, v]) => v !== "" && v > 0).map(([c, v]) => [`${c}${row}`, v]);
  pairs.push([`B${row}`, tech]); state.techName = tech;
  try { await writeCells(tab, pairs); patchRow(tab, row, { delivered: true, tech, pulses: [vals[0] || null, vals[1] || null, vals[2] || null] }); toast("Treatment saved."); } catch (e) { fail(e); }
}
async function saveMeasures(tab, row, el) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const vals = {}; el.closest(".meas").querySelectorAll("input[data-m]").forEach(i => { if (i.value !== "") vals[i.dataset.m] = +i.value; });
  if (vals.P != null && (vals.P < 0 || vals.P > 27)) return toast("PHQ-9 must be 0 to 27.", true);
  if (vals.Q != null && (vals.Q < 0 || vals.Q > 21)) return toast("GAD-7 must be 0 to 21.", true);
  const pairs = Object.entries(vals).map(([c, n]) => [`${c}${row}`, n]); if (!pairs.length) return toast("Enter a score first.", true);
  try {
    await writeCells(tab, pairs);
    const phq = vals.P != null ? String(vals.P) : r.phq, gad = vals.Q != null ? String(vals.Q) : r.gad;
    patchRow(tab, row, { phq, gad, measX: phq.toLowerCase() === "x" || gad.toLowerCase() === "x", measVal: true }); toast("Measures saved.");
  } catch (e) { fail(e); }
}
async function moveMeasures(tab, row, el) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const dest = p.rows.find(x => x.row === +el.closest(".meas").querySelector("select[data-mvto]").value); if (!dest) return;
  const note = (r.note ? r.note + "; " : "") + `Measures not collected ${fmtDate(r.date)}; moved to ${fmtDate(dest.date)}`;
  const pairs = [[`S${row}`, note]];
  if (r.phq.toLowerCase() === "x") pairs.push([`P${row}`, ""]); if (r.gad.toLowerCase() === "x") pairs.push([`Q${row}`, ""]);
  if (!dest.phq) pairs.push([`P${dest.row}`, "x"]); if (!dest.gad) pairs.push([`Q${dest.row}`, "x"]);
  try { await writeCells(tab, pairs); state.open.delete(`meas|${tab}|${row}`); toast(`Measures moved to ${fmtDate(dest.date)}.`); await refresh(); } catch (e) { fail(e); }
}
async function saveMT(tab, row, el) {
  const [p, r] = findPR(tab, row); if (!r) return;
  const parts = [...el.closest("details").querySelectorAll("input[data-mt]")].filter(i => i.value.trim()).map(i => `${i.dataset.mt}: ${i.value.trim()}`);
  if (!parts.length) return toast("Enter at least one MT detail.", true);
  const keep = r.note.split("; ").filter(x => !/^(hotspot|mt %|amps|baseline sleep|baseline caffeine):/i.test(x));
  const note = [...keep.filter(Boolean), ...parts].join("; ");
  try { await writeCells(tab, [[`S${row}`, note]]); patchRow(tab, row, { note }); toast("MT details saved to Notes."); } catch (e) { fail(e); }
}
/* protocol rows 2-4: name from the Protocols sheet fills J:M (N calculates), site label in O */
async function saveProtocol(tab, el) {
  const box = el.closest(".pe"), pairs = [];
  for (const n of [1, 2, 3]) {
    const name = box.querySelector(`select[data-pname="${n}"]`).value, site = box.querySelector(`select[data-psite="${n}"]`).value, row = n + 1;
    const q = state.protocols.find(x => x.name === name);
    if (q) pairs.push([`J${row}`, q.rep], [`K${row}`, q.pulses], [`L${row}`, q.trains], [`M${row}`, q.iti]);
    else if (!name) pairs.push([`J${row}`, ""], [`K${row}`, ""], [`L${row}`, ""], [`M${row}`, ""]);
    pairs.push([`O${row}`, q || site ? site : ""]);
  }
  try { await writeCells(tab, pairs); toast("Protocol saved."); await refresh(); } catch (e) { fail(e); }
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
async function noteAction(a, tab, row, el) {
  const p = state.patients.find(x => x.tab === tab), r = p && p.rows.find(x => x.row === row); if (!r) return;
  const dr = draftOf(p, r);
  if (a === "resetNote") { dr.edited = false; dr.text = ""; return renderDash(); }
  if (!dr.edited) {
    const missing = NOTE_QS.filter(q => q.flag && dr.a[q.k] === "Yes" && !(dr.d[q.k] || "").trim()).map(q => q.flag);
    if (missing.length) return toast(`Specify details for: ${missing.join(", ")}.`, true);
    if (!dr.initials.trim()) return toast("Enter your initials.", true);
  }
  const text = dr.edited ? dr.text : noteText(p, r, dr), ta = el.closest("details").querySelector(".ntext");
  const ok = await copyText(text, ta);
  if (!ok) return toast("Copy was blocked. The note is selected: press Ctrl+C, then tick Treatment report complete.", true);
  state.initials = dr.initials.trim();
  dr.open = false; renderDash(); toast("Note copied.");
}
function draftInput(el) {
  const det = el.closest("details[data-draft]"); if (!det) return false;
  const dr = state.drafts[det.dataset.draft]; if (!dr) return false;
  const dk = el.dataset.dk, v = el.value;
  if (dk === "text") { dr.text = v; dr.edited = true; return true; }
  if (dk.startsWith("a.")) dr.a[dk.slice(2)] = v; else if (dk.startsWith("d.")) dr.d[dk.slice(2)] = v; else dr[dk] = v;
  if (!dr.edited) { const [tab, row] = dr.key.split("|"), p = state.patients.find(x => x.tab === tab), r = p.rows.find(x => x.row === +row); det.querySelector(".ntext").value = noteText(p, r, dr); }
  return true;
}

/* ---------- measures placeholders (x in PHQ-9 and GAD-7) ----------
   Always on MT, MTR, F/U and the last treatment. Between them, every 5 treatments when the gap is 8 or more,
   and at least every 7 calendar days when visits are sparse (tapering), but never within 2 days or 2 visits before a milestone.
   Only blank cells on undelivered, non-missed rows are filled. Existing x's and scores never move. As on New Template, a Daily
   (or Taper) row with measures also carries the Measures action flag in O; MT, MTR and F/U rows only get the x's. */
function planMeasures(p) {
  const total = p.txApproved || 36, rows = p.rows.filter(r => r.type || typeof r.tx === "number");
  const isMs = r => STD_VISIT.includes(r.type) || (typeof r.tx === "number" && r.tx === total);
  const hasMeas = r => r.measX || r.measVal;
  const target = new Set(rows.filter(isMs).map(r => r.row)), cadence = new Set();
  let lastIdx = -1, lastDate = null;
  rows.forEach((r, i) => {
    if (isMs(r) || hasMeas(r)) { lastIdx = i; if (r.date != null) lastDate = r.date; return; }
    const ni = rows.findIndex((x, j) => j > i && isMs(x)), nxt = ni >= 0 ? rows[ni] : null;
    const since = i - lastIdx, toNext = ni >= 0 ? ni - i : Infinity, gap = since + toNext - 1;
    const byCount = gap >= 8 && since >= 5 && toNext > 2;
    const byDays = lastDate != null && r.date != null && r.date - lastDate >= 7 && !(nxt && nxt.date != null && nxt.date - r.date <= 2);
    if (byCount || byDays) { cadence.add(r.row); lastIdx = i; if (r.date != null) lastDate = r.date; }
  });
  const ops = [];
  for (const r of p.rows) {
    if (r.delivered || r.miss) continue;
    const want = target.has(r.row) || cadence.has(r.row), daily = r.type === "Daily" || r.type === "Taper";
    if (want && !r.phq) ops.push([`P${r.row}`, "x"]);
    if (want && !r.gad) ops.push([`Q${r.row}`, "x"]);
    if (daily && (want || r.measX) && !r.action) ops.push([`O${r.row}`, "Measures"]);
  }
  return ops;
}
async function placeMeasures(patients) {
  const plans = patients.map(p => ({ tab: p.tab, ops: planMeasures(p) })).filter(x => x.ops.length);
  if (!plans.length) return 0;
  await applyOps(plans.flatMap(x => x.ops.map(([a, v]) => ({ tab: x.tab, a, v: [[v]] }))));
  return plans.reduce((n, x) => n + x.ops.filter(o => o[0][0] !== "O").length, 0);
}
async function placeMeasuresClick() {
  try { const n = await placeMeasures(state.patients.filter(p => p.d.active)); toast(n ? `Placed ${n} measure placeholder${n > 1 ? "s" : ""}.` : "Measure placeholders are already in place."); await refresh(); } catch (e) { fail(e); }
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
  /* 3. visit log: the PR formula over the whole log, and the template's formulas on rows that lost them */
  const F = lg.formulas;
  ws.getRange(`A${LOG_FIRST}:A${end}`).formulas = F.slice(0, end - LOG_FIRST + 1).map((v, i) => [isPrFormula(v[0]) ? prF(LOG_FIRST + i, end) : v[0]]);
  for (let i = 1; i <= end - LOG_FIRST; i++) for (const [c, fx] of Object.entries(ROW_F)) if (F[i][colIndex(c)] === "") ws.getRange(`${c}${LOG_FIRST + i}`).formulas = [[fx(LOG_FIRST + i)]];
  /* 4. the TX SUMMARY row: New Template's formats (merge, checkbox), FALSE in A until it's ticked */
  if (!isTpl) ws.getRange(`A${s}:T${s}`).copyFrom(tpl.getRange(`A${TPL_SUM}:T${TPL_SUM}`), Excel.RangeCopyType.formats);
  const sv = lg.values[s - LOG_FIRST];
  if (typeof sv[0] !== "boolean") ws.getRange(`A${s}`).values = [[false]];
  if (!SUM_RE.test(String(sv[1] ?? ""))) ws.getRange(`B${s}`).values = [[SUM_LABEL]];
  /* 5. colors by visit type and the dropdowns, over every log row */
  applyRules(ws, end, s);
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
  const bad = [...wrote].filter(([a, v]) => { const m = a.match(/^([A-Z]+)(\d+)$/), c = colIndex(m[1]); return !same(v, back.values[+m[2] - 1][c]); }).map(([a]) => a);
  if (!bad.length) { old.delete(); await ctx.sync(); }
  return { name, deleted: !bad.length, bad };
}
/* rebuilding deletes the earlier tabs, so it asks first. force: every active tracker, not only the ones that look off */
function setupClick(force) {
  if (!state.templateOk) return toast("New Template must have Time in T8 and Freebie Count in J6 before tabs can match it.", true);
  if (!state.old.length) return setupColumns(force === true);
  const n = state.old.length;
  showModal(`<h3>Match New Template</h3><p class="sub">${n} tab${n > 1 ? "s" : ""} (${state.old.map(esc).join(", ")}) will be rebuilt from New Template. Each earlier tab is deleted once every copied entry checks out on the new tab. If anything doesn't match, the earlier tab is kept as "(old)".</p>
    <p class="sub">To get a deleted tab back, use the workbook's version history (File > Info > Previous versions).</p>
    <div class="actions"><button class="pri" data-act="setupGo">Rebuild and delete</button><button data-mclose>Cancel</button></div>`);
}
async function setupColumns(force) {
  if (POPOUT) return callPane("setupColumns", [force]);
  try {
    if (!state.templateOk) return toast("New Template must have Time in T8 and Freebie Count in J6 before tabs can match it.", true);
    const stale = new Set(state.stale), courseOf = n => (state.patients.find(p => p.tab === n) || { course: 0 }).course;
    /* oldest course first, so the newest course of an MRN writes its Benefits row last */
    const names = (force ? state.patients.filter(p => !stale.has(p.tab)).map(p => p.tab) : [...state.fix.keys()].filter(n => n !== "New Template")).sort((a, b) => courseOf(a) - courseOf(b));
    const res = [], matched = [];
    await Excel.run(async ctx => {
      await benSheet(ctx);   // first: a PR formula written before the Benefits tab exists would turn into #REF!
      if (force || state.fix.has("New Template")) await matchTab(ctx, "New Template");
      for (const n of state.old) res.push(await rebuildTab(ctx, n));
      for (const n of names) matched.push((await matchTab(ctx, n)).name);
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
      ? `${parts.join(", ") || "Tabs updated"}. Kept ${kept.map(x => `${x.name} (old)`).join(", ")}: ${kept.map(x => x.bad.slice(0, 3).join(", ")).join("; ")} didn't copy over. Check those cells, then delete the "(old)" tab.`
      : `${parts.join(", ") || "Tabs already match New Template"}.`, kept.length > 0);
    await refresh();
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
  $("apHome").value = "SRL"; $("apTxs").value = 36; $("apRedos").value = 0; $("apCourse").value = 1; $("apNote").textContent = ""; state.apRanges = [{}]; renderApRanges();
  state.codeEdited = state.courseEdited = false; updateApTab();
}
/* copies New Template and fills the header, benefits, cost by code and availability in the template's own cells */
async function addPatient(d) {
  if (POPOUT) { const ok = await callPane("addPatient", [d]); if (ok) { resetAddPatient(); if (state.file) { show("imp"); await buildPlan(); } } return ok; }
  const code = d.code, course = +d.apCourse || 1, name = course > 1 ? `${code} ${course}` : code;
  if (!code) return toast("Enter a name or a LASFIR code.", true);
  if (!state.templateOk) return toast("New Template does not match the current layout, so a new tab cannot be created from it.", true);
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
      const ws = tpl.copy(Excel.WorksheetPositionType.end); ws.name = name;
      const set = (a, val) => { if (val !== "" && val != null) ws.getRange(a).values = [[val]]; };
      set("A2", code); set("A4", d.apMrn); set("D2", +d.apTxs || 36); set("D4", d.apRedos === "" ? "" : +d.apRedos); set("D6", d.apMd); set("E6", d.apIns); set("N7", d.apAssess);
      set("A6", inputToSerial(d.apDob)); set("F4", inputToSerial(d.apAuthStart)); set("G4", inputToSerial(d.apAuthExp));
      ws.getRange("D7").values = [["Pending Start"]];
      if (d.apHome) ws.getRange("S9").values = [[`Home Clinic: ${d.apHome}`]];
      /* benefits and cost by code go on the patient's Benefits row */
      const mrn = normMrn(d.apMrn), bv = { W1: bidf, W4: ded, W5: oop, W6: d.apMtr, W7: 0 };
      d.cost.forEach(([x, y], i) => { bv[`X${COST_FIRST + i}`] = x === "" ? "" : +x; bv[`Y${COST_FIRST + i}`] = y === "" ? "" : +y; });
      if (mrn) await writeBen(ctx, Object.entries(bv).filter(([, v]) => v !== "" && v != null).map(([key, v]) => ({ mrn, mrnVal: d.apMrn, tab: name, key, v })));
      d.week.forEach(([f, t], i) => { if (f || t) ws.getRange(`W${WK_FIRST + i}:X${WK_FIRST + i}`).values = [[tf(f), tf(t)]]; });
      const rg = d.ranges.map(w => ["Unavailable", inputToSerial(w.from), inputToSerial(w.to) || inputToSerial(w.from), w.note || ""]).slice(0, AV_LAST - AV_FIRST + 1);
      if (rg.length) ws.getRange(`V${AV_FIRST}:Y${AV_FIRST + rg.length - 1}`).values = rg;
      ws.activate(); await ctx.sync();
    });
    state.pendingNew = state.pendingNew.filter(x => x.mrn !== normMrn(d.apMrn));
    if (!POPOUT) resetAddPatient();
    await refresh();
    const np = state.patients.find(x => x.tab === name), meas = !!np && planMeasures(np).length > 0; if (meas) await placeMeasures([np]);
    toast(`Created ${name}.`);
    // buildPlan() refreshes on its own, so refresh here only when measures were written and no preview rebuild follows
    if (state.file && state.file.rows.some(r => normMrn(r[state.file.map.mMrn]) === normMrn(d.apMrn))) { show("imp"); await buildPlan(); toast(`Created ${name}. The preview now includes its dates.`); }
    else if (meas) await refresh();
    return true;
  } catch (e) { fail(e); }
}
function startTracker(mrn) {
  const x = state.pendingNew.find(y => y.mrn === mrn); if (!x) return;
  ["apLast", "apFirst", "apDob", "apIns", "apAuthStart", "apAuthExp", "apAssess"].forEach(i => $(i).value = "");
  $("apMrn").value = x.raw; $("apCode").value = x.code; state.codeEdited = !!x.code; state.courseEdited = false;
  $("apNote").textContent = `From NextGen: ${x.n} visit${x.n > 1 ? "s" : ""} today or later, first ${fmtDateY(x.first.date)} (${x.first.type}). Enter coverage and billing, then create the tab. The import preview reruns afterward to fill its dates.`;
  updateApTab(); show("add");
}

/* ---------- NextGen import: column layout is fixed by the "1 Appt List - ALL TMS" report, so it is detected, not mapped ---------- */
const MAP_RE = { mMrn: /^mrn$/i, mDate: /appt\s*date|appointment\s*date/i, mTime: /^time$/i, mSts: /^sts$|^status$/i, mEvent: /^event$|appointment type/i, mProv: /^provider$/i, mChg: /charged/i, mMod: /^mod\s*dt|modified\s*date/i, mLoc: /^location$/i };
async function onFile(e) {
  const f = e.target.files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    state.file = parseReport(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" }), f.name);
    await buildPlan();
  } catch (err) { fail(err); }
}
function parseReport(all, name, saved) {
  const hi = all.findIndex(r => r.filter(c => String(c).trim() !== "").length >= 5);
  if (hi < 0) throw new Error("Could not find the report's header row.");
  let asOf = null, asOfMin = null;
  for (const r of all.slice(0, hi)) for (const c of r) { const m = String(c).match(/(\d{1,2}\/\d{1,2}\/\d{2,4})\s+(\d{1,2}:\d{2}\s*[AaPp][Mm])/); if (m) { asOf = toSerial(m[1]); asOfMin = toMinutes(m[2]); } }
  const headers = all[hi].map(h => String(h).trim());
  const map = Object.fromEntries(Object.entries(MAP_RE).map(([k, re]) => [k, headers.findIndex(h => re.test(h))]));
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
    const u = ws.getUsedRange(true); u.load("values"); await ctx.sync(); vals = u.values;
  });
  if (!vals || vals.length <= LIST_HEAD) return null;
  const name = String((vals.find(r => String(r[0]).trim() === "File") || [])[1] || LIST_SHEET);
  return parseReport(vals, name, true);
}
function planTab(p, appts, from, to, o) {
  const R = p.rows.map(r => ({ ...r, _o: r }));
  const st = { cleared: 0, miss: 0, none: 0, unrecorded: 0, exempt: 0, unclass: 0, types: 0, sameCancel: 0, early: 0, earlyNew: 0, noMod: 0, elsewhere: 0, unmissed: 0, dup: 0 };
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
    let r = R.find(x => x.date === a.date && !x.delivered && (o.ps || !x.alt));
    if (!r) { r = R.find(x => !x.delivered && !x.alt && !x.miss && x.date == null && x.type !== "F/U"); if (r) { r.date = a.date; r.time = a.time; } }
    if (!r) { st.unrecorded++; continue; }
    if (!r.miss) { r.miss = true; st.miss++; if (kind === "Same-day cancel") st.sameCancel++; }
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
  const fill = rows => {
    const f = { cleared: 0, types: 0, elsewhere: 0, past: 0 }, left = [];
    for (const pool of strict ? ["Daily", "MT", "MTR", "F/U"] : ["fu", "tx"]) {
      const free = rows.filter(r => poolOf(r) === pool && (inWin(r) || over(r)));
      const done = rows.filter(r => poolOf(r) === pool && r.delivered && r.date != null && r.date >= from);   // a visit already treated (pulses in) answers its own event, so it never takes a second row that day
      for (const e of events.filter(e => evPool(e) === pool)) {
        const k = done.findIndex(r => r.date === e.date); if (k >= 0) { done.splice(k, 1); continue; }
        if (e.date < from && rows.some(r => poolOf(r) === pool && r.date === e.date)) continue;
        const r = free.shift(); if (!r) { left.push(e); continue; }
        if (e.date < from) f.past++;
        r.date = e.date; r.time = e.time; r.miss = false; r.freebie = "";
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
  Object.assign(st, res.f); st.none = res.left.length;
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
  let n = 0; const txNo = order.map(i => isTxRow(rows[i]) ? ++n : "");
  const txWas = r => r._o && isNum(r._o.tx) ? r._o.tx : r._o && String(r._o.tx ?? "").trim() !== "" ? String(r._o.tx) : "";
  const resorted = order.some((s, i) => s !== i), renum = rows.filter((r, i) => txWas(r) !== txNo[i]).length;
  const arrange = resorted || renum ? { order, tx: txNo } : null;
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
  // from the ALL TMS list only: each visit's chair (the resource booked), and per patient the cancellations, whether anything is still booked, the last kept visit
  if (!ps) {
    const v = {}, pt = {}, c = {}, names = new Set(), cnt = {}, ap = [], apKey = new Set();
    for (const r of state.file.rows) {
      const mrn = normMrn(cell(r, "mMrn")), pv = String(cell(r, "mProv") || "").trim();
      if (!mrn || isOther(mrn) || dev.other.test(pv)) continue;
      const date = toSerial(cell(r, "mDate")), sts = mapSts(cell(r, "mSts")), ev = mapEvent(cell(r, "mEvent"));
      if (date == null || sts === "Skip" || !ev) continue;
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
        if (dev.ours.test(pv)) { v[`${mrn}|${date}`] = pv; names.add(pv); const k = cnt[mrn] || (cnt[mrn] = {}); k[pv] = (k[pv] || 0) + 1; }
      }
    }
    for (const x of Object.values(c)) { x.cx = [...new Set(x.cx)].filter(d => !x.on.includes(d)); delete x.on; }   // rebooked the same day: not a cancellation
    for (const [mrn, k] of Object.entries(cnt)) pt[mrn] = Object.entries(k).sort((a, b) => b[1] - a[1])[0][0];
    const ds = ap.map(a => a.date);
    state.lx = { v, pt, c, chairs: [...names].sort(), ap, lo: ds.length ? Math.min(...ds) : null, hi: ds.length ? Math.max(...ds) : null };
    state.patients.forEach(p => p.d = derive(p));
  }
  const by = new Map(), codes = new Map(), rawMrn = new Map(), clin = new Map(), keptPast = new Set(); let maxDate = from, ignored = 0;
  for (const r of state.file.rows) {
    const mrn = normMrn(cell(r, "mMrn")), pv = String(cell(r, "mProv") || "").trim();
    if (!mrn) continue;
    if (isOther(mrn) || dev.other.test(pv)) { ignored++; continue; }
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
    by.get(mrn).push({ date, time, status: sts, type: ev.type, tele: ev.tele, charged, mod: m.mMod >= 0 ? toSerial(cell(r, "mMod")) : null, loc: String(cell(r, "mLoc") || "").trim() });
    if (pv && !past && !/san rafael|magv|brainsway|mag\s?venture/i.test(pv)) { const c = clin.get(mrn) || {}; c[pv] = (c[pv] || 0) + 1; clin.set(mrn, c); }
    const c = String(cell(r, "mPat") ?? "").trim().toUpperCase(); if (/^[A-Z]{6}$/.test(c)) codes.set(mrn, c);
  }
  const latest = new Map(), latestCode = new Map();
  for (const p of state.patients) {
    if (p.mrn) { const c = latest.get(p.mrn); if (!c || p.course > c.course) latest.set(p.mrn, p); }
    const cc = latestCode.get(p.code); if (!cc || p.course > cc.course) latestCode.set(p.code, p);
  }
  const o = { lastAsOf: state.report && state.report.asOf != null ? state.report.asOf : null, today: from, ps, keptPast, ours: loc => !loc || loc.toLowerCase().includes(state.cfg.location.toLowerCase()) };
  const plans = [], disc = [];
  for (const [mrn, appts] of by) {
    let p = latest.get(mrn); const code = codes.get(mrn);
    if (!p && code) { const q = latestCode.get(code); if (q && !q.mrn) p = q; }
    if (!p) continue;
    // a Discontinued tab is left as it is; the preview and the Day view's list check name any visits NextGen still has booked
    if (!onSched(p)) { if (appts.some(a => a.status === "Scheduled" && a.date >= from)) disc.push(p.tab); continue; }
    const pl = planTab(p, appts, from, maxDate, { ...o, mrn, grow: p.canGrow });
    // the report is the source of truth for the provider name, so E2 follows it even when something was typed there
    const c = clin.get(mrn), top = c && Object.entries(c).sort((a, b) => b[1] - a[1])[0][0]; if (top && top !== p.provider) pl.e2 = top;
    if (pl.e2) pl.changed = true;
    plans.push(pl);
  }
  const noTab = [...by.entries()].filter(([k]) => !latest.has(k) && !(codes.get(k) && latestCode.has(codes.get(k)) && !latestCode.get(codes.get(k)).mrn));
  state.pendingNew = noTab.map(([k, a]) => { const fut = a.filter(x => x.status !== "No-show" && x.status !== "Cancelled" && x.date >= from).sort(cmp); return { mrn: k, raw: rawMrn.get(k) || k, last4: k.slice(-4), code: codes.get(k) || "", n: fut.length, first: fut[0] }; }).filter(x => x.n > 0);
  state.plan = { ps, plans, from, to: maxDate, ignored, disc };
  if (!quiet) renderPlan(); renderBanners(); renderSavedList();
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
  if (state.pendingNew.length) h += `<div class="sub" style="margin-top:6px">${state.pendingNew.length} patient${state.pendingNew.length > 1 ? "s" : ""} in the report ${state.pendingNew.length > 1 ? "have" : "has"} no tracker yet. See the top of the dashboard.</div>`;
  h += `<div class="actions"><button class="pri" data-act="apply">Apply</button></div>`;
  $("planBody").innerHTML = h; $("planBox").hidden = false;
}
async function applyPlan() {
  const plan = state.plan, ch = plan.plans.filter(p => p.changed), reflow = new Set(plan.plans.filter(p => p.reflow).map(p => p.name));
  try {
    await growLogs(ch.filter(p => p.ins.length).map(p => ({ tab: p.name, ins: p.ins })));   // rows first: the ops address the grown log
    await applyOps(ch.flatMap(p => [...p.ops, ...(p.e2 ? [["E2", p.e2]] : [])].map(([a, v]) => ({ tab: p.name, a, v: [[v]] }))));
    await arrangeLogs(ch.filter(p => p.arrange).map(p => ({ tab: p.name, ...p.arrange })));   // it moves the rows the ops just wrote
    await formatLogs(ch.map(p => ({ tab: p.name, touched: p.touched })));                       // last: formatting follows the dates where they ended up
    const f = state.file;
    if (!plan.ps && !f.saved) { const rep = { asOf: f.asOf, asOfMin: f.asOfMin, importedOn: todaySerial() }; state.report = rep; await saveSetting("tms_report", rep); await saveList(f); }
    state.plan = null; $("planBox").hidden = true; $("file").value = ""; state.file = plan.ps ? null : { ...f, saved: true };
    await refresh();
    const sug = state.patients.filter(p => p.d.suggest);
    if (sug.length) await applyOps(sug.map(p => ({ tab: p.tab, a: "D7", v: [[p.d.suggest]] })));
    const nx = reflow.size ? await placeMeasures(state.patients.filter(p => reflow.has(p.tab))) : 0;
    toast(`Updated ${ch.length} tab${ch.length === 1 ? "" : "s"}${sug.length ? `, ${sug.length} status${sug.length > 1 ? "es" : ""}` : ""}${nx ? `, measures re-planned for ${reflow.size}` : ""}.`);
    if (sug.length || reflow.size) await refresh();   // the refresh above is current unless statuses or measures were just written
    show("sch");
  } catch (e) { fail(e); }
}

/* ---------- patient schedule: opens in a dialog window (schedule.html) to print or save as PDF ---------- */
const pad2 = n => String(n).padStart(2, "0");
const apptTypeName = r => r.type === "MT" ? (device() === "Brainsway" ? "MT BW" : "MT MagV") : r.type === "MTR" ? "MT Redo" : r.type === "F/U" ? (r.tele ? "Telepsych TMS F/U" : "TMS Follow Up") : "TMS Treatment";
function scheduleRows(p, from, to) {
  const t = todaySerial(), loc = state.cfg.location, lo = Math.max(from ?? t, t), hi = to ?? Infinity;
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
  showModal(`<h3>Schedule for ${esc(p.tab)}</h3><p class="sub">Choose the visits to include. The schedule opens in a window where you can print it or save it as a PDF.</p>
    <div class="fchips" style="margin:8px 0">${quick.map(q => `<button class="sm${q === sel ? " on" : ""}" data-prq="${q[2]}">${q[1]}</button>`).join("")}</div>
    <div class="row2"><label>From<input type="date" id="prFrom" value="${serialToInput(t)}"></label><label>To<input type="date" id="prTo" value="${serialToInput(sel[2])}"></label></div>
    <div class="sub" id="prCount"></div>
    <div class="actions"><button class="pri" data-act="doPrint" data-t="${esc(tab)}">Open schedule</button><button data-mclose>Cancel</button></div>`);
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
function show(view) {
  state.view = view;
  if (view !== "settings") state.lastView = ["sch", "pts", "clinic"].includes(view) ? view : state.lastView;
  ["sch", "pts", "clinic", "imp", "add", "settings"].forEach(v => $("v-" + v).hidden = v !== view);
  $("btnSettings").classList.toggle("on", view === "settings");
  if (["sch", "pts", "clinic"].includes(view)) renderDash(); else document.querySelectorAll("#mainTabs button").forEach(b => b.classList.toggle("on", view === "add" && b.dataset.view === "pts"));
}
/* ---------- pop-out window: full screen or a separate window, driven by the task pane ---------- */
const CHUNK = 20000;
function toChild(m) { if (pop.dlg) try { pop.dlg.messageChild(JSON.stringify(m)); } catch { /* window closed */ } }
function sendSnap() {
  const s = JSON.stringify({ cfg: state.cfg, patients: state.patients.map(({ d, ...p }) => p), tabs: state.tabs, protocols: state.protocols, fix: [...state.fix], old: state.old, stale: state.stale,
    skipped: state.skipped, templateOk: state.templateOk, pendingNew: state.pendingNew, closures: [...state.closures], wbName: state.wbName, clinic: state.clinic, report: state.report, lx: state.lx });
  const n = Math.ceil(s.length / CHUNK) || 1, seq = Date.now();
  for (let i = 0; i < n; i++) toChild({ k: "snap", seq, i, n, part: s.slice(i * CHUNK, (i + 1) * CHUNK) });
}
function applySnap(o) {
  Object.assign(state, { cfg: o.cfg, patients: o.patients, tabs: o.tabs, protocols: o.protocols, fix: new Map(o.fix), old: o.old, stale: o.stale || [], skipped: o.skipped, templateOk: o.templateOk,
    pendingNew: o.pendingNew, closures: new Map(o.closures), wbName: o.wbName, clinic: o.clinic, report: o.report, lx: o.lx || null });
  state.patients.forEach(p => p.d = derive(p));
  syncCfgInputs(); renderBanners(); renderDash(); updateApTab(); fillAddPatientForm();
  rpc.snapWait.splice(0).forEach(f => f());
}
function callPane(fn, args) {
  return new Promise((res, rej) => {
    const id = ++rpc.id; rpc.wait.set(id, { res, rej });
    try { Office.context.ui.messageParent(JSON.stringify({ k: "call", id, fn, args })); } catch (e) { rpc.wait.delete(id); rej(e); }
  });
}
/* what a pop-out may ask the pane to do: reads and writes need the workbook, settings need the document */
const PANE_FNS = {
  refresh: () => refresh(), applyOps: ops => applyOps(ops), growLogs: x => growLogs(x), arrangeLogs: x => arrangeLogs(x), formatLogs: x => formatLogs(x), writeList: x => writeList(x), rebuildActive: () => rebuildActiveClick(), mkSchedSheet: () => mkSchedSheet(),
  writeChairSheet: o => writeChairSheet(o), setupColumns: force => setupColumns(force), addPatient: d => addPatient(d), openTabName: (t, r) => openTabName(t, r),
  saveSetting: (k, v) => saveSetting(k, v)
};
async function onPopMessage(arg) {
  let m; try { m = JSON.parse(arg.message); } catch { return; }
  if (m.k === "ready") return sendSnap();
  if (m.k === "close") { closePopout(); return; }
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
  Office.context.ui.displayDialogAsync(url.href, { height: 95, width: 95, displayInIframe: true }, res => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) return toast(res.error.code === 12007 ? "Close the other add-in window first." : `Full screen couldn't open (${res.error.message}). Try again.`, true);
    pop.dlg = res.value; $("popNote").hidden = false;
    pop.dlg.addEventHandler(Office.EventType.DialogMessageReceived, onPopMessage);
    pop.dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { pop.dlg = null; $("popNote").hidden = true; });
  });
}
function closePopout() { if (pop.dlg) { try { pop.dlg.close(); } catch { /* already closed */ } pop.dlg = null; } $("popNote").hidden = true; }
function initPopout() {
  document.body.classList.add("popout");
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, a => {
    let m; try { m = JSON.parse(a.message); } catch { return; }
    if (m.k === "snap") {
      if (!rpc.parts.length || rpc.parts.seq !== m.seq) { rpc.parts = Array(m.n); rpc.parts.seq = m.seq; }
      rpc.parts[m.i] = m.part;
      if (rpc.parts.filter(x => x != null).length === m.n) { const s = rpc.parts.join(""); rpc.parts = []; applySnap(JSON.parse(s)); }
    } else if (m.k === "ack") {
      const w = rpc.wait.get(m.id); if (!w) return; rpc.wait.delete(m.id);
      if (m.ok) w.res(m.result); else w.rej(new Error(m.err || "The task pane could not do that."));
    } else if (m.k === "toast") toast(m.msg, m.err);
  }, () => Office.context.ui.messageParent(JSON.stringify({ k: "ready" })));
}
/* Full screen: the tracker floats inside the Excel page over the workbook; the menu sits on frosted glass over a blurred page (body.fs).
   In full screen the same button closes it. */
function fullScreen() {
  if (POPOUT) { try { Office.context.ui.messageParent(JSON.stringify({ k: "close" })); } catch { /* pane gone */ } return; }
  openPopout();
}
const FULL = POPOUT;
function syncFs() {
  const b = $("btnFull");
  document.body.classList.toggle("fs", FULL);
  const tip = FULL ? "Exit full screen" : "Full screen"; b.title = tip; b.setAttribute("aria-label", tip);
  syncAtt();
}
/* the ⋯ menus blur the add-in behind them; the shade also catches the click that closes them */
function syncShade() { $("menuShade").hidden = !document.querySelector("details.kebab[open]"); }
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
function openPatient(tab, sec) { state.ptTab = tab; state.ptMode = "pt"; state.avDraft = null; if (sec) state.open.add(`${sec}|${tab}`); show("pts"); window.scrollTo(0, 0); }
/* a patient name on today's schedule also opens their Excel tab at that visit, unless turned off in Settings */
function openPtLink(el) { if (el.dataset.sched && state.cfg.openTab !== "off") openTabName(el.dataset.pt, +el.dataset.sched); openPatient(el.dataset.pt, el.dataset.sec); }
const CFG_IDS = { cfgLoc: "location", cfgDevice: "device", cfgAuth: "authDays", cfgProto: "protoDays", cfgRecent: "recentDays", cfgOpenTab: "openTab", cfgCutoff: "cutoff" };
function syncCfgInputs() { Object.entries(CFG_IDS).forEach(([id, k]) => { if ($(id) && document.activeElement !== $(id)) $(id).value = state.cfg[k]; }); }
const benData = el => Object.fromEntries([...el.closest("details").querySelectorAll("[data-benf],[data-cost]")].map(i => [i.dataset.benf || i.dataset.cost, i.value]));
Office.onReady(info => {
  if (!POPOUT && info.host !== Office.HostType.Excel) return;
  if (!POPOUT) { const saved = Office.context.document.settings.get("tms_cfg"); if (saved) Object.assign(state.cfg, saved); }
  syncCfgInputs();
  Object.entries(CFG_IDS).forEach(([id, k]) => { $(id).onchange = () => {
    const v = $(id).value; state.cfg[k] = k === "cutoff" ? (v || CUTOFF_DEFAULT) : ["location", "device", "openTab"].includes(k) ? (v.trim() || { location: "San Rafael", device: "auto", openTab: "on" }[k]) : (+v || { authDays: 14, protoDays: 2, recentDays: 10 }[k]);
    saveSetting("tms_cfg", { ...state.cfg }); state.patients.forEach(p => p.d = derive(p)); renderBanners(); renderDash(); if (k === "cutoff") refresh(); }; });
  $("mainTabs").onclick = e => { const b = e.target.closest("[data-view]"); if (!b) return; if (b.dataset.view === "pts") state.ptMode = "list"; show(b.dataset.view); };
  $("schToggle").onclick = e => { const b = e.target.closest("[data-sv]"); if (b) { state.dash = b.dataset.sv; renderDash(); } };
  $("clinicTabs").onclick = e => { const b = e.target.closest("[data-cv]"); if (b) { state.clinicTab = b.dataset.cv; renderDash(); } };
  $("btnImport").onclick = () => show("imp"); $("apBack").onclick = () => show("pts");
  $("btnSettings").onclick = () => show("settings"); $("setBack").onclick = () => show(state.lastView || "sch");
  $("btnFull").onclick = fullScreen;
  $("popClose").onclick = closePopout;
  window.addEventListener("resize", syncAtt); syncFs();
  document.addEventListener("toggle", e => { if (e.target.matches && e.target.matches("details.kebab")) syncShade(); }, true);
  new MutationObserver(() => { if (!$("menuShade").hidden) syncShade(); }).observe(document.body, { childList: true, subtree: true }); // a re-render removed the open menu
  window.addEventListener("message", e => { if (e.origin === location.origin && e.data && e.data.sched) onSchedMessage(e.data.sched); });
  document.addEventListener("click", e => {
    const el = e.target;
    document.querySelectorAll("details.kebab[open]").forEach(k => { if (!k.contains(el) || el.closest(".menu button")) k.open = false; }); // outside click, shade click, or a menu choice
    syncShade();
    const bn = el.closest("[data-banner]"); if (bn) { if (bn.dataset.bact === "snooze") snooze(bn.dataset.banner); else { state.dismissed.add(bn.dataset.banner); renderBanners(); } return; }
    if (el.closest("[data-mclose]") || el.id === "modal") return closeModal();
    const mf = el.closest("[data-mf]"); if (mf) return setFreebie(mf.dataset.t, +mf.dataset.row, mf.dataset.mf);
    const pq = el.closest("[data-prq]"); if (pq) { $("prFrom").value = serialToInput(todaySerial()); $("prTo").value = serialToInput(+pq.dataset.prq); document.querySelectorAll("[data-prq]").forEach(b => b.classList.toggle("on", b === pq)); return prCount($("modalBody").querySelector("[data-act=doPrint]").dataset.t); }
    const as = el.closest("#attWrap > summary");
    if (as && document.body.classList.contains("wide")) { e.preventDefault(); state.attCollapsed = !state.attCollapsed; document.body.classList.toggle("attc", state.attCollapsed); $("attWrap").open = true; return; }
    const sm = el.closest("details > summary");
    if (sm) { const det = sm.parentElement; if (det.dataset.draft && state.drafts[det.dataset.draft]) state.drafts[det.dataset.draft].open = !det.open; if (det.dataset.open) { if (det.open) state.open.delete(det.dataset.open); else state.open.add(det.dataset.open); } return; }
    const ap = el.closest("[data-apact]"); if (ap) { document.querySelectorAll("[data-apr]").forEach(i => { state.apRanges[+i.dataset.apr][i.dataset.k] = i.value; }); if (ap.dataset.apact === "add") state.apRanges.push({}); else state.apRanges.splice(+ap.dataset.v, 1); renderApRanges(); return; }
    const a = el.dataset.act;
    if (a) {
      const t = el.dataset.t, v = el.dataset.v, row = +el.dataset.row;
      const acts = {
        apply: () => applyPlan(), previewSaved: () => buildPlan(), setupGo: () => { closeModal(); setupColumns(); }, newTracker: () => startTracker(v), status: () => setStatus(t, v), mkSched: () => mkSchedSheet(), schedSheet: () => writeChairSheet(),
        classify: () => freebieModal(t, row), missKind: () => markMiss(t, row, v), undoMiss: () => undoMiss(t, row), pulses: () => markPulses(t, row, el), saveMeas: () => saveMeasures(t, row, el), moveMeas: () => moveMeasures(t, row, el),
        measToggle: () => { const k = `meas|${t}|${row}`; if (state.open.has(k)) state.open.delete(k); else state.open.add(k); renderDash(); },
        mtSave: () => saveMT(t, row, el), protoSave: () => saveProtocol(t, el), print: () => printModal(t, v), reimport: () => { show("imp"); buildPlan(); },
        doPrint: () => { const [f, to] = prRange(); if (!f || !to || to < f) return toast("Pick a valid date range.", true); closeModal(); openSchedule(t, f, to); },
        copyNote: () => noteAction("copyNote", t, row, el), resetNote: () => noteAction("resetNote", t, row, el),
        ptBack: () => { state.ptMode = "list"; renderDash(); }, addPt: () => show("add"), ptFilter: () => { state.ptFilter = v; renderDash(); }, ptAll: () => { state.ptAll = !state.ptAll; renderDash(); },
        benSave: () => saveBenefits(t, benData(el)),
        extend: () => { const box = el.closest("details"); addExtension(t, { n: +box.querySelector("input[data-extn]").value, exp: inputToSerial(box.querySelector("input[data-extd]").value) }); },
        availSave: () => { const av = state.avDraft; if (!av || av.tab !== t) return; readAvailEditor($("ptsBody"), av); saveAvailability(t, av); },
        avAdd: () => { readAvailEditor($("ptsBody"), state.avDraft); state.avDraft.ranges.push({ type: "Unavailable", from: null, to: null, note: "" }); renderDash(); },
        avDel: () => { readAvailEditor($("ptsBody"), state.avDraft); state.avDraft.ranges.splice(+v, 1); renderDash(); },
        prev: () => { state.weekOffset--; renderDash(); }, next: () => { state.weekOffset++; renderDash(); }, this: () => { state.weekOffset = 0; renderDash(); },
        sthis: () => { state.schedOffset = 0; renderDash(); },
        chair: () => { state.chair = v; try { localStorage.setItem("tms_chair", v); } catch { /* session only */ } renderDash(); },
        dthis: () => { state.dayOffset = 0; renderDash(); },
        dprev: () => { do { state.dayOffset--; } while ([0, 6].includes(dow(todaySerial() + state.dayOffset))); renderDash(); },
        dnext: () => { do { state.dayOffset++; } while ([0, 6].includes(dow(todaySerial() + state.dayOffset))); renderDash(); },
        sprev: () => { do { state.schedOffset--; } while ([0, 6].includes(dow(todaySerial() + state.schedOffset))); renderDash(); },
        snext: () => { do { state.schedOffset++; } while ([0, 6].includes(dow(todaySerial() + state.schedOffset))); renderDash(); }
      };
      if (acts[a]) { e.preventDefault(); acts[a](); }
      return;
    }
    const pt = el.closest("[data-pt]"); if (pt) return openPtLink(pt);
    const tb = el.closest("[data-tab],[data-sheet]"); if (tb && tb.closest("section")) openTabName(tb.dataset.tab || tb.dataset.sheet, tb.dataset.row ? +tb.dataset.row : null);
  });
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && !$("schedOverlay").hidden) return closeSchedOverlay();
    if (e.key === "Escape" && !$("modal").hidden) return closeModal();
    if (e.key === "Enter") { const pt = e.target.closest && e.target.closest("[data-pt]"); if (pt && e.target === pt) return openPtLink(pt); const el = e.target.closest && e.target.closest("[data-tab]"); if (el && e.target === el) openTabName(el.dataset.tab, el.dataset.row ? +el.dataset.row : null); }
  });
  document.addEventListener("change", e => {
    const el = e.target;
    if (!el.closest("section") || el.closest("#v-add") || el.closest("#v-imp") || el.closest("#v-settings")) return;
    if (el.dataset.miss && el.value) return markMiss(el.dataset.t, +el.dataset.miss, el.value);
    if (el.dataset.dk && el.tagName === "SELECT" && draftInput(el)) renderDash();
  });
  document.addEventListener("input", e => { if (e.target.dataset.dk) draftInput(e.target); if (e.target.id === "prFrom" || e.target.id === "prTo") { document.querySelectorAll("[data-prq]").forEach(b => b.classList.remove("on")); prCount($("modalBody").querySelector("[data-act=doPrint]").dataset.t); } });
  $("refresh").onclick = () => refresh(); $("btnApRebuild").onclick = () => rebuildActiveClick(); $("btnSetup").onclick = () => setupClick(false); $("btnMeas").onclick = placeMeasuresClick;
  $("btnMatch").onclick = () => setupClick(true);
  $("file").onchange = onFile; $("benFile").onchange = onBenFile; document.querySelectorAll("input[name=rtype]").forEach(r => r.onchange = () => state.file && !state.file.saved && buildPlan());
  ["apLast", "apFirst"].forEach(i => $(i).oninput = () => { if (!state.codeEdited) $("apCode").value = lasfir($("apLast").value, $("apFirst").value); updateApTab(); });
  $("apCode").oninput = () => { state.codeEdited = true; updateApTab(); };
  $("apCourse").oninput = () => { state.courseEdited = true; updateApTab(); };
  $("btnAdd").onclick = () => addPatient(collectAddPatient());
  $("btnFmt").onclick = formatLogsClick;
  fillAddPatientForm(); show("sch");
  if (POPOUT) initPopout(); else refresh();
});
