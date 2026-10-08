/* TMS Tracker: Excel task-pane add-in for the San Rafael MagVenture tracker.
   Works on the existing layout: one tab per course (LASFIR or "LASFIR 2"), cloned from "New Template".
   PHI: MRNs are read into browser memory to match NextGen rows to tabs. They are never displayed in full,
   written anywhere, or sent to a server. Patient names typed in Add patient become LASFIR and are discarded. */
"use strict";

/* ---------- layout of a patient tab (existing, plus columns T:Z added by this add-in) ---------- */
const LOG_FIRST = 10, LOG_LAST = 49;              // Tx log rows
const C = { tx: 2, date: 4, L: 5, R: 7, O: 9, type: 11, miss: 12, taper: 13, action: 14, alt: 17, note: 18, freebie: 19, time: 20 }; // 0-based in A:U
const TAB_RE = /^([A-Z]{1,6})(?: (\d+))?$/;
const ACTIVE = ["Pending Start", "In Progress", "Tapering"];
const DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const AV_TYPE = ["Unavailable", "Available"];
const FREEBIE = ["Used", "Waived"];
const CANCEL = ["Cancelled", "No-show"];

const state = { cfg: { authDays: 14, protoDays: 7 }, patients: [], tabs: [], protocols: [], needCols: [], weekOffset: 0, dash: "att", file: null, plan: null, codeEdited: false, courseEdited: false };
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- dates, times ---------- */
const utcSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5);
const todaySerial = () => { const d = new Date(); return utcSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); };
const serialToDate = s => new Date(Date.UTC(1899, 11, 30) + s * 864e5);
const fmtDate = s => { if (typeof s !== "number") return ""; const d = serialToDate(s); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}`; };
const fmtDateY = s => { if (typeof s !== "number") return ""; const d = serialToDate(s); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`; };
const fmtTime = m => { if (m == null) return ""; const h = Math.floor(m / 60); return `${((h + 11) % 12) + 1}:${String(m % 60).padStart(2, "0")} ${h < 12 ? "AM" : "PM"}`; };
const weekStart = s => s - serialToDate(s).getUTCDay();   // weeks run Sunday to Saturday
const inputToSerial = v => { if (!v) return null; const [y, m, d] = v.split("-").map(Number); return utcSerial(y, m, d); };
const serialToInput = s => serialToDate(s).toISOString().slice(0, 10);
const cmp = (a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0);
const isNum = v => typeof v === "number" && v > 0;

function toSerial(v) {
  if (v instanceof Date) { const d = new Date(v.getTime() + 432e5); return utcSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  if (typeof v === "number") return v > 20000 ? Math.floor(v) : null;
  const s = String(v ?? "").trim(); let m;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/))) { let y = +m[3]; if (y < 100) y += 2000; return utcSerial(y, +m[1], +m[2]); }
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return utcSerial(+m[1], +m[2], +m[3]);
  const d = new Date(s); return isNaN(d) ? null : toSerial(d);
}
function toMinutes(v) {
  if (v instanceof Date) { const d = new Date(v.getTime() + 3e4); return d.getHours() * 60 + d.getMinutes() || null; }
  if (typeof v === "number") { const f = v < 1 ? v : v - Math.floor(v); return Math.round(f * 1440) || null; }
  const m = String(v ?? "").match(/(\d{1,2}):(\d{2})\s*([AaPp][Mm])?/);
  if (!m) return null;
  let h = +m[1]; const ap = (m[3] || "").toLowerCase();
  if (ap === "pm" && h < 12) h += 12; if (ap === "am" && h === 12) h = 0;
  return h * 60 + +m[2];
}
const clean = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z]/g, "").toUpperCase();
const lasfir = (last, first) => clean(last).slice(0, 3) + clean(first).slice(0, 3);
const normMrn = s => String(s ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "").replace(/^0+/, "");
function mapStatus(raw) {
  const s = String(raw ?? "").toLowerCase();
  if (/resched/.test(s)) return "Skip";                  // the slot moved; the new slot is its own row
  if (/no[\s-]?show/.test(s)) return "No-show";
  if (/cancel/.test(s)) return "Cancelled";
  if (/complete|checked out|check-out|seen|arrived|checked in/.test(s)) return "Completed";
  return "Scheduled";
}

/* ---------- UI helpers ---------- */
let toastTimer;
function toast(msg, err) { const t = $("toast"); t.textContent = msg; t.className = err ? "err" : ""; t.style.display = "block"; clearTimeout(toastTimer); toastTimer = setTimeout(() => t.style.display = "none", err ? 9000 : 3500); }
const fail = e => { console.error(e); toast(e.message || String(e), true); };

/* ---------- read the workbook ---------- */
function parseLog(vals) {
  return vals.map((v, i) => ({
    row: LOG_FIRST + i, tx: v[C.tx], date: isNum(v[C.date]) ? Math.floor(v[C.date]) : null,
    time: isNum(v[C.time]) ? Math.round(v[C.time] * 1440) : null,
    delivered: [v[C.L], v[C.R], v[C.O]].some(x => x !== "" && x != null),
    type: String(v[C.type] || ""), miss: String(v[C.miss]).trim().toLowerCase() === "x",
    action: String(v[C.action] || ""), alt: String(v[C.alt]).trim().toLowerCase() === "x",
    note: String(v[C.note] || ""), freebie: String(v[C.freebie] || "")
  }));
}
async function refresh() {
  try {
    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets; sheets.load("items/name"); await ctx.sync();
      const L = sheets.items.filter(s => TAB_RE.test(s.name) || s.name === "New Template").map(s => {
        const g = a => { const r = s.getRange(a); r.load("values"); return r; };
        return { name: s.name, a1: g("A1"), mrn: g("A4"), d2: g("D2"), d67: g("D6:D7"), fg4: g("F4:G4"), g2: g("G2"), f7h7: g("F7:H7"), prot: g("J2:N5"), pref: g("M6"), log: g(`A${LOG_FIRST}:U${LOG_LAST}`), t8: g("T8"), fb: g("X1"), av: g("W9:Z30") };
      });
      const pr = ctx.workbook.worksheets.getItemOrNullObject("Protocols"); await ctx.sync();
      let prv = null; if (!pr.isNullObject) { prv = pr.getRange("A2:I40"); prv.load("values"); }
      await ctx.sync();
      state.patients = []; state.tabs = []; state.needCols = [];
      for (const l of L) {
        if (l.t8.values[0][0] !== "Freebie" && (l.a1.values[0][0] === "PATIENT:" || l.name === "New Template")) state.needCols.push(l.name);
        if (l.name === "New Template" || l.a1.values[0][0] !== "PATIENT:") continue;
        const m = l.name.match(TAB_RE), pv = l.prot.values;
        const mins = [0, 1, 2].map(i => pv[i][4]).filter(x => typeof x === "number");
        state.tabs.push({ name: l.name, code: m[1], course: +(m[2] || 1) });
        state.patients.push({
          tab: l.name, code: m[1], course: +(m[2] || 1), mrn: normMrn(l.mrn.values[0][0]),
          txApproved: l.d2.values[0][0], mdReq: String(l.d67.values[0][0] || ""), status: String(l.d67.values[1][0] || ""),
          authStart: isNum(l.fg4.values[0][0]) ? l.fg4.values[0][0] : null, authExp: isNum(l.fg4.values[0][1]) ? l.fg4.values[0][1] : null,
          updateBy: isNum(l.g2.values[0][0]) ? l.g2.values[0][0] : null,
          txEnd: isNum(l.f7h7.values[0][0]) ? l.f7h7.values[0][0] : null, txRem: typeof l.f7h7.values[0][2] === "number" ? l.f7h7.values[0][2] : null,
          pref: String(l.pref.values[0][0] || ""),
          proto: pv[0][0] ? `${pv[0][0]} Hz, ${pv[0][2]} trains` : "", protoMin: mins.length ? mins.reduce((a, b) => a + b, 0) : null,
          freebiesAllowed: typeof l.fb.values[0][0] === "number" ? l.fb.values[0][0] : 3,
          rows: parseLog(l.log.values),
          avail: l.av.values.filter(v => AV_TYPE.includes(v[0]) && isNum(v[1])).map(v => ({ type: v[0], from: Math.floor(v[1]), to: isNum(v[2]) ? Math.floor(v[2]) : Math.floor(v[1]), note: String(v[3] || "") }))
        });
      }
      state.patients.forEach(p => p.d = derive(p));
      state.patients.sort((a, b) => a.code.localeCompare(b.code) || a.course - b.course);
      state.protocols = prv ? prv.values.filter(r => r[1] !== "").map(r => ({ type: r[0], name: String(r[1]).trim(), rep: r[2], pulses: r[3], trains: r[4], iti: r[5], typed: r[7], inten: r[8] })) : [];
    });
    $("setup").hidden = !state.needCols.length;
    $("setupMsg").textContent = `${state.needCols.length} tab${state.needCols.length === 1 ? "" : "s"} (including New Template) need the freebie and availability columns to the right of Notes.`;
    renderDash(); updateApTab();
  } catch (e) { fail(e); }
}

function derive(p) {
  const today = todaySerial(), cfg = state.cfg, rows = p.rows;
  const active = p.status === "" || ACTIVE.includes(p.status);
  const sched = rows.filter(r => r.date != null && r.date >= today && !r.delivered && !r.miss).sort(cmp);
  const next = sched.find(r => r.type !== "F/U" && !r.alt) || sched[0];
  const overdue = rows.filter(r => r.date != null && r.date < today && !r.delivered && !r.miss && !r.alt);
  const decisions = rows.filter(r => r.miss && !r.freebie);
  const used = rows.filter(r => r.freebie === "Used").length, waived = rows.filter(r => r.freebie === "Waived").length;
  const fu = rows.filter(r => r.type === "F/U");
  const fuNext = fu.filter(r => r.date != null && r.date >= today && !r.delivered).sort(cmp)[0];
  const fuStatus = !fu.length ? "none" : fuNext ? "scheduled" : fu.every(r => r.delivered) ? "done" : "not scheduled";
  const protoActs = rows.filter(r => r.action === "Protocol" && !r.delivered && r.date != null && r.date >= today).sort(cmp);
  const completed = rows.filter(r => r.delivered && r.type !== "F/U").length;
  const authFlags = [];
  if (active) {
    if (p.authExp == null) authFlags.push("No auth expiry entered");
    else {
      const days = p.authExp - today;
      if (days < 0) authFlags.push(`Auth expired ${fmtDate(p.authExp)}`);
      else if (days <= cfg.authDays) authFlags.push(`Auth expires in ${days} d (${fmtDate(p.authExp)})`);
      if (p.txEnd != null && p.txEnd > p.authExp) authFlags.push(`Tx end ${fmtDate(p.txEnd)} is after auth expiry`);
      if (sched.some(r => r.date > p.authExp)) authFlags.push("Visits booked past auth expiry");
    }
  }
  const protoFlags = [];
  if (active) {
    if (p.updateBy != null && p.updateBy - today <= cfg.protoDays) protoFlags.push(p.updateBy < today ? `Update was due ${fmtDate(p.updateBy)}` : `Update by ${fmtDate(p.updateBy)}`);
    const pa = protoActs.find(r => r.date - today <= cfg.protoDays + 7);
    if (pa) protoFlags.push(`Protocol action at the ${fmtDate(pa.date)} visit`);
  }
  return { active, sched, next, overdue, decisions, used, waived, left: Math.max(0, p.freebiesAllowed - used), fu, fuNext, fuStatus, completed, authFlags, protoFlags };
}
const isUnavail = (p, day) => p.avail.some(w => w.type === "Unavailable" && w.from <= day && day <= w.to);
function availText(p) {
  const t = todaySerial(), parts = [];
  const un = p.avail.filter(w => w.type === "Unavailable" && w.to >= t).sort((a, b) => a.from - b.from)[0];
  if (un) parts.push(un.from <= t ? `Unavailable until ${fmtDate(un.to)}` : `Unavailable ${fmtDate(un.from)}\u2013${fmtDate(un.to)}`);
  const av = p.avail.filter(w => w.type === "Available" && w.to >= t).sort((a, b) => a.from - b.from)[0];
  if (av) parts.push(`Available ${fmtDate(av.from)}\u2013${fmtDate(av.to)}`);
  if (p.pref) parts.push(`Prefers ${p.pref}`);
  return parts.join("; ");
}

/* ---------- dashboard ---------- */
const item = (p, main, sub, cls, row) =>
  `<div class="item ${cls || ""}" tabindex="0" data-tab="${esc(p.tab)}"${row ? ` data-row="${row}"` : ""}><b>${esc(p.tab)}</b> ${main || ""}${sub ? `<div class="sub">${esc(sub)}</div>` : ""}</div>`;
const group = (title, items, empty) =>
  `<div class="group"><h2>${esc(title)}<span class="n">${items.length}</span></h2>${items.length ? items.join("") : `<div class="none">${esc(empty)}</div>`}</div>`;

function renderDash() {
  const body = $("dashBody");
  if (state.dash === "proto") { body.innerHTML = viewProtocols(); return; }
  if (!state.patients.length) { body.innerHTML = `<div class="none">No patient tabs found. A patient tab has "PATIENT:" in A1 and a LASFIR name such as SMIJAN or SMIJAN 2.</div>`; return; }
  body.innerHTML = { att: viewAttention, week: viewWeek, fu: viewFollowUp, roster: viewRoster }[state.dash]();
}

function viewAttention() {
  const P = state.patients, A = P.filter(p => p.d.active);
  return group("Protocol update needed", A.filter(p => p.d.protoFlags.length).map(p => item(p, "", p.d.protoFlags.join("; "), "amb")), "None due")
    + group("Auth flags", A.filter(p => p.d.authFlags.length).map(p => item(p, "", p.d.authFlags.join("; "), "red")), "No auth issues")
    + group("Freebie decision needed", P.flatMap(p => p.d.decisions.map(r => item(p, `<span class="chip amb">Missed</span> ${fmtDateY(r.date)}`, "Mark the Freebie column Used or Waived", "amb", r.row))), "Every missed visit is marked")
    + group("Out of freebies", A.filter(p => p.d.left === 0).map(p => item(p, "", `${p.d.used} of ${p.freebiesAllowed} used`, "red")), "Nobody")
    + group("Past date, no outcome recorded", A.filter(p => p.d.overdue.length).map(p => item(p, "", `${p.d.overdue.length} visit${p.d.overdue.length > 1 ? "s" : ""}, latest ${fmtDate(p.d.overdue[p.d.overdue.length - 1].date)}. Enter pulses or mark Miss.`, "amb", p.d.overdue[p.d.overdue.length - 1].row)), "Nothing outstanding");
}

function viewWeek() {
  const t = todaySerial(), start = weekStart(t) + 7 * state.weekOffset, P = state.patients;
  let html = `<div class="wk"><button data-act="prev">\u2039 Prev</button><b>${fmtDate(start)} \u2013 ${fmtDate(start + 6)}</b><button data-act="next">Next \u203A</button></div>`;
  if (state.weekOffset) html += `<div style="margin:-4px 0 10px"><button data-act="this">Back to this week</button></div>`;
  for (let i = 0; i < 7; i++) {
    const day = start + i;
    const rows = P.flatMap(p => p.rows.filter(r => r.date === day).map(r => ({ p, r }))).sort((x, y) => (x.r.time ?? 0) - (y.r.time ?? 0));
    const sched = rows.filter(({ r }) => !r.delivered && !r.miss && !r.alt).length;
    html += `<div class="day${day === t ? " today" : ""}"><h2><span>${DAYS[i]} ${fmtDate(day)}</span><span class="n">${sched} scheduled</span></h2>`;
    html += rows.length ? rows.map(({ p, r }) => {
      const chips = [r.type && r.type !== "Daily" ? `<span class="chip">${esc(r.type)}</span>` : "", r.action ? `<span class="chip">${esc(r.action)}</span>` : "",
        r.delivered ? `<span class="chip">Done</span>` : "", r.miss ? `<span class="chip amb">Missed</span>` : "", r.alt ? `<span class="chip amb">Other clinic</span>` : "",
        !r.delivered && !r.miss && isUnavail(p, day) ? `<span class="chip red">marked unavailable</span>` : ""].filter(Boolean).join(" ");
      return item(p, `${fmtTime(r.time)} ${chips}`, "", "", r.row);
    }).join("") : `<div class="none">No visits</div>`;
    html += `</div>`;
  }
  const out = P.filter(p => p.d.active && p.avail.some(w => w.type === "Unavailable" && w.from <= start + 6 && w.to >= start));
  return html + group("Unavailable this week", out.map(p => {
    const w = p.avail.find(w => w.type === "Unavailable" && w.from <= start + 6 && w.to >= start);
    return item(p, `${fmtDate(w.from)}\u2013${fmtDate(w.to)}`, w.note, "");
  }), "Nobody");
}

function viewFollowUp() {
  const A = state.patients.filter(p => p.d.active);
  const prog = p => `${p.d.completed}${typeof p.txApproved === "number" ? "/" + p.txApproved : ""} Tx${p.txRem != null ? `, ${p.txRem} left` : ""}`;
  const notFu = A.filter(p => p.d.fuStatus === "not scheduled"), fuOk = A.filter(p => p.d.fuStatus === "scheduled").sort((a, b) => a.d.fuNext.date - b.d.fuNext.date);
  const noNext = A.filter(p => !p.d.next && p.txRem !== 0);
  return group("Follow-up not scheduled", notFu.map(p => item(p, "", [prog(p), availText(p)].filter(Boolean).join("; "), "red")), "Every follow-up has a date")
    + group("Follow-up scheduled", fuOk.map(p => item(p, `${DAYS[serialToDate(p.d.fuNext.date).getUTCDay()]} ${fmtDate(p.d.fuNext.date)} ${fmtTime(p.d.fuNext.time)}`, prog(p), "grn", p.d.fuNext.row)), "None")
    + group("No upcoming treatment visit", noNext.map(p => item(p, "", [prog(p), availText(p)].filter(Boolean).join("; "), "amb")), "Everyone has a next visit");
}

function viewRoster() {
  const rows = state.patients.map(p => `<tr data-tab="${esc(p.tab)}" tabindex="0"><td>${esc(p.tab)}</td><td>${esc(p.status || "")}</td>
    <td class="num">${p.d.completed}${typeof p.txApproved === "number" ? "/" + p.txApproved : ""}</td><td>${p.d.next ? fmtDate(p.d.next.date) : ""}</td><td>${fmtDate(p.authExp)}</td><td class="num">${p.d.left}</td><td class="num">${p.protoMin != null ? p.protoMin.toFixed(1) : ""}</td></tr>`).join("");
  return `<table><thead><tr><th>Patient</th><th>Status</th><th class="num">Tx</th><th>Next</th><th>Auth exp</th><th class="num">Free</th><th class="num">Min</th></tr></thead><tbody>${rows}</tbody></table>
    <div class="hint">Free = freebies left. Min = treatment minutes from the tab's Hz/Dur/Trains/Wait block. Select a row to open the tab.</div>`;
}

function hhmm(v) { if (typeof v !== "number") return ""; const h = Math.floor(v * 24 + 1e-6), m = Math.round((v * 24 - h) * 60); return `${h}:${String(m).padStart(2, "0")}`; }
function viewProtocols() {
  if (!state.protocols.length) return `<div class="none">No Protocols sheet found.</div>`;
  const rows = state.protocols.map(p => {
    const calc = [p.rep, p.pulses, p.trains, p.iti].every(x => typeof x === "number" && x > 0) ? ((p.pulses / p.rep) + p.iti) * p.trains / 60 : null;
    const mm = calc == null ? "" : `${Math.floor(calc)}:${String(Math.round((calc % 1) * 60)).padStart(2, "0")}`;
    return `<tr><td>${esc(p.name)}</td><td class="num">${p.rep}</td><td class="num">${p.pulses}</td><td class="num">${p.trains}</td><td class="num">${p.iti}</td><td class="num">${mm}</td><td class="num">${hhmm(p.typed)}</td></tr>`;
  }).join("");
  return `<table><thead><tr><th>Protocol</th><th class="num">Hz</th><th class="num">Pulses</th><th class="num">Trains</th><th class="num">ITI</th><th class="num">Calc m:ss</th><th class="num">Typed</th></tr></thead><tbody>${rows}</tbody></table>
    <div class="hint">Calc uses the same formula as each tab's N2: ((pulses / Hz) + ITI) x trains / 60. Typed is the Tx Time cell on the Protocols sheet as entered. Edit protocols on that sheet.</div>
    <div class="actions"><button data-sheet="Protocols">Open Protocols sheet</button></div>`;
}

/* ---------- copy dashboard to a sheet; fill the Active Patients grid ---------- */
function sections() {
  const t = todaySerial(), A = state.patients.filter(p => p.d.active), start = weekStart(t);
  const week = state.patients.flatMap(p => p.rows.filter(r => r.date >= start && r.date <= start + 6).map(r => ({ p, r }))).sort((x, y) => cmp(x.r, y.r))
    .map(({ p, r }) => [fmtDateY(r.date), DAYS[serialToDate(r.date).getUTCDay()], fmtTime(r.time), p.tab, r.type, r.action, r.delivered ? "Done" : r.miss ? "Missed" : r.alt ? "Other clinic" : "Scheduled"]);
  return [
    { title: "Protocol update needed", headers: ["Patient", "Why"], rows: A.filter(p => p.d.protoFlags.length).map(p => [p.tab, p.d.protoFlags.join("; ")]) },
    { title: "Auth flags", headers: ["Patient", "Auth expires", "Flags"], rows: A.filter(p => p.d.authFlags.length).map(p => [p.tab, fmtDateY(p.authExp), p.d.authFlags.join("; ")]) },
    { title: "Freebie decision needed", headers: ["Patient", "Missed date"], rows: state.patients.flatMap(p => p.d.decisions.map(r => [p.tab, fmtDateY(r.date)])) },
    { title: "Follow-up not scheduled", headers: ["Patient", "Tx done", "Availability"], rows: A.filter(p => p.d.fuStatus === "not scheduled").map(p => [p.tab, p.d.completed, availText(p)]) },
    { title: `Week of ${fmtDateY(start)} (Sunday to Saturday)`, headers: ["Date", "Day", "Time", "Patient", "Type", "Action", "Status"], rows: week }
  ];
}
async function publish() {
  try {
    await Excel.run(async ctx => {
      let ws = ctx.workbook.worksheets.getItemOrNullObject("Dashboard"); await ctx.sync();
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add("Dashboard"); else ws.getRange().clear();
      ws.getRange("A1").values = [[`Snapshot ${fmtDateY(todaySerial())}. Run it again from the add-in to refresh.`]];
      let r = 2;
      for (const s of sections()) {
        ws.getRangeByIndexes(r, 0, 1, 1).values = [[s.title]]; ws.getRangeByIndexes(r, 0, 1, 1).format.font.bold = true;
        const h = ws.getRangeByIndexes(r + 1, 0, 1, s.headers.length); h.values = [s.headers]; h.format.font.bold = true; h.format.fill.color = "#e2f1f1";
        if (s.rows.length) ws.getRangeByIndexes(r + 2, 0, s.rows.length, s.headers.length).values = s.rows; else ws.getRangeByIndexes(r + 2, 0, 1, 1).values = [["None"]];
        r += Math.max(1, s.rows.length) + 4;
      }
      ws.getRange("A:G").format.columnWidth = 110; ws.activate(); await ctx.sync();
    });
    toast("Dashboard sheet updated.");
  } catch (e) { fail(e); }
}
const gridCode = r => r.miss ? "Missed" : r.alt ? "Different Clinic" : r.type === "MT" ? "MT" : r.type === "MTR" ? "MTR" : r.type === "F/U" ? "F/U" : "Daily Tx";
async function fillGrid() {
  try {
    const n = await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem("Active Patients");
      if ($("gridMonday").checked) { const t = todaySerial(); ws.getRange("N2").values = [[t - ((serialToDate(t).getUTCDay() + 6) % 7)]]; await ctx.sync(); }
      const hdr = ws.getRange("N2:W2"), names = ws.getRange("B3:B39"), grid = ws.getRange("N3:W39");
      hdr.load("values"); names.load("values"); grid.load("formulas"); await ctx.sync();
      const dates = hdr.values[0], out = grid.formulas.map(r => r.slice());
      let count = 0;
      names.values.forEach((nm, i) => {
        const p = state.patients.find(x => x.tab === String(nm[0]).trim()); if (!p) return;
        dates.forEach((d, j) => {
          if (out[i][j] !== "" || !isNum(d)) return;
          const r = p.rows.find(x => x.date === Math.floor(d)); if (r) { out[i][j] = gridCode(r); count++; }
        });
      });
      grid.formulas = out; await ctx.sync(); return count;
    });
    toast(`Filled ${n} cells.`);
  } catch (e) { fail(e); }
}

/* ---------- tracker columns (right of Notes) ---------- */
function ensureColumns(ws) {
  ws.getRange("T8:U8").values = [["Freebie", "Time"]];
  ws.getRange("T8:U8").copyFrom("S8", Excel.RangeCopyType.formats);
  ws.getRange(`T${LOG_FIRST}:U${LOG_LAST}`).copyFrom(`R${LOG_FIRST}:R${LOG_LAST}`, Excel.RangeCopyType.formats);
  ws.getRange(`U${LOG_FIRST}:U${LOG_LAST}`).numberFormat = [["h:mm AM/PM"]];
  ws.getRange(`T${LOG_FIRST}:T${LOG_LAST}`).dataValidation.rule = { list: { inCellDropDown: true, source: FREEBIE.join(",") } };
  const cf = ws.getRange(`T${LOG_FIRST}:T${LOG_LAST}`).conditionalFormats.add(Excel.ConditionalFormatType.custom);
  cf.custom.rule.formula = `=AND($M${LOG_FIRST}="x",$T${LOG_FIRST}="")`; cf.custom.format.fill.color = "#FDE9B8";
  ws.getRange("W1:X4").values = [["Freebies allowed", 3], ["Used", ""], ["Waived", ""], ["Remaining", ""]];
  ws.getRange("X2:X4").formulas = [[`=COUNTIF($T$${LOG_FIRST}:$T$${LOG_LAST},"Used")`], [`=COUNTIF($T$${LOG_FIRST}:$T$${LOG_LAST},"Waived")`], ["=MAX(0,X1-X2)"]];
  ws.getRange("W1:W4").format.font.bold = true; ws.getRange("X1").format.fill.color = "#fff8dc"; ws.getRange("X1:X4").format.horizontalAlignment = "Left";
  ws.getRange("W6").values = [["AVAILABILITY WINDOWS"]]; ws.getRange("W6").format.font.bold = true;
  ws.getRange("W8:Z8").values = [["Type", "From", "To", "Note"]];
  ws.getRange("W8:Z8").format.font.bold = true; ws.getRange("W8:Z8").format.fill.color = "#e2f1f1";
  ws.getRange("W9:W30").dataValidation.rule = { list: { inCellDropDown: true, source: AV_TYPE.join(",") } };
  ws.getRange("X9:Y30").numberFormat = [["m/d/yyyy"]];
  ws.getRange("T:U").format.columnWidth = 70; ws.getRange("V:V").format.columnWidth = 16; ws.getRange("W:Y").format.columnWidth = 100; ws.getRange("Z:Z").format.columnWidth = 220;
}
async function setupColumns() {
  try {
    await Excel.run(async ctx => { state.needCols.forEach(n => ensureColumns(ctx.workbook.worksheets.getItem(n))); await ctx.sync(); });
    toast("Columns added."); await refresh();
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
async function addPatient() {
  const code = apCode(), course = +$("apCourse").value || 1, name = course > 1 ? `${code} ${course}` : code;
  if (!code) return toast("Enter a name or a LASFIR code.", true);
  try {
    await Excel.run(async ctx => {
      const wsAll = ctx.workbook.worksheets;
      const ex = wsAll.getItemOrNullObject(name), tpl = wsAll.getItemOrNullObject("New Template"), ap = wsAll.getItemOrNullObject("Active Patients"); await ctx.sync();
      if (!ex.isNullObject) throw new Error(`A tab named ${name} already exists.`);
      if (tpl.isNullObject) throw new Error("There is no New Template tab to copy.");
      const names = ap.isNullObject ? null : ap.getRange("B3:B39"); if (names) names.load("values"); await ctx.sync();
      const ws = tpl.copy(Excel.WorksheetPositionType.end); ws.name = name;
      ws.getRange("A2").values = [[code]];
      if ($("apMrn").value.trim()) ws.getRange("A4").values = [[$("apMrn").value.trim()]];
      if ($("apTxs").value) ws.getRange("D2").values = [[+$("apTxs").value]];
      ws.getRange("D7").values = [["Pending Start"]];
      ensureColumns(ws);
      if (names) { const i = names.values.findIndex(r => String(r[0]).trim() === ""); if (i >= 0) ap.getRange(`B${3 + i}`).values = [[name]]; else toast("Tab created, but Active Patients has no empty row.", true); }
      ws.activate(); await ctx.sync();
    });
    ["apLast", "apFirst", "apCode", "apMrn", "apTxs"].forEach(i => $(i).value = "");
    state.codeEdited = state.courseEdited = false;
    toast(`Created ${name}.`); await refresh();
  } catch (e) { fail(e); }
}

/* ---------- NextGen import ---------- */
const MAP_IDS = ["mMrn", "mDate", "mTime", "mStatus", "mType", "mFilterCol"];
async function onFile(e) {
  const f = e.target.files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    const all = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" });
    const hi = all.findIndex(r => r.filter(c => String(c).trim() !== "").length >= 3);
    if (hi < 0) throw new Error("Could not find a header row in that file.");
    const headers = all[hi].map((h, i) => String(h).trim() || `Column ${i + 1}`);
    state.file = { headers, rows: all.slice(hi + 1).filter(r => r.some(c => String(c).trim() !== "")) };
    state.plan = null; $("planBox").hidden = true;
    const saved = Office.context.document.settings.get("tms_map") || {};
    const guess = (id, re) => { const s = saved[id]; return s && headers.includes(s) ? headers.indexOf(s) : headers.findIndex(h => re.test(h)); };
    const opts = '<option value="-1">(none)</option>' + headers.map((h, i) => `<option value="${i}">${esc(h)}</option>`).join("");
    MAP_IDS.forEach(id => $(id).innerHTML = opts);
    $("mMrn").value = guess("mMrn", /mrn|chart|account|patient\s*(id|#|no)/i); $("mDate").value = guess("mDate", /date/i);
    $("mTime").value = guess("mTime", /time|start/i); $("mStatus").value = guess("mStatus", /status/i);
    $("mType").value = guess("mType", /type|reason|visit/i); $("mFilterCol").value = guess("mFilterCol", /^$/);
    if (saved.mFuTxt) $("mFuTxt").value = saved.mFuTxt; if (saved.mFilterTxt) $("mFilterTxt").value = saved.mFilterTxt;
    if (!$("fromDate").value) $("fromDate").value = serialToInput(todaySerial());
    $("mapBox").hidden = false; toast(`Read ${state.file.rows.length} rows.`);
  } catch (err) { fail(err); }
}

/* Put NextGen appointments into a tab's Tx rows. Delivered rows and Miss rows are never moved (unless asked). */
function planTab(p, appts, from, to, o, importDate) {
  const R = p.rows.map(r => ({ ...r })), orig = p.rows;
  const eligible = r => !r.delivered && !r.alt && (o.clearMiss || !r.miss);
  const inWin = r => eligible(r) && r.date != null && r.date >= from && r.date <= to;
  const over = r => !r.delivered && !r.alt && !r.miss && (r.date == null || r.date > to);
  const keptMiss = new Set(orig.filter(r => r.miss && !(o.clearMiss && r.date != null && r.date >= from && r.date <= to)).map(r => r.date));
  const events = []; let unrecorded = 0;
  for (const a of appts) {
    if (a.status === "Scheduled") events.push(a);
    else if (CANCEL.includes(a.status) && !keptMiss.has(a.date)) { if (o.addMiss) events.push({ ...a, miss: true }); else unrecorded++; }
  }
  const slotOf = r => !o.hasType ? "all" : r.type === "F/U" ? "fu" : "tx", evOf = a => !o.hasType ? "all" : a.fu ? "fu" : "tx";
  const st = { cleared: 0, miss: 0, none: 0, unrecorded };
  for (const pool of o.hasType ? ["fu", "tx"] : ["all"]) {
    const win = R.filter(r => slotOf(r) === pool && inWin(r)), ov = R.filter(r => slotOf(r) === pool && over(r));
    for (const e of events.filter(e => evOf(e) === pool).sort(cmp)) {
      const r = win.shift() || ov.shift(); if (!r) { st.none++; continue; }
      r.date = e.date; r.time = e.time; r.freebie = "";
      if (e.miss) { r.miss = true; r.note = (r.note ? r.note + "; " : "") + `${e.status} per NextGen ${fmtDate(e.date)}`; r.freebie = ""; st.miss++; } else r.miss = false;
    }
    win.forEach(r => { r.date = null; r.time = null; r.miss = false; r.freebie = ""; st.cleared++; });
  }
  const ops = [];
  R.forEach((r, i) => {
    const b = orig[i];
    if (r.date !== b.date) ops.push([`E${r.row}`, r.date ?? ""]);
    if (r.time !== b.time) ops.push([`U${r.row}`, r.time == null ? "" : r.time / 1440]);
    if (r.miss !== b.miss) ops.push([`M${r.row}`, r.miss ? "x" : ""]);
    if (r.note !== b.note) ops.push([`S${r.row}`, r.note]);
    if (r.freebie !== b.freebie) ops.push([`T${r.row}`, r.freebie]);
  });
  const dates = R.filter((r, i) => r.date !== orig[i].date && r.date != null).map(r => r.date).sort((a, b) => a - b);
  return { name: p.tab, ops, dates, ...st, set: dates.length, changed: ops.length > 0 };
}

async function buildPlan() {
  const g = id => +$(id).value, from = inputToSerial($("fromDate").value);
  if (!from) return toast("Pick a start date.", true);
  const m = { mrn: g("mMrn"), date: g("mDate"), time: g("mTime"), status: g("mStatus"), type: g("mType"), fcol: g("mFilterCol"), ftxt: $("mFilterTxt").value.trim().toLowerCase() };
  if (m.mrn < 0 || m.date < 0) return toast("Pick the MRN and date columns.", true);
  const H = state.file.headers, nm = i => i >= 0 ? H[i] : "";
  const s = Office.context.document.settings;   // column names only, no patient data
  s.set("tms_map", { mMrn: nm(m.mrn), mDate: nm(m.date), mTime: nm(m.time), mStatus: nm(m.status), mType: nm(m.type), mFilterCol: nm(m.fcol), mFuTxt: $("mFuTxt").value, mFilterTxt: $("mFilterTxt").value }); s.saveAsync();
  let fuRe = null; try { fuRe = new RegExp($("mFuTxt").value || "F/U|follow", "i"); } catch { return toast("The follow-up text is not a valid pattern.", true); }

  await refresh();
  const by = new Map(); let maxDate = from, skipped = 0;
  for (const r of state.file.rows) {
    if (m.fcol >= 0 && m.ftxt && !String(r[m.fcol]).toLowerCase().includes(m.ftxt)) continue;
    const mrn = normMrn(r[m.mrn]), date = toSerial(r[m.date]), status = m.status >= 0 ? mapStatus(r[m.status]) : "Scheduled";
    if (!mrn || date == null || date < from || status === "Skip" || status === "Completed") { skipped++; continue; }
    let time = m.time >= 0 ? toMinutes(r[m.time]) : null; if (time == null) time = toMinutes(r[m.date]);
    maxDate = Math.max(maxDate, date);
    if (!by.has(mrn)) by.set(mrn, []);
    by.get(mrn).push({ date, time, status, fu: m.type >= 0 && fuRe.test(String(r[m.type])) });
  }
  const latest = new Map();
  for (const p of state.patients) if (p.mrn) { const c = latest.get(p.mrn); if (!c || p.course > c.course) latest.set(p.mrn, p); }
  const o = { addMiss: $("chkAddMiss").checked, clearMiss: $("chkClearMiss").checked, hasType: m.type >= 0 };
  const plans = [];
  for (const p of latest.values()) {
    const appts = by.get(p.mrn) || [];
    if (!appts.length && !($("chkAbsent").checked && p.d.active)) continue;
    plans.push(planTab(p, appts, from, maxDate, o, todaySerial()));
  }
  const unmatched = [...by.entries()].filter(([k]) => !latest.has(k)).map(([k, a]) => [k.slice(-4), a.length]);
  const noMrn = state.patients.filter(p => !p.mrn).map(p => p.tab);
  state.plan = { plans, unmatched, noMrn, from, to: maxDate, skipped };
  renderPlan();
}
function renderPlan() {
  const { plans, unmatched, noMrn, from, to, skipped } = state.plan, ch = plans.filter(p => p.changed);
  const tot = k => ch.reduce((n, p) => n + p[k], 0);
  let h = `<h2>Preview</h2><div class="sub">Window ${fmtDateY(from)} to ${fmtDateY(to)}. Delivered visits and existing Miss rows are kept${$("chkClearMiss").checked ? ", except Miss rows inside the window (you chose to clear them)" : ""}. ${skipped} file rows skipped (filtered, before the start date, rescheduled, or completed).</div>`;
  h += ch.length ? `<table><thead><tr><th>Tab</th><th class="num">Dates set</th><th class="num">Cleared</th><th class="num">Miss added</th></tr></thead><tbody>${ch.map(p => `<tr><td>${esc(p.name)}</td><td class="num">${p.set}</td><td class="num">${p.cleared}</td><td class="num">${p.miss}</td></tr>`).join("")}
    <tr><th>Total</th><th class="num">${tot("set")}</th><th class="num">${tot("cleared")}</th><th class="num">${tot("miss")}</th></tr></tbody></table>` : `<div class="none">No changes.</div>`;
  if (tot("cleared")) h += `<div class="hint">Cleared = rows in the window with a date that NextGen no longer shows. Their date is removed.</div>`;
  const warn = [];
  const none = plans.filter(p => p.none), unrec = plans.filter(p => p.unrecorded);
  if (none.length) warn.push(`No free Tx row for ${none.reduce((n, p) => n + p.none, 0)} appointment(s): ${none.map(p => p.name).join(", ")}. Add rows or check the tab.`);
  if (unrec.length) warn.push(`NextGen shows ${unrec.reduce((n, p) => n + p.unrecorded, 0)} no-show/cancellation(s) not on the tab: ${unrec.map(p => p.name).join(", ")}. Tick "Add NextGen no-shows" to record them.`);
  if (unmatched.length) warn.push(`No tab matches MRN ending ${unmatched.map(([k, n]) => `${esc(k)} (${n})`).join(", ")}. They may be non-TMS visits; the "Only rows where" filter can exclude them.`);
  if (noMrn.length) warn.push(`Tabs with no MRN in A4 cannot be matched: ${noMrn.join(", ")}.`);
  h += warn.map(w => `<div class="item amb" style="margin-top:8px">${w}</div>`).join("");
  $("planBody").innerHTML = h; $("planBox").hidden = false; $("btnApply").disabled = !ch.length;
}
async function applyPlan() {
  const ch = state.plan.plans.filter(p => p.changed);
  try {
    await Excel.run(async ctx => {
      for (const p of ch) { const ws = ctx.workbook.worksheets.getItem(p.name); p.ops.forEach(([a, v]) => ws.getRange(a).values = [[v]]); }
      await ctx.sync();
    });
    toast(`Updated ${ch.length} tab${ch.length === 1 ? "" : "s"}.`);
    state.plan = null; $("planBox").hidden = true; await refresh();
  } catch (e) { fail(e); }
}

/* ---------- wiring ---------- */
function show(view) {
  document.querySelectorAll("#mainTabs button").forEach(b => b.classList.toggle("on", b.dataset.view === view));
  ["dash", "imp", "add"].forEach(v => $("v-" + v).hidden = v !== view);
}
async function openTab(el) {
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(el.dataset.tab || el.dataset.sheet); ws.activate();
      if (el.dataset.row) ws.getRange(`A${el.dataset.row}:U${el.dataset.row}`).select();
      await ctx.sync();
    });
  } catch (e) { fail(e); }
}
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) return;
  const saved = Office.context.document.settings.get("tms_cfg"); if (saved) Object.assign(state.cfg, saved);
  $("cfgAuth").value = state.cfg.authDays; $("cfgProto").value = state.cfg.protoDays;
  const saveCfg = () => { state.cfg.authDays = +$("cfgAuth").value || 14; state.cfg.protoDays = +$("cfgProto").value || 7; Office.context.document.settings.set("tms_cfg", state.cfg); Office.context.document.settings.saveAsync(); state.patients.forEach(p => p.d = derive(p)); renderDash(); };
  $("cfgAuth").onchange = $("cfgProto").onchange = saveCfg;
  $("mainTabs").onclick = e => e.target.dataset.view && show(e.target.dataset.view);
  $("dashTabs").onclick = e => { const d = e.target.dataset.d; if (!d) return; state.dash = d; document.querySelectorAll("#dashTabs button").forEach(b => b.classList.toggle("on", b.dataset.d === d)); renderDash(); };
  $("dashBody").addEventListener("click", e => {
    const a = e.target.dataset.act;
    if (a) { state.weekOffset = a === "prev" ? state.weekOffset - 1 : a === "next" ? state.weekOffset + 1 : 0; return renderDash(); }
    const el = e.target.closest("[data-tab],[data-sheet]"); if (el) openTab(el);
  });
  $("dashBody").addEventListener("keydown", e => { if (e.key === "Enter") { const el = e.target.closest("[data-tab]"); if (el) openTab(el); } });
  $("refresh").onclick = refresh; $("publish").onclick = publish; $("btnGrid").onclick = fillGrid; $("btnSetup").onclick = setupColumns;
  $("file").onchange = onFile; $("btnPreview").onclick = () => buildPlan().catch(fail);
  $("btnApply").onclick = applyPlan; $("btnCancelPlan").onclick = () => { state.plan = null; $("planBox").hidden = true; };
  ["apLast", "apFirst"].forEach(i => $(i).oninput = () => { if (!state.codeEdited) $("apCode").value = lasfir($("apLast").value, $("apFirst").value); updateApTab(); });
  $("apCode").oninput = () => { state.codeEdited = true; updateApTab(); };
  $("apCourse").oninput = () => { state.courseEdited = true; updateApTab(); };
  $("btnAdd").onclick = addPatient;
  refresh();
});
