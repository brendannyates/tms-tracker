/* core.js: the TMS Tracker floor rules. Pure functions (no Office.js, no DOM), shared with tests/.
   A visit's phase, a chair's state on the Today board, chair moves against NextGen, missed-visit streaks, and the needs list
   that Patients sorts by, and the visit log's own checks (Tx #, duplicate days, lost formulas, the appointment list). taskpane.js builds the inputs from the visit logs, the saved appointment list and the Sessions sheet. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TmsCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const LATE_MIN = 10, SOON_MIN = 3;
  const byTime = (a, b) => (a.time ?? 1e4) - (b.time ?? 1e4);

  /* one Sessions row per visit: date, tab and TMS or F/U ("TMS#2" for a second one that day). Not the log row: an import can re-sort the log. */
  const pool = type => type === "F/U" ? "F/U" : "TMS";
  const sessKey = (date, tab, visit) => `${date}|${tab}|${visit}`;
  /* the visit names for one patient's rows on one day, in time order */
  function visitNames(rows) {
    const seen = {};
    return rows.slice().sort(byTime).map(r => { const k = pool(r.type), n = seen[k] = (seen[k] || 0) + 1; return { r, visit: n > 1 ? `${k}#${n}` : k }; });
  }

  /* missed, then done (pulses entered, or a follow-up marked done), running (started), waiting (checked in), booked */
  function phase(v) {
    if (v.miss) return "missed";
    if (v.done) return "done";
    if (v.start != null) return "running";
    if (v.checkin != null) return "waiting";
    return "booked";
  }

  /* A chair row on Today. open: nothing running; track: running, more than 3 min left; soon: 3 min or less;
     late: the patient booked here checked in and is 10+ min past their time, not started. Patients wait in the lobby, never in a room. */
  function chairState(chair, visits, now) {
    const vs = visits.filter(v => v.chairVisit && v.chair === chair && phase(v) !== "missed").sort(byTime);
    const run = vs.find(v => phase(v) === "running");
    if (run) { const end = run.start + run.len, left = Math.ceil(end - now) || 0; return { state: left <= SOON_MIN ? "soon" : "track", v: run, end, left }; }
    const wait = vs.find(v => phase(v) === "waiting");
    if (wait) { const late = wait.time == null ? 0 : now - wait.time; return { state: late >= LATE_MIN ? "late" : "open", v: wait, late }; }
    return { state: "open", v: null, next: vs.find(v => phase(v) === "booked") || null };
  }
  function chairStatus(cs) {
    if (cs.state === "track") return `${cs.left} min left`;
    if (cs.state === "soon") return cs.left > 0 ? `Ends in ${cs.left} min` : cs.left === 0 ? "Ends now" : `Over by ${-cs.left} min`;
    if (cs.v) return cs.late > 0 ? `Waiting ${cs.late} min` : "Checked in";
    return "";
  }
  /* waiting on staff: checked in, 10+ min past their time, not started (wherever they are listed) */
  const isLate = (v, now) => phase(v) === "waiting" && v.time != null && now - v.time >= LATE_MIN;

  /* Chair moves. rc is the chair NextGen has on the saved appointment list, listAt the minute today's list was run (null if it
     wasn't run today). A Sessions row (s) holds the tracker's chair, the chair NextGen had when it was moved (ng), the minute it
     moved (at) and the "Moved in NextGen" tick. The reminder stays until the tick, a list run after the move showing the new
     chair, or a move back to the chair NextGen has. */
  function moveState(s, rc, listAt) {
    if (!s || !s.chair) return { moved: false, settled: false, chair: rc || "", from: "" };
    const back = s.chair === (s.ng || ""), listed = !!rc && s.chair === rc && listAt != null && s.at != null && listAt > s.at;
    return { moved: !s.tick && !back && !listed, settled: (!!s.tick || listed) && !back, chair: s.chair, from: s.ng || "" };
  }
  /* what to write as NextGen's chair when a visit moves again: once the last move was ticked, or a list run after it showed it,
     NextGen has the chair it was moved to */
  function ngForMove(s, rc, listAt) {
    if (s && s.chair) return moveState(s, rc, listAt).settled ? s.chair : (s.ng || "");
    return rc || "";
  }

  /* missed treatments in a row, most recent first; visits not yet entered (no pulses, not missed) neither count nor break it */
  function missStreak(rows, today) {
    const past = rows.filter(r => r.date != null && r.date <= today && r.type && r.type !== "F/U" && !r.alt && (r.miss || r.delivered))
      .sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
    let n = 0;
    for (let i = past.length - 1; i >= 0 && past[i].miss; i--) n++;
    return n;
  }

  /* Needs for one course, from the checks taskpane.js already makes (p.d) plus today's floor (x). Red (bad) is kept for what blocks
     care or billing: an expired or overrun authorization, no visit in the recent-visit window, 3 missed in a row. Quiet labels
     show muted and don't count. Each need carries the action that resolves it. */
  function needs(p, x) {
    const d = p.d, f = x.fmtDate, n = [], add = (k, t, act, o) => n.push(Object.assign({ k, t, act }, o));
    const onSched = !/^discontinued/i.test(String(p.status).trim());
    if (!onSched) return n;
    const A = d.active, rec = x.recentDays;
    if (d.stale) add("stale", `No visit in ${rec}+ days, last ${d.lastTx != null ? f(d.lastTx) : "none"}`, ["Pause or discontinue", "statuspick"], { bad: true });
    const streak = missStreak(p.rows, x.today);
    if (streak >= 3) add("missed", `Missed ${streak} in a row`, ["Call to reschedule", "call"], { bad: true });
    if (A) d.authFlags.forEach(t => add("auth", t, /^Auth starts|MT redo/.test(t) ? ["Open", "opentab"] : ["Add extension", "ext"], { bad: /expired|past auth|authorized \(D2\)/i.test(t) }));
    (x.moved || []).forEach(v => add("ng", `Moved to ${v.chairLabel}${v.dayLabel ? ` on ${v.dayLabel}` : ""}, not yet in NextGen`, ["Mark moved in NextGen", "ngdone"], { key: v.key }));
    (x.noChair || []).forEach(v => add("nochair", `No chair for ${v.timeLabel}`, ["Choose a chair", "move"], { key: v.key }));
    if (A && d.bidfToday) add("bidf", "BIDF info needed today", ["Enter benefits", "ben"]);
    if (d.suggest) add("status", `Status: ${p.status || "blank"} to ${d.suggest}`, [`Set to ${d.suggest}`, "status"], { v: d.suggest });
    if (d.cancelled) add("cancelled", d.cancelled, ["Update status", "statuspick"]);
    d.decisions.forEach(r => {
      const ex = r.note.includes(x.exemptTag), un = r.note.includes(x.unclassTag);
      add("freebie", ex ? `Exempt in NextGen on ${f(r.date)}: freebie or waived?` : un ? `No-show ${f(r.date)} not classified in NextGen` : `Freebie decision for ${f(r.date)}`, ["Decide freebie", "freebie"], { row: r.row });
    });
    d.overdue.forEach(r => add("pulses", `No pulses for ${f(r.date)}`, ["Enter pulses", "pastpulses"], { row: r.row }));
    d.noTech.forEach(r => add("tech", `No technician on ${f(r.date)}`, ["Add technician", "pasttech"], { row: r.row }));
    (x.noAppt || []).forEach(r => add("noappt", `Pulses on ${f(r.date)}, no visit in NextGen that day`, ["Edit pulses", "editPulses"], { row: r.row }));
    /* earlier days' measures can't be collected now: one reminder to look in NextGen, only while the course is active */
    if (A && d.measDue.length) add("measures", "Check NextGen for PHQ-9 and GAD-7", ["Enter scores", "pastmeas"], { row: d.measDue.reduce((a, b) => b.date > a.date ? b : a).row });
    if (A && d.noProto) add("noproto", "Protocol not confirmed", ["Confirm protocol", "proto"]);
    if (A) d.protoFlags.forEach(t => add("protocol", t, ["Update protocol", "proto"]));
    if (A && d.rollover) add("rollover", `Re-verify benefits for ${d.rollover.year}`, ["Update BIDF", "ben"], { bad: d.rollover.bad });
    /* one line per absence, naming every visit booked inside it */
    if (A) [...new Set(d.awayConflicts.map(c => c.why))].forEach(why => { const cs = d.awayConflicts.filter(c => c.why === why); add("away", `Booked ${cs.map(c => f(c.r.date)).join(", ")} while ${why}`, ["Edit availability", "avail"], { row: cs[0].r.row }); });
    d.closedConflicts.forEach(c => add("closed", `Booked ${f(c.r.date)} on a closure (${c.why})`, ["Open", "opentab"], { row: c.r.row }));
    if (A && d.sameDay.length) add("sameday", `More than one visit on ${d.sameDay.map(f).join(", ")}`, ["Open", "opentab"]);
    if (A && x.mtToday && !x.schedGiven) add("sched", "Course schedule not given", ["Print course schedule", "print"], { v: "all" });
    if (A) d.benFlags.forEach(t => add("cost", t, ["Review benefits", "ben"]));
    if (d.oopmMet) add("oopm", `${d.oopmMet} OOPM met`, null, { quiet: true });
    if (x.mtToday) add("start", "Starts today", null, { quiet: true });
    if (/^tapering$/i.test(p.status)) add("taper", "Tapering", null, { quiet: true });
    if (/^paused$/i.test(p.status)) add("paused", "Paused", null, { quiet: true });
    if (p.external) add("home", `Home clinic ${p.home}`, null, { quiet: true });
    const rank = y => y.quiet ? 2 : y.bad ? 0 : 1;
    return n.map((y, i) => [y, i]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map(a => a[0]);
  }
  const loud = list => list.filter(y => !y.quiet);

  /* ---------- copays: the one "Copay ..." line the add-in keeps in a visit's Notes (S) ----------
     "Copay $25.00 collected by Priya 11:32 AM", "Copay $25.00 not collected by ...", "Copay $0, OOPM met by ...",
     "Copay $25.00 collected (confirmed 10/27 by Priya)", "Copay waived by Priya 10/27". A later answer replaces the line. */
  const COPAY_RE = /^copay\b/i;
  function copayOf(note) {
    const s = String(note ?? "").split("; ").map(x => x.trim()).find(x => COPAY_RE.test(x));
    if (!s) return null;
    const m = s.match(/\$([\d,]+(?:\.\d+)?)/), amt = m ? +m[1].replace(/,/g, "") : null;
    const kind = /oopm met/i.test(s) ? "oopm" : /waived/i.test(s) ? "waived" : /not collected/i.test(s) ? "not" : /collected/i.test(s) ? "collected" : "other";
    return { line: s, kind, amt, by: (s.match(/\bby ([A-Za-z][\w'-]*)/) || [])[1] || "", at: (s.match(/\b(\d{1,2}:\d{2} ?[AP]M)\b/i) || [])[1] || "", confirmed: /confirmed/i.test(s) };
  }
  const withCopayLine = (note, line) => [...String(note ?? "").split("; ").filter(x => x && !COPAY_RE.test(x.trim())), ...(line ? [line] : [])].join("; ");
  const isTreatment = r => !r.miss && !r.alt && ["MT", "Daily", "Taper", "MTR"].includes(r.type);   // TMS Follow Ups (F/U) never count
  /* delivered treatment visits before today, with a payment due (pay(r) > 0) and no copay line yet */
  const copaysToConfirm = (rows, pay, today) => rows.filter(r => isTreatment(r) && r.delivered && r.date != null && r.date < today && !copayOf(r.note) && (pay(r) ?? 0) > 0).sort((a, b) => a.date - b.date);
  /* copays logged as collected on treatment visits dated from..to (the plan year) */
  function copayTally(rows, from, to) {
    let sum = 0, n = 0;
    rows.forEach(r => { if (!isTreatment(r) || r.date == null || r.date < from || r.date > to) return; const c = copayOf(r.note); if (c && c.kind === "collected" && c.amt != null) { sum += c.amt; n++; } });
    return { sum: Math.round(sum * 100) / 100, n };
  }
  /* benefits reminders only matter while copays are still being collected: not once the OOPM is met (0 remaining) or the BIDF quote is $0 */
  const owes = ben => !(ben.oopRem === 0 && !ben.oopNA) && ben.bidf !== 0;
  /* plan-year rollover: a visit booked after Dec 31 of Last updated's year needs benefits re-verified. Amber from Dec 1, red once a
     new-year visit has happened (or checked in) without it. y: { updated, dec1, jan1 } serials for Last updated's year. */
  function rollover(y, rows, today, checkedIn) {
    if (y.updated == null || y.updated >= y.jan1) return null;
    const ahead = rows.some(r => r.date != null && r.date >= y.jan1 && !r.miss);
    if (!ahead || today < y.dec1) return null;
    const late = rows.some(r => r.date != null && r.date >= y.jan1 && r.date <= today && (r.delivered || checkedIn(r)));
    return { bad: late };
  }

  /* ---------- protocol ---------- */
  /* site in O: F3 (left), F4 (right) or another site; the template's Left / Right / Other read the same */
  const siteOf = o => { const s = String(o ?? "").trim(); return /^(f3|left)$/i.test(s) ? "F3" : /^(f4|right)$/i.test(s) ? "F4" : s ? "Other" : ""; };
  const SITE_LABEL = { F3: "F3 (left)", F4: "F4 (right)", Other: "Other" };
  const SITE_CELL = { F3: "Left", F4: "Right", Other: "Other" };     // what O holds (the template's dropdown)
  const pulseCol = site => site === "F3" ? 0 : site === "F4" ? 1 : 2;  // F, H, J on the visit row
  /* a protocol row (J:M) against the Protocols table: the same Hz and values is that protocol; the same Hz with other values is
     "{name} (Custom)" with what changed; anything else "Custom" */
  const PF = [["dur", "pulses", "pulses per train"], ["trains", "trains", "trains"], ["wait", "iti", "s wait"]];
  function protoMatch(x, table) {
    if (typeof x.hz !== "number") return null;
    let best = null;
    for (const q of table.filter(q => q.rep === x.hz)) {
      const diffs = PF.filter(([a, b]) => x[a] !== q[b]).map(([a, b, label]) => ({ k: a, label, v: x[a], std: q[b] }));
      if (!best || diffs.length < best.diffs.length) best = { q, diffs };
    }
    if (!best) return { name: "Custom", base: null, custom: true, diffs: [] };
    const name = String(best.q.name).trim();
    return { name: best.diffs.length ? `${name} (Custom)` : name, base: best.q, custom: best.diffs.length > 0, diffs: best.diffs };
  }

  /* the provider's last name for E2, from the report's name: "Smith, Jane MD" or "Dr. Jane Smith, PMHNP" both give "Smith".
     With a comma, the part before it; otherwise the last word once credentials are dropped. */
  const CRED = /^(md|do|np|pmhnp|pa-c|phd|psyd|dr)$/i;
  function providerLast(name) {
    const parts = String(name ?? "").split(",").map(x => x.trim().split(/\s+/).filter(t => t && !CRED.test(t.replace(/\./g, "")))).filter(a => a.length);
    if (!parts.length) return "";
    return parts.length > 1 ? parts[0].join(" ") : parts[0][parts[0].length - 1];
  }

  /* "Castro MagV1" -> "MagV 1": the clinic's own name is dropped, a chair number gets a space */
  function chairLabel(name, location) {
    let s = String(name || "").trim();
    const loc = String(location || "").trim();
    if (loc && s.toLowerCase().startsWith(loc.toLowerCase())) s = s.slice(loc.length).trim();
    return s.replace(/^(magv|mag\s?venture|brainsway|bw)\s*(\d+)$/i, (m, a, b) => `${/^b/i.test(a) ? (a.length <= 2 ? "BW" : "BrainsWay") : "MagV"} ${b}`) || String(name || "");
  }

  /* ---------- measures placeholders (x in PHQ-9 and GAD-7, P and Q) ----------
     Required on MT, MTR, F/U and the last treatment. Between them, on the 5th treatment after the last measures (MT and MTR count
     as treatments, F/U doesn't), unless a required one is 7 days away or less: one collection covers both, so patients aren't asked
     too often. A placeholder from an earlier day that wasn't collected is cleared. If the next row has scores, they were taken a
     day late and that row is the new start (it gets the Measures flag); if nothing was collected after it, the next visit takes the
     measures and the count restarts there. An x already on the next visit stays, so a second run doesn't move it. Only dated, upcoming, unmissed visits get placeholders; scores are never touched, and
     an x or Measures flag on an upcoming visit the plan doesn't want is removed. Daily, Taper and F/U rows carry the Measures flag
     in O (when O is free); MT and MTR rows only get the x's. Returns [cell, value] pairs. */
  const MEAS_EVERY = 5, MEAS_WINDOW = 7;
  function planMeasures(rows, total, today) {
    rows = rows.filter(r => r.type || typeof r.tx === "number");
    const isTx = r => r.type !== "F/U" && !!r.type && !r.miss;
    const isReq = r => ["MT", "MTR", "F/U"].includes(r.type) || (typeof r.tx === "number" && r.tx === total);
    const upcoming = r => !r.delivered && !r.miss && (r.date == null || r.date >= today);
    const flag = r => ["Daily", "Taper", "F/U"].includes(r.type);
    const xOnly = r => r.measX && !r.measVal;
    const ops = [], clear = r => { if (/^x$/i.test(r.phq)) ops.push([`P${r.row}`, ""]); if (/^x$/i.test(r.gad)) ops.push([`Q${r.row}`, ""]); if (r.action === "Measures") ops.push([`O${r.row}`, ""]); };
    let force = false;
    rows.forEach((r, i) => {
      if (upcoming(r) || !xOnly(r) || r.date == null || r.date >= today) return;
      clear(r);
      const nx = rows[i + 1];
      if (nx && nx.measVal) { if (flag(nx) && !nx.action) ops.push([`O${nx.row}`, "Measures"]); }
      else if (!rows.slice(i + 1).some(x => x.measVal)) force = true;
    });
    const want = new Set();
    let n = 0, first = true;
    rows.forEach((r, i) => {
      if (r.measVal) { n = 0; return; }
      const open = upcoming(r) && r.date != null;
      if (!open) { if (isTx(r)) n++; return; }
      const next = first; first = false;
      if (force || isReq(r) || (next && xOnly(r))) { want.add(r.row); force = false; n = 0; return; }
      if (!isTx(r) || ++n < MEAS_EVERY) return;
      const later = rows.slice(i + 1), ni = later.findIndex(isReq), nxt = ni >= 0 ? later[ni] : null;
      const close = nxt && (nxt.date != null ? nxt.date - r.date <= MEAS_WINDOW : later.slice(0, ni).filter(isTx).length < 2);
      if (!close) { want.add(r.row); n = 0; }
    });
    for (const r of rows) {
      if (!upcoming(r) || r.measVal) continue;
      if (want.has(r.row)) {
        if (!r.phq) ops.push([`P${r.row}`, "x"]);
        if (!r.gad) ops.push([`Q${r.row}`, "x"]);
        if (flag(r) && !r.action) ops.push([`O${r.row}`, "Measures"]);
      } else clear(r);
    }
    return ops;
  }

  /* ---------- the visit log's own checks ---------- */
  /* Tx # (column C), top to bottom: the MT is always 1, and each Daily, Taper or MTR that wasn't missed is one more than the
     treatment above it. Follow-ups and missed visits have none. rows: { type, miss } in row order. */
  const TX_TYPES = ["MT", "Daily", "Taper", "MTR"];
  function txNumbers(rows) {
    let n = 0;
    return rows.map(r => { if (r.miss || !TX_TYPES.includes(r.type)) return ""; n = r.type === "MT" ? 1 : n + 1; return n; });
  }
  const txText = v => typeof v === "number" ? String(v) : String(v ?? "").trim();
  /* New Template's last follow-up row carries this reminder. It isn't needed on a visit: it is template text that comes along
     with the row when an import gives the row a date */
  const KAISER_NOTE = "**Send final progress note/summary to Kaiser**";
  const dropKaiser = note => String(note || "").split("; ").filter(x => x.trim() !== KAISER_NOTE).join("; ");
  /* rows: the parsed log (row, date, type, miss, delivered, tx, note, noA, noD). o: { pr (the tab has the PR formula), approved (D2) }.
     dup: the same day and kind of visit on two or more rows, unless every one was treated (an accelerated day).
     tx: rows whose Tx # isn't the count. order: dated rows above an earlier date. formulas: rows that lost the PR formula in A
     or the Tx-left formula in D. kaiser: dated rows still holding the template's Kaiser reminder. over: dated treatments numbered
     past # of TXs (D2), for example a 37th booked on a 36-treatment course. */
  function logChecks(rows, o = {}) {
    const groups = new Map();
    rows.forEach(r => { if (r.date == null) return; const k = `${r.date}|${pool(r.type)}`; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); });
    const dup = [...groups.values()].filter(g => g.length > 1 && g.filter(r => r.delivered).length < g.length).map(g => ({ date: g[0].date, pool: pool(g[0].type), rows: g.map(r => r.row) }));
    const want = txNumbers(rows), tx = rows.map((r, i) => ({ row: r.row, was: txText(r.tx), want: want[i] })).filter(x => x.was !== String(x.want));
    let last = -Infinity; const order = [];
    rows.forEach(r => { if (r.date == null) return; if (r.date < last) order.push(r.row); else last = r.date; });
    const formulas = rows.filter(r => (o.pr && r.noA) || r.noD).map(r => ({ row: r.row, a: !!(o.pr && r.noA), d: !!r.noD }));
    const kaiser = rows.filter(r => r.date != null && dropKaiser(r.note) !== String(r.note || "")).map(r => r.row);
    const over = o.approved > 0 ? rows.map((r, i) => ({ row: r.row, date: r.date, n: want[i] })).filter(x => x.date != null && typeof x.n === "number" && x.n > o.approved) : [];
    return { dup, tx, order, formulas, kaiser, over, n: dup.length + (tx.length ? 1 : 0) + (order.length ? 1 : 0) + (formulas.length ? 1 : 0) + (kaiser.length ? 1 : 0) + (over.length ? 1 : 0) };
  }
  /* the saved appointment list against one tab. appts: { date, time, status (Scheduled, Kept, No-show, Cancelled), type } for this
     patient. Booked from today on (o.from) and kept since the course's first dated row each need a row that day; an open row from
     today through the list's last day (o.to) needs a booking. The same visit listed twice (day, type, time) counts once. */
  function apptGaps(rows, appts, o) {
    const dated = rows.filter(r => r.date != null), start = dated.length ? Math.min(...dated.map(r => r.date)) : o.from, seen = new Set(), want = new Map();
    for (const a of appts) {
      if (!(a.status === "Scheduled" && a.date >= o.from) && !(a.status === "Kept" && a.date >= start)) continue;
      const k = `${a.date}|${a.type}|${a.time}`; if (seen.has(k)) continue; seen.add(k);
      const g = `${a.date}|${pool(a.type)}`; if (!want.has(g)) want.set(g, []); want.get(g).push(a);
    }
    const on = (d, p) => rows.filter(r => r.date === d && pool(r.type) === p && !r.miss && !r.alt), notOnTab = [], notOnList = [];
    for (const [g, list] of want) { const [d, p] = g.split("|"); list.slice(on(+d, p).length).forEach(a => notOnTab.push(a)); }
    const days = new Set(dated.filter(r => r.date >= o.from && r.date <= o.to && !r.miss && !r.alt && !r.delivered).map(r => `${r.date}|${pool(r.type)}`));
    for (const g of days) {
      const [d, p] = g.split("|"), rs = on(+d, p).sort((x, y) => x.delivered - y.delivered), n = (want.get(g) || []).length;
      rs.slice(0, Math.max(0, rs.length - n)).filter(r => !r.delivered).forEach(r => notOnList.push(r.row));
    }
    return { notOnTab: notOnTab.sort((a, b) => a.date - b.date), notOnList };
  }
  /* pulses on a day the list covers (lo through today) with no visit of that kind in NextGen that day, cancelled ones aside:
     a treatment NextGen can't bill, or pulses entered on the wrong row. appts: every booking on the list for the patient. Rows. */
  function pulsesNoAppt(rows, appts, lo, today) {
    if (lo == null) return [];
    return rows.filter(r => r.delivered && !r.alt && r.date != null && r.date >= lo && r.date <= today
      && !appts.some(a => a.date === r.date && pool(a.type) === pool(r.type) && a.status !== "Cancelled"));
  }
  /* the clinic in a tracker's name: "Castro MagV Patient Tracker" -> "Castro", "Walnut Creek TMS Pt Tracker 2026" -> "Walnut Creek" */
  function clinicFromName(name) {
    const s = String(name || "").replace(/\.(xlsx|xlsm|xls)$/i, "").replace(/[_-]+/g, " ");
    const m = s.match(/^(.*?)\s+(magv|mag\s?venture|brainsway|bw|tms|esk|ivk|pt|patient|tracker)\b/i);
    return (m ? m[1] : "").replace(/\s{2,}/g, " ").trim();
  }
  /* which clinic carries a patient's tab. visits: { chair, here, where, type } per booking (no cancellations), where chair is a
     device resource, here is a resource or location naming this tracker's clinic (null when the report doesn't say), and where names
     any other clinic. The MT and MT redo decide: the clinic holding most of them is home (F/U don't count). With them elsewhere, a
     patient with ASK_MIN or more daily visits here is "ask" (offer a tracker here), fewer is a "guest", none is left out. With no MT
     or MT redo on the list, the clinic with the most chair visits carries it, a tie staying here.
     role: "home" (a tab here), "ask", "guest" (tracked at home, which names the other clinic) or "none" (left out). */
  const ASK_MIN = 16;
  function clinicRole(visits) {
    const chair = visits.filter(v => v.chair), here = chair.filter(v => v.here).length, by = {};
    chair.filter(v => !v.here).forEach(v => { const w = v.where || "Another clinic"; by[w] = (by[w] || 0) + 1; });
    const most = list => { const c = {}; list.forEach(v => { const w = v.where || "Another clinic"; c[w] = (c[w] || 0) + 1; }); return Object.entries(c).sort((a, b) => b[1] - a[1])[0] || ["", 0]; };
    const top = most(chair.filter(v => !v.here)), daily = chair.filter(v => v.here && (v.type === "Daily" || v.type === "Taper")).length;
    const out = (role, home) => ({ role, here, daily, home: home || "", homeCount: home ? by[home] || 0 : top[1], by });
    const mts = visits.filter(v => (v.type === "MT" || v.type === "MTR") && v.here != null), mtOut = mts.filter(v => !v.here);
    if (mts.length) {
      if (mts.length - mtOut.length >= mtOut.length) return out("home");
      const away = most(mtOut)[0];
      return !here ? out("none", away) : daily >= ASK_MIN ? out("ask", away) : out("guest", away);
    }
    if (!chair.length) return visits.some(v => v.here) || !visits.some(v => v.here === false) ? out("home") : out("none", most(visits)[0]);
    if (!here) return out("none", top[0]);
    return here >= top[1] ? out("home") : out("guest", top[0]);
  }
  /* whether a patient with no tab here is offered one (the New patients banner). appts: { date, status, type, here } per booking,
     here false when it names another clinic. An MT or MT redo here (kept or booked), or ASK_MIN or more treatments booked here from
     today on. F/U alone never does, nor does an MT at another clinic. */
  function needsTracker(appts, from) {
    const ok = appts.filter(a => a.status !== "No-show" && a.status !== "Cancelled" && a.here !== false), fut = ok.filter(a => a.date >= from);
    return ok.some(a => a.type === "MT" || a.type === "MTR") || fut.filter(a => a.type === "Daily" || a.type === "Taper").length >= ASK_MIN;
  }
  /* a patient with no tab here and no tracker due (needsTracker) but treatments booked here is a guest: tracked at their home clinic,
     whose tab link the Guest Patients banner asks for. F/U alone stays as it was. g: clinicRole's answer, or { role: "home" }. */
  function noTabRole(g, visits, from) {
    const tx = visits.filter(v => v.here !== false && (v.type === "Daily" || v.type === "Taper"));
    if (g.role !== "home" || !tx.length || needsTracker(visits, from)) return g;
    const top = Object.entries(g.by || {}).sort((a, b) => b[1] - a[1])[0];
    return { ...g, role: "guest", here: tx.length, daily: tx.length, home: top ? top[0] : "", homeCount: top ? top[1] : 0, by: g.by || {} };
  }
  /* what a tracker cell holds, for the rebuild's side-by-side check: "A6" -> { label: "DOB", date: true }; a log cell names its
     visit row ("E12" -> "Date, log row 3"). Unknown cells are labeled by address. */
  const CELL_HEAD = { A2: "LASFIR", A4: "MRN", A6: ["DOB"], D2: "# of TXs", D4: "# of MT redos", D6: "MD required on-site", D7: "Status",
    E2: "Provider", E4: "Care navigator", E6: "Insurance", F4: ["Auth start"], G4: ["Auth expiry"], G2: ["Update protocol by"], G6: ["Progress note sent"],
    N6: "Freebie count", N7: "Auth assessment", S9: "Home clinic" };
  const CELL_LOG = { B: "Technician", C: "Tx #", E: ["Date"], F: "Pulses left", H: "Pulses right", J: "Pulses other", L: "Visit type", M: "Miss",
    N: "Miss", O: "Action", P: "PHQ-9", Q: "GAD-7", R: "Alt clinic", S: "Notes", T: "Time" };
  const CELL_PROT = { J: "Hz", K: "Pulses per train", L: "Trains", M: "Wait", O: "Site" };
  function cellLabel(a) {
    const m = String(a).match(/^([A-Z]+)(\d+)$/); if (!m) return { label: String(a), date: false };
    const c = m[1], r = +m[2], one =x => ({ label: [].concat(x)[0], date: Array.isArray(x) });
    if (CELL_HEAD[a]) return one(CELL_HEAD[a]);
    if (r >= 2 && r <= 4 && CELL_PROT[c]) return { label: `Protocol row ${r - 1}: ${CELL_PROT[c]}`, date: false };
    if (r >= 10 && r <= 49 && CELL_LOG[c]) { const x = one(CELL_LOG[c]); return { ...x, label: `${x.label}, log row ${r - 9}` }; }
    return { label: a, date: false };
  }
  /* a short fingerprint of a log's typed cells, so a saved state can be recognised later (an undo only runs on the state it left) */
  function hashRows(rows) {
    const s = JSON.stringify(rows); let h = 5381;
    for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
    return `${rows.length}:${(h >>> 0).toString(36)}`;
  }

  return { LATE_MIN, SOON_MIN, pool, sessKey, visitNames, txNumbers, KAISER_NOTE, dropKaiser, logChecks, apptGaps, pulsesNoAppt, hashRows, phase, chairState, chairStatus, isLate, moveState, ngForMove, missStreak, needs, loud, chairLabel, planMeasures, providerLast, clinicFromName, clinicRole, ASK_MIN, needsTracker, noTabRole, cellLabel,
    copayOf, withCopayLine, isTreatment, copaysToConfirm, copayTally, owes, rollover, siteOf, SITE_LABEL, SITE_CELL, pulseCol, protoMatch };
});
