/* core.js: the TMS Tracker floor rules. Pure functions (no Office.js, no DOM), shared with tests/.
   A visit's phase, a chair's state on the Today board, chair moves against NextGen, missed-visit streaks, and the needs list
   that Patients sorts by. taskpane.js builds the inputs from the visit logs, the saved appointment list and the Sessions sheet. */
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
    if (A) d.authFlags.forEach(t => add("auth", t, /^Auth starts|MT redo/.test(t) ? ["Open tab", "opentab"] : ["Add extension", "ext"], { bad: /expired|past auth|authorized \(D2\)/i.test(t) }));
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
    /* earlier days' measures can't be collected now: one reminder to look in NextGen, only while the course is active */
    if (A && d.measDue.length) add("measures", "Check NextGen for PHQ-9 and GAD-7", ["Enter scores", "pastmeas"], { row: d.measDue.reduce((a, b) => b.date > a.date ? b : a).row });
    if (A) d.protoFlags.forEach(t => add("protocol", t, ["Update protocol", "proto"]));
    /* one line per absence, naming every visit booked inside it */
    if (A) [...new Set(d.awayConflicts.map(c => c.why))].forEach(why => { const cs = d.awayConflicts.filter(c => c.why === why); add("away", `Booked ${cs.map(c => f(c.r.date)).join(", ")} while ${why}`, ["Edit availability", "avail"], { row: cs[0].r.row }); });
    d.closedConflicts.forEach(c => add("closed", `Booked ${f(c.r.date)} on a closure (${c.why})`, ["Open tab", "opentab"], { row: c.r.row }));
    if (A && d.sameDay.length) add("sameday", `More than one visit on ${d.sameDay.map(f).join(", ")}`, ["Open tab", "opentab"]);
    if (A && x.mtToday && !x.schedGiven) add("sched", "Course schedule not given", ["Print course schedule", "print"], { v: "all" });
    if (A) d.benFlags.forEach(t => add("cost", t, ["Review benefits", "ben"]));
    if (x.mtToday) add("start", "Starts today", null, { quiet: true });
    if (/^tapering$/i.test(p.status)) add("taper", "Tapering", null, { quiet: true });
    if (/^paused$/i.test(p.status)) add("paused", "Paused", null, { quiet: true });
    if (p.external) add("home", `Home clinic ${p.home}`, null, { quiet: true });
    const rank = y => y.quiet ? 2 : y.bad ? 0 : 1;
    return n.map((y, i) => [y, i]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map(a => a[0]);
  }
  const loud = list => list.filter(y => !y.quiet);

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

  return { LATE_MIN, SOON_MIN, pool, sessKey, visitNames, phase, chairState, chairStatus, isLate, moveState, ngForMove, missStreak, needs, loud, chairLabel, planMeasures, providerLast };
});
