/* core.js: the TMS Tracker floor rules. Pure functions (no Office.js, no DOM), shared with tests/.
   A visit's phase, a chair's state on the Today board, chair moves against NextGen, missed-visit streaks, and the needs list
   that Courses sorts by. taskpane.js builds the inputs from the visit logs, the saved appointment list and the Sessions sheet. */
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
  /* what to write as NextGen's chair when a visit moves again: after a tick, NextGen has the chair it was moved to */
  function ngForMove(s, rc) {
    if (s && s.chair) return s.tick ? s.chair : (s.ng || "");
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
    (x.moved || []).forEach(v => add("ng", `Moved to ${v.chairLabel}, not yet in NextGen`, ["Mark moved in NextGen", "ngdone"], { key: v.key }));
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
    if (d.measDue.length) add("measures", `PHQ-9 and GAD-7 not collected ${f(d.measDue[0].date)}${d.measDue.length > 1 ? ` and ${d.measDue.length - 1} more` : ""}`, ["Enter scores", "pastmeas"], { row: d.measDue[0].row });
    if (A) d.protoFlags.forEach(t => add("protocol", t, ["Update protocol", "proto"]));
    /* one line per absence, naming every visit booked inside it */
    if (A) [...new Set(d.awayConflicts.map(c => c.why))].forEach(why => { const cs = d.awayConflicts.filter(c => c.why === why); add("away", `Booked ${cs.map(c => f(c.r.date)).join(", ")} while ${why}`, ["Edit availability", "avail"], { row: cs[0].r.row }); });
    d.closedConflicts.forEach(c => add("closed", `Booked ${f(c.r.date)} on a closure (${c.why})`, ["Open tab", "opentab"], { row: c.r.row }));
    if (A && d.sameDay.length) add("sameday", `More than one visit on ${d.sameDay.map(f).join(", ")}`, ["Open tab", "opentab"]);
    if (A && x.mtToday && !x.schedGiven) add("sched", "Course schedule not given", ["Print course schedule", "print"]);
    if (A) d.benFlags.forEach(t => add("cost", t, ["Review benefits", "ben"]));
    if (x.mtToday) add("start", "Starts today", null, { quiet: true });
    if (/^tapering$/i.test(p.status)) add("taper", "Tapering", null, { quiet: true });
    if (/^paused$/i.test(p.status)) add("paused", "Paused", null, { quiet: true });
    if (p.external) add("home", `Home clinic ${p.home}`, null, { quiet: true });
    const rank = y => y.quiet ? 2 : y.bad ? 0 : 1;
    return n.map((y, i) => [y, i]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map(a => a[0]);
  }
  const loud = list => list.filter(y => !y.quiet);

  /* "Castro MagV1" -> "MagV 1": the clinic's own name is dropped, a chair number gets a space */
  function chairLabel(name, location) {
    let s = String(name || "").trim();
    const loc = String(location || "").trim();
    if (loc && s.toLowerCase().startsWith(loc.toLowerCase())) s = s.slice(loc.length).trim();
    return s.replace(/^(magv|mag\s?venture|brainsway|bw)\s*(\d+)$/i, (m, a, b) => `${/^b/i.test(a) ? (a.length <= 2 ? "BW" : "BrainsWay") : "MagV"} ${b}`) || String(name || "");
  }

  return { LATE_MIN, SOON_MIN, pool, sessKey, visitNames, phase, chairState, chairStatus, isLate, moveState, ngForMove, missStreak, needs, loud, chairLabel };
});
