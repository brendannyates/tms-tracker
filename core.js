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

  /* NextGen's event names: "TMS Treatment Unilateral", "TMS Theta Tx Bilateral" (theta burst says Tx, not Treatment), "MT MagV",
     "MT Redo", "TMS Follow Up", "Telepsych TMS F/U". Anything else is not a TMS visit (null). */
  function mapEvent(ev) {
    const s = String(ev ?? "").toLowerCase();
    if (/tele/.test(s)) return { type: "F/U", tele: true };
    if (/follow|f\/u/.test(s)) return { type: "F/U", tele: false };
    if (/redo|mtr/.test(s)) return { type: "MTR", tele: false };
    if (/\bmt\b/.test(s)) return { type: "MT", tele: false };
    if (/treatment|\btx\b|theta|\btbs\b/.test(s)) return { type: "Daily", tele: false };
    return null;
  }

  /* missed treatments in a row, most recent first; visits not yet entered (no pulses, not missed) neither count nor break it */
  const pastTx = (rows, today) => rows.filter(r => r.date != null && r.date <= today && r.type && r.type !== "F/U" && !r.alt && (r.miss || r.delivered))
    .sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
  function missRun(rows, today) {
    const past = pastTx(rows, today), run = [];
    for (let i = past.length - 1; i >= 0 && past[i].miss; i--) run.unshift(past[i]);
    return run;
  }
  const missStreak = (rows, today) => missRun(rows, today).length;

  /* ---------- the rendering clinician, with their credential: not everyone is a doctor ----------
     NextGen writes "Rivera MD, Ana", "Chen, MD Sam", "Okafor PMHNP, Jo", "Berg, Lena PMHNP", "Singh Md, Ravi Kumar".
     clinician() gives { first, last, cred, name: "Ana Rivera, MD" }; with no credential the name alone. */
  const CREDS = /^(md|do|np|pmhnp|pmhnp-bc|fnp|fnp-c|fnp-bc|aprn|dnp|pa|pa-c|phd|psyd|lcsw|lmft|mft|rn|bc)$/i;
  function clinician(raw) {
    const s = String(raw ?? "").replace(/\s+/g, " ").trim(); if (!s) return { first: "", last: "", cred: "", name: "" };
    const cred = [], words = part => part.split(" ").map(w => w.replace(/[.]/g, "")).filter(w => {
      if (!w || /^dr$/i.test(w)) return false;
      if (CREDS.test(w)) { const c = w.toUpperCase(); if (!cred.includes(c)) cred.push(c); return false; }
      return true; });
    const parts = s.split(",").map(x => words(x.trim()));
    let last, first;
    if (parts.length > 1) { last = parts[0].join(" "); first = parts.slice(1).flat().join(" "); }
    else { const w = parts[0]; last = w.length ? w[w.length - 1] : ""; first = w.slice(0, -1).join(" "); }
    const name = `${[first, last].filter(Boolean).join(" ")}${cred.length ? `, ${cred.join(", ")}` : ""}`;
    return { first, last, cred: cred.join(", "), name };
  }
  const clinicianName = raw => clinician(raw).name;

  /* ---------- telephone notes to the clinician (written in NextGen; the add-in logs that it was done in Notes, S) ----------
     prov is E2, the rendering clinician with their credential ("Ana Rivera, MD"). Missed treatments: each missed row covered gets
     "Tel note to Ana Rivera, MD 11/11 (missed 2 in a row)"; a later miss gets "Tel note entry 11/12 (missed 3 in a row)" on its
     own row. "(cancelled 11/10, 11/12)" is for 2+ treatment cancellations within 7 days; "(F/U 11/6 not rebooked)" for a follow-up
     missed or cancelled with none booked within 5 days. */
  const TEL_MISS = /^tel note (?:to|entry)\b[^;]*\(missed (\d+)(?: in a row)?\)/i, TEL_CX = /^tel note (?:to|entry)\b[^;]*\(cancelled ([^)]*)\)/i;
  const telTo = prov => `Tel note to ${prov || "the clinician"}`;
  const telMissLine = (prov, day, n, entry) => entry ? `Tel note entry ${day} (missed ${n} in a row)` : `${telTo(prov)} ${day} (missed ${n} in a row)`;
  const telFuLine = (prov, day, fuDay) => `${telTo(prov)} ${day} (F/U ${fuDay} not rebooked)`;
  const telCxLine = (prov, day, days, entry) => entry ? `Tel note entry ${day} (cancelled ${days.join(", ")})` : `${telTo(prov)} ${day} (cancelled ${days.join(", ")})`;
  const notesOf = r => String(r.note ?? "").split("; ").map(x => x.trim());
  /* two or more missed treatments in a row need a telephone note to the provider; each miss after that, a new entry on it.
     Returns null under 2, else { n, logged (the miss count the last logged line covered), due, first, run, add (misses not yet on it) } */
  function telNote(rows, today) {
    const run = missRun(rows, today), n = run.length;
    if (n < 2) return null;
    let logged = 0;
    run.forEach(r => notesOf(r).forEach(s => { const m = s.match(TEL_MISS); if (m) logged = Math.max(logged, +m[1]); }));
    logged = Math.min(logged, n);
    return { n, logged, due: logged < n, first: logged === 0, run, add: run.slice(logged) };
  }
  /* business days after from, up to to (serials; Saturday and Sunday don't count) */
  const wday = s => (s + 6) % 7;   // 0 Sunday, for serials counted from Dec 30, 1899
  function bizDays(from, to) { let n = 0; for (let d = from + 1; d <= to; d++) { const w = wday(d); if (w > 0 && w < 6) n++; } return n; }
  /* 2 or more treatment cancellations within a rolling 7 days (by appointment date) need a telephone note to the clinician; each
     cancellation after that, within 7 days of the one before, a new entry. cx: cancelled treatment dates from the saved list (a
     same-day rebooking already left out). lbl: serial -> "m/d", as the logged lines write it. Cancellations from 30 days back on.
     Returns the runs still due: [{ dates, add (not yet logged), first }] */
  const CX_WINDOW = 7, CX_LOOK = 30;
  function cancelNote(rows, cx, today, lbl) {
    const ds = [...new Set((cx || []).filter(d => d != null && d >= today - CX_LOOK))].sort((a, b) => a - b), runs = [];
    ds.forEach(d => { const r = runs[runs.length - 1]; if (r && d - r[r.length - 1] < CX_WINDOW) r.push(d); else runs.push([d]); });
    const logged = new Set();
    rows.flatMap(notesOf).forEach(s => { const m = s.match(TEL_CX); if (m) m[1].split(",").map(x => x.trim()).forEach(x => logged.add(x)); });
    return runs.filter(r => r.length >= 2).map(r => ({ dates: r, add: r.filter(d => !logged.has(lbl(d))), first: r.every(d => !logged.has(lbl(d))) })).filter(r => r.add.length);
  }
  /* a follow-up missed (x in M) or cancelled in NextGen counts unless the next follow-up is booked within FU_NEXT days after it:
     rebook it; after REBOOK_DAYS business days still not rebooked within that, a telephone note goes to the clinician. cx:
     cancelled F/U from the saved list { date, mod }. lbl: serial -> "m/d", to find the logged line. Only the last 30 days count.
     next: the next follow-up after it, if any (booked too late to count). */
  const REBOOK_DAYS = 2, REBOOK_LOOK = 30, FU_NEXT = 5;
  function fuRebook(rows, cx, today, lbl) {
    const evs = rows.filter(r => r.type === "F/U" && r.miss && r.date != null && r.date <= today).map(r => ({ date: r.date, kind: "missed", start: r.date, row: r.row }));
    (cx || []).forEach(c => {
      if (c.date == null || evs.some(e => e.date === c.date)) return;
      if (rows.some(r => r.type === "F/U" && !r.miss && r.date === c.date)) return;   // rebooked the same day
      evs.push({ date: c.date, kind: "cancelled", start: Math.min(today, c.mod ?? c.date), row: null });   // the clock starts when it was cancelled
    });
    const fu = rows.filter(r => r.type === "F/U" && !r.miss && r.date != null), all = rows.flatMap(notesOf);
    return evs.filter(e => e.start >= today - REBOOK_LOOK && !fu.some(r => r.date > e.date && r.date <= e.date + FU_NEXT))
      .map(e => ({ ...e, next: (fu.filter(r => r.date > e.date).sort((a, b) => a.date - b.date)[0] || {}).date ?? null, telDue: bizDays(e.start, today) >= REBOOK_DAYS, logged: all.some(s => /^tel note to\b/i.test(s) && s.includes(`(F/U ${lbl(e.date)} not rebooked)`)) }))
      .sort((a, b) => a.date - b.date);
  }
  /* follow-ups through the course: one between the MT and the MT redo (or treatment 15 with no MT redo booked), one around
     treatments 20 to 27, and one within 7 days of the last treatment (D2), either side. A window counts once its treatments are
     dated; a gap shows only while the window hasn't passed. Returns [{ k, label, from, to }] */
  const FU_END = 7;
  function fuGaps(rows, total, today) {
    const live = r => r.date != null && !r.miss;
    const txd = rows.filter(r => live(r) && ["MT", "Daily", "Taper", "MTR"].includes(r.type)).sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
    const mt = txd.find(r => r.type === "MT"); if (!mt) return [];
    const at = k => { const r = txd.slice(txd.indexOf(mt))[k - 1]; return r ? r.date : null; };
    const mtr = txd.find(r => r.type === "MTR" && r.date > mt.date), fus = rows.filter(r => live(r) && r.type === "F/U");
    const win = [];
    const aTo = mtr ? mtr.date : at(15);
    if (aTo != null) win.push({ k: "mtr", label: mtr ? "between the MT and the MT redo" : "between the MT and treatment 15", from: mt.date, to: aTo });
    if (total >= 27 && at(20) != null) win.push({ k: "mid", label: "around treatments 20 to 27", from: at(20), to: at(27) ?? at(20) + 9 });
    const last = at(total);
    if (last != null) win.push({ k: "end", label: `within ${FU_END} days of the last treatment`, from: last - FU_END, to: last + FU_END });
    return win.filter(w => w.to >= today && !fus.some(f => f.date >= w.from && f.date <= w.to));
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
    /* telephone notes to the provider: 2+ missed in a row, then an entry for each miss after; a follow-up not rebooked */
    const dr = p.provider || "the clinician", tel = d.tel;
    if (tel && tel.due) add("telnote", tel.first ? `Missed ${tel.n} in a row: telephone note to ${dr}` : `Missed again ${tel.add.map(r => f(r.date)).join(", ")}: add to the telephone note`,
      [tel.first ? "Write telephone note" : "Add to telephone note", "telnote"], { bad: true });
    (d.fuRebook || []).forEach(e => {
      if (e.telDue && e.logged) return;
      const nx = e.next != null ? `, next one ${f(e.next)}` : "";
      add("furebook", e.telDue ? `Follow-up ${f(e.date)} ${e.kind}, not rebooked within ${FU_NEXT} days${nx}: telephone note to ${dr}` : `Follow-up ${f(e.date)} ${e.kind}${nx}: rebook within ${FU_NEXT} days`,
        e.telDue ? ["Write telephone note", "telfu"] : ["Rebook in NextGen", "opentab"], { bad: e.telDue, v: e.date, row: e.row });
    });
    (d.cxNote || []).forEach(c => add("telcx", c.first ? `Cancelled ${c.dates.map(f).join(", ")} (${c.dates.length} within ${CX_WINDOW} days): telephone note to ${dr}` : `Cancelled again ${c.add.map(f).join(", ")}: add to the telephone note`,
      [c.first ? "Write telephone note" : "Add to telephone note", "telcx"], { bad: true, v: c.dates[0] }));
    if (A) (d.fuGaps || []).forEach(w => add("fugap", `No follow-up booked ${w.label} (${f(w.from)} to ${f(w.to)})`, ["Book in NextGen", "opentab"], { v: w.k }));
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
    const mis = x.misdated || [];
    mis.forEach(m => add("misdated", `Pulses on ${f(m.date)}, NextGen has the visit on ${f(m.to)}`, [`Move to ${f(m.to)}`, "movedate"], { row: m.row, v: m.to }));
    (x.noAppt || []).filter(r => !mis.some(m => m.row === r.row)).forEach(r => add("noappt", `Pulses on ${f(r.date)}, no visit in NextGen that day`, ["Edit pulses", "editPulses"], { row: r.row }));
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
  /* card and consent: the one line a technician leaves on a visit's Notes after asking during the session.
     "LumaPay card on file, verbal copay consent, checked by Priya 9:42 AM"; either half can be "No LumaPay card on file" /
     "no verbal copay consent". A later answer replaces the line. */
  const CONSENT_RE = /^(no )?lumapay card\b/i;
  const consentLine = (card, ok, by, at) => `${card ? "LumaPay card on file" : "No LumaPay card on file"}, ${ok ? "verbal copay consent" : "no verbal copay consent"}${by ? `, checked by ${by}${at ? ` ${at}` : ""}` : ""}`;
  function consentOf(note) {
    const s = String(note ?? "").split("; ").map(x => x.trim()).find(x => CONSENT_RE.test(x));
    if (!s) return null;
    return { line: s, card: !/^no /i.test(s), ok: !/\bno verbal copay consent\b/i.test(s), by: (s.match(/\bchecked by ([A-Za-z][\w'-]*)/) || [])[1] || "", at: (s.match(/\b(\d{1,2}:\d{2} ?[AP]M)\b/i) || [])[1] || "" };
  }
  const withConsentLine = (note, line) => [...String(note ?? "").split("; ").filter(x => x && !CONSENT_RE.test(x.trim())), ...(line ? [line] : [])].join("; ");
  /* a Benefits-tab flag cell: Yes (or x, TRUE, a tick) is ticked; blank or a missing column is not */
  const flagOn = v => v === true || /^(y|yes|x|true|✓|✔)$/i.test(String(v ?? "").trim());
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
  /* pulses a protocol row gives per session: pulses per train x trains, scaled by the Protocols table's Total Pulses where it says
     more. Theta burst lists bursts per train (3 pulses each), so iTBS 600 is 10 x 20 x 3. A custom row takes its base protocol's ratio. */
  function protoPulses(x, table) {
    if (typeof x.dur !== "number" || typeof x.trains !== "number") return 0;
    const m = protoMatch(x, table || []), q = m && m.base, per = q ? q.pulses * q.trains : 0;
    return Math.round(x.dur * x.trains * (typeof q?.total === "number" && per > 0 ? q.total / per : 1));
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
  /* a patient tab's name: "DOEJAN", "DOEJAN 4", and the same written differently by hand ("DOEJAN (4)", "ROEBOB(2)", "LEEANN #2",
     "KIMSAM2", "TAYMAX " with a trailing space, "ParKim" in mixed case). Returns { code, course, std } or null; std is false for the
     hand-written forms, which are read only when the tab's layout is the tracker's. */
  function tabName(name) {
    const s = String(name ?? ""), m = s.match(/^([A-Z]{1,6})(?: (\d+))?$/);
    if (m) return { code: m[1], course: +(m[2] || 1), std: true };
    const l = s.trim().match(/^([A-Za-z]{3,6})\s*(?:\(\s*(\d{1,2})\s*\)|#\s*(\d{1,2})|(\d{1,2}))?$/);
    if (!l) return null;
    return { code: l[1].toUpperCase(), course: +(l[2] || l[3] || l[4] || 1), std: false };
  }
  /* pulses entered a day or so off: a treated row on a day NextGen has no visit of that kind, and a kept NextGen visit of that kind
     within MISDATED_DAYS calendar days with no row of its own. Each kept visit pairs with the nearest such row once. These are
     usually pulses typed on the wrong row; the import leaves them for someone to check instead of adding a second row for the
     NextGen day. appts: { date, type, status } from the list; lo..hi: the days the list covers, through today. Returns
     [{ row, date, to, type }] */
  const MISDATED_DAYS = 3;
  function misdated(rows, appts, lo, hi) {
    if (lo == null) return [];
    const on = (list, d, p) => list.some(x => x.date === d && pool(x.type) === p);
    const ng = (appts || []).filter(a => a.status !== "Cancelled" && a.date != null);
    const off = rows.filter(r => r.delivered && !r.alt && r.date != null && r.date >= lo && r.date <= hi && !on(ng, r.date, pool(r.type)))
      .sort((a, b) => a.date - b.date);
    const free = ng.filter(a => a.status === "Kept" && a.date >= lo && a.date <= hi && !rows.some(r => r.date === a.date && pool(r.type) === pool(a.type) && !r.miss));
    /* nearest pairs first, so a kept visit goes to the row closest to it */
    const pairs = [];
    off.forEach(r => free.forEach(a => { const g = Math.abs(a.date - r.date); if (pool(a.type) === pool(r.type) && g <= MISDATED_DAYS) pairs.push({ r, a, g }); }));
    pairs.sort((x, y) => x.g - y.g || x.r.date - y.r.date);
    const usedA = new Set(), usedR = new Set(), out = [];
    for (const { r, a } of pairs) {
      const k = `${a.date}|${pool(a.type)}`; if (usedA.has(k) || usedR.has(r.row)) continue;
      usedA.add(k); usedR.add(r.row); out.push({ row: r.row, date: r.date, to: a.date, type: r.type });
    }
    return out.sort((x, y) => x.date - y.date);
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
    const mtHere = mts.length > 0 && mts.length - mtOut.length >= mtOut.length;
    if (!chair.length) {
      if (mts.length) return mtHere ? out("home") : out("none", most(mtOut)[0]);
      return visits.some(v => v.here) || !visits.some(v => v.here === false) ? out("home") : out("none", most(visits)[0]);
    }
    /* most treatments booked at another clinic: the tab is there, whatever the MT */
    if (top[1] > here) return here ? out("guest", top[0]) : out("none", top[0]);
    if (!mts.length || mtHere) return out("home");
    /* the MT elsewhere, most treatments here: a tie stays with the MT; 16+ here is a course booked here (a tracker here), fewer a guest */
    const away = most(mtOut)[0];
    if (top[1] === here) return out("guest", away);
    return daily >= ASK_MIN ? out("ask", away) : out("guest", away);
  }
  /* whether a patient with no tab here is offered one (the New patients banner). appts: { date, status, type, here } per booking,
     here false when it names another clinic. An MT or MT redo here (kept or booked), or ASK_MIN or more treatments booked here from
     today on. F/U alone never does, nor does an MT at another clinic. */
  function needsTracker(appts, from) {
    const ok = appts.filter(a => a.status !== "No-show" && a.status !== "Cancelled" && a.here !== false), fut = ok.filter(a => a.date >= from);
    const tx = a => a.type === "Daily" || a.type === "Taper";
    /* an MT here, 16+ treatments booked from today, or 16+ kept and booked here in all (a course mostly done here) */
    return ok.some(a => a.type === "MT" || a.type === "MTR") || fut.filter(tx).length >= ASK_MIN || ok.filter(tx).length >= ASK_MIN;
  }
  /* a patient with no tab here and no tracker due (needsTracker) but treatments booked here is a guest: intermittent treatments
     here, tracked at their home clinic. F/U alone stays as it was. Follow-ups never decide (they're on the provider's schedule,
     often at another location), but with no chair visit elsewhere on the list they name the guest's home clinic.
     g: clinicRole's answer, or { role: "home" }. */
  function noTabRole(g, visits, from) {
    const tx = visits.filter(v => v.here !== false && (v.type === "Daily" || v.type === "Taper"));
    if (g.role !== "home" || !tx.length || needsTracker(visits, from)) return g;
    const live = visits.filter(v => v.status !== "No-show" && v.status !== "Cancelled"), fu = {};
    live.filter(v => v.type === "F/U" && v.here === false && v.where).forEach(v => { fu[v.where] = (fu[v.where] || 0) + 1; });
    const chairTop = Object.entries(g.by || {}).sort((a, b) => b[1] - a[1])[0], fuTop = Object.entries(fu).sort((a, b) => b[1] - a[1])[0];
    const top = chairTop || (fuTop ? [fuTop[0], 0] : null);
    return { ...g, role: "guest", here: tx.length, daily: tx.length, home: top ? top[0] : "", homeCount: top ? top[1] : 0, by: g.by || {}, fuAway: !chairTop && !!fuTop };
  }
  /* the MRN on a tab's first rows, wherever it sits: "MRN: 0012345" in one cell, or a value right of or below a cell reading MRN
     (A3 "MRN:" over A4 on New Template, B3 beside it on older tabs). A bare number elsewhere isn't taken: auth dates and counts
     are numbers too. grid: the rows' values (A1:T4). Returned as normMrn would: letters and digits, no leading zeros. */
  function mrnsIn(grid) {
    const out = new Set(), norm = v => String(v ?? "").toUpperCase().replace(/\.0+$/, "").replace(/[^A-Z0-9]/g, "").replace(/^0+/, "");
    const ok = v => /^\d{2,12}$/.test(norm(v)) || /^[A-Z]{0,3}\d{3,12}$/.test(norm(v));
    const g = grid || [];
    g.forEach((row, i) => (row || []).forEach((c, j) => {
      const s = String(c ?? "").trim(), m = s.match(/^mrn\s*[:#]?\s*(\S.*)$/i);
      if (m && ok(m[1])) out.add(norm(m[1]));
      else if (/^mrn\s*[:#]?$/i.test(s)) [(row || [])[j + 1], (g[i + 1] || [])[j]].forEach(v => { if (String(v ?? "").trim() && ok(v)) out.add(norm(v)); });
    }));
    return [...out];
  }
  /* "Their tab is in this workbook": tabs ranked by how close their name is to the patient's LASFIR (same code, then the same last or
     first three letters, then the fewest letters apart), filtered by what was typed. tabs: { name, code, ... }. Returns the tabs with
     near: true for the close ones. */
  function lev(a, b) {
    const d = Array.from({ length: b.length + 1 }, (_, j) => j);
    for (let i = 1; i <= a.length; i++) { let p = d[0]; d[0] = i; for (let j = 1; j <= b.length; j++) { const t = d[j]; d[j] = Math.min(d[j] + 1, d[j - 1] + 1, p + (a[i - 1] === b[j - 1] ? 0 : 1)); p = t; } }
    return d[b.length];
  }
  function rankTabs(code, query, tabs) {
    const c = String(code || "").toUpperCase(), q = String(query || "").trim().toUpperCase();
    const score = t => {
      const k = String(t.code || t.name.split(" ")[0] || "").toUpperCase();
      if (!c) return 10;
      if (k === c) return 0;
      const n = lev(k, c);
      if (k.slice(0, 3) === c.slice(0, 3)) return 1 + n / 10;
      if (k.slice(3, 6) === c.slice(3, 6)) return 2 + n / 10;
      return n <= 2 ? 3 + n : 10 + n;
    };
    return tabs.filter(t => !q || t.name.toUpperCase().includes(q)).map(t => ({ ...t, score: score(t) }))
      .sort((a, b) => a.score - b.score || a.name.localeCompare(b.name)).map(t => ({ ...t, near: t.score < 10 }));
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

  /* ---------- this tracker's clinic and its people. Nothing defaults to one clinic: an unknown clinic stays unknown ---------- */
  /* the names a clinic goes by, lowercased: "SRL" and "San Rafael" are the same clinic, only when the clinic is one of them */
  function clinicNames(c) {
    const x = String(c || "").trim().toLowerCase(); if (!x) return [];
    return x === "srl" || x === "san rafael" ? ["san rafael", "srl"] : [x];
  }
  /* a location or resource name at this clinic; with the clinic unknown everything counts as here, as before */
  function atClinic(s, c) {
    const n = clinicNames(c), x = String(s || "").toLowerCase();
    return !n.length || n.some(a => x.includes(a));
  }
  /* a tab's home clinic (S9) is this one: the same name, a short form of it ("CAS" for Castro) or one naming it. Unknown clinic: here */
  function homeHere(h, c) {
    const x = String(h || "").trim().toLowerCase(), n = clinicNames(c);
    if (!n.length || !x) return true;
    return n.some(a => x === a || (x.length >= 3 && a.startsWith(x)) || x.includes(a));
  }
  /* the people list on the hidden "TMS User Settings" sheet: Name | Added | Updated | Settings (JSON); first row is the header */
  const USER_SHEET = "TMS User Settings", USER_HDR = ["Name", "Added", "Updated", "Settings (JSON)"];
  const userKey = n => String(n || "").trim().replace(/\s+/g, " ").toLowerCase();
  function parseUsers(values) {
    const out = [], seen = new Set();
    (values || []).slice(1).forEach((r, i) => {
      const name = String((r || [])[0] ?? "").trim().replace(/\s+/g, " "), k = userKey(name);
      if (!name || seen.has(k)) return; seen.add(k); out.push({ name, row: i + 2, added: r[1] ?? "", updated: r[2] ?? "", set: userSet(r[3]) });
    });
    return out;
  }
  /* a person's own settings: the Settings (JSON) cell, anything unreadable counts as none */
  function userSet(json) { try { const o = JSON.parse(String(json ?? "").trim() || "{}"); return o && typeof o === "object" && !Array.isArray(o) ? o : {}; } catch { return {}; } }
  /* "Open in full screen whenever I open this workbook": each person's own choice (autoFull in their settings). Excel can only open
     the add-in with the workbook for everyone (document setting Office.AutoShowTaskpaneWithDocument), so it's on while anyone has
     theirs on. users = parseUsers(); name/on = a change not saved yet. */
  const AUTO_SHOW = "Office.AutoShowTaskpaneWithDocument";
  const autoFullOf = u => !!(u && u.set && u.set.autoFull === true);
  const autoAny = (users, name, on) => (users || []).some(u => name != null && userKey(u.name) === userKey(name) ? !!on : autoFullOf(u)) || (name != null && !!on);
  /* a typed name for the list: letters, spaces, apostrophes, periods and hyphens, 2 to 40 characters, tidied */
  function cleanUserName(s) {
    const n = String(s || "").trim().replace(/\s+/g, " ");
    return n.length >= 2 && n.length <= 40 && /^[A-Za-z][A-Za-z .'-]*$/.test(n) ? n : "";
  }
  /* the picker's order: people who have used this computer in this workbook first (most recent first, tagged), then everyone
     else A to Z. Nobody is pre-selected. here: { lowercased name: ISO last used } from this computer's memory */
  function pickOrder(users, here) {
    const h = here || {}, on = [], off = [];
    (users || []).forEach(u => { const at = h[userKey(u.name)]; (at ? on : off).push({ name: u.name, here: !!at, at: at || "" }); });
    on.sort((a, b) => String(b.at).localeCompare(String(a.at)) || a.name.localeCompare(b.name));
    off.sort((a, b) => a.name.localeCompare(b.name));
    return on.concat(off);
  }
  /* this computer's memory of who used the tracker is kept per workbook: a key from the document's URL, else from the
     workbook's own random id (saved in its document settings) */
  function wbKey(url, id) {
    const s = String(url || "").trim().split(/[?#]/)[0].toLowerCase() || (id ? "id:" + id : "");
    if (!s) return "";
    let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
    return `tms_who_${(h >>> 0).toString(36)}_${s.length.toString(36)}`;
  }
  /* local calendar day "2026-11-05": the daily pick lasts until it changes */
  const dayKey = d => { const x = d || new Date(); return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`; };
  /* the admin code. Only its salted SHA-256 is here; the repo is public, so this keeps honest people honest, nothing more */
  const ADMIN_SALT = "tms-admin-v1|", ADMIN_HASH = "1d869ce3250fd8b16b9ddce733cfc7815e4762e8b457043ebfaf477815553800", ADMIN_MIN = 15;
  async function adminOk(code) {
    const c = String(code || "").trim(); if (!/^\d{4}$/.test(c)) return false;
    const subtle = (typeof crypto !== "undefined" && crypto.subtle) || (globalThis.crypto && globalThis.crypto.subtle);
    if (!subtle) return false;
    const buf = await subtle.digest("SHA-256", new TextEncoder().encode(ADMIN_SALT + c));
    return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, "0")).join("") === ADMIN_HASH;
  }

  return { LATE_MIN, SOON_MIN, pool, clinicNames, atClinic, homeHere, USER_SHEET, USER_HDR, userKey, parseUsers, userSet, AUTO_SHOW, autoFullOf, autoAny, cleanUserName, pickOrder, wbKey, dayKey, adminOk, ADMIN_MIN, sessKey, visitNames, txNumbers, KAISER_NOTE, dropKaiser, logChecks, apptGaps, pulsesNoAppt, hashRows, phase, chairState, chairStatus, isLate, moveState, ngForMove, missStreak, missRun, mapEvent, clinician, clinicianName, telNote, telMissLine, telFuLine, telCxLine, cancelNote, CX_WINDOW, bizDays, fuRebook, REBOOK_DAYS, FU_NEXT, fuGaps, mrnsIn, needs, loud, chairLabel, planMeasures, providerLast, clinicFromName, clinicRole, ASK_MIN, needsTracker, noTabRole, rankTabs, cellLabel, tabName, misdated, MISDATED_DAYS,
    copayOf, withCopayLine, consentLine, consentOf, withConsentLine, flagOn, isTreatment, copaysToConfirm, copayTally, owes, rollover, siteOf, SITE_LABEL, SITE_CELL, pulseCol, protoMatch, protoPulses };
});
