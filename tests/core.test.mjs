import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
const T = createRequire(import.meta.url)("../core.js");

// Synthetic data only.
const v = o => ({ chairVisit: true, chair: "MagV1", time: 600, len: 19, miss: false, done: false, start: null, checkin: null, ...o });

test("phase: missed, done, running, waiting, booked", () => {
  assert.equal(T.phase(v({ miss: true, done: true })), "missed");
  assert.equal(T.phase(v({ done: true, start: 600 })), "done");
  assert.equal(T.phase(v({ start: 600, checkin: 590 })), "running");
  assert.equal(T.phase(v({ checkin: 590 })), "waiting");
  assert.equal(T.phase(v({})), "booked");
});

test("chair state: track, soon, over time, late, open", () => {
  const now = 640;
  assert.deepEqual([T.chairState("MagV1", [v({ start: 630 })], now)].map(s => [s.state, s.left]), [["track", 9]]);
  assert.equal(T.chairState("MagV1", [v({ start: 623 })], now).state, "soon");
  const over = T.chairState("MagV1", [v({ start: 610 })], now);
  assert.equal(T.chairStatus(over), "Over by 11 min");
  assert.equal(T.chairStatus(T.chairState("MagV1", [v({ start: 621 })], now)), "Ends now");
  const late = T.chairState("MagV1", [v({ time: 625, checkin: 620 })], now);
  assert.deepEqual([late.state, T.chairStatus(late)], ["late", "Waiting 15 min"]);
  const early = T.chairState("MagV1", [v({ time: 660, checkin: 638 })], now);
  assert.deepEqual([early.state, T.chairStatus(early)], ["open", "Checked in"]);
  const open = T.chairState("MagV1", [v({ time: 700 }), v({ time: 680 }), v({ time: 600, done: true })], now);
  assert.equal(open.state, "open"); assert.equal(open.next.time, 680);
});

test("chair state: a running session wins over a waiting patient; missed and other chairs are ignored", () => {
  const s = T.chairState("MagV1", [v({ checkin: 600, time: 600 }), v({ start: 630, time: 615 }), v({ chair: "MagV2", start: 635 }), v({ miss: true, start: 600 })], 640);
  assert.equal(s.state, "track"); assert.equal(s.v.start, 630);
});

test("late anywhere: waiting 10+ min past the booked time", () => {
  assert.equal(T.isLate(v({ checkin: 600, time: 600 }), 610), true);
  assert.equal(T.isLate(v({ checkin: 600, time: 600 }), 609), false);
  assert.equal(T.isLate(v({ checkin: 600, time: null }), 700), false);
  assert.equal(T.isLate(v({ checkin: 600, time: 600, start: 605 }), 700), false);
});

test("moves: reminder until ticked, listed, or moved back", () => {
  assert.deepEqual(T.moveState(null, "MagV1"), { moved: false, settled: false, chair: "MagV1", from: "" });
  const s = { chair: "MagV3", ng: "MagV1", tick: false, at: 640 };
  assert.equal(T.moveState(s, "MagV1", 425).moved, true);
  assert.equal(T.moveState({ ...s, tick: true }, "MagV1", 425).moved, false);
  assert.equal(T.moveState({ ...s, tick: true }, "MagV1", 425).settled, true);
  assert.equal(T.moveState(s, "MagV3", 700).moved, false, "a list run after the move shows the new chair");
  assert.equal(T.moveState(s, "MagV3", 700).settled, true);
  assert.equal(T.moveState(s, "MagV3", 600).moved, true, "a list run before the move proves nothing");
  assert.equal(T.moveState(s, "MagV3", null).moved, true, "no list today");
  assert.equal(T.moveState({ chair: "MagV1", ng: "MagV1", at: 640 }, "MagV1", 425).moved, false, "moved back");
});

test("moves: NextGen's chair for the next move follows the tick", () => {
  assert.equal(T.ngForMove(null, "MagV1"), "MagV1");
  assert.equal(T.ngForMove({ chair: "MagV3", ng: "MagV1", tick: false }, "MagV1"), "MagV1");
  assert.equal(T.ngForMove({ chair: "MagV3", ng: "MagV1", tick: true }, "MagV1"), "MagV3");
  assert.equal(T.ngForMove({ chair: "MagV3", ng: "MagV1", tick: false, at: 640 }, "MagV3", 700), "MagV3", "a list run after the move confirmed it");
  assert.equal(T.ngForMove({ chair: "MagV3", ng: "MagV1", tick: false, at: 640 }, "MagV3", 600), "MagV1", "an older list doesn't");
  // ticked at MagV3, then moved back to MagV1: NextGen has MagV3 now, and the morning's list (MagV1) predates the move
  const back = { chair: "MagV1", ng: T.ngForMove({ chair: "MagV3", ng: "MagV1", tick: true }, "MagV1"), tick: false, at: 700 };
  assert.equal(back.ng, "MagV3");
  assert.equal(T.moveState(back, "MagV1", 425).moved, true);
  assert.equal(T.moveState(back, "MagV1", 720).moved, false, "a list run after the move settles it");
});

test("missed streak counts trailing misses and skips visits not entered yet", () => {
  const r = (date, o) => ({ date, type: "Daily", miss: false, delivered: true, alt: false, ...o });
  const rows = [r(1), r(2, { miss: true, delivered: false }), r(3, { delivered: false }), r(4, { miss: true, delivered: false }), r(5, { miss: true, delivered: false }), r(9, { delivered: false })];
  assert.equal(T.missStreak(rows, 6), 3);
  assert.equal(T.missStreak([...rows, r(6)], 6), 0);
  assert.equal(T.missStreak([...rows, r(6, { type: "F/U" })], 6), 3, "follow-ups don't break it");
});

test("visit names: a second visit in the same pool that day gets #2", () => {
  const out = T.visitNames([{ type: "Daily", time: 700 }, { type: "F/U", time: 800 }, { type: "MT", time: 600 }]).map(x => x.visit);
  assert.deepEqual(out, ["TMS", "TMS#2", "F/U"]);
  assert.equal(T.sessKey(46316, "ALVMAR", "TMS"), "46316|ALVMAR|TMS");
});

const D = o => ({ active: true, stale: false, lastTx: null, authFlags: [], bidfToday: false, suggest: null, cancelled: null, decisions: [], overdue: [], noTech: [], measDue: [],
  protoFlags: [], awayConflicts: [], closedConflicts: [], sameDay: [], benFlags: [], ...o });
const P = o => ({ status: "In Progress", rows: [], external: false, home: "", ...o, d: D(o.d || {}) });
const X = o => ({ today: 100, fmtDate: s => `d${s}`, recentDays: 10, exemptTag: "NextGen: Exempt, confirm", unclassTag: "NextGen: no-show not classified", ...o });

test("needs: red first, then amber in order, quiet last; each carries its action", () => {
  const p = P({ status: "Tapering", d: { authFlags: ["Auth expires in 9 d (10/30)", "Auth expired 10/18"], bidfToday: true, decisions: [{ row: 12, date: 95, note: "" }] } });
  const n = T.needs(p, X({}));
  assert.deepEqual(n.map(x => x.k), ["auth", "auth", "bidf", "freebie", "taper"]);
  assert.equal(n[0].t, "Auth expired 10/18"); assert.equal(n[0].bad, true);
  assert.equal(n[1].bad, false);
  assert.deepEqual(n[3].act, ["Decide freebie", "freebie"]); assert.equal(n[3].row, 12);
  assert.equal(T.loud(n).length, 4);
});

test("needs: moves, chairless visits and exempt misses", () => {
  const p = P({ d: { decisions: [{ row: 14, date: 98, note: "No-show; NextGen: Exempt, confirm" }] } });
  const n = T.needs(p, X({ moved: [{ key: "k1", chairLabel: "MagV 3" }], noChair: [{ key: "k2", timeLabel: "11:30" }] }));
  assert.deepEqual(n.map(x => x.t), ["Moved to MagV 3, not yet in NextGen", "No chair for 11:30", "Exempt in NextGen on d98: freebie or waived?"]);
  assert.equal(T.needs(p, X({ moved: [{ key: "k0", chairLabel: "MagV 2", dayLabel: "Oct 20" }] }))[0].t, "Moved to MagV 2 on Oct 20, not yet in NextGen");
  assert.equal(n[0].key, "k1"); assert.deepEqual(n[0].act, ["Mark moved in NextGen", "ngdone"]);
});

test("needs: active-only checks skip a paused course; discontinued has none", () => {
  const paused = P({ status: "Paused", d: { active: false, authFlags: ["Auth expired 10/1"], benFlags: ["x"], overdue: [{ row: 11, date: 90 }] } });
  assert.deepEqual(T.needs(paused, X({})).map(x => x.k), ["pulses", "paused"]);
  assert.deepEqual(T.needs(P({ status: "Discontinued", d: { stale: true } }), X({})), []);
});

test("needs: 3 missed in a row is red with Call to reschedule", () => {
  const rows = [95, 96, 97].map(date => ({ date, type: "Daily", miss: true, delivered: false, alt: false }));
  const n = T.needs(P({ rows }), X({}));
  assert.equal(n[0].t, "Missed 3 in a row"); assert.equal(n[0].bad, true); assert.deepEqual(n[0].act, ["Call to reschedule", "call"]);
});

test("needs: MT today without a schedule given", () => {
  const n = T.needs(P({ status: "Pending Start" }), X({ mtToday: true, schedGiven: false }));
  assert.deepEqual(n.map(x => x.k), ["sched", "start"]); assert.equal(n[0].v, "all");
  assert.deepEqual(T.needs(P({}), X({ mtToday: true, schedGiven: true })).map(x => x.k), ["start"]);
});

test("chair labels", () => {
  assert.equal(T.chairLabel("Castro MagV1", "Castro"), "MagV 1");
  assert.equal(T.chairLabel("San Rafael MagV2", "san rafael"), "MagV 2");
  assert.equal(T.chairLabel("Brainsway2", ""), "BrainsWay 2");
  assert.equal(T.chairLabel("MagV1"), "MagV 1");
  assert.equal(T.chairLabel("Other Chair", "Castro"), "Other Chair");
});

test("needs: one line per absence, naming each visit booked in it", () => {
  const r = (row, date) => ({ row, date });
  const p = P({ d: { awayConflicts: [{ r: r(20, 102), why: "away d102 to d105 (vacation)" }, { r: r(21, 105), why: "away d102 to d105 (vacation)" }, { r: r(30, 120), why: "away d120 to d120" }] } });
  const n = T.needs(p, X({})).filter(x => x.k === "away");
  assert.deepEqual(n.map(x => x.t), ["Booked d102, d105 while away d102 to d105 (vacation)", "Booked d120 while away d120 to d120"]);
  assert.equal(n[0].row, 20);
});

/* measures: a log row as parseLog reads it; dates are day numbers (Oct 1 = 1, Nov 2 = 33) */
const lr = (row, o) => ({ row, type: "Daily", tx: null, date: null, delivered: false, miss: false, phq: "", gad: "", action: "", ...o,
  measX: [o.phq, o.gad].some(x => /^x$/i.test(x || "")), measVal: [o.phq, o.gad].some(x => x && !/^x$/i.test(x)) });
const applyMeas = (rows, ops) => { const m = new Map(rows.map(r => [r.row, { ...r }])); for (const [a, v] of ops) { const r = m.get(+a.slice(1)), k = { P: "phq", Q: "gad", O: "action" }[a[0]]; r[k] = v; } return [...m.values()].map(r => lr(r.row, r)); };
const measRows = out => out.filter(r => /^x$/i.test(r.phq) || (r.phq && !/^x$/i.test(r.phq))).map(r => r.date);
const MX = { phq: "x", gad: "x" }, XM = { ...MX, action: "Measures" };
/* the second half of a 36-treatment course on Oct 25, as the old planner left it */
function course() {
  const d = [[16, 14], [17, 15, "MTR", MX], [18, 16, null, { phq: "9", gad: "5" }], [19, 21], [20, 22], [21, 23, null, XM], [22, 26], [23, 28], [24, 29], [25, 30, null, XM],
    [null, 33, "F/U", MX], [26, 35, null, XM], [27, 36], [28, 37], [29, 40], [30, 42, null, XM], [31, 43, null, XM], [32, 44], [33, 47], [34, 48, null, XM], [35, 49], [36, 51, null, XM],
    [null, 54, "F/U", MX], [null, null, "F/U", MX]];
  return d.map(([tx, date, type, m], i) => lr(10 + i, { tx, date, type: type || "Daily", delivered: date != null && date < 25 && type !== "F/U", ...(m || {}) }));
}

test("measures: late collection moves the start, an uncollected day rolls to the next visit, then every 5th treatment", () => {
  const rows = course(), out = applyMeas(rows, T.planMeasures(rows, 36, 25));
  assert.deepEqual(measRows(out), [16, 26, 33, 42, 51, 54]);   // Oct 16 (taken), Oct 26, Nov 2 F/U, Nov 11, Nov 20 last, Nov 23 F/U
  const at = d => out.find(r => r.date === d);
  assert.equal(at(15).phq, "");                                   // the MTR's x's go: they were taken the next day
  assert.equal(at(16).action, "Measures");                        // and that day carries the flag
  assert.equal(at(23).phq, "");                                   // Oct 23 wasn't collected
  assert.equal(at(33).action, "Measures");                        // F/U rows carry the flag too
  assert.equal(at(49).phq, "");                                   // Nov 18 is 2 days before the last treatment
  assert.equal(out[out.length - 1].phq, "");                      // an undated row gets none until it has a date
  assert.equal(at(16).phq, "9");                                  // scores are never touched
  assert.deepEqual(T.planMeasures(out, 36, 25), []);              // running it again changes nothing
});

test("measures: overdue ones on the next visit stay there when it runs again", () => {
  const rows = [lr(10, { tx: 1, date: 1, type: "MT", delivered: true }), lr(11, { tx: 2, date: 4, delivered: true, ...MX }), lr(12, { tx: 3, date: 5, delivered: true }),
    ...[6, 7, 8, 11, 12, 13, 14].map((d, i) => lr(13 + i, { tx: 4 + i, date: d }))];
  const once = applyMeas(rows, T.planMeasures(rows, 36, 6));
  assert.deepEqual(measRows(once), [6, 13]);                     // the next visit, then the 5th treatment after it
  assert.deepEqual(T.planMeasures(once, 36, 6), []);
});

test("measures: a regular collection within 7 days of a required one is skipped", () => {
  const rows = [lr(10, { tx: 1, date: 1, type: "MT", delivered: true, phq: "12", gad: "8" }), ...[2, 3, 4, 5].map((d, i) => lr(11 + i, { tx: i + 2, date: d + 1 })),
    lr(15, { tx: 6, date: 8 }), lr(16, { tx: 7, date: 9, type: "MTR" }), lr(17, { tx: 8, date: 10 })];
  assert.deepEqual(measRows(applyMeas(rows, T.planMeasures(rows, 36, 2))), [1, 9]);             // 6th treatment Thursday, MTR Friday: just the MTR
  const far = rows.map(r => r.row === 16 ? { ...r, date: 20 } : r);
  assert.deepEqual(measRows(applyMeas(far, T.planMeasures(far, 36, 2))), [1, 8, 20]);          // MTR 12 days later: both
});

test("needs: earlier measures are one Check NextGen line, and none once the course is over", () => {
  const meas = [{ row: 12, date: 90 }, { row: 15, date: 95 }];
  const on = T.needs(P({ d: { measDue: meas } }), X({})).filter(y => y.k === "measures");
  assert.deepEqual(on.map(y => [y.t, y.act[1], y.row]), [["Check NextGen for PHQ-9 and GAD-7", "pastmeas", 15]]);
  assert.equal(T.needs(P({ status: "Complete", d: { active: false, measDue: meas } }), X({})).some(y => y.k === "measures"), false);
});

test("providerLast: E2 gets the last name from either report shape", () => {
  assert.equal(T.providerLast("Testlast, Pat MD"), "Testlast");
  assert.equal(T.providerLast("Van Testa, Pat"), "Van Testa");
  assert.equal(T.providerLast("Dr. Pat Testlast"), "Testlast");
  assert.equal(T.providerLast("Pat Q. Testlast, PMHNP"), "Testlast");
  assert.equal(T.providerLast("Pat Testlast PA-C"), "Testlast");
  assert.equal(T.providerLast("Pat Testlast M.D."), "Testlast");
  assert.equal(T.providerLast("Pat Testlast, Ph.D."), "Testlast");
  assert.equal(T.providerLast("Testlast"), "Testlast");
  assert.equal(T.providerLast(""), "");
  assert.equal(T.providerLast(null), "");
});

// copays, benefits and protocol (synthetic rows only)
const row = o => ({ row: 10, type: "Daily", date: 100, delivered: true, miss: false, alt: false, note: "", ...o });

test("copayOf: reads every line the add-in writes, and none", () => {
  assert.equal(T.copayOf("Tele; Copay $25.00 collected by Pat 11:32 AM").kind, "collected");
  assert.deepEqual((({ amt, by, at }) => ({ amt, by, at }))(T.copayOf("Copay $1,025.50 collected by Pat 11:32 AM")), { amt: 1025.5, by: "Pat", at: "11:32 AM" });
  assert.equal(T.copayOf("Copay $25.00 not collected").kind, "not");
  assert.equal(T.copayOf("Copay $0, OOPM met by Pat 9:05 AM").kind, "oopm");
  assert.equal(T.copayOf("Copay waived by Pat 10/27").kind, "waived");
  const c = T.copayOf("Copay $20.00 collected (confirmed 10/27 by Pat)"); assert.ok(c.kind === "collected" && c.confirmed && c.by === "Pat");
  assert.equal(T.copayOf("Tele; Freebie used"), null);
  assert.equal(T.withCopayLine("Tele; Copay $20.00 not collected", "Copay $20.00 collected by Pat 1:00 PM"), "Tele; Copay $20.00 collected by Pat 1:00 PM");
  assert.equal(T.withCopayLine("Copay waived by Pat 10/27", ""), "");
});

test("copays to confirm: delivered treatment rows before today, owed, no copay line; F/U never", () => {
  const pay = r => r.pay;
  const rows = [row({ row: 10, pay: 20 }), row({ row: 11, pay: 20, note: "Copay $20.00 collected" }), row({ row: 12, type: "F/U", pay: 40 }),
    row({ row: 13, pay: 0 }), row({ row: 14, pay: 20, date: 105 }), row({ row: 15, pay: 20, delivered: false }), row({ row: 16, pay: 20, miss: true }), row({ row: 17, type: "MTR", pay: 50, date: 99 })];
  assert.deepEqual(T.copaysToConfirm(rows, pay, 105).map(r => r.row), [17, 10]);
});

test("copay tally: collected amounts on treatment rows in the plan year only", () => {
  const rows = [row({ date: 100, note: "Copay $20.00 collected" }), row({ date: 101, note: "Copay $20.00 collected (confirmed 10/27 by Pat)" }), row({ date: 102, note: "Copay $20.00 not collected" }),
    row({ date: 103, type: "F/U", note: "Copay $40.00 collected" }), row({ date: 50, note: "Copay $20.00 collected" }), row({ date: 104, note: "Copay $0, OOPM met" })];
  assert.deepEqual(T.copayTally(rows, 90, 200), { sum: 40, n: 2 });
});

test("owes: no benefits reminders once the OOPM is met or the BIDF quote is $0", () => {
  assert.equal(T.owes({ oopRem: 0, oopNA: false, bidf: 900 }), false);
  assert.equal(T.owes({ oopRem: 500, bidf: 0 }), false);
  assert.equal(T.owes({ oopRem: null, oopNA: true, bidf: 900 }), true);
  assert.equal(T.owes({ oopRem: 500, bidf: null }), true);
});

test("rollover: amber from Dec 1 with a new-year visit booked, red once one happens, cleared by a new-year update", () => {
  const y = { updated: 10, dec1: 100, jan1: 131 }, booked = [row({ date: 140, delivered: false })], none = () => false;
  assert.equal(T.rollover(y, booked, 99, none), null);
  assert.deepEqual(T.rollover(y, booked, 100, none), { bad: false });
  assert.equal(T.rollover(y, [row({ date: 120, delivered: false })], 110, none), null);
  assert.deepEqual(T.rollover(y, [row({ date: 132, delivered: true }), ...booked], 133, none), { bad: true });
  assert.deepEqual(T.rollover(y, [row({ date: 133, delivered: false })], 133, r => r.date === 133), { bad: true });
  assert.equal(T.rollover({ ...y, updated: 131 }, booked, 133, none), null);
  assert.equal(T.rollover({ ...y, updated: null }, booked, 133, none), null);
});

test("protocol: exact, custom with what changed, none; sites and pulse columns", () => {
  const table = [{ name: "HF 10Hz ", rep: 10, pulses: 40, trains: 75, iti: 11 }, { name: "LF 1Hz", rep: 1, pulses: 1800, trains: 1, iti: 0 }];
  assert.deepEqual(T.protoMatch({ hz: 10, dur: 40, trains: 75, wait: 11 }, table), { name: "HF 10Hz", base: table[0], custom: false, diffs: [] });
  const c = T.protoMatch({ hz: 1, dur: 1200, trains: 1, wait: 0 }, table);
  assert.equal(c.name, "LF 1Hz (Custom)"); assert.deepEqual(c.diffs.map(d => [d.k, d.v, d.std]), [["dur", 1200, 1800]]);
  assert.equal(T.protoMatch({ hz: 5, dur: 40, trains: 75, wait: 11 }, table).name, "Custom");
  assert.equal(T.protoMatch({ hz: "", dur: 40 }, table), null);
  assert.deepEqual(["F3", "Left", "f4", "Right", "Other", "Cz", ""].map(T.siteOf), ["F3", "F3", "F4", "F4", "Other", "Other", ""]);
  assert.deepEqual(["F3", "F4", "Other"].map(T.pulseCol), [0, 1, 2]);
});

test("needs: Choose a protocol, rollover, OOPM met is quiet", () => {
  const d = { active: true, authFlags: [], decisions: [], overdue: [], noTech: [], measDue: [], protoFlags: [], awayConflicts: [], closedConflicts: [], sameDay: [], benFlags: [], noProto: true, rollover: { year: 2027, bad: true }, oopmMet: 2026 };
  const n = T.needs({ status: "In Progress", rows: [], d }, { today: 100, fmtDate: String, recentDays: 10 });
  assert.deepEqual(n.map(x => [x.k, !!x.bad, !!x.quiet]), [["rollover", true, false], ["noproto", false, false], ["oopm", false, true]]);
  assert.equal(n[0].t, "Re-verify benefits for 2027");
});
