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

test("protocol pulses: theta burst counts 3 per burst, custom keeps its base's ratio, no total is pulses x trains", () => {
  const table = [{ name: "X", rep: 10, pulses: 40, trains: 75, iti: 10, total: 3000 }, { name: "iTBS 600 ", rep: 5, pulses: 10, trains: 20, iti: 8, total: 600 },
    { name: "cTBS", rep: 5, pulses: 300, trains: 1, iti: 1, total: 300 }, { name: "LF 1Hz", rep: 1, pulses: 50, trains: 20, iti: 1 }];
  assert.equal(T.protoPulses({ hz: 10, dur: 40, trains: 75, wait: 10 }, table), 3000);
  assert.equal(T.protoPulses({ hz: 5, dur: 10, trains: 20, wait: 8 }, table), 600);
  assert.equal(T.protoPulses({ hz: 5, dur: 10, trains: 60, wait: 8 }, table), 1800);   // iTBS custom, more trains
  assert.equal(T.protoPulses({ hz: 1, dur: 50, trains: 20, wait: 1 }, table), 1000);
  assert.equal(T.protoPulses({ hz: 20, dur: 40, trains: 75, wait: 10 }, table), 3000);   // no protocol at that rate
  assert.equal(T.protoPulses({ hz: 10, dur: "", trains: 75 }, table), 0);
});

test("needs: Confirm protocol, rollover, OOPM met is quiet", () => {
  const d = { active: true, authFlags: [], decisions: [], overdue: [], noTech: [], measDue: [], protoFlags: [], awayConflicts: [], closedConflicts: [], sameDay: [], benFlags: [], noProto: true, rollover: { year: 2027, bad: true }, oopmMet: 2026 };
  const n = T.needs({ status: "In Progress", rows: [], d }, { today: 100, fmtDate: String, recentDays: 10 });
  assert.deepEqual(n.map(x => [x.k, !!x.bad, !!x.quiet]), [["rollover", true, false], ["noproto", false, false], ["oopm", false, true]]);
  assert.equal(n[0].t, "Re-verify benefits for 2027");
});

/* ---------- visit-log checks (synthetic rows) ---------- */
const L = (row, o) => ({ row, date: null, type: "Daily", miss: false, delivered: false, tx: "", note: "", alt: false, noA: false, noD: false, ...o });

test("Tx #: the MT is 1, each Daily, Taper or MTR one more, misses and F/U blank", () => {
  const rows = [L(10, { type: "MT" }), L(11), L(12, { miss: true }), L(13, { type: "F/U" }), L(14, { type: "MTR" }), L(15, { type: "Taper" }), L(16, { type: "" })];
  assert.deepEqual(T.txNumbers(rows), [1, 2, "", "", 3, 4, ""]);
  assert.deepEqual(T.txNumbers([L(10), L(11, { type: "MT" }), L(12)]), [1, 1, 2]);   // a second course tab may start on a Daily
});

test("log checks: two rows on a day, Tx # off, order, lost formulas, Kaiser reminder, past # of TXs", () => {
  const K = T.KAISER_NOTE, rows = [
    L(10, { type: "MT", date: 100, tx: 1, delivered: true }), L(11, { date: 101, tx: 2, delivered: true }),
    L(12, { type: "F/U", date: 104, miss: true, note: `${K}; No-show` }), L(13, { type: "F/U", date: 104, miss: true, note: `${K}; No-show` }),
    L(14, { date: 103, tx: "", delivered: true, noA: true }), L(15, { date: 105, tx: 3, noD: true }), L(16, { date: 106, tx: 4 })];
  const c = T.logChecks(rows, { pr: true, approved: 4 });
  assert.deepEqual(c.dup, [{ date: 104, pool: "F/U", rows: [12, 13] }]);
  assert.deepEqual(c.tx.map(x => [x.row, x.was, x.want]), [[14, "", 3], [15, "3", 4], [16, "4", 5]]);
  assert.deepEqual(c.order, [14]);
  assert.deepEqual(c.formulas, [{ row: 14, a: true, d: false }, { row: 15, a: false, d: true }]);
  assert.deepEqual(c.kaiser, [12, 13]);
  assert.deepEqual(c.over.map(x => [x.row, x.n]), [[16, 5]]);
  assert.equal(T.dropKaiser(`${K}; Tele`), "Tele");
  // every row treated that day (an accelerated day) is not a duplicate; no PR formula on the tab means A isn't checked
  assert.equal(T.logChecks([L(10, { date: 1, delivered: true, tx: 1 }), L(11, { date: 1, delivered: true, tx: 2 })]).dup.length, 0);
  assert.equal(T.logChecks([L(10, { date: 1, tx: 1, noA: true })], { pr: false }).formulas.length, 0);
});

test("appointment list against the tab: missing rows, open rows not booked, same visit listed twice", () => {
  const rows = [L(10, { type: "MT", date: 90, delivered: true }), L(11, { date: 100 }), L(12, { date: 101 }), L(13, { type: "F/U" })];
  const appts = [{ date: 90, time: 600, status: "Kept", type: "MT" }, { date: 100, time: 600, status: "Scheduled", type: "Daily" }, { date: 100, time: 600, status: "Scheduled", type: "Daily" },
    { date: 102, time: 660, status: "Scheduled", type: "F/U" }, { date: 80, time: 600, status: "Kept", type: "Daily" }];
  const g = T.apptGaps(rows, appts, { from: 95, to: 110 });
  assert.deepEqual(g.notOnTab.map(a => [a.date, a.type]), [[102, "F/U"]]);   // kept on day 80 is before this course
  assert.deepEqual(g.notOnList, [12]);
  assert.notEqual(T.hashRows([["a"]]), T.hashRows([["b"]]));
});

test("clinicFromName: the clinic before the device or Tracker", () => {
  assert.equal(T.clinicFromName("Castro MagV Patient Tracker.xlsx"), "Castro");
  assert.equal(T.clinicFromName("Walnut Creek TMS Pt Tracker"), "Walnut Creek");
  assert.equal(T.clinicFromName("Post MagV Tracker 2026"), "Post");
  assert.equal(T.clinicFromName("Book1"), "");
});

test("clinicRole: the MT and MT redo decide home; 16+ dailies here asks; no MT falls back to the majority", () => {
  const c = (here, where = "Post", type = "Daily") => ({ chair: true, here, where: here ? "" : where, type });
  const mt = (here, type = "MT") => ({ chair: false, here, where: here ? "" : "Post", type });
  const dailies = (n, here) => Array.from({ length: n }, () => c(here));
  assert.equal(T.ASK_MIN, 16);
  assert.equal(T.clinicRole([mt(true), c(false), c(false), c(true)]).role, "guest");          // MT here, most treatments at Post: Post carries the tab
  assert.equal(T.clinicRole([mt(true), c(false), c(true)]).role, "home");                     // MT here, a tie: the MT decides
  assert.equal(T.clinicRole([mt(false), c(false), c(true)]).role, "guest");                   // MT at Post, a tie: Post's
  assert.equal(T.clinicRole([mt(false), mt(true, "MTR")]).role, "home");                      // tie between MT and MTR stays here
  const g = T.clinicRole([mt(false), ...dailies(15, true), c(false)]);
  assert.deepEqual([g.role, g.home, g.daily], ["guest", "Post", 15]);                          // MT at Post, 15 dailies here
  const a = T.clinicRole([mt(false), ...dailies(16, true)]);
  assert.deepEqual([a.role, a.home, a.daily], ["ask", "Post", 16]);                            // 16 dailies here: ask
  assert.equal(T.clinicRole([mt(false, "MTR"), ...dailies(20, true)]).role, "ask");           // an MT redo elsewhere counts too
  assert.equal(T.clinicRole([mt(false), c(false)]).role, "none");                             // MT at Post, nothing here
  assert.equal(T.clinicRole([{ chair: false, here: false, where: "Post", type: "F/U" }, c(true)]).role, "home");   // F/U don't decide
  assert.equal(T.clinicRole([mt(null), c(true), c(false)]).role, "home");                     // MT location unknown: majority, tie here
  assert.equal(T.clinicRole([c(true), c(false), c(false)]).role, "guest");                    // no MT: most chair visits at Post
  assert.equal(T.clinicRole([c(false), c(false)]).role, "none");                              // no chair visit here
  assert.equal(T.clinicRole([mt(true)]).role, "home");                                         // a new patient's MT here
});

test("needsTracker: an MT here or 16+ treatments here offers a tracker; F/U alone and an MT elsewhere don't", () => {
  const b = (type, here = null, date = 10, status = "Scheduled") => ({ type, here, date, status });
  const tx = (n, here = null) => Array.from({ length: n }, (_, i) => b("Daily", here, 10 + i));
  assert.equal(T.needsTracker([b("F/U"), b("F/U", true, 12), b("F/U", false, 14)], 10), false);   // F/U only, in office or telehealth
  assert.equal(T.needsTracker(tx(16), 10), true);                                              // 16 booked treatments
  assert.equal(T.needsTracker(tx(15), 10), false);                                             // 15 isn't enough
  assert.equal(T.needsTracker([...tx(15), b("Daily", null, 30, "Cancelled")], 10), false);     // cancellations don't count
  assert.equal(T.needsTracker([...tx(16, false)], 10), false);                                 // treatments at another clinic
  assert.equal(T.needsTracker([b("MT", true)], 10), true);                                     // MT at this tracker's clinic
  assert.equal(T.needsTracker([b("MT")], 10), true);                                           // MT, location not given
  assert.equal(T.needsTracker([b("MT", false), b("F/U", true)], 10), false);                   // MT at another clinic
  assert.equal(T.needsTracker([b("MT", false), ...tx(16, true)], 10), true);                   // MT elsewhere, but 16 here
  assert.equal(T.needsTracker([b("MT", true, 5, "Kept")], 10), true);                          // an MT already done here
  assert.equal(T.needsTracker([b("MT", true, 5, "No-show")], 10), false);                      // a missed MT
});

test("noTabRole: no tab, no MT here and under 16 treatments here is a guest; F/U alone and a tracker due stay home", () => {
  const b = (type, here = null, date = 10) => ({ type, here, date, status: "Scheduled" });
  const tx = (n, here = null) => Array.from({ length: n }, (_, i) => b("Daily", here, 10 + i));
  const g = T.noTabRole({ role: "home" }, tx(5), 10);
  assert.deepEqual([g.role, g.home, g.here], ["guest", "", 5]);                                 // 5 treatments: a guest, home clinic unknown
  assert.equal(T.noTabRole({ role: "home", by: { Post: 3 } }, tx(5), 10).home, "Post");        // the clinic with most chair visits elsewhere
  assert.equal(T.noTabRole({ role: "home" }, tx(16), 10).role, "home");                        // 16: create tracker
  assert.equal(T.noTabRole({ role: "home" }, [b("MT", true), ...tx(5)], 10).role, "home");     // MT here: create tracker
  assert.equal(T.noTabRole({ role: "home" }, [b("F/U"), b("F/U", true)], 10).role, "home");    // F/U only: no tracker, no guest
  assert.equal(T.noTabRole({ role: "ask", home: "Post" }, tx(5), 10).role, "ask");             // clinicRole's answer stands
});

test("cellLabel: names a tracker cell for the rebuild's side-by-side check", () => {
  assert.deepEqual(T.cellLabel("A6"), { label: "DOB", date: true });
  assert.deepEqual(T.cellLabel("A4"), { label: "MRN", date: false });
  assert.deepEqual(T.cellLabel("E12"), { label: "Date, log row 3", date: true });
  assert.equal(T.cellLabel("K3").label, "Protocol row 2: Pulses per train");
  assert.equal(T.cellLabel("Z99").label, "Z99");
});

test("card and consent line: written, read back, replaced, and kept apart from the copay line", () => {
  const yes = T.consentLine(true, true, "Pat", "9:42 AM");
  assert.equal(yes, "LumaPay card on file, verbal copay consent, checked by Pat 9:42 AM");
  assert.deepEqual(T.consentOf(`Tele; Copay $20.00 collected by Pat 9:40 AM; ${yes}`), { line: yes, card: true, ok: true, by: "Pat", at: "9:42 AM" });
  const no = T.consentLine(false, false, "Sam", "1:05 PM");
  assert.equal(no, "No LumaPay card on file, no verbal copay consent, checked by Sam 1:05 PM");
  const c = T.consentOf(no); assert.equal(c.card, false); assert.equal(c.ok, false); assert.equal(c.by, "Sam");
  const half = T.consentOf(T.consentLine(true, false, "Sam", "1:05 PM")); assert.equal(half.card, true); assert.equal(half.ok, false);
  assert.equal(T.consentOf("Copay $20.00 collected by Pat 9:40 AM"), null);
  assert.equal(T.copayOf(yes), null);
  const n = T.withConsentLine(`Copay $20.00 collected by Pat 9:40 AM; ${no}`, yes);
  assert.equal(n, `Copay $20.00 collected by Pat 9:40 AM; ${yes}`);
  assert.equal(T.withCopayLine(n, "Copay $20.00 not collected by Pat 9:50 AM"), `${yes}; Copay $20.00 not collected by Pat 9:50 AM`);
  assert.equal(T.withConsentLine(yes, ""), "");
});

test("Benefits flag cells: Yes, x, TRUE and ticks are on; blank and No are off", () => {
  for (const v of ["Yes", "yes", "Y", "x", true, "TRUE", "✓"]) assert.equal(T.flagOn(v), true, String(v));
  for (const v of ["", null, undefined, "No", false, "n"]) assert.equal(T.flagOn(v), false, String(v));
});

test("noTabRole: a F/U at another clinic names an intermittent guest's home, but never outweighs a course booked here", () => {
  const b = (type, here = true, where = "", status = "Scheduled") => ({ type, here, where: here ? "" : where, date: 10, status });
  const tx = n => Array.from({ length: n }, () => b("Daily"));
  const g = T.noTabRole({ role: "home" }, [...tx(20), b("F/U", false, "Post")], 10);
  assert.equal(g.role, "home");                                                                   // 20 treatments here: their tracker is here, the F/U location doesn't decide
  const h = T.noTabRole({ role: "home" }, [...tx(5), b("F/U", false, "Post")], 10);
  assert.deepEqual([h.role, h.home, h.fuAway], ["guest", "Post", true]);                         // 5 here: a guest, home named from the F/U
  assert.equal(T.noTabRole({ role: "home" }, [b("MT"), ...tx(5), b("F/U", false, "Post")], 10).role, "home");   // MT here: still this clinic's
  assert.equal(T.noTabRole({ role: "home" }, [...tx(20), b("F/U", false, "Post", "Cancelled")], 10).role, "home");   // a cancelled F/U doesn't count
  assert.equal(T.noTabRole({ role: "home" }, [...tx(20), b("F/U", true)], 10).role, "home");     // F/U here: unchanged
  assert.equal(T.noTabRole({ role: "home" }, [...tx(20), b("F/U", null)], 10).role, "home");     // report doesn't name the clinic: unchanged
  assert.equal(T.noTabRole({ role: "home" }, [b("F/U", false, "Post")], 10).role, "home");       // F/U only, nothing here: unchanged
});

test("rankTabs: same LASFIR first, then the same last or first three letters, filtered by what's typed", () => {
  const tabs = ["ZEDKIM", "SMIJAN 2", "SMIBOB", "SMYJAN", "ALVMAR", "DOEJAN", "SMIJAM"].map(name => ({ name, code: name.split(" ")[0] }));
  const r = T.rankTabs("SMIJAN", "", tabs);
  assert.equal(r[0].name, "SMIJAN 2");
  assert.deepEqual(r.slice(1, 3).map(t => t.name), ["SMIJAM", "SMIBOB"]);                    // same last three, fewest letters apart first
  assert.ok(r.find(t => t.name === "DOEJAN").near && r.find(t => t.name === "SMYJAN").near);
  assert.equal(r.find(t => t.name === "ZEDKIM").near, false);
  assert.deepEqual(T.rankTabs("SMIJAN", "alv", tabs).map(t => t.name), ["ALVMAR"]);            // typing filters by name
  assert.equal(T.rankTabs("", "", tabs).length, 7);
});

test("mapEvent: theta burst \"Tx\" visits are treatments, as are the other NextGen TMS events", () => {
  const t = e => (T.mapEvent(e) || {}).type ?? null;
  assert.equal(t("TMS Theta Tx Bilateral"), "Daily");
  assert.equal(t("TMS Theta Tx Unilateral"), "Daily");
  assert.equal(t("TMS Treatment Unilateral"), "Daily");
  assert.equal(t("MT MagV"), "MT");
  assert.equal(t("MT Redo BW H7"), "MTR");
  assert.equal(t("TMS Follow Up"), "F/U");
  assert.equal(T.mapEvent("Telepsych TMS F/U").tele, true);
  assert.equal(t("Med Management"), null);
});

test("needsTracker: 16+ treatments here in all (a course mostly done here) needs a tracker", () => {
  const b = (type, date, status = "Kept", here = true) => ({ type, here, date, status });
  const past = Array.from({ length: 14 }, (_, i) => b("Daily", i + 1)), fut = [b("Daily", 20, "Scheduled"), b("Daily", 21, "Scheduled")];
  assert.equal(T.needsTracker([...past, ...fut], 20), true);
  assert.equal(T.needsTracker([...past.slice(2), ...fut], 20), false);
  assert.equal(T.noTabRole({ role: "home" }, [...past, ...fut], 20).role, "home");
});

test("mrnsIn: the MRN by its label in the first rows, never a bare number such as a date serial", () => {
  assert.deepEqual(T.mrnsIn([["PATIENT:", "", "", 36], ["SMIJAN"], ["MRN:", "", "", "", "", 46300], [12345]]), ["12345"]);   // A3 over A4
  assert.deepEqual(T.mrnsIn([["MRN: 0045678"]]), ["45678"]);
  assert.deepEqual(T.mrnsIn([["", ""], ["", ""], ["MRN", "00789"]]), ["789"]);                 // B3 beside it
  assert.deepEqual(T.mrnsIn([[46300, 36], ["Auth", 46400]]), []);
});

test("telNote: 2 missed in a row starts a telephone note; each miss after adds an entry; a treatment ends the run", () => {
  const r = (date, o = {}) => ({ date, type: "Daily", delivered: !o.miss, miss: false, note: "", ...o });
  assert.equal(T.telNote([r(1), r(2, { miss: true })], 5), null);
  const a = T.telNote([r(1), r(2, { miss: true }), r(3, { miss: true })], 5);
  assert.deepEqual([a.n, a.due, a.first], [2, true, true]);
  const logged = T.telMissLine("Smith", "1/3", 2);
  assert.equal(logged, "Tel note to Dr. Smith 1/3 (missed 2)");
  const b = T.telNote([r(1), r(2, { miss: true }), r(3, { miss: true, note: logged })], 5);
  assert.equal(b.due, false);
  const c = T.telNote([r(1), r(2, { miss: true }), r(3, { miss: true, note: logged }), r(4, { miss: true })], 5);
  assert.deepEqual([c.n, c.due, c.first, c.add.map(x => x.date)], [3, true, false, [4]]);
  assert.equal(T.telMissLine("Smith", "1/4", 3, true), "Tel note entry 1/4 (missed 3)");
  assert.equal(T.telNote([r(1, { miss: true }), r(2, { miss: true }), r(3)], 5), null);       // treated since
});

test("fuRebook: a missed or cancelled F/U not rebooked asks to rebook, then after 2 business days a telephone note", () => {
  const lbl = s => `d${s}`, mon = 46328;   // Monday, Nov 2, 2026
  assert.equal(T.bizDays(mon, mon + 4), 4); assert.equal(T.bizDays(mon + 4, mon + 7), 1);   // Fri to Mon is one business day
  const fu = (date, o = {}) => ({ date, type: "F/U", miss: false, note: "", row: 20, ...o });
  const a = T.fuRebook([fu(mon, { miss: true })], [], mon + 1, lbl);
  assert.deepEqual(a.map(e => [e.kind, e.telDue]), [["missed", false]]);
  assert.equal(T.fuRebook([fu(mon, { miss: true })], [], mon + 2, lbl)[0].telDue, true);
  assert.deepEqual(T.fuRebook([fu(mon, { miss: true }), fu(mon + 7)], [], mon + 3, lbl), []);   // rebooked
  const c = T.fuRebook([], [{ date: mon + 7, mod: mon }], mon + 2, lbl);                       // cancelled ahead: the clock starts at the cancellation
  assert.deepEqual(c.map(e => [e.kind, e.telDue]), [["cancelled", true]]);
  const line = T.telFuLine("Smith", "x", lbl(mon + 7));
  assert.equal(T.fuRebook([{ date: mon, type: "Daily", note: line }], [{ date: mon + 7, mod: mon }], mon + 2, lbl)[0].logged, true);
  assert.deepEqual(T.fuRebook([fu(mon - 40, { miss: true })], [], mon, lbl), []);               // older than 30 days
});

test("fuGaps: a follow-up between the MT and MT redo, around treatments 20 to 27, and within 7 days of the last", () => {
  const rows = [{ date: 100, type: "MT" }];
  for (let i = 2; i <= 36; i++) rows.push({ date: 100 + i, type: i === 10 ? "MTR" : "Daily" });
  const at = k => rows[k - 1].date;
  assert.deepEqual(T.fuGaps(rows, 36, 90).map(w => w.k), ["mtr", "mid", "end"]);
  const ok = [...rows, { date: at(5), type: "F/U" }, { date: at(23), type: "F/U" }, { date: at(36) + 5, type: "F/U", tele: true }];
  assert.deepEqual(T.fuGaps(ok, 36, 90), []);
  assert.deepEqual(T.fuGaps([...ok.slice(0, -1), { date: at(36) + 9, type: "F/U" }], 36, 90).map(w => w.k), ["end"]);   // 9 days after: too late
  assert.deepEqual(T.fuGaps([...ok.slice(0, -1), { date: at(36) - 6, type: "F/U", miss: true }], 36, 90).map(w => w.k), ["end"]);   // a missed one doesn't count
  assert.deepEqual(T.fuGaps(rows, 36, at(28)).map(w => w.k), ["end"]);                       // windows already past don't show
  assert.deepEqual(T.fuGaps(rows.slice(0, 18), 36, 90).map(w => w.k), ["mtr"]);              // treatment 20 not booked yet
  assert.equal(T.fuGaps(rows.filter(r => r.type !== "MTR"), 36, 90)[0].label, "between the MT and treatment 15");
  assert.deepEqual(T.fuGaps([{ date: 100, type: "Daily" }], 36, 90), []);                   // no MT, no plan
});
