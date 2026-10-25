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
  assert.deepEqual(n.map(x => x.k), ["sched", "start"]);
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
