import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
const T = createRequire(import.meta.url)("../core.js");

// Synthetic names only.
test("clinic: no San Rafael fallback when the workbook name says nothing", () => {
  assert.equal(T.clinicFromName("Book1.xlsx"), "");
  assert.equal(T.clinicFromName("Castro MagV Tracker"), "Castro");
  assert.deepEqual(T.clinicNames(""), []);
  assert.deepEqual(T.clinicNames("Castro"), ["castro"]);
  assert.deepEqual(T.clinicNames("SRL"), ["san rafael", "srl"]);
  assert.deepEqual(T.clinicNames("San Rafael"), ["san rafael", "srl"]);
});

test("atClinic: SRL alias only when the clinic is San Rafael", () => {
  assert.equal(T.atClinic("San Rafael MagV1", "Castro"), false);
  assert.equal(T.atClinic("Castro MagV1", "Castro"), true);
  assert.equal(T.atClinic("SRL TMS", "San Rafael"), true);
  assert.equal(T.atClinic("San Rafael MagV1", "SRL"), true);
  assert.equal(T.atClinic("Anything", ""), true);   // unknown clinic: everything is here, no clinic assumed
});

test("homeHere: an unknown clinic never makes SRL home over others", () => {
  assert.equal(T.homeHere("SRL", ""), T.homeHere("Castro", ""));
  assert.equal(T.homeHere("SRL", "Castro"), false);
  assert.equal(T.homeHere("CAS", "Castro"), true);
  assert.equal(T.homeHere("Castro TMS", "Castro"), true);
  assert.equal(T.homeHere("SRL", "San Rafael"), true);
  assert.equal(T.homeHere("San Rafael", "SRL"), true);
  assert.equal(T.homeHere("Walnut Creek", "SRL"), false);
});

test("parseUsers: header skipped, blanks and duplicates dropped, rows kept", () => {
  const u = T.parseUsers([T.USER_HDR, ["Alex P.", "a", "b", ""], ["", 1], ["alex  p.", "x"], ["Sam K.", "", "", "{}"]]);
  assert.deepEqual(u.map(x => [x.name, x.row]), [["Alex P.", 2], ["Sam K.", 5]]);
  assert.deepEqual(T.parseUsers(null), []);
});

test("cleanUserName", () => {
  assert.equal(T.cleanUserName("  Jo   D. "), "Jo D.");
  assert.equal(T.cleanUserName("O'Neil-Smith"), "O'Neil-Smith");
  assert.equal(T.cleanUserName("J"), "");
  assert.equal(T.cleanUserName("x@y.com"), "");
  assert.equal(T.cleanUserName("1234"), "");
});

test("pickOrder: this computer's people first (latest first), then A to Z, none pre-selected", () => {
  const users = [{ name: "Zed A." }, { name: "Bea C." }, { name: "Alex P." }, { name: "Sam K." }];
  const o = T.pickOrder(users, { "sam k.": "2026-11-04T10:00:00Z", "zed a.": "2026-11-05T09:00:00Z" });
  assert.deepEqual(o.map(x => [x.name, x.here]), [["Zed A.", true], ["Sam K.", true], ["Alex P.", false], ["Bea C.", false]]);
  assert.ok(o.every(x => !("selected" in x)));
  assert.deepEqual(T.pickOrder(users, {}).map(x => x.here), [false, false, false, false]);   // a new computer: nobody tagged
});

test("wbKey: per workbook, URL query ignored, id fallback", () => {
  const a = T.wbKey("https://x.sharepoint.com/Castro MagV Tracker.xlsx?web=1"), b = T.wbKey("https://x.sharepoint.com/Castro MagV Tracker.xlsx");
  assert.equal(a, b);
  assert.notEqual(a, T.wbKey("https://x.sharepoint.com/Post MagV Tracker.xlsx"));
  assert.match(T.wbKey("", "abc123"), /^tms_who_/);
  assert.notEqual(T.wbKey("", "abc123"), T.wbKey("", "abc124"));
  assert.equal(T.wbKey("", ""), "");
});

test("dayKey: local date", () => {
  assert.equal(T.dayKey(new Date(2026, 0, 5, 23, 59)), "2026-01-05");
});

/* the code itself isn't in this public repo: set TMS_ADMIN_CODE in the terminal to check it unlocks */
const CODE = process.env.TMS_ADMIN_CODE || "";
test("adminOk: the admin code unlocks", { skip: !CODE && "set TMS_ADMIN_CODE to check the right code" }, async () => {
  assert.equal(await T.adminOk(CODE), true);
  assert.equal(await T.adminOk(` ${CODE} `), true);
});
test("adminOk: wrong codes don't", async () => {
  for (const c of ["", "12", "12345", "1234", "abcd", "0000", "9999"]) assert.equal(await T.adminOk(c), false, c);
});
