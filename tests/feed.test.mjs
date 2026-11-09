import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
const B = createRequire(import.meta.url)("../benefits.js");
// Synthetic data only: invented LASFIR codes and fake ids.
const row = (o) => B.FEED_COLS.map((c) => o[c] ?? "");
const q = (v) => (/[",\n]/.test(v) ? `"${String(v).replace(/"/g, '""')}"` : v);
const csv = (rows) => "﻿" + [B.FEED_COLS, ...rows.map(row)].map((r) => r.map(q).join(",")).join("\r\n") + "\r\n";
const R = [
  { "Pulled At": "2026-11-02T10:00:00Z", LASFIR: "RIVANA", MRN: "0900", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": "v-old", Copay: "20" },
  { "Pulled At": "2026-11-09T10:00:00Z", LASFIR: "RIVANA", MRN: "900", Kind: "Verification", "Sohar Status": "error.member.id", "Sohar ID": "v-new", "Estimate Note": 'a, "quoted" note' },
  { "Pulled At": "2026-11-08T10:00:00Z", LASFIR: "OKACHI", MRN: "901", Kind: "Discovery", "Sohar Status": "No coverage found", "Sohar ID": "d-1" },
];
test("parseFeed reads the rotation's CSV, BOM and quotes included", () => {
  const f = B.parseFeed(csv(R)); assert.equal(f.rows.length, 3); assert.equal(f.rows[1]["Estimate Note"], 'a, "quoted" note');
});
test("parseFeed accepts a JSON-wrapped body and rejects other files", () => {
  assert.equal(B.parseFeed(JSON.stringify({ body: csv(R) })).rows.length, 3);
  assert.ok(B.parseFeed("MRN,Copay\n1,2").error);
});
test("parseFeed merges the webhook table rows by Sohar ID, newer wins, blank rows skipped", () => {
  const rows = [
    { "Pulled At": "2026-11-09T15:00:00Z", MRN: "900", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": "v-new", Copay: "25", "@odata.etag": "x" },
    { "Pulled At": "2026-11-10T09:00:00Z", MRN: "902", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": "v-web" },
    { "Pulled At": "", MRN: "", "Sohar ID": "" },
  ];
  const f = B.parseFeed(JSON.stringify({ csv: csv(R), rows }));
  assert.equal(f.rows.length, 4);
  assert.equal(f.rows.find((r) => r["Sohar ID"] === "v-new").Copay, "25");
  assert.equal(f.rows.find((r) => r["Sohar ID"] === "v-web").MRN, "902");
  assert.equal(B.parseFeed(JSON.stringify({ csv: "", rows })).rows.length, 2);
});
test("currentByMrn keeps the newest eligible verification over a newer error", () => {
  const c = B.currentByMrn(B.parseFeed(csv(R)).rows);
  assert.equal(c.get("900")["Sohar ID"], "v-old"); assert.equal(c.get("901")["Sohar ID"], "d-1");
});
test("pullsFor lists one MRN newest first, matching leading zeros", () => {
  assert.deepEqual(B.pullsFor(B.parseFeed(csv(R)).rows, "900").map((r) => r["Sohar ID"]), ["v-new", "v-old"]);
});
test("trackerKind tells TMS from ESK and IVK workbooks", () => {
  assert.equal(B.trackerKind(["New Template", "Benefits", "RIVANA"], "Castro MagV Patient Tracker.xlsx"), "TMS");
  assert.equal(B.trackerKind(["Master Patient Tracker", "Inventory", "Benefits"], "SanRafael Esketamine Master Tracker.xlsx"), "ESK");
  assert.equal(B.trackerKind(["Master Patient Tracker", "IVK Tracker Data"], "Castro IVK Patient Tracker.xlsx"), "IVK");
  assert.equal(B.trackerKind(["Sheet1"], "Book1"), "unknown");
  assert.equal(B.trackerKind(["Sheet1"], "Post Esketamine Master Tracker.xlsx"), "ESK");
  assert.equal(B.trackerKind(["H7 Template", "RIVANA 2"], "San Rafael Pt. Tracking - BW.xlsx"), "TMS");
});
