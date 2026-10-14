import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
const B = createRequire(import.meta.url)("../benefits.js");

// Synthetic data only. Header matches the Sohar Eligibility add-in's BENEFITS_EXPORT_COLS.
const HDR = ["Per Nbr", "MRN", "Last Name", "First Name", "DOB", "Sohar Payer", "Plan", "Network", "Coverage Level", "Coverage Start", "Coverage End",
  "Copay", "Coinsurance", "Deductible", "Deductible Remaining", "Deductible Status", "OOP Max", "OOP Remaining", "OOP Status", "Benefit Flag",
  "Sohar Status", "Sohar Date", "Sohar Ref"];
const rec = (o) => HDR.map((h) => o[h] ?? "");
const ben = (o) => ({ dedNA: false, dedRem: null, dedBlank: true, oopNA: false, oopRem: null, ...o });
const PATIENTS = [
  { tab: "RIVANA", mrn: "900", mrnRaw: "0900", course: 1, status: "Active", insurance: "", ben: ben({}) },
  { tab: "OKACHI", mrn: "901", mrnRaw: "901", course: 1, status: "Active", insurance: "Cigna", ben: ben({ dedRem: 500, dedBlank: false, oopRem: 3000 }) },
  { tab: "LINSVE", mrn: "902", mrnRaw: "902", course: 1, status: "Discontinued", insurance: "", ben: ben({}) },
  { tab: "LINSVE 2", mrn: "902", mrnRaw: "902", course: 2, status: "Active", insurance: "", ben: ben({}) },
  { tab: "MENAFI", mrn: "903", mrnRaw: "903", course: 1, status: "Active", insurance: "Aetna", ben: ben({}) },
];

function file() {
  return [HDR,
    rec({ "Per Nbr": "100", MRN: "900", "Sohar Payer": "Blue Shield of California", Plan: "SYNTH PPO", Copay: "40", Coinsurance: "20%", Deductible: "1500", "Deductible Remaining": "0", "Deductible Status": "Met", "OOP Max": "5000", "OOP Remaining": "800", "Sohar Status": "complete.eligible", "Sohar Date": "2026-10-14" }),
    rec({ "Per Nbr": "101", MRN: "901", "Sohar Payer": "Cigna", Deductible: "500", "Deductible Remaining": "500", "OOP Max": "3000", "OOP Remaining": "2500", "Sohar Status": "complete.eligible", "Sohar Date": "2026-10-14" }),
    rec({ "Per Nbr": "102", MRN: "0902", "Sohar Payer": "Aetna", Deductible: "0", "Deductible Remaining": "0", "Deductible Status": "None", "OOP Max": "4000", "OOP Remaining": "4000", "Sohar Status": "complete.eligible", "Sohar Date": "2026-10-14" }),
    rec({ "Per Nbr": "103", MRN: "903", "Sohar Payer": "Aetna", "Sohar Status": "complete.ineligible (INACTIVE)", "Sohar Date": "2026-10-14" }),
    rec({ "Per Nbr": "104", MRN: "999", "Sohar Payer": "UMR", "Deductible Remaining": "10", "Sohar Status": "complete.eligible" }),
  ];
}

test("parseSoharBenefits finds the header and rejects other files", () => {
  assert.equal(B.parseSoharBenefits(file()).records.length, 5);
  assert.match(B.parseSoharBenefits([["Appt Date", "MRN"], ["1/1/2026", "1"]]).error, /not a Sohar benefits file/);
});

test("planBenefits fills blanks, offers changes unticked, skips ineligible, reports unmatched", () => {
  const plan = B.planBenefits(B.parseSoharBenefits(file()).records, PATIENTS);
  const by = (t) => plan.rows.find((r) => r.tab === t);

  const a = by("RIVANA");
  assert.deepEqual(a.changes.map((c) => [c.cell, c.to, c.kind, c.on]), [["W4", 0, "fill", true], ["W5", 800, "fill", true], ["E6", "Blue Shield of California (SYNTH PPO)", "fill", true]]);

  const b = by("OKACHI");
  assert.deepEqual(b.changes.map((c) => [c.cell, c.from, c.to, c.kind, c.on]), [["W4", 500, 500, "same", false], ["W5", 3000, 2500, "change", false]]);
  assert.match(b.note, /start of the course/);

  assert.equal(by("LINSVE"), undefined, "the discontinued course is not chosen");
  assert.deepEqual(by("LINSVE 2").changes.find((c) => c.cell === "W4").to, "N/A", "no deductible becomes N/A; MRN leading zeros ignored");

  assert.equal(by("MENAFI").changes.length, 0);
  assert.match(by("MENAFI").note, /INACTIVE/);

  assert.deepEqual(plan.unmatched.map((r) => r.mrn), ["999"]);
});

test("benefitOps writes only ticked, non-identical changes", () => {
  const plan = B.planBenefits(B.parseSoharBenefits(file()).records, PATIENTS);
  plan.rows.find((r) => r.tab === "OKACHI").changes.find((c) => c.cell === "W5").on = true;
  const ops = B.benefitOps(plan);
  assert.ok(ops.some((o) => o.tab === "OKACHI" && o.a === "W5" && o.v[0][0] === 2500));
  assert.ok(!ops.some((o) => o.tab === "OKACHI" && o.a === "W4"), "identical values are not rewritten");
  assert.ok(ops.some((o) => o.tab === "RIVANA" && o.a === "E6"));
});

test("Sohar Results rows from discovery are ignored", () => {
  const rows = [["MRN", "Sent As", "Deductible Remaining", "Sohar Status"], ["900", "Discovery", "", "Coverage found"], ["900", "Verification", "100", "complete.eligible"]];
  assert.equal(B.parseSoharBenefits(rows).records.length, 1);
});
