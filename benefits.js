/* benefits.js: import Sohar benefits into tracker tabs. Pure functions (no Office.js, no DOM), shared with tests/.
   Input is the Sohar Eligibility add-in's "Benefits (TMS Tracker)" CSV, or its Sohar Results sheet saved as CSV.
   Matching is by MRN. Nothing here writes; planBenefits returns the proposed cell changes for the pane to show and apply. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SoharBenefits = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const norm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const normMrn = (s) => String(s ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "").replace(/^0+/, "");
  const blank = (v) => v == null || String(v).trim() === "";
  const num = (v) => { if (blank(v)) return null; const n = Number(String(v).replace(/[$,%\s]/g, "")); return isNaN(n) ? null : n; };

  // Columns of the Sohar add-in's benefits export (BENEFITS_EXPORT_COLS). The Sohar Results sheet uses the same names.
  const FIELDS = {
    perNbr: "Per Nbr", mrn: "MRN", last: "Last Name", first: "First Name", dob: "DOB", payer: "Sohar Payer", plan: "Plan", network: "Network",
    level: "Coverage Level", covStart: "Coverage Start", covEnd: "Coverage End", copay: "Copay", coins: "Coinsurance",
    ded: "Deductible", dedRem: "Deductible Remaining", dedStatus: "Deductible Status", oop: "OOP Max", oopRem: "OOP Remaining",
    oopStatus: "OOP Status", flag: "Benefit Flag", status: "Sohar Status", date: "Sohar Date", sentAs: "Sent As",
  };

  /** 2D rows from the CSV/xlsx -> {records} or {error}. The header row is the first one with MRN and a benefit column. */
  function parseSoharBenefits(rows) {
    const hi = (rows || []).findIndex((r) => { const h = (r || []).map(norm); return h.includes("mrn") && (h.includes("deductibleremaining") || h.includes("oopremaining")); });
    if (hi < 0) return { error: 'This is not a Sohar benefits file. Use "Benefits (TMS Tracker)" from the Sohar Eligibility add-in; it needs MRN, Deductible Remaining and OOP Remaining columns.' };
    const idx = {};
    rows[hi].forEach((h, i) => { const k = Object.keys(FIELDS).find((f) => norm(FIELDS[f]) === norm(h)); if (k && idx[k] == null) idx[k] = i; });
    const records = [];
    for (const r of rows.slice(hi + 1)) {
      if (!(r || []).some((c) => !blank(c))) continue;
      const o = {};
      for (const [k, i] of Object.entries(idx)) o[k] = String(r[i] ?? "").trim();
      // The Sohar Results sheet also has discovery rows; only verified benefits belong in a tracker.
      if (o.sentAs && o.sentAs !== "Verification") continue;
      records.push(o);
    }
    return { records };
  }

  /** What the tracker shows today for W4 / W5: "N/A", a number, or "" when blank. */
  const current = (na, amount, isBlank) => (na ? "N/A" : amount != null ? amount : isBlank ? "" : "");
  const same = (a, b) => (a === "N/A" || b === "N/A" ? a === b : num(a) != null && num(b) != null && Math.abs(num(a) - num(b)) < 0.005);

  /**
   * records: from parseSoharBenefits. patients: tracker tabs as the pane reads them
   *   ({ tab, mrn (normalized), mrnRaw, course, status, insurance, ben: { dedNA, dedRem, dedBlank, oopNA, oopRem } }).
   * Returns { rows: [{ tab, mrnRaw, rec, changes: [{ cell, label, from, to, kind: 'fill'|'change'|'same', on }], note }], unmatched: [rec] }.
   * W4 / W5 hold the deductible and OOP max remaining at the start of the course (column A's PR formula subtracts visits
   * charged since), so blank cells are filled by default and different values are offered unticked for staff to decide.
   */
  function planBenefits(records, patients) {
    const latest = new Map();
    for (const r of records || []) {
      const k = normMrn(r.mrn);
      if (!k) continue;
      const prev = latest.get(k);
      if (!prev || String(r.date) > String(prev.date)) latest.set(k, r);
    }
    const closed = (p) => /discontinued|complete/i.test(p.status || "");
    const rows = [], unmatched = [];
    for (const [k, rec] of latest) {
      const tabs = (patients || []).filter((p) => p.mrn === k);
      if (!tabs.length) { unmatched.push(rec); continue; }
      // The current course: an open tab first, then the highest course number.
      const p = tabs.slice().sort((a, b) => (closed(a) - closed(b)) || (b.course - a.course))[0];
      const out = { tab: p.tab, mrnRaw: p.mrnRaw || rec.mrn, rec, changes: [], note: "" };
      const st = String(rec.status || "").toLowerCase();
      if (st && !st.startsWith("complete.eligible")) {
        out.note = `Sohar: ${rec.status}. Nothing imported; check coverage before the next visit.`;
        rows.push(out); continue;
      }
      const add = (cell, label, from, to) => {
        if (to === "" || to == null) return;
        const kind = from === "" || from == null ? "fill" : same(from, to) ? "same" : "change";
        out.changes.push({ cell, label, from, to, kind, on: kind === "fill" });
      };
      const ded = rec.dedStatus === "None" ? "N/A" : num(rec.dedRem);
      const oop = num(rec.oopRem) != null ? num(rec.oopRem) : num(rec.oop) === 0 ? "N/A" : null;
      add("W4", "Deductible remaining", current(p.ben.dedNA, p.ben.dedRem, p.ben.dedBlank), ded);
      add("W5", "Out-of-pocket max remaining", current(p.ben.oopNA, p.ben.oopRem, p.ben.oopRem == null && !p.ben.oopNA), oop);
      if (blank(p.insurance) && !blank(rec.payer)) add("E6", "Insurance", "", rec.plan ? `${rec.payer} (${rec.plan})` : rec.payer);
      if (!out.changes.length) out.note = "No deductible or OOP amounts in the Sohar file.";
      else if (out.changes.some((c) => c.kind === "change")) out.note = "The tracker already has a different amount. W4 and W5 are the amounts at the start of the course, so only replace them if the tracker's number is wrong.";
      rows.push(out);
    }
    rows.sort((a, b) => a.tab.localeCompare(b.tab));
    return { rows, unmatched };
  }

  /** Ticked changes -> write ops for applyOps ({ tab, a, v: [[value]] }). */
  function benefitOps(plan) {
    const ops = [];
    for (const r of plan.rows) for (const c of r.changes) if (c.on && c.kind !== "same") ops.push({ tab: r.tab, a: c.cell, v: [[c.to]] });
    return ops;
  }

  return { parseSoharBenefits, planBenefits, benefitOps, normMrn };
});
