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
    // Final tracker file only: which service the patient is booked for (TMS / ESK), and where the policy came from.
    service: "Service", source: "Source",
  };

  /** 2D rows from the CSV/xlsx -> {records} or {error}. The header row is the first one with MRN and a benefit column. */
  function parseSoharBenefits(rows) {
    const hi = (rows || []).findIndex((r) => { const h = (r || []).map(norm); return h.includes("mrn") && (h.includes("deductibleremaining") || h.includes("oopremaining")); });
    if (hi < 0) return { error: 'This is not a Sohar benefits file. Use "Benefits (TMS Tracker)" from the Sohar Eligibility add-in; it needs MRN, Deductible Remaining and OOP Remaining columns.' };
    const idx = {};
    rows[hi].forEach((h, i) => { const k = Object.keys(FIELDS).find((f) => norm(FIELDS[f]) === norm(h)); if (k && idx[k] == null) idx[k] = i; });
    const records = [];
    let otherService = 0;
    for (const r of rows.slice(hi + 1)) {
      if (!(r || []).some((c) => !blank(c))) continue;
      const o = {};
      for (const [k, i] of Object.entries(idx)) o[k] = String(r[i] ?? "").trim();
      // The Sohar Results sheet also has discovery rows; only verified benefits belong in a tracker.
      if (o.sentAs && o.sentAs !== "Verification") continue;
      // The final tracker file covers TMS and ESK; rows booked only for another service are left out.
      if (o.service && !/\btms\b/i.test(o.service)) { otherService++; continue; }
      records.push(o);
    }
    return { records, otherService };
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
      else if (out.changes.some((c) => c.kind === "change")) out.note = "The tracker already has a different amount. Deductible and out-of-pocket max remaining are the amounts at the start of the course, so only replace them if the tracker's number is wrong.";
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

  // ---------- Eligibility feed (contract v1: Eligibility History.csv from the 7 AM rotation, served by a Power Automate flow) ----------
  const FEED_COLS = ["Pulled At", "LASFIR", "MRN", "Per Nbr", "Clinic", "Kind", "Sohar Status", "Sohar Payer", "Plan", "Network", "Coverage Level", "Coverage Start", "Coverage End",
    "Copay", "Coinsurance", "Deductible", "Deductible Remaining", "Deductible Status", "OOP Max", "OOP Remaining", "OOP Status", "Benefit Flag",
    "Sohar ID", "Sohar Link", "Related Sohar IDs", "Projected Patient Responsibility", "Estimate Note", "Source"];
  /* The common benefits block every tracker keeps, appended after the TMS-owned Benefits columns: [header, feed column].
     Deductible and OOP remaining here are Sohar's current numbers; the TMS-owned columns D and E stay the amounts at the
     start of the course (column A's PR formula reads those), so the feed never writes D or E. */
  const COMMON = [["Sohar Payer", "Sohar Payer"], ["Plan", "Plan"], ["Network", "Network"], ["Coverage Level", "Coverage Level"], ["Coverage Start", "Coverage Start"],
    ["Coverage End", "Coverage End"], ["Copay", "Copay"], ["Coinsurance", "Coinsurance"], ["Deductible", "Deductible"], ["Deductible remaining (Sohar)", "Deductible Remaining"],
    ["Deductible Status", "Deductible Status"], ["OOP Max", "OOP Max"], ["OOP remaining (Sohar)", "OOP Remaining"], ["OOP Status", "OOP Status"], ["Benefit Flag", "Benefit Flag"],
    ["Sohar Status", "Sohar Status"], ["Sohar Date", "Pulled At"], ["Sent As", "Kind"], ["Sohar ID", "Sohar ID"], ["Sohar Link", "Sohar Link"], ["Related Sohar IDs", "Related Sohar IDs"],
    ["Projected Patient Responsibility", "Projected Patient Responsibility"], ["Estimate Note", "Estimate Note"], ["Last Checked", null]];

  /** Quoted CSV text -> 2D rows (BOM, CRLF, "" escapes). */
  function parseCsv(text) {
    const t = String(text ?? "").replace(/^﻿/, ""), rows = []; let row = [], f = "", q = false;
    for (let i = 0; i < t.length; i++) {
      const c = t[i];
      if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
      else if (c === '"') q = true; else if (c === ",") { row.push(f); f = ""; }
      else if (c === "\n") { row.push(f.replace(/\r$/, "")); rows.push(row); row = []; f = ""; } else f += c;
    }
    if (f || row.length) { row.push(f.replace(/\r$/, "")); rows.push(row); }
    return rows.filter((r) => r.some((c) => c !== ""));
  }
  /** The flow's response -> {rows: [{col: value}]} or {error}. Accepts CSV text, JSON wrapping it as {body} / {csv} / {content},
      and {csv, rows}: rows are the Sohar webhook's table rows (tblEligFeed), merged by Sohar ID, the newer Pulled At winning.
      Table rows without a Sohar ID (Excel's blank first row) are skipped. */
  function parseFeed(text) {
    let t = String(text ?? ""), extra = [];
    if (/^\s*[{"]/.test(t)) { try { const j = JSON.parse(t); t = typeof j === "string" ? j : j.body ?? j.csv ?? j.content ?? ""; if (j && Array.isArray(j.rows)) extra = j.rows; } catch { /* plain CSV */ } }
    const g = parseCsv(t), h = (g[0] || []).map((x) => String(x).trim());
    if (!["MRN", "Sohar ID", "Pulled At", "Kind"].every((c) => h.includes(c)) && !extra.length) return { error: "That isn't the eligibility feed. The flow should return Eligibility History.csv. Nothing changed." };
    const rows = g.slice(1).map((r) => Object.fromEntries(h.map((k, i) => [k, String(r[i] ?? "").trim()])));
    const byId = new Map(); const keep = (r) => { const id = r["Sohar ID"]; if (!id) return; const c = byId.get(id); if (!c || pulledAt(r) >= pulledAt(c)) byId.set(id, r); };
    rows.forEach(keep);
    extra.forEach((o) => keep(Object.fromEntries(FEED_COLS.map((c) => [c, String((o && o[c]) ?? "").trim()]))));
    return { rows: [...byId.values()] };
  }
  const pulledAt = (r) => Date.parse(r["Pulled At"]) || 0;
  /** Feed rows -> Map(normalized MRN -> the row that counts): the newest eligible verification, else the newest of any kind. */
  function currentByMrn(rows) {
    const out = new Map();
    for (const r of [...(rows || [])].sort((a, b) => pulledAt(b) - pulledAt(a))) {
      const k = normMrn(r.MRN); if (!k) continue;
      const elig = r.Kind === "Verification" && /^complete\.eligible/i.test(r["Sohar Status"]);
      const c = out.get(k); if (!c || (elig && !c.elig)) out.set(k, { row: r, elig });
    }
    return new Map([...out].map(([k, v]) => [k, v.row]));
  }
  /** One MRN's pulls, newest first. */
  const pullsFor = (rows, mrn) => (rows || []).filter((r) => normMrn(r.MRN) === normMrn(mrn)).sort((a, b) => pulledAt(b) - pulledAt(a));
  /** Sohar page for an id. The feed doesn't say whether a related id is a verification or a discovery, so related ids link as
      verifications unless the caller knows better (Sohar shows "not found" for the wrong kind; the row's own link is exact). */
  const soharUrl = (id, kind) => `https://app.soharhealth.com/${/disc/i.test(kind || "") ? "discovery" : "verifications"}/${encodeURIComponent(id)}`;

  /** Which tracker this workbook is: sheet names first (strongest), then the file name. "unknown" is a new or empty workbook. */
  function trackerKind(sheetNames, wbName) {
    const s = new Set((sheetNames || []).map(norm)), w = String(wbName || "");
    if (s.has("ivktrackerdata")) return "IVK";
    if (s.has("eskvisitlog") || (s.has("inventory") && s.has("masterpatienttracker"))) return "ESK";
    if (["newtemplate", "h1template", "h7template", "protocols", "clinicschedule"].some((n) => s.has(n))) return "TMS";
    if (/esketamine|spravato|\besk\b/i.test(w)) return "ESK";
    if (/\bivk\b|iv ketamine/i.test(w)) return "IVK";
    if (/\btms\b|magv|brainsway|pt\.? tracking/i.test(w)) return "TMS";
    return "unknown";
  }

  return { parseSoharBenefits, planBenefits, benefitOps, normMrn, FEED_COLS, COMMON, parseCsv, parseFeed, currentByMrn, pullsFor, soharUrl, trackerKind };
});
