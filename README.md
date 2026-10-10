# TMS Tracker (Excel on the web add-in)

A light augmentation over the organization's TMS patient tracker template. Information stays where the template already keeps it, and the add-in adds a small, recognizable layer. It runs in **Excel on the web**; desktop add-ins are not available at the organization.

## Status

Tested in Node and in a headless browser with sample data:

- event and status mapping against the real 10/9 export
- the import planner, including cancellations, telehealth tags and measures re-planning
- the measures cadence
- payments and benefit flags
- the Today, Patients and patient views

**Not yet run in Excel.** Test on a copy of the real workbook first.

## Repo contents

| Path | What it is |
|---|---|
| `manifest.xml` | Office add-in manifest (version 1.1). Requires ExcelApi 1.9, which Excel on the web supports. Points to the GitHub Pages pilot host, `https://brendannyates.github.io/tms-tracker`, served from `main`. |
| `taskpane.html`, `taskpane.js` | The add-in. A single static page with no build step. |
| `schedule.html` | The patient schedule window: print or save as PDF. |
| `assets/` | Ribbon icons (from the logo mark), the clinic logo used on printouts, and the NextGen export icon shown in the report directions. |
| `demo/tms-tracker-demo.html` | A self-contained demo with fictional patients. Open it in a browser; no Excel needed. |
| `scripts/set-host.mjs` | Rewrites whatever host the manifest currently points to. |
| `package.json` | Helper scripts: `certs`, `start`, `validate`, `check`, `set-host`. |

The files contain no PHI. `.gitignore` blocks spreadsheets and CSVs so a real workbook or NextGen export is never committed by accident.

## Setup

**Pilot (GitHub Pages):** the committed manifest already points to the Pages host. Upload it in Excel on the web: Insert > Add-ins > Upload My Add-in. Changes go live when they merge to `main`.

**Try it locally:**
1. `npm run set-host -- https://localhost:3000` (don't commit that change), `npm run certs` (once), then `npm start`.
2. In Excel on the web: Insert > Add-ins > Upload My Add-in > `manifest.xml`.
3. Open the add-in from the ribbon (Home > TMS). If tabs need updating, the banner offers **Update tabs**.

**Deploy for the team:**
1. Host the folder on an internal HTTPS location approved by Product/Engineering.
2. Run `npm run set-host -- https://YOUR-HOST/tms-tracker-addin`.
3. Run `npm run validate`.
4. Have the Microsoft 365 admin deploy `manifest.xml` through Integrated Apps, or share it for Upload My Add-in.
5. Bump `<Version>` in the manifest whenever the hosted files change, so Excel picks up the update.

## Where things live on a patient tab

| Cells | Purpose |
|---|---|
| A10:A49 (PR) | Patient responsibility per visit. A formula picks the amount by visit type in L, and for follow-ups by whether the note says "tele". Only blank cells get the formula. |
| B / C / E / F, H, J | Treating technician, Tx #, date, and pulses Left / Right / Other. Pulses entered = treatment delivered. |
| L / M / O / P, Q / R / S | Visit type, Miss x, Action, PHQ-9 and GAD-7 (x = placeholder), Alt Clinic x, Notes. |
| S (Notes) | "Tele" marks a telehealth follow-up. The MT row also holds the MT details: hotspot, MT %, amps, baseline sleep and baseline caffeine. "Schedule provided m/d" is written here when a schedule is given. |
| S9 | Home Clinic. Anything other than SRL makes the patient external. |
| D6 / D7 / E2 / E4 / E6 | MD required on-site, Status, Provider, Care navigator (TAS), Insurance. |
| J2:N4, O1:O4 | Protocol rows 1 to 3 (Hz, pulses per train, trains, wait; N calculates minutes). O labels each row Left, Right or Other. The legend moved one column right (swatches in P, labels in Q:S). |
| T1:U4, T, U | Freebies allowed (U1, default 3), used, waived and remaining. Per-visit freebie; appointment time. |
| Y6:Z15 | Benefits: plan type, insurance type, deductible remaining, out-of-pocket max remaining, MTR approved, extensions added, last schedule provided, and the BIDF estimated course cost (Z15, directly above the codes). |
| Y16:AE22 | Auth by code: 90867, 90868, 90869, 99214 in-office and 99214 telehealth. Columns are auth qty, used, left, charge and patient pays. |
| Y25:AB48, AD25:AG31 | Dates away (with a reason) and weekly availability for Monday to Friday. |
| AI / AJ | Treatment report complete; Payment collected. Both are optional checklist marks. |
| AM1:AN1 | Date the external patient-specific report was last uploaded. |

The earlier pilot columns V (tele), W (cost), X (provider) and AK (location) are no longer used. Update tabs clears them.

## Layout

- **Header:** the clinic logo and the workbook's name (for example "San Rafael Pt Tracking - MagV"), with the tracker device under it. **Import list**, **Settings** (gear) and **Refresh** sit at the top right.
- **Banners:** the stale appointment list (with the NextGen directions), new patients needing a tracker, tabs left untouched, and tabs needing updates. Each banner can be dismissed for the session or snoozed for 3 hours.
- **Main menu:**

| Menu | What it holds |
|---|---|
| Schedule | A Today / Week toggle. **Attention** is a section on the left when the pane is wide enough (about 640 px); on a narrow pane it collapses to the top with a count. Only groups with something in them show. |
| Patients | The roster with filters: All, Active, Follow-up not scheduled, No upcoming visit. **Add patient** is here. Selecting a patient opens the patient view. |
| Clinic | Chair and Protocols. |

**Settings** is its own screen (gear, top right): clinic location, device, warning days, recent-visit window, external report stale days, the Update Active Patients sheet button and Place measure placeholders.

### Today

- **Schedule variances:** "X of Y patients in treatment are being treated", with the reasons: in follow-up, missed, away, at another clinic, not scheduled, and graduated this week.
- **Visit cards** are colored like the tab's row highlight: MT light blue, MTR green, F/U pink, other clinic orange. Each shows **Payment $X** or **Payment N/A**.
- **Record missed** (No-show or Same-day cancel) opens a pop-up with **Use freebie** or **Waived**. It warns when the patient will have 1 or 0 freebies left. The card then shows the result with Undo.
- **Mark pulses delivered** shows what the protocol calls for and writes F, H, J and the technician.
- **Payment collected** is greyed out when there is no payment. **Treatment report complete** is optional and never nags.
- **Measures needed** opens PHQ-9 and GAD-7 entry. **Move measures** sends them to a chosen upcoming visit and notes it on the original row, suggesting an MT, MTR or F/U within 7 days.
- **MT** cards add MT details and the protocol picker. Daily cards add the note drafter, with the MT baseline sleep and caffeine as a quiet reference.

### Week

Monday to Friday visits with a one-line variance summary per day.

### Attention

- Status prompts and auth flags: auth start after the first TMS visit, expiry within 14 calendar days, treatments past expiry, quantity by code.
- Protocol renewal: G2 or an Action = Protocol within 2 business days.
- Benefits: projected cost above the BIDF estimate, likely to meet the out-of-pocket max.
- Booked while the patient is away (date ranges only), external report stale, schedule out, closures, same-day appointments (not Kaiser).
- Freebie decisions, pulses or technician missing, and measures not collected.

### Patients

- **Roster:** MRN, patient and status, a Monday-to-Friday column per day (MT, MTR, F/U, an action, Alt or X for missed), next follow-up, graduation date, MD on-site, and the date a schedule was last given.
- **Patient view:**
  - **Summary:** MRN, insurance, MD on-site, provider, care navigator, authorization, treatments, graduation, next follow-up, freebies and home clinic.
  - **Protocol editor.**
  - **Billing and benefits:** estimated balance (record a payment, or clear it, which marks the visits collected), auth by code, and an **extension** (+15 daily visits by default, with a new auth expiry).
  - **Measures over time.**
  - **Appointments:** upcoming by default, with all visits on request.
  - **Availability editor:** weekly hours plus as many date ranges as needed.
  - **Buttons:** Print or save schedule, Open tab.

### Clinic

Chair (hours and holiday closures come from the Clinic Schedule sheet) and Protocols.

## Import

- **No setup needed:** the report's columns are detected automatically. There is no mapping and no ignore pattern to manage.
- **Preview:** upload the file and the preview appears with Apply below it.
- **Device:** each patient's device is decided by where their daily treatments are booked. A Brainsway patient's MT on a provider schedule is ignored on a MagV tracker.
- **Today and later:** dates, times and types are updated. Telehealth follow-ups get "Tele" in Notes. E2 (Provider) is set to the provider named on the report, replacing anything typed there. The schedule uses E2.
- **Past visits:** never change, except no-shows and same-day cancels (cancelled with Mod Dt on the visit date). Those are always recorded, including in an empty row if needed. Earlier cancellations are noted and their date cleared.
- **Measures:** they are re-planned only after a significant disruption: an MTR or F/U moved, or 2 or more missed visits within a week.

## Measures placeholders

- **Always:** MT, MTR, every F/U, and the last treatment.
- **In between:**
  - Every 5 treatments, when the gap between milestones is 8 or more.
  - At least every 7 calendar days when visits are sparse, as in tapering.
- **Never placed:** within 2 days or 2 visits before a milestone.
- **Existing entries:** placeholders only fill blank cells. Scores and existing x's never move.

## Schedule printout

**Print schedule** (in a visit card's ⋯ menu, or **Print or save schedule** in the patient view) opens a pop-up to choose the visits: **Next 14 days** (default), **Next 30 days**, **Rest of course**, or any From / To range. It shows how many visits fall in the range.

**Open schedule** opens the schedule in its own window (`schedule.html`), not in a workbook tab. Nothing is added to the workbook. The window shows the clinic logo centered over the table, the report's "As of" time, and the visits (no patient-name column).

- **Print** opens the browser's print dialog, Letter landscape. Long schedules run onto more pages with the header row repeated.
- **Save as PDF** downloads `SMIJOH schedule 10.6.pdf`. If it fails, use Print and choose Save as PDF.

Either one writes "Schedule provided m/d" to Notes and the date to Z13. The schedule contains the MRN, so keep it only where PHI is allowed. Excel may ask you to allow the window the first time.

## Open items

- Holiday calendar: Brendan to provide it to pre-fill the Clinic Schedule closures.
- Schedule planner (finding 60-minute chair slots for an MT or MTR against provider and patient availability) comes later.
- The UI map (`tms-tracker-ui-map.html`) reflects the earlier design and needs regenerating for this version.
