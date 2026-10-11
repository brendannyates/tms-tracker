# TMS Tracker (Excel on the web add-in)

A light augmentation over the organization's TMS patient tracker template. Information stays where the template already keeps it, and the add-in adds a small, recognizable layer. It runs in **Excel on the web**; desktop add-ins are not available at the organization.

## Status

Version 1.2 follows the reformatted MagV template (the "MagV Template" workbook): the add-in writes only into the template's own cells.

Tested with fictional patients:

- in a headless browser, against an in-memory copy of the real template's cells: reading tabs, BIDF and freebie flags, benefits, extensions, availability, Add patient, rebuilding an earlier-layout tab, the import planner, the schedule window, and the pop-out window relaying every read and write through the task pane
- in desktop Excel, on a scratch copy of the template: the column A PR formula, with and without a deductible

**Not yet run in Excel on the web.** Test on a copy of the real workbook first.

## Repo contents

| Path | What it is |
|---|---|
| `manifest.xml` | Office add-in manifest (version 1.2). Requires ExcelApi 1.9, which Excel on the web supports. Points to the GitHub Pages pilot host, `https://brendannyates.github.io/tms-tracker`, served from `main`. |
| `taskpane.html`, `taskpane.js` | The add-in. A single static page with no build step. |
| `schedule.html` | The patient schedule window: print, save as PDF, or full screen. |
| `assets/` | Ribbon icons (from the logo mark), the clinic logo used on printouts, and the NextGen export icon shown in the report directions. |
| `demo/tms-tracker-demo.html` | A self-contained demo with fictional patients, built for version 1.1 (the earlier layout). Not yet updated for 1.2. |
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

The template owns the layout. The add-in reads and writes these cells and adds only the **Before ded.** column (Y11:Y16).

| Cells | Purpose |
|---|---|
| A2, A4, A6 | LASFIR, MRN, DOB. |
| D2 / D4 / D6 / D7 | # of TXs (authorized treatments, extensions included), # of MT redos, MD required on-site, Status. |
| E2 / E4 / E6 | Provider (set from the appointment list on import), care navigator (TAS), insurance. |
| F4:G4, G2, G6 | Auth start and expiry, update protocol by, progress note sent on. |
| J2:O4, N5 | Protocol rows 1 to 3 (Hz, pulses per train, trains, wait; N calculates minutes; O is the site) and the total duration. |
| N6 / N7 | Auth assessment; **Remaining Freebies**. Using a freebie takes one off N7, undoing it gives it back. |
| A10:A49 (PR) | Patient responsibility per visit, by visit type in L (follow-ups by whether Notes says "tele"). With a deductible in W4, visits charge the Before ded. cost in date order until the deductible is used up; the visit that crosses it pays what is left plus PR on the rest; then PR. Missed visits show no charge. |
| B / C / E / F, H, J | Treating technician, Tx #, date, pulses Left / Right / Other. Pulses entered = treatment delivered. |
| L / M / O / P, Q / R / S / T | Visit type, Miss x, Action, PHQ-9 and GAD-7 (x = placeholder), Alt Clinic x, Notes, appointment time. |
| S (Notes) | "Tele" marks a telehealth follow-up. A missed visit's freebie decision is "Freebie used" or "Freebie waived". The MT row also holds the MT details. "Schedule provided m/d" is written when a schedule is given. |
| S9 | Home Clinic. Anything other than SRL makes the patient external. |
| W1 / W4 / W5 | BIDF estimated course cost; deductible remaining and out-of-pocket max remaining, each an amount or **N/A**. |
| W6 / W7 / W8 | MTR approved (with D4, decides whether 90869 is authorized); extensions added; last schedule provided. |
| V11:Y16 | Cost by code: 90867 MT, 90868 Daily, 90869 MT redo, 99214 F/U in-office and telehealth. X is PR, Y is the cost before the deductible (only when one applies). |
| W20:X24, V29:Y48 | Weekly availability (earliest and latest start, Monday to Friday) and dates (Unavailable / Available, from, to, reason). |

**Update tabs** (the banner) adds the Before ded. column and the PR formula to New Template and every tab. A PR cell is only replaced when it is blank or still holds a cost-table lookup; a typed amount stays. Tabs in an earlier layout (for example the pilot's Freebie and Time columns in T and U) are rebuilt: a fresh copy of New Template gets their entries, benefits, costs and availability, and the earlier tab is kept as "NAME (old)" so you can check it and delete it.

## Layout

- **Header:** the clinic logo and the workbook's name (for example "San Rafael Pt Tracking - MagV"), with the tracker device under it. **Import list**, **Open in a new window**, **Full screen**, **Settings** (gear) and **Refresh** sit at the top right.
- **Banners:** the stale appointment list (with the NextGen directions), new patients needing a tracker, tabs left untouched, and tabs needing updates. Each banner can be dismissed for the session or snoozed for 3 hours.
- **Main menu:**

| Menu | What it holds |
|---|---|
| Schedule | A Today / Week toggle. **Attention** is a section on the left when the pane is wide enough (about 640 px); on a narrow pane it collapses to the top with a count. Only groups with something in them show. |
| Patients | The roster with filters: All, Active, Follow-up not scheduled, No upcoming visit. **Add patient** is here. Selecting a patient opens the patient view. |
| Clinic | Chair and Protocols. |

**Settings** is its own screen (gear, top right): clinic location, device, warning days, recent-visit window, the Update Active Patients sheet button and Place measure placeholders.

### New window and full screen

**Open in a new window** (⧉) opens the whole tracker in its own window; **Full screen** (⛶) opens it at full size, and inside that window the same button switches the browser to full screen. In Excel on the web only the task pane can reach the workbook, so the window sends every read and save through the pane: keep the pane open while you use it. The pane shows a notice, and **Close the other window** closes it.

### Today

- **Schedule variances:** "X of Y patients in treatment are being treated", with the reasons: in follow-up, missed, away, at another clinic, not scheduled, and graduated this week.
- **Visit cards** are colored like the tab's row highlight: MT light blue, MTR green, F/U pink, other clinic orange. Each shows **Payment $X** (from column A), or **Payment not set** when the cost for that visit type is missing. A patient missing BIDF info shows **BIDF info needed** on the card.
- **Record missed** (No-show or Same-day cancel) opens a pop-up with **Use freebie** or **Waived**. It warns when the patient will have 1 or 0 freebies left. The decision goes in Notes and Remaining Freebies (N7) counts down. The card shows the result with Undo.
- **Mark pulses delivered** shows what the protocol calls for and writes F, H, J and the technician.
- **Draft treatment note** builds the daily note; **Copy note** puts it on the clipboard.
- **Measures needed** opens PHQ-9 and GAD-7 entry. **Move measures** sends them to a chosen upcoming visit and notes it on the original row, suggesting an MT, MTR or F/U within 7 days.
- **MT** cards add MT details and the protocol picker. Daily cards add the note drafter, with the MT baseline sleep and caffeine as a quiet reference.

### Week

Monday to Friday visits with a one-line variance summary per day.

### Attention

- **BIDF info needed today:** a patient scheduled today without cost by code, a deductible (or N/A), the before-deductible costs when a deductible applies, or the out-of-pocket max (or N/A). **Enter benefits** opens their Benefits and cost section.
- Status prompts and auth flags: auth start after the first TMS visit, expiry within 14 calendar days, treatments past expiry, more treatments booked than D2 allows, an MT redo booked without authorization.
- Protocol renewal: G2 or an Action = Protocol within 2 business days.
- Benefits: projected cost above the BIDF estimate, likely to meet the out-of-pocket max.
- Booked while the patient is away (date ranges only), closures, same-day appointments (not Kaiser).
- Freebie decisions, pulses or technician missing, and measures not collected.

### Patients

- **Roster:** MRN, patient and status, a Monday-to-Friday column per day (MT, MTR, F/U, an action, Alt or X for missed), next follow-up, graduation date, MD on-site, and the date a schedule was last given.
- **Patient view:**
  - **Summary:** MRN, insurance, MD on-site, provider, care navigator, authorization, treatments and MT redos, graduation, next follow-up, freebies, home clinic, auth assessment and progress note sent.
  - **Protocol editor.**
  - **Benefits and cost:** cost so far and projected, the deductible and out-of-pocket max (amount or N/A), the BIDF estimate, and the cost by code (PR and Before ded.), all editable here. **Add extension** adds daily visits to D2 and W7 (+15 by default) and can move the auth expiry.
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

Either one writes "Schedule provided m/d" to Notes and the date to W8. **Full screen** fills the screen with the schedule. The schedule contains the MRN, so keep it only where PHI is allowed. Excel may ask you to allow the window the first time. When the tracker is already open in its own window, the schedule shows over the page instead, since Excel allows one add-in window at a time.

## Open items

- Holiday calendar: Brendan to provide it to pre-fill the Clinic Schedule closures.
- Schedule planner (finding 60-minute chair slots for an MT or MTR against provider and patient availability) comes later.
- The UI map (`tms-tracker-ui-map.html`) and the demo reflect the earlier design and need regenerating for this version.
