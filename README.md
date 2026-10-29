# TMS Tracker (Excel on the web add-in)

A light augmentation over the organization's TMS patient tracker template. Information stays where the template already keeps it, and the add-in adds a small, recognizable layer. It runs in **Excel on the web**; desktop add-ins are not available at the organization.

Staff install and day-to-day use: see [GUIDE.md](GUIDE.md).

## Status

Version 1.3 (v2 screens: chair board, Patients, Sessions sheet, chair moves) follows the reformatted MagV template (the "MagV Template" workbook): on the patient tabs the add-in writes only into the template's own cells.

Tested with fictional patients:

- in a headless browser, against an in-memory copy of the real template's cells: reading tabs, BIDF and freebie flags, benefits, extensions, availability, Add patient, rebuilding an earlier-layout tab, the import planner, the schedule window, and the pop-out window relaying every read and write through the task pane
- in desktop Excel, on a scratch copy of the template: the column A PR formula, with and without a deductible
- v2 screens, in headless Edge against a fake Office.js and an invented three-chair clinic day: every Today step (check in, motor threshold, start, end with pulses and measures, note, undo), missed visits and freebies, chair moves and the Moved in NextGen tick, chairs marked unavailable, each Patients need's action and the cells it writes, Schedule, Import apply, Add patient, Settings, the full-screen snapshot, keyboard focus, and a pass that presses every button on every screen; `core.js` has unit tests

**Not yet run in Excel on the web.** Test on a copy of the real workbook first.

## Repo contents

| Path | What it is |
|---|---|
| `manifest.xml` | Office add-in manifest (version 1.2). Requires ExcelApi 1.9, which Excel on the web supports. Points to the GitHub Pages pilot host, `https://brendannyates.github.io/tms-tracker`, served from `main`. |
| `taskpane.html`, `taskpane.js` | The add-in. A single static page with no build step. |
| `core.js` | The floor rules (pure functions): visit phases, chair board states, chair moves against NextGen, missed streaks, and the Patients needs list. |
| `benefits.js` | Sohar benefits import: matching and plan (pure functions). |
| `tests/` | Tests for `core.js` and `benefits.js` (`npm test`). |
| `schedule.html` | The patient schedule printout, shown over the add-in: print or save as PDF. |
| `assets/` | Ribbon icons (from the logo mark), the clinic logo used on printouts, and the NextGen export icon shown in the report directions. |
| `scripts/set-host.mjs` | Rewrites whatever host the manifest currently points to. |
| `package.json` | Helper scripts: `certs`, `start`, `validate`, `check`, `set-host`. |

The files contain no PHI. `.gitignore` blocks spreadsheets and CSVs so a real workbook or NextGen export is never committed by accident.

## Setup

**Pilot (GitHub Pages):** the committed manifest already points to the Pages host. Upload it in Excel on the web: Insert > Add-ins > Upload My Add-in. Changes go live when they merge to `main`.

**Try it locally:**
1. `npm run set-host -- https://localhost:3000` (don't commit that change), `npm run certs` (once), then `npm start`.
2. In Excel on the web: Insert > Add-ins > Upload My Add-in > `manifest.xml`.
3. Open the add-in from the ribbon (Home > TMS). If tabs don't match New Template, the banner offers **Match New Template**.

**Deploy for the team:**
1. Host the folder on an internal HTTPS location approved by Product/Engineering.
2. Run `npm run set-host -- https://YOUR-HOST/tms-tracker-addin`.
3. Run `npm run validate`.
4. Have the Microsoft 365 admin deploy `manifest.xml` through Integrated Apps, or share it for Upload My Add-in.
5. Bump `<Version>` in the manifest whenever the hosted files change, so Excel picks up the update.

## Where things live on a patient tab

New Template owns the layout. The add-in reads and writes these cells; benefits and cost by code live on the **Benefits** tab, one row per MRN.

| Cells | Purpose |
|---|---|
| A2, A4, A6 | LASFIR, MRN, DOB. |
| D2 / D4 / D6 / D7 | # of TXs (authorized treatments, extensions included), # of MT redos (number format `0`, set on every refresh, Add patient and Match New Template), MD required on-site, Status. |
| E2 / E4 / E6 | Provider's last name (set from the appointment list on import), care navigator (TAS), insurance. |
| F4:G4, G2, G6 | Auth start and expiry, update protocol by, progress note sent on. |
| J2:O4, N5 | Protocol rows 1 to 3 (Hz, pulses per train, trains, wait; N calculates minutes; O is the site) and the total duration. |
| N6 / N7 | **Freebie Count** (freebies left, 3 on a new tab); auth assessment. Using a freebie takes one off N6, undoing it gives it back. |
| A10:A49 (PR) | The visit log is A10:T49 in the template and can grow on import (see Import). Patient responsibility per visit, by visit type in L (follow-ups by whether Notes says "tele"). It reads the patient's Benefits row by the MRN in A4 (blank when there is none). With a deductible remaining, visits charge the before-deductible cost in date order until the deductible is used up; the visit that crosses it pays what is left plus PR on the rest; then PR. Missed visits show no charge. |
| B / C / E / F, H, J | Treating technician, Tx #, date, pulses Left / Right / Other. Pulses entered = treatment delivered. |
| L / M / O / P, Q / R / S / T | Visit type, Miss x, Action, PHQ-9 and GAD-7 (x = placeholder), Alt Clinic x, Notes, appointment time. |
| S (Notes) | "Tele" marks a telehealth follow-up. A missed visit's freebie decision is "Freebie used" or "Freebie waived". The MT row also holds the MT details. "Schedule provided m/d" is written when a schedule is given. |
| S9 | Home Clinic. Anything other than SRL makes the patient external. |
| Row after the log (B: TX SUMMARY) | Closes the log. Ticking the checkbox in A turns the row green. |
| W20:X24, V29:Y48 | Weekly availability (earliest and latest start, Monday to Friday) and dates (Unavailable / Available, from, to, reason). |

**Benefits tab** (one row per MRN, the newest course's tab in Patient): BIDF estimated course cost; deductible remaining and out-of-pocket max remaining (each an amount or **N/A**); MTR approved (with D4, decides whether 90869 is authorized; set to Yes on refresh when D4 is 1 or more); extensions added; last schedule provided; cost by code (PR, then before the deductible) for 90867 MT, 90868 Daily, 90869 MT redo, 99214 F/U in-office and telehealth; **Last updated**, set to today whenever the add-in or a Sohar import changes or verifies the row. The add-in still calls these W1, W4 to W8 and X12:Y16 inside (the cells they used to have), and sends those writes to the Benefits row.

**Match New Template** (the banner, or Settings for every active tab) brings active trackers (see the cutoff in Settings) in line with New Template: header formats, Freebie Count in N6 and the auth assessment in N7, benefits moved from V1:Y16 to the Benefits tab (availability stays on the tab), the PR formula over the whole log, the template's formulas on log rows that lost them, the TX SUMMARY row and its checkbox, New Template's conditional formats (colors by visit type in L, Action in O, Miss, Alt Clinic, auth dates) and dropdowns stretched to the last log row, then visit-log formatting and measures. The log's last row is the row above TX SUMMARY, so rows added by hand count. Rows an import adds get the same rules. A PR cell is only replaced when it is blank or still holds a cost-table lookup; a typed amount stays. Tracker tabs from before the Benefits tab are still read until they're matched; finished ones are left as they are, and their benefits are copied to the Benefits tab so it lists every patient. Tabs in an earlier layout (for example the pilot's Freebie and Time columns in T and U) are rebuilt: a fresh copy of New Template gets their entries, benefits, costs and availability, and every copied entry is read back from the new tab. When they all match, the earlier tab is deleted; if anything does not, it stays as "NAME (old)" and the add-in names the cells to check. Match New Template asks before rebuilding. A deleted tab can be recovered from the workbook's version history.

## Layout (v2)

The screens follow the MHS tracker design shared with ESK Tracker (tokens, chair board, visit and course cards, copy rules). They're built for the full-screen window at 1366x768 and up; in the task pane the menu becomes a bar across the top. Times on the floor screens read 24-hour (9:05) and dates "Oct 21". Patients show by LASFIR only; MRNs appear only in the import previews, the new-patient banner (last four digits) and printed schedules.

- **Menu:** the workbook's name and device, then **Today**, **Schedule**, **Patients** (with a count of patients that need attention, red when one needs action now), **Chairs**, **Import**, **Settings**, **Refresh** and **Full screen**.
- **Banners:** the stale appointment list (with the NextGen directions), new patients needing a tracker, tracker tabs left untouched, and tabs that don't match New Template. Each can be dismissed for the session or snoozed for 3 hours.

### Today

- **View:** Clinic, or one chair (saved per computer; the appointment list check and treatment counts follow it).
- **Chair board:** one row per chair from the saved appointment list (its MagV or BrainsWay resources; one chair named for the device until a list is imported). States: open (names who is next), in session with time left and a progress bar, ending (3 min or less, or over time) in amber, and waiting in red (only for a visit checked in on an earlier version): 10+ min past their time with nothing started. The lamp is lit only while a session runs. Chair rules are pure functions in `core.js`.
- **Still to come:** everyone not started, with time, chair, and the top need on their course or "Not started, 25 min"; **Start session** or **Choose a chair** beside each.
- **Appointment list check** and **Treatments today** (the variances: in follow-up, missed, away, at another clinic, not scheduled, graduated this week).
- **Visit card:** the meta line has the protocol time ("Daily, treatment 12 of 36 · protocol 19 min", from N5, else N2:N4 added up); before the session starts the lead adds when it's projected to end (after the chair frees up, if it's in use). Two steps, no separate check-in. **Start session in MagV 2** writes the check-in and start times together (Sessions H and I) and opens the treatment note draft, or **Record motor threshold** on an MT or MT redo (MT details to Notes; at the first MT also the protocol to J2:O4). **Move to another chair** replaces it when that chair is busy and another is open. **Complete session**: pulses pre-filled from the protocol, technician, pain and muscle movements, PHQ-9 and GAD-7 when the row has the x (or move them to the next measures visit), and, when payment today is over $0, Collected or Not collected, saved as "Copay $X collected" in Notes (S). The note draft then reopens for **Copy note**. Also **Record missed** (No-show or Same-day cancel with Use freebie, Waived or Decide later, warning at 1 or 0 freebies and at 3 missed in a row), **Undo missed visit**, **Undo complete session** (clears the pulses, technician and copay line it wrote, only while the chair hasn't started another session), **Print schedule**, **Open tab**. Payment today (column A) or **Not set**; **BIDF info needed** with **Enter benefits**. Follow-ups: **Enter PHQ-9 and GAD-7**, **Mark follow-up done**. Selecting a visit also opens its tab at that row unless Settings turns it off.
- **Chair moves:** **Move to another chair** on any visit not started. The dialog shows each chair's current state and any booking the move would overlap. Every move prompts **Moved in NextGen**, in the dialog and on the card. Until it's ticked, Patients lists the move and the import keeps the tracker's chair; a list run after the move that shows the new chair, or a move back to NextGen's chair, also settles it (`core.js` moveState).

### Schedule

**Day by chair:** columns per chair plus a provider column (MT, MT redo and F/U), 1.6 px per minute over the Clinic Schedule hours, with a now-line today; selecting a visit opens it on Today (another day: its row on the tab). Notes for closures, visits outside chair hours, and visits with no time. **Week:** Monday to Friday cards with "X of Y" treated or booked and each day's variances. **Write this week to the Chair Schedule sheet** is here.

### Patients

One row per tab with its two most urgent needs; filters Needs attention (default), All active, Follow-up not scheduled, No upcoming visit, Tapering, Paused, Complete, Discontinued, Copay unconfirmed; **Add patient**. Every check from the earlier Attention list is a need on its course, red for what blocks care or billing (auth expired or overrun, no visit in the recent-visit window, 3 missed in a row), amber otherwise, each with its action: Add extension, Enter or Review benefits, Decide freebie, Enter pulses, Add technician, Enter scores, Update protocol, Edit availability, Set status, Call to reschedule (logged in Notes), Mark moved in NextGen, Choose a chair, Print course schedule, or Open tab. Choose a protocol (no protocol row with a site in O, from the MT on) and Re-verify benefits for {next year} are needs too; {year} OOPM met is a quiet tag.

**Patient view.** Selecting a row, or **Open** anywhere in the add-in (it replaces Open tab), opens the patient over the list (ink at 28% and a 6 px blur behind it; it fills the task pane). Focus stays inside it; Esc or × closes it and the row stays selected. The header has the LASFIR tab name and the MRN (A4), then "Dr. {E2} · {done} of {D2} done · {status}". A visit today comes first, on one band: time, visit, chair, its state and the next step (Start session, Complete session, Edit pulses, follow-up steps), **On Today**, and the copay line under it. Then the top need with its button and **Also needs**. A **sidebar** picks one section at a time, saved on this computer: **Schedule** (with Past visits), **Protocol**, **Benefits**, **Financial**; Edit availability and Change status sit under it.

- **Today's copay** (when there's a visit today), on one line: column A's PR, **Collected** or **Not collected** and **By**. The answer writes one copay line to S with the name and time, for example "Copay $25.00 collected by Priya 11:32 AM", and the line then shows "✓ Collected $25.00 · Priya, 11:32 AM" with **Change**. Complete session asks the same question and starts from this answer; either one replaces the line.
- **Benefits**: the Benefits row in two columns with **Edit benefits** and **Re-run eligibility**. An OOPM met is entered here, as 0 in Out-of-pocket max remaining.
- **Financial**: the payment row, the copay tally against what's left, and Copays to confirm.
- **Edit pulses** (a done visit's card on Today, the band, and each Past visits row with pulses) rewrites F, H, J and the technician; clearing a site clears it on the row.
- **Last treatment:** Complete session on treatment {D2} of {D2} asks **Mark Complete** or **Not yet**. Refresh no longer writes Complete on its own; until it's set, the course shows "Set to Complete".
- **Pulses with no NextGen visit:** a dated row with pulses, from the saved list's first day through today, with no visit of that kind in NextGen that day (cancelled ones aside) is a need and an Edit log check, with **Edit pulses**.
- Today's Still to come rows and the visit card show "MRN: 1234" (last four) in grey next to the LASFIR, where the chair used to be (the chair is in the card's meta line).
- "Choose a protocol" is now **Confirm protocol** (need "Protocol not confirmed"). The duplicate Add extension in the patient view's action bar is gone; an auth warning still offers it.
- **Protocol**: one line per J2:O4 row with a site in O, as "F3 (left) · 10 Hz · 120% MT · 3,000 pulses" (O reads Left or F3 as F3, Right or F4 as F4; % MT comes from the Protocols table). The name comes from the Protocols table: the same Hz and values is that protocol, the same Hz with other values is "{name} (Custom)" with what changed against the standard, anything else "Custom". Then treatment (Unilateral, or Bilateral "F3 then F4", and the minutes from N5), MT (from the MT row's Notes) and the next MT redo, device, course shape, authorization (F4:G4, treatments left of D2) and update-by (G2). **Edit**: a protocol per row from the table (its values fill in and stay editable), the site, and the **next expiration date**, required, which writes G2. Saving writes J:M and O (as Left, Right or Other, the template's dropdown) and adds "Protocol renewed, next by m/d" to today's Notes. With no protocol the section opens on the dropdown, values greyed.
- **Schedule**: the course strip with a payment row under it (amber: a delivered treatment with a payment due and no copay line; pale green: logged), the count to confirm, PHQ-9 and GAD-7, the next eight booked visits (date, time, chair from the Sessions sheet or the appointment list, Tx #, note, **Move** for today's visit), the weekly start windows and dates away, **Open in Schedule**, **Book the rest of the course** (goes to Import: book in NextGen, then import the list; the planner comes later), **Pause course**, **Print schedule**, and **Past visits**: every logged visit, newest first, with Tx #, chair, pulses, PR, the copay and who logged it, PHQ-9, GAD-7 and the note.
- **Copay and benefits**: the Benefits row (insurance from E6, the 90868 Daily cost as the copay per treatment visit, deductible and OOPM remaining, BIDF estimate and projected cost, Last updated, the plan year) with **Update BIDF** (Enter benefits) and **Re-run eligibility** (Import, for the Sohar file); the **copay tally** for the plan year (Last updated's year) against the OOPM and deductible remaining, the out-of-pocket flag with its caveat, and **Mark {year} OOPM met**; and **Copays to confirm** with **Mark selected paid**, **Mark all paid** and **Waive**.

**Copays and benefits rules** (`core.js`):
- **Copays to confirm**: delivered treatment visits (MT, Daily, Taper, MTR; never F/U) before today, with a payment due (column A, else the cost table) and no copay line in S. Marking writes "Copay $X collected (confirmed m/d by {first name})"; Waive writes "Copay waived by {first name} m/d".
- **Tally**: the "Copay $X collected" amounts on treatment visits in the plan year. When the tally plus today's payment reaches the OOPM remaining, the existing need "Out-of-pocket max likely met: set the copay to $0 at pre-check and update Insurance Maintenance in NextGen" shows (amber). Logged copays are a floor, so confirm in NextGen or re-run eligibility first.
- **TMS Follow Ups** keep their own 99214 PR in column A, but stay out of the projected course cost, the BIDF estimate check and the OOPM tally. Column A's PR formula is unchanged.
- **No benefits reminders when nothing is owed**: with the OOPM remaining at 0 on the Benefits row (entered as 0 in Edit benefits, which sets Last updated) or a BIDF estimate of $0, there is no BIDF info needed, no over-estimate or OOPM flag and no copays to confirm. OOPM met shows the quiet tag "{year} OOPM met".
- **Plan-year rollover**: a visit booked after Dec 31 of Last updated's year raises "Re-verify benefits for {next year}", amber from Dec 1, red once a new-year visit has happened or checked in. A Sohar import or Enter benefits in the new year clears it. It shows even when the OOPM is met, since the new year resets it.

**Active Patients** gains a **Copays unconfirmed** count (column O).

### Chairs

Chair cards (resource name, device, coil, today's count, last service) with **Mark unavailable** / **Mark available** and **Log service or coil**; chair hours and upcoming closures from the Clinic Schedule sheet (or **Create Clinic Schedule sheet**); the Protocols table. A chair marked unavailable leaves the board and the schedule, and its bookings show **No chair yet** until moved.

### Settings

Its own screen: clinic location, device, warning days, recent-visit window, whether selecting a visit opens its tab, Appearance (Light, Dark or Match Windows; saved on this computer, also in the rail as Dark mode), the cutoff for updates, and the workbook tools (**Match New Template**, **Fix visit-log formatting**, **Replan measures and pulse totals**, **Rebuild Active Patients sheet**).

### Full screen

**Full screen** floats the tracker inside the Excel page, over the workbook (an Office dialog with `displayInIframe`). It never opens a separate browser window. Dialogs blur the add-in behind them. The in-page window can't reach the workbook in Excel on the web, so it sends every read and save through the pane (`callPane`, including the Sessions sheet writes): keep the pane open while you use it. The pane shows a notice, and **Close full screen** closes it.

## Active Patients sheet

The add-in rebuilds this sheet in the Patients roster layout for the current week (Sunday to Saturday, showing Monday to Friday). It runs on open and after any change, and skips the rewrite when nothing changed.

- **Who is listed:** every tab whose D7 status is Pending Start, In Progress, Tapering, Paused or blank. Nobody keeps a list; Complete and Discontinued drop off.
- **Columns:** patient (links to the tab), MRN, status, Monday to Friday codes colored like the add-in (MT, MTR, F/U, Alt, X, · for a daily treatment, or the action), next F/U, graduation, auth expiry (red past, amber within the warning days), tx left, MD on-site, schedule given.
- **Above the grid:** "X of Y treated" and the variances for each day, as in the Week view. Today's column is outlined.
- **View only:** the sheet is protected without a password. Change things on the patient tabs or in the add-in. Typing on the sheet would be overwritten anyway.

## Sessions sheet

The add-in creates a **Sessions** sheet the first time it records a check-in, session time, chair move or chair setting, and reads it back on every refresh so everyone sees the same board.

| Cells | Purpose |
|---|---|
| A:J, from row 2 | One row per visit: date, tab, visit (TMS or F/U; TMS#2 for a second that day), chair (when moved), NextGen chair at the move, Moved in NextGen (TRUE/FALSE), moved at, checked in, started, ended (times of day). Keyed by date, tab and visit, not the log row, because an import can re-sort the log. |
| L:P, rows 2 to 21 | Chairs by NextGen resource name: device, coil, last service, unavailable (Yes). |

Every write reads the sheet again first and adds rows after what is there, so two people can use it at once; the board re-reads it every 30 seconds. Rows older than 45 days are dropped when the sheet reaches row 2000. Writes to a visit-log row check first that E still holds that visit's date (another person's import may have re-sorted the log); if not, nothing is saved and the add-in asks for a Refresh. Nothing about sessions is written to the patient tabs; pulses, technician, measures, misses and MT details still go to the visit log as before.

## Edit Log sheet

Every add-in write to a patient tab or the Benefits tab goes through `applyOps`, which reads what each cell held first and adds one row per tab to the hidden **Edit Log** sheet: A id, B when, C tab, D what (worded from the cells' meaning, `describeCells`), E kind, F by, G state ("" can be undone, "Undone ...", "No undo"), H on the undo data as JSON in 30,000-character parts. Imports, sorts and removed rows store the log's typed cells (B:T without D, G, I, K) before, the rows added or removed, and a hash of the log after. **Undo** runs only when the cells (or the whole log, by hash and row count) still hold what the edit left. The newest 300 show at the bottom of Patients; the sheet keeps about 2,000. Tab creation and Match New Template are logged without undo; the Sessions sheet, formatting and the Appt List tab are not logged.

The same panel lists checks per active tab (`core.js` logChecks and apptGaps): two rows for one day and kind of visit, Tx # (MT = 1, each Daily, Taper or MTR one more; F/U and misses blank), date order, rows whose A or D is empty (Range.valueTypes "Empty"), New Template's Kaiser reminder on a dated row, treatments numbered past D2, and the saved Appt List against the tab.

## Import

- **One Apply, on the logs it was made from:** a plan addresses rows by number and moves them, so applying it twice or after a tab changed doubled rows and put visits on the wrong rows. Apply is locked while it runs (one import at a time through the pane, pop-out included); before anything is written each tab is read again and compared with the log the preview was built from (`logSig`), and a changed tab rebuilds the preview instead. arrangeLogs moves nothing unless every E holds the date the plan wrote there. A no-show on a day already treated for that kind of visit is not recorded as a second, missed row. New Template's Kaiser reminder is dropped from any row an import dates.
- **No setup needed:** the report's columns are detected automatically. There is no mapping and no ignore pattern to manage.
- **Preview:** upload the file and the preview appears with Apply below it.
- **Device:** each patient's device is decided by where their daily treatments are booked. A Brainsway patient's MT on a provider schedule is ignored on a MagV tracker.
- **Today and later:** dates, times and types are updated. Telehealth follow-ups get "Tele" in Notes. E2 (Provider) is set to the last name of the provider on the report, replacing anything typed there: the part before a comma ("Smith, Jane MD"), otherwise the last word once MD, DO, NP, PMHNP, PA-C, PhD, PsyD and Dr. are dropped (`core.js` providerLast). The schedule uses E2.
- **Every Expected visit gets a row:** free rows fill in row order, so the log stays in date order. When a pool runs out (no-shows and MT redos use rows), rows are inserted in A:T only, after the visits they follow and never below the last log row, so Excel stretches the PR formulas, header counts and conditional formats over them. Each new row copies a row of the same visit type (fill, borders, validation, height, formulas) with the typed cells blank. The add-in finds a grown log from the range in A10's PR formula, so a tab without that formula (Match New Template adds it) is not grown and the preview says so.
- **Past visits:** never change, except no-shows and same-day cancels (cancelled with Mod Dt on the visit date). Those are always recorded, including in an empty row if needed. Earlier cancellations (any date) have their date and time cleared so the row is free for another visit, unless the patient is still booked or seen that day. A Kept visit on or after the tab's first dated row that has no row that day gets one, with its date and time, so the pulses can be entered.
- **Order and Tx #:** after every import the visit log reads top to bottom by date (open rows with no date last), so a follow-up sits between the treatments around it. Only the typed cells move (B, C, E, F, H, J, L to T); the formulas in A, D, G, I and K stay on their rows. C numbers the treatments down the log (`core.js` txNumbers): the MT is 1, each Daily, Taper or MTR one more than the treatment above it, missed visits and follow-ups blank, so D (D2 minus C) is the treatments left. The preview lists the tabs re-sorted or renumbered.
- **Formatting follows the dates:** after the writes, every log row with a date in E takes row 10's formatting (fill, font, borders, alignment, number formats), and a row without one takes New Template's blank (gray) row for that position, the last two rows mapping to its last two. Rows the import inserted, re-dated or re-sorted always follow; other rows only when they have no fill of their own, so a hand-colored row is left alone. Cell properties only: the conditional formats (MT, MTR, F/U, missed highlights) and validation stay. The same pass runs once each time the pane opens and from **Settings > Fix visit-log formatting**.
- **Duplicates:** the same visit listed twice (same day, type and time, for example on the provider's schedule and the chair) gets one row; the preview counts them. A no-show or same-day cancel an earlier list recorded is undone, freebie included, when NextGen now shows that day's visit Expected or Kept with no no-show; misses marked in the add-in stay.
- **Missed rows keep their visit:** a row marked missed (x in M) that NextGen hasn't recorded as a no-show answers that day's booking at the same time even while NextGen still lists it as Expected or Kept, so the visit never gets a second row and the schedule below isn't pushed down. A booking at another time that day, or a rebook after a NextGen no-show, is a second visit and gets its own row. Leave the type in L on a missed row; if it was emptied, the import puts back the type of that day's booking no other row holds (a treatment before a follow-up).
- **No-shows land on the right row:** a NextGen no-show or same-day cancel is recorded on that day's row for the same kind of visit (treatment or follow-up), so a follow-up no-show never marks the treatment row missed.
- **Times** read as 2:30 PM everywhere: on screen, in T on the tabs (set with the formatting below), and on the Sessions sheet.
- **Discontinued tabs** are not updated. If NextGen still has them booked, the preview names them, and so does the Day view's appointment list check.
- **Saved list:** applying an ALL TMS list copies it to the **Appt List** tab (replaced each time). The add-in reads it back on open, so the new-patient banner and a new tab's dates still work without uploading again, and **Preview again** in Import list reruns it. Patient-specific reports are not saved there.
- **Measures:** they are re-planned only after a significant disruption: an MTR or F/U moved, the log re-sorted, or 2 or more missed visits within a week.

## Clinic, guests and first launch

- **The clinic** comes from the workbook name ("Castro MagV Patient Tracker" is Castro; `core.js` clinicFromName) unless Settings names one. A chair is the clinic's when its NextGen resource or location names it ("Castro MagV1"). Only those chairs show on the board, the schedule and Chairs.
- **Who is imported** (`core.js` clinicRole): chair visits are counted by clinic, cancellations left out. No chair visit here: the patient is left out (an MT at Post with nothing here is Post's). Most chair visits at another clinic: a **guest**. A tie goes to where the MT was. A patient with a tab here is always imported. A report whose chairs never name the clinic isn't split, and the preview says so.
- **Guests** get no tab. They're listed on the **Guest Patients** sheet (one row per MRN: home clinic, chair visits here and there, next and last visit here, the link, On latest list, notes) and in a banner asking for the link to their tab on the home clinic's tracker, where pulses and the rest are entered. **Track here instead** opens Add patient. Rows stay when a patient drops off the list; delete them by hand. On Today's list check a guest shows as "Guest from Post" with **Open their tab**.
- **Blank MRN:** a banner lists tabs still in progress (Pending Start, In Progress, Tapering or blank) with a visit in the recent-visit window or later and nothing in A4, with the MRN from the saved list filled in when the LASFIR matches. **Save to A4** writes it (Edit log, undoable).
- **First launch:** a workbook without New Template gets it (and Protocols, if missing) from `assets/templates/MagV Template.xlsx` or `BrainsWay Template.xlsx` (ExcelApi 1.13), and Active Patients and Guest Patients are made. Without the file the add-in says to copy New Template in by hand. **The template file must be blank: no patients, no PHI, and no sensitivity label (an encrypted file can't be inserted).**
- The report's columns now also match Res Name (provider), Loc Name (location) and Beg Tm (time).

## Sohar benefits import

Import list > **Import Sohar benefits** takes the "Benefits (TMS Tracker)" file from the Sohar Eligibility add-in (or its Sohar Results sheet saved as CSV; discovery rows are ignored). Matching and the plan live in `benefits.js`, which has tests in `tests/` (`npm test`).

- **Matching:** by MRN (A4), ignoring punctuation and leading zeros. If a patient has more than one course, the open tab with the highest course number is used. The latest Sohar result per MRN wins.
- **What it writes:** deductible remaining (N/A when the plan has no deductible) and out-of-pocket max remaining on the patient's Benefits row, and E6 insurance (payer and plan) only when E6 is blank. Every verified patient gets **Last updated** = today, even when the amounts already match. Copay and coinsurance are shown for reference; the cost table is not changed.
- **Blank cells are filled by default.** Deductible and out-of-pocket max remaining are the amounts at the start of the course (column A's PR formula subtracts visits charged since), so a tab that already has a different amount is shown with both values and left unticked. Tick it only if the tracker's number is wrong.
- **Not imported:** ineligible or errored verifications (the preview says so), and patients with no tracker tab (listed by the last four MRN digits).
- Nothing changes until **Apply**.

## Measures placeholders

The rule is `planMeasures` in core.js, with tests.

- **Required:** MT, MTR, every F/U, and the last treatment.
- **In between:** the 5th treatment after the last measures. MT and MTR count as treatments (there's a treatment at both), F/U doesn't.
- **Skipped:** an in-between collection when a required one is 7 days away or less (by the dates in E). One collection covers both, so patients aren't asked too often.
- **Taken a day late:** x's on an earlier day with scores on the next row means the scores came a day late. The x's are cleared and the row with the scores is where the count starts again (it gets the Measures flag).
- **Not collected:** x's on an earlier day with no scores after them are cleared, and the next visit takes the measures instead. The count restarts there.
- **On Patients:** earlier measures that weren't entered show as one need, **Check NextGen for PHQ-9 and GAD-7**, on active courses only.

**Status:** when the treatments done reach D2 (36 of 36, 15 of 15), refresh writes Complete in D7. Paused and Discontinued are left alone. Freebie decisions are asked for misses in the last 10 days.
- **Upcoming visits** with a date get the plan: x's added where it wants them and removed where it doesn't. Rows with no date get none until an import dates them.
- **Flag:** Daily, Taper and F/U rows with measures also get Measures in O (when O is empty). MT and MTR rows get the x's only.
- **Never touched:** scores.
- **Pulse totals:** the same pass (Settings > **Replan measures and pulse totals**, and after imports) puts New Template's running-total formula (`=F10 + INDEX(G:G, ROW()-1)`, likewise I and K) back on any log row where G, I or K is empty. Without it on one row, such as a hand-typed F/U, every total below restarts at 0.

## Schedule printout

**Print schedule** (on a visit card or a patient's card on Patients; **Print course schedule** on an MT visit and as the need on its course until a schedule is given) opens a pop-up to choose the visits: **Next 14 days** (default), **Next 30 days**, **Rest of course**, or any From / To range. It shows how many visits fall in the range.

**Open schedule** shows the schedule over the add-in on a blurred backdrop (`schedule.html` in a frame), not in a workbook tab or a separate window. Nothing is added to the workbook. It shows the clinic logo centered over the table, the report's "As of" time, and the visits (no patient-name column).

- **Print** opens the browser's print dialog, Letter landscape. Long schedules run onto more pages with the header row repeated.
- **Save as PDF** downloads `SMIJOH schedule 10.6.pdf`. If it fails, use Print and choose Save as PDF.

Either one writes "Schedule provided m/d" to Notes and the date to W8. The schedule contains the MRN, so keep it only where PHI is allowed.

## Open items

- Holiday calendar: Brendan to provide it to pre-fill the Clinic Schedule closures.
- Schedule planner (finding 60-minute chair slots for an MT or MTR against provider and patient availability) comes later.
