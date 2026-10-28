# TMS Tracker add-in: install and user guide

The TMS Tracker is an Excel add-in for the MHS TMS patient tracker workbook. It opens as a panel beside the workbook (or full screen over it) and shows the chairs, today's visits and what each course needs, and it writes back into the same cells the template already uses. Nothing about the workbook changes for people who don't use the add-in.

This guide is for anyone at MHS who uses a TMS tracker workbook: technicians, care navigators, and clinic leads.

## Before you start

- **Excel on the web.** The add-in runs in Excel in a browser (Edge or Chrome). It does not run in the desktop Excel app at MHS.
- **A tracker workbook in OneDrive or SharePoint,** built from the current MagV template. It needs the **New Template** tab; **Protocols** and **Clinic Schedule** are used when present.
- **NextGen PM access** to run the appointment report.

## Install

You install it once per Microsoft account, not per workbook.

1. Download the manifest file: open https://brendannyates.github.io/tms-tracker/manifest.xml in your browser and save it (right-click > Save as, keep the name `manifest.xml`).
2. Open your tracker workbook in Excel on the web.
3. Go to **Home > Add-ins > More Add-ins** (in some versions, **Insert > Add-ins**), choose the **My Add-ins** tab, then **Upload My Add-in**.
4. Browse to `manifest.xml` and select **Upload**.
5. A **TMS** group with a **TMS Tracker** button appears on the Home tab. Select it to open the panel.

The manifest file only tells Excel where the add-in lives. It contains no patient information.

**Updates** arrive on their own. When a new version is released, close and reopen the panel (or reload the browser tab). You only need to upload the manifest again if you are told the manifest changed.

**If the TMS button disappears,** for example after clearing your browser's cache, repeat steps 2 to 4.

**To remove it,** go to **Home > Add-ins > More Add-ins > My Add-ins**, open the ⋯ menu on TMS Tracker and choose **Remove**.

## First time in a workbook

Do this once per clinic workbook. Settings save with the workbook, so everyone using it shares them.

1. Open the panel and choose **Settings** in the menu.
2. Check **This clinic (NextGen location)**, for example San Rafael. Visits booked at other locations are marked as another clinic, and the clinic's name is dropped from chair names (San Rafael MagV1 shows as MagV 1).
3. Check **Tracker device**. "From the workbook name" works when the name contains MagV or Brainsway; otherwise pick MagVenture or BrainsWay.
4. Leave the warning windows at their defaults unless your clinic wants different ones: auth warning 14 calendar days, protocol warning 2 business days, recent visit window 10 days.
5. If a banner says tabs don't match New Template, select **Match New Template**. Active trackers get New Template's layout, colors by visit type, dropdowns, borders and TX SUMMARY checkbox over every log row, and their benefits move to the **Benefits** tab (one row per MRN, with Last updated). Entries aren't touched. **Settings > Match New Template (all active tabs)** runs it on every active tracker. Tabs in an older layout are rebuilt on a fresh copy and checked cell by cell before the old one is removed; anything that doesn't match is kept as "NAME (old)" for you to review.
6. On the **Clinic Schedule** tab, enter chair hours and holiday closures (column G) if they are not there yet. **Chairs** shows them, and the Schedule and closure warnings use them.
7. Import the appointment list (below). The chairs come from it: each NextGen resource for your device (for example San Rafael MagV1, MagV2) is a chair.

## Finding your way around

The menu runs down the left in full screen and across the top in the panel.

| Menu | Use it for |
|---|---|
| Today | The chair board, who is still to come, and the visit card for whoever you select. **View** switches between **Clinic** (every chair) and one chair's day; your pick is remembered on your computer. |
| Schedule | **Day by chair**: one weekday's visits in columns by chair, plus the provider's MT, MT redo and follow-up visits, with a line at the current time. **Earlier** and **Later** move between weekdays. **Week**: Monday to Friday with "X of Y" treated or booked, the reasons for each day, and each day's visits (select one to open its row). On a day other than today, the appointment list check for that day is under the grid. |
| Patients | A course is marked Complete on its own once every approved treatment is done (36 of 36, 15 of 15). Every tracker tab, sorted by what needs doing. Filters: Needs attention (the default), All active, Follow-up not scheduled, No upcoming visit, Tapering, Paused, Complete, Discontinued. **Add patient** is here. The number beside Patients in the menu counts patients that need attention; it turns red when one needs action now. |
| Chairs | Each chair's device, coil, last service and today's count, with **Mark unavailable**; the chair hours from Clinic Schedule; the protocol list. |
| Import | The NextGen appointment list and Sohar benefits. |
| Settings | Clinic settings, Appearance (Light, Dark or Match Windows, for this computer) and the workbook tools. **Dark mode** in the menu switches it too. |

**Refresh** re-reads every tab. **Full screen** floats the tracker over the Excel page; **Exit full screen** closes it. Keep the panel open while you use it: full screen saves through the panel.

Banners at the top flag an out-of-date appointment list, new NextGen patients without a tab, and tabs needing an update. Each can be dismissed or snoozed for 3 hours.

## Today

**The chair board** has one row per chair. The lamp is lit while a session is running.

| Row | Means |
|---|---|
| Plain, **13 min left** | A session is running. The bar shows how far along it is. |
| Amber, **Ends in 2 min** or **Over by 4 min** | The session is about to end, or its time is up. |
| Red, **Waiting 12 min** | Only for a visit checked in on an earlier version of the add-in: 10 or more minutes past their time with no session started. There is no separate check-in now, so a late patient shows as **Not started** under Still to come instead. |
| **Open** | Nothing running. It names who is booked next. |

The line under **Today** counts what matters: in session, waiting to start (in red when anyone is), still to come.

**Still to come** lists everyone not yet started, with their time, chair and the most urgent thing on their course. **Start session** is beside each one. Below it, **Appointment list check** compares today with the last imported list (a NextGen visit missing from the tracker, or a visit on a tab NextGen doesn't have), and **Treatments today** gives "X of Y in treatment are booked" with the reasons for the rest.

**The visit card** follows a visit from start to the note. Under the name it shows the protocol time (for example "protocol 19 min"), and until the session starts, when it's projected to end. There are two steps:

1. **Start session in MagV 2** stamps the check-in and start times together and opens the treatment note draft: mood on arrival, the screening questions and technician comments. The answers become sentences and the questions stay out. **Save draft** keeps it while the treatment runs. On an MT or MT redo, it opens **Record motor threshold** instead: MT %, hotspot, amps, and at the first MT the baseline sleep and caffeine and the protocol chosen (MT row's Notes and J2:O4). If the chair is in use and another is open, the button becomes **Move to another chair**.
2. **Complete session** asks for the pulses (filled in from the protocol), the treating technician, pain or discomfort and muscle movements, and PHQ-9 and GAD-7 when the visit has the x for them. If they weren't collected, tick **Not collected today** and pick the visit to move them to (it suggests the next MT, MT redo or follow-up within 7 days). When a copay is due today, choose **Collected** or **Not collected**; it's saved to the visit's Notes, e.g. "Copay $25.00 collected". Pulses entered means the treatment happened. The note draft then opens again for **Copy note**, which puts it on the clipboard for the chart.

Other actions on the card: **Record missed** (no-show or same-day cancel, then **Use freebie**, **Waived** or **Decide later**; it warns when the patient is down to 1 or 0 freebies and when this would be 3 missed in a row), **Undo missed visit**, **Undo complete session** (clears the pulses, technician and copay line; only until the chair's next session starts), **Print schedule** (**Print course schedule** on an MT), and **Open tab**, which opens the patient's tab at that visit's row. Selecting a visit opens its tab too, unless **Settings > Selecting a visit on Today** is set to the card only. The card also shows the payment due today, or **Not set** when that visit type has no cost entered, and **BIDF info needed** with **Enter benefits** when benefits are missing.

Follow-ups show on Today too: **Enter PHQ-9 and GAD-7** when they're due, then **Mark follow-up done**.

## Moving a patient to another chair

**Move to another chair** is on any visit that hasn't started. The list shows what each chair is doing now and, in amber, any booking the move would overlap. Overlaps are a warning, not a block.

Moving in the tracker doesn't move the appointment in NextGen, so every move asks you to do that too:

- Right after the move, tick **Moved in NextGen** once you've moved it there. The same checkbox stays on the visit card ("Moved from MagV 1 to MagV 3. NextGen still has MagV 1; move it there too").
- Until it's ticked, Patients lists "Moved to MagV 3, not yet in NextGen", and the import keeps the tracker's chair.
- It also clears on its own when an appointment list run after the move shows the new chair, or if you move the patient back to the chair NextGen has.

**Mark unavailable** (Chairs) takes a chair off the board and the schedule. Its bookings for the rest of the day show **No chair yet** with **Choose a chair**, which is a move like any other. **Mark available** puts them back.

## Patients

Each patient row shows its two most urgent needs. Red is for what blocks care or billing: an expired or overrun authorization, no visit in the recent-visit window, 3 missed in a row. Amber is everything else that needs doing. Muted labels such as Tapering don't count.

Select a patient to open them over the list. Esc or × closes it. The header shows the tab name, the MRN, the provider and today's session and chair. The big button resolves the top need; **Also needs** lists the rest, each with its own button. **Show** picks the sections (Protocol, Schedule, Copay and benefits); your choice is saved on this computer.

| Need | Action |
|---|---|
| Auth ending, expired, or more treatments booked than authorized | **Add extension**: adds daily visits to D2 (15 by default) and can move the auth end. |
| BIDF info needed today, projected cost over the BIDF estimate, out-of-pocket max likely met | **Enter benefits** / **Review benefits**: deductible and out-of-pocket max (an amount or N/A), BIDF estimate, and cost by code. |
| Freebie decision, or Exempt in NextGen | **Decide freebie**: Use freebie or Waived. Asked for misses in the last 10 days only. |
| No pulses, or no technician, on a past visit | **Enter pulses** (or **It was missed**) / **Add technician**. |
| Check NextGen for PHQ-9 and GAD-7 | One line for any earlier visit whose measures weren't entered, on active courses only. If NextGen has them, **Enter scores**; if not, the next replan moves them to the next visit. |
| Protocol update due, Choose a protocol | **Update protocol** / **Choose a protocol**: opens the protocol Edit dialog. |
| Re-verify benefits for next year | **Update BIDF**. Amber from Dec 1 when a visit is booked in the new year; red once one happens. |
| Booked while away | **Edit availability**: weekly hours and dates away. |
| Status suggested, cancelled in NextGen, no visit in 10+ days | **Set to [status]** / **Update status** / **Pause or discontinue**. |
| Missed 3 in a row | **Call to reschedule**: logs the outcome in Notes on the last missed visit. |
| Moved, not yet in NextGen; no chair | **Mark moved in NextGen** / **Choose a chair**. |
| MT today, course schedule not given | **Print course schedule**. |
| Booked on a closure, more than one visit in a day | **Open tab**. |

**Today's copay.** When the patient has a visit today, the line under the header shows the payment due and whether it's logged. Put your name in **By**, then choose **Collected**, **Not collected** or **$0, OOPM met**. It goes in the visit's Notes with your name and the time, and Complete session starts from the same answer.

**Protocol.** Each site as "F3 (left) · 10 Hz · 120% MT · 3,000 pulses", the protocol name from the Protocols table ("(Custom)" and what changed when a value differs), treatment, MT, device, course, authorization and update-by. **Edit** picks the protocol per row (the values fill in; change one and it reads Custom), the site, and the **next expiration date**, which is required and goes in G2. Saving also notes "Protocol renewed, next by m/d" on today's visit. A patient with no protocol shows the dropdown instead.

**Schedule.** The course strip with a thin payment row under it (amber = copay not logged, pale green = logged), PHQ-9 and GAD-7, the next eight visits (with **Move** for today's), availability, **Open in Schedule**, **Book the rest of the course** (book in NextGen, then import the list), **Pause course**, **Print schedule**, and **Past visits** (every logged visit with pulses, PR, the copay and who logged it, and the note).

**Copay and benefits.** The Benefits row with **Update BIDF** and **Re-run eligibility**; the copay tally for the plan year against the OOPM and deductible remaining, with **Mark {year} OOPM met**; and **Copays to confirm**: tick visits and **Mark selected paid**, or **Mark all paid**, or **Waive**. The **Copay unconfirmed** filter lists everyone with one to confirm.

Follow-ups (99214) keep their own PR, but don't count toward the course cost or the OOPM tally. When the OOPM is met (0 remaining) or the BIDF quote is $0, the copay and benefits reminders stop for that patient; only the new-year re-verify still shows.

## Import the appointment list

Run this whenever the schedule changes, at least daily. The banner tells you when the list on file is out of date.

1. In NextGen PM, open **Reports** (top left) > **Memorized**.
2. Set Report Type to **Appointment Reports**, select **1 Appt List - ALL TMS**, and select **OK**.
3. Export the report with the x icon beside the bar graph icon.
4. Copy the file from your NextGen File Explorer to your computer.
5. In the add-in, choose **Import**, leave **ALL TMS for this clinic** selected, and choose the file.
6. Review the preview (dates set, types changed, visits cleared, misses recorded per tab), then select **Apply**. Nothing changes until you apply.

What the import does:

- Today and later: dates, times and visit types follow NextGen. Telehealth follow-ups get "Tele" in Notes. The provider on the tab (E2) follows the report, last name only.
- Every Expected visit gets a row. When the visit log runs out (each no-show or MT redo uses a row), rows are added inside the log with the same formatting as rows of that visit type, so the last treatments and the final follow-up still show. The preview lists the tabs that grew.
- Past visits never change, except no-shows and same-day cancels, which are always recorded so the freebie decision isn't missed. A kept visit with no row on the tab gets one with its date; enter the pulses.
- A missed visit keeps its row: leave the visit type in L when you put x in M. The import won't add the day again while NextGen still shows it booked, and it puts the type back if L was emptied.
- The visit log is put in date order, top to bottom, and Tx # (C) is renumbered: MT, Daily and MTR rows count, follow-ups and missed visits don't. Tx Remaining (D) follows.
- Rows with an appointment date take row 10's formatting; empty rows go back to the template's gray. Rows you colored yourself are left alone. **Settings > Fix visit-log formatting** runs the same check on every tab.
- A visit listed twice on the report gets one row. If NextGen corrects a no-show to kept or booked, the next import undoes the miss and gives the freebie back.
- Discontinued tabs are not updated. The preview names any that NextGen still has booked.
- Patients booked on the other device are left out (a BrainsWay patient doesn't appear on a MagV tracker).
- Measures are re-planned only after a real disruption: an MTR or F/U moved, or 2 or more misses in a week.
- PHQ-9 and GAD-7 are due on the MT, MTR, every F/U and the last treatment, and on the 5th treatment after the last measures in between. One that would land within 7 days of a required one is skipped. If they were taken a day late, enter the scores on the day they were taken: the next replan clears the earlier day's x's and counts from there. If they weren't taken at all, the next visit gets them. **Settings > Replan measures and pulse totals** runs it now, and also fixes running pulse totals that restart at 0.
- The list is kept on the **Appt List** tab, replaced each time you apply. Closing the panel doesn't lose it: the chairs, the new-patient banner and a new tab's dates keep working, and **Preview again** reruns it.

After applying, delete the exported file from your computer. The workbook keeps what it needs.

For an external patient seen at several locations, run their patient-specific report and choose **Patient-specific** before selecting the file. That report is not kept on the Appt List tab.

## Import benefits from Sohar

After a Sohar verification run, the Sohar Eligibility add-in can fill in each tracker's deductible and out-of-pocket max for you.

1. In the Sohar Eligibility workbook, select **Download benefits CSV** (step 5), or **Benefits (TMS Tracker)** in full screen. In Excel on the web the rows land on the **Export** sheet; use File > Export > Download as CSV.
2. In this add-in, choose **Import** and choose the file under **Sohar benefits**.
3. Review the preview. Each tab shows the payer, plan, copay and coinsurance, and what will change:
   - Blank **Deductible remaining**, **Out-of-pocket max remaining** (both on the Benefits tab) and **Insurance** (E6) are ticked to fill.
   - If the tracker already has a different amount, it is shown unticked with both numbers. These are the amounts at the start of the course, so only tick it if the tracker's number is wrong.
   - Applying sets **Last updated** to today on the Benefits tab for every verified patient, even when nothing changed.
   - Inactive or errored verifications are not imported; the preview says to check coverage.
4. Select **Apply**. Copay and coinsurance are for reference; set the cost table yourself.

Delete the downloaded file afterward.

## Add a patient

When NextGen has a patient without a tab, the banner (and Today's appointment list check) lists them; select **Create tracker** to start one with the MRN and LASFIR filled in. Otherwise, go to **Patients > Add patient**.

Fill in what you have: name (the LASFIR tab name builds from it; names themselves aren't saved), course number, MRN, DOB, home clinic, MD required on-site, insurance, auth assessment, auth start and expiry, number of treatments and MT redos, MTR approved, deductible and out-of-pocket max remaining, BIDF estimated course cost, cost by code, availability, and dates away. In cost by code, **PR** is the patient's cost per visit; fill **Before ded.** only when a deductible applies. Select **Create tab**. The protocol is chosen at the MT.

The add-in copies New Template, fills the tab, places measure placeholders, and, if the patient is on the saved appointment list, opens the import preview with their dates; otherwise it opens the new course.

## Print a patient schedule

From **Print course schedule** on an MT visit, **Print schedule** on any visit card, or a patient's card on Patients, pick **Next 14 days**, **Next 30 days**, **Rest of course**, or your own dates; the count of visits updates as you choose. **Open schedule** shows it over the add-in with **Print** and **Save as PDF**. Either one notes "Schedule provided" on the tab.

The schedule includes the MRN. Hand it to the patient or save it only where patient information is allowed.

## The Sessions sheet

The add-in keeps check-in, session start and end times, chair moves (with the Moved in NextGen tick) and the chair list (device, coil, last service, unavailable) on a **Sessions** sheet it creates the first time it needs it. One row per visit; rows older than 45 days are dropped when the sheet fills up. Nothing about sessions is written on the patient tabs. Leave the sheet as it is: the add-in reads it back, so everyone sees the same board.

## The Active Patients sheet

The add-in keeps the **Active Patients** sheet up to date. It shows this week (Monday to Friday) for every tab whose status is Pending Start, In Progress, Tapering, Paused or blank: the day codes colored like the template, next F/U, graduation, auth expiry (red when past, amber when close), treatments left, MD on-site and when a schedule was last given. "X of Y treated" and each day's reasons sit above the grid, and each patient name links to their tab.

The sheet is view only and is rebuilt whenever something changes. Make changes on the patient tab or in the add-in. **Settings > Rebuild Active Patients sheet** forces a rebuild. Complete and Discontinued patients drop off on their own.

## Protecting patient information

- Patient information stays in the workbook. The add-in's own files hold none.
- The add-in's screens show patients by LASFIR (last three + first three letters, so Jane Smith is SMIJAN). MRNs only appear where a step needs them: new-patient banners show the last four digits, and the import previews and printed schedules show the MRN.
- Name tabs by LASFIR, as the template does.
- Delete NextGen exports from your computer after importing.
- Keep printed or saved schedules only where patient information is allowed. Don't email them outside MHS systems.

## Troubleshooting

| What you see | What to do |
|---|---|
| No TMS button on the Home tab | Upload the manifest again (Install, steps 2 to 4). Make sure you're in Excel on the web, not the desktop app. |
| "This doesn't look like the 1 Appt List - ALL TMS report" | Run the memorized report named exactly **1 Appt List - ALL TMS** and upload that export unchanged. |
| Only one chair on Today | The chairs come from the appointment list. Import the list; until then the board shows the device's one chair. |
| A patient shows on the wrong chair | Move them (Move to another chair), then move the appointment in NextGen and tick **Moved in NextGen**. |
| A banner says tabs were left untouched | Its layout doesn't match the template. Use **Match New Template**, or check that A1 says PATIENT: and A3 says MRN:. |
| Match New Template says New Template needs fixing | The message (on the banner, the button and Settings) says which cells are off and what they read. Fix them on New Template, for example Time in T8 and Freebie Count in J6, then select Refresh. |
| The schedule doesn't show | Choose Print schedule again; it opens over the add-in, not in a separate window. |
| Numbers look out of date | Select **Refresh** in the menu. The board moves on its own every 30 seconds and picks up other people's check-ins and moves then. |
| "The visit on row N has moved on the tab" | Someone imported a list that re-sorted the tab. Nothing was saved. Select **Refresh** and do it again. |
| Active Patients didn't change | It skips the rewrite when nothing changed. Use **Settings > Rebuild Active Patients sheet**. |
| The Appt List tab couldn't be read | Import the list again. |
| A tab was deleted during Match New Template and you need it back | Restore it from the workbook's version history (File > Info > Previous versions). |

## Help

Contact Brendan Yates with questions or problems. Include the workbook name and what you selected, but not patient details.
