# TMS Tracker add-in: install and user guide

The TMS Tracker is an Excel add-in for the MHS TMS patient tracker workbook. It opens as a panel beside the workbook and shows today's visits, what needs attention, and each patient's course, and it writes back into the same cells the template already uses. Nothing about the workbook changes for people who don't use the add-in.

This guide is for anyone at MHS who uses a TMS tracker workbook: technicians, care navigators, and clinic leads.

## Before you start

- **Excel on the web.** The add-in runs in Excel in a browser (Edge or Chrome). It does not run in the desktop Excel app at MHS.
- **A tracker workbook in OneDrive or SharePoint,** built from the current MagV template. It needs the **New Template** tab; **Protocols** and **Clinic Schedule** are used when present.
- **Pop-ups allowed for Excel** in your browser, for schedule printouts and the separate tracker window.
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

1. Open the panel and select the gear (top right) for **Settings**.
2. Check **This clinic (NextGen location)**, for example San Rafael. Visits booked at other locations are marked as another clinic.
3. Check **Tracker device**. "From the workbook name" works when the name contains MagV or Brainsway; otherwise pick MagVenture or Brainsway.
4. Leave the warning windows at their defaults unless your clinic wants different ones: auth warning 14 calendar days, protocol warning 2 business days, recent visit window 10 days.
5. If a banner says tabs need updating, select **Update tabs**. It adds the template's newest columns to every patient tab without touching entries. Tabs in an older layout are rebuilt on a fresh copy and checked cell by cell before the old one is removed; anything that doesn't match is kept as "NAME (old)" for you to review.
6. On the **Clinic Schedule** tab, enter chair hours and holiday closures (column G) if they are not there yet. The add-in uses them for the chair view and closure warnings.

## Finding your way around

Across the top: **Import list**, **Open in a new window** (⧉), **Full screen** (⛶), **Settings** (gear) and **Refresh**.

| Menu | Use it for |
|---|---|
| Schedule | **Day** shows one weekday's visits; **Prev** and **Next** move between weekdays, **Back to today** returns. **Week** shows Monday to Friday with "X of Y treated" and the reasons for each day. Discontinued patients are left off both, and off Attention and the Patients list (use its **Discontinued** filter to find them). When the appointment list has more than one chair (for example Castro MagV1 and Castro MagV2), buttons above the schedule show one chair or both; your pick is remembered on your computer. |
| Patients | The roster for the week, with filters (All, Active, Follow-up not scheduled, No upcoming visit) and **Add patient**. Select a patient for their full view. |
| Clinic | The chair schedule and the protocol list. |

**Attention** lists what needs action: auth expiring or exceeded, protocol renewals due, missing BIDF information for patients seen today, freebie decisions, missing pulses or technician, measures not collected, visits booked while a patient is away or on a closure date. It sits beside the schedule on a wide panel and collapses to the top with a count on a narrow one. Only groups with something in them show.

Banners at the top flag an out-of-date appointment list, new NextGen patients without a tab, and tabs needing an update. Each can be dismissed or snoozed for 3 hours.

**A bigger view:** drag the panel's edge wider, or use ⧉ or ⛶ to open the tracker in its own window. Keep the panel open while the window is in use; the window saves through it.

## Daily use

Each visit card is colored like the tab's row highlight: MT light blue, MTR green, F/U pink, another clinic orange. It shows the payment due (**Payment $X**), or **Payment not set** when that visit type has no cost entered.

- **Complete treatment** opens with the pulses the protocol calls for already filled in. Adjust them if needed, check the treating technician, and select **Save treatment**. Pulses entered means the treatment happened; the card then shows **Delivered**.
- **Missed visits:** open the card's ⋯ menu and choose **Record no-show** or **Record same-day cancel**. A pop-up asks whether to **Use freebie** or mark it **Waived** (or **Decide later**), and warns when the patient is down to 1 or 0 freebies. Remaining freebies count down on the tab. If you decide later, the card shows **Classify freebie**. **Undo missed visit** is in the same ⋯ menu.
- **Measures needed** opens PHQ-9 and GAD-7 entry; select **Save measures**. If they weren't collected, **Move measures** shifts them to an upcoming visit (it suggests an MT, MTR or F/U within 7 days) and notes the move on the original row.
- **Draft treatment note** (daily treatments) builds the note from mood on arrival, observations and technician comments. The yes/no observation answers become sentences in the note; the questions are not copied into it. **Copy note** puts it on the clipboard for the chart.
- **Cancelled in NextGen: update status?** (Attention) lists patients whose cancelled appointments are still on the tracker, or who have nothing left booked, with Paused, Discontinued and Complete buttons. When every appointment is cancelled and the last visit was more than 10 days ago (the recent-visit window in Settings), Discontinued is suggested and applied with the next import.
- A number in **# of MT redos** (D4) sets **MTR approved** (W6) to Yes on its own.
- **MT** cards also take the MT details (hotspot, MT %, amps, baseline sleep and caffeine); select **Save MT details**.
- The ⋯ menu also has **Print schedule** and **Open patient**.

## Import the appointment list

Run this whenever the schedule changes, at least daily. The banner tells you when the list on file is out of date.

1. In NextGen PM, open **Reports** (top left) > **Memorized**.
2. Set Report Type to **Appointment Reports**, select **1 Appt List - ALL TMS**, and select **OK**.
3. Export the report with the x icon beside the bar graph icon.
4. Copy the file from your NextGen File Explorer to your computer.
5. In the panel, select **Import list**, leave **ALL TMS for this clinic** selected, and choose the file.
6. Review the preview (dates set, types changed, visits cleared, misses recorded per tab), then select **Apply**. Nothing changes until you apply.

What the import does:

- Today and later: dates, times and visit types follow NextGen. Telehealth follow-ups get "Tele" in Notes. The provider on the tab follows the report.
- Every Expected visit gets a row. When the visit log runs out (each no-show or MT redo uses a row), rows are added inside the log with the same formatting as rows of that visit type, so the last treatments and the final follow-up still show. The preview lists the tabs that grew.
- Past visits never change, except no-shows and same-day cancels, which are always recorded so the freebie decision isn't missed. A kept visit with no row on the tab gets one with its date; enter the pulses.
- The visit log is put in date order, top to bottom, and Tx # (C) is renumbered: MT, Daily and MTR rows count, follow-ups and missed visits don't. Tx Remaining (D) follows.
- Patients booked on the other device are left out (a Brainsway patient doesn't appear on a MagV tracker).
- Measures are re-planned only after a real disruption: an MTR or F/U moved, or 2 or more misses in a week.
- The list is kept on the **Appt List** tab, replaced each time you apply. Closing the panel no longer loses it: the new-patient banner and a new tab's dates keep working, and **Preview again** reruns it.

After applying, delete the exported file from your computer. The workbook keeps what it needs.

For an external patient seen at several locations, run their patient-specific report and choose **Patient-specific** before selecting the file. That report is not kept on the Appt List tab.

## Add a patient

When NextGen has a patient without a tab, the banner lists them; select one to start a tracker with the MRN and LASFIR filled in. Otherwise, go to **Patients > Add patient**.

Fill in what you have: name (the LASFIR tab name builds from it; names themselves aren't saved), course number, MRN, DOB, home clinic, MD required on-site, insurance, auth assessment, auth start and expiry, number of treatments and MT redos, MTR approved, deductible and out-of-pocket max remaining, BIDF estimated course cost, cost by code, availability, and dates away. In cost by code, **PR** is the patient's cost per visit; fill **Before ded.** only when a deductible applies. Select **Create tab**. The protocol is chosen at the MT.

The add-in copies New Template, fills the tab, places measure placeholders, and, if the patient is on the saved appointment list, opens the import preview with their dates.

## The patient view

Select a patient on the Patients menu to see:

- **Summary:** MRN, insurance, provider, care navigator, authorization, treatments, graduation, next follow-up, freebies left, home clinic.
- **Protocol** editor.
- **Benefits and cost:** cost so far and projected, deductible and out-of-pocket max, BIDF estimate and cost by code, all editable. **Add extension** adds daily visits (15 by default) and can move the auth expiry.
- **Measures over time.**
- **Appointments:** upcoming, or all visits on request.
- **Availability:** weekly hours and date ranges away.
- **Print or save schedule** and **Open tab**.

## Print a patient schedule

From a visit card's ⋯ menu or the patient view, choose **Print schedule**. Pick **Next 14 days**, **Next 30 days**, **Rest of course**, or your own dates; the count of visits updates as you choose. **Open schedule** shows it in its own window with **Print** and **Save as PDF**. Either one notes "Schedule provided" on the tab.

The schedule includes the MRN. Hand it to the patient or save it only where patient information is allowed.

## The Active Patients sheet

The add-in keeps the **Active Patients** sheet up to date. It shows this week (Monday to Friday) for every tab whose status is Pending Start, In Progress, Tapering, Paused or blank: the day codes colored like the add-in, next F/U, graduation, auth expiry (red when past, amber when close), treatments left, MD on-site and when a schedule was last given. "X of Y treated" and each day's reasons sit above the grid, and each patient name links to their tab.

The sheet is view only and is rebuilt whenever something changes. Make changes on the patient tab or in the add-in. **Settings > Rebuild Active Patients sheet** forces a rebuild. Complete and Discontinued patients drop off on their own.

## Protecting patient information

- Patient information stays in the workbook. The add-in's own files hold none.
- Name tabs by LASFIR (last three + first three letters, so Jane Smith is SMIJAN), as the template does.
- Delete NextGen exports from your computer after importing.
- Keep printed or saved schedules only where patient information is allowed. Don't email them outside MHS systems.

## Troubleshooting

| What you see | What to do |
|---|---|
| No TMS button on the Home tab | Upload the manifest again (Install, steps 2 to 4). Make sure you're in Excel on the web, not the desktop app. |
| "This doesn't look like the 1 Appt List - ALL TMS report" | Run the memorized report named exactly **1 Appt List - ALL TMS** and upload that export unchanged. |
| A banner says tabs were left untouched | Its layout doesn't match the template. Use **Update tabs**, or check that A1 says PATIENT: and A3 says MRN:. |
| The schedule window doesn't open | Allow pop-ups for Excel in your browser. If the tracker is already open in its own window, the schedule shows over the page instead. |
| Numbers look out of date | Select **Refresh** (top right). |
| Active Patients didn't change | It skips the rewrite when nothing changed. Use **Settings > Rebuild Active Patients sheet**. |
| The Appt List tab couldn't be read | Import the list again. |
| A tab was deleted during Update tabs and you need it back | Restore it from the workbook's version history (File > Info > Previous versions). |

## Help

Contact Brendan Yates with questions or problems. Include the workbook name and what you selected, but not patient details.
