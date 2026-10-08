# TMS Tracker (Excel add-in) for the San Rafael MagVenture workbook

Static task-pane add-in. No backend, no PHI in the code. It works on your existing layout: one tab per course (`SMIJAN`, `SMIJAN 2`) cloned from **New Template**, plus **Active Patients** and **Protocols**.

## Status

The supplied documentation reports prior import-planner and merge-rule tests, but those tests were not included in this repository. The Excel and Office.js calls have **not** been run in Excel. Test on a copy of the workbook first.

## Run it locally (pilot)

1. `npx office-addin-dev-certs install`
2. In this folder: `npx http-server -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000`
3. Excel on the web: Insert > Add-ins > Upload My Add-in > `manifest.xml`.

For the team, ask Product/Engineering for an internal HTTPS host and replace `https://localhost:3000` in `manifest.xml`. The files contain no PHI.

## What it adds to your workbook (right of Notes, every tab and New Template)

Click **Add tracker columns** once. Nothing existing moves.

| Cells | Purpose |
|---|---|
| T8:T49 Freebie | `Used` or `Waived` for a Miss row. Amber when a Miss row has no decision. |
| U8:U49 Time | Appointment time from NextGen. |
| W1:X4 | Freebies allowed (3, editable), Used, Waived, Remaining (formulas). |
| W8:Z30 | Availability windows: Type (Unavailable/Available), From, To, Note. "Preferred Times" in M6 is also shown. |

## What each part reads

- **MRN** (A4) matches NextGen rows to tabs. With several courses for one MRN, the highest course number gets the import.
- **Active** = status (D7) is Pending Start, In Progress, or Tapering.
- **Delivered** = pulses in F, H, or J. **Missed** = M is `x`. **Other clinic** = R is `x`.
- **Protocol update needed**: G2 (Update protocol by) within the warning window or past, or an upcoming row with Action = Protocol.
- **Auth flags**: no expiry in G4; expired; expiring within the warning window; Tx End (F7) after expiry; visits booked past expiry.
- **Follow-up**: rows with Type = F/U. Not scheduled = F/U rows exist, none delivered, none dated today or later.
- **Past date, no outcome**: dated before today, no pulses, not marked Miss.
- **Protocols tab** shows your Protocols sheet with minutes recalculated by the same formula as each tab's N2, next to your typed Tx Time.

## Import rules

1. Window = start date (default today) to the last date in the file. Earlier dates are never touched.
2. Scheduled NextGen visits fill the tab's undelivered rows in order, writing the date in E and time in U. Delivered rows and existing Miss rows are not moved.
3. Rows in the window with no matching NextGen visit have their date cleared. Rows after the window keep their planned dates unless NextGen has more visits than free window rows, in which case they are overwritten.
4. With an appointment-type column, follow-ups go to F/U rows and treatments to the rest.
5. No-shows and cancellations in NextGen that are not on the tab are listed in the preview. Ticking the option records each as a Miss row (`x` in M, note in S). **A recorded miss uses up one Tx row.** If your team handles misses differently, leave that option off.
6. "Clear existing Miss rows in the window" removes the `x` and the freebie mark. Off by default.
7. Patients missing from the file are left alone unless you tick the option.

## Other buttons

- **Fill Active Patients week grid**: fills blank cells in N3:W39 with MT, Daily Tx, MTR, F/U, Missed, or Different Clinic from each tab's dates. The checkbox sets N2 to this week's Monday. N2 is a typed date today.
- **Copy dashboard to a sheet**: writes a snapshot to a Dashboard sheet.
- **Add patient**: copies New Template, writes the LASFIR to A2, the MRN to A4, status Pending Start, and adds the tab to the first empty row of Active Patients column B. Protocol numbers (J2:M4) are left as the template has them.

## Caveats

- LASFIR collisions are possible. MRN matching avoids that for imports.
- Weeks run Sunday to Saturday in the dashboard. The Active Patients grid is workdays from N2.
- Tab log is rows 10-49. More than 40 slots needs a layout change.

## Repository setup

This project is separate from the patient billing portal. The supplied add-in files are preserved; manifest icons and development scripts are included.

```bash
npm install
npm run check
npm run certs
npm start
```

The syntax check does not validate workbook behavior. Use a copy of the workbook for the Excel pilot. Keep patient workbooks and appointment exports out of version control.
