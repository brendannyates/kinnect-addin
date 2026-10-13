# Kinnect (Excel on the web add-in)

Full brief (uses, features, automations, safety): [BRIEF.md](BRIEF.md).

Front desk and clinician notifications inside one shared workbook, plus one-click **Join Zoom** and **Open Nudge AI** for clinicians. Skeleton, version 0.1.

## How it works

- **Messages** are rows in the `KinnectLog` table on the **Kinnect Log** sheet. Anyone with the pane open sees new ones within about 10 seconds, with a chime, a red banner and a count in the tab title, until they click **Got it**.
- **Clinician links** (Zoom room, Nudge AI, default room) are rows in `KinnectClinicians` on the **Clinicians** sheet. Clinicians can edit their own links from the pane. Only `https://` links open.
- **Today's appointments** come from the NextGen report pasted on the **Appt List** sheet. The pane finds the header row by name (`APPT_HDR` in `taskpane.js`) and lists today's non-cancelled visits with Arrived, Roomed, Running late and No-show buttons.
- **Who am I** (front desk or a clinician's name) is saved per computer in the browser, not in the workbook.

Patients appear by **LASFIR code only** (last three + first three letters). Names from the report are converted in the pane and never written to the log. The hosted files contain no PHI; `.gitignore` blocks spreadsheets and CSVs.

## Limits of the skeleton

- Alerts only show while the pane is open in Excel on the web. There are no OS pop-ups.
- One front-desk or clinician "Got it" marks a message seen for everyone it was sent to (including **All clinicians**).
- The log keeps growing; the pane only shows today's rows. An archive step is still to do.
- `APPT_HDR` is a guess until the NextGen report exists.

## Try it

- **Demo, no Excel:** open `taskpane.html?demo=1` (fictional data).
- **In Excel on the web:** Insert > Add-ins > Upload My Add-in > `manifest.xml`, then Home > Kinnect > **Set up workbook**.
