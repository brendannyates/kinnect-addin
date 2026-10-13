# Kinnect: project brief

Version 0.1 (skeleton), October 2026. Owner: Brendan Yates.

## Why it exists

A patient who arrives at the front desk shouldn't have to wait because their clinician didn't hear about it. A clinician running a telehealth visit shouldn't be hunting through bookmarks for a Zoom room or the Nudge AI scribe. Kinnect is a single pane where the front desk and clinicians tell each other what's happening with each visit, and where clinicians launch the two tools they open every visit.

It runs as a task pane inside Excel on the web, in a workbook the clinic shares. That keeps every message inside Microsoft 365, which MHS already covers for patient information, and avoids a new vendor, login or install.

## Who uses it and how

**Front desk**
- Opens the day's appointment list, generated from the NextGen report pasted on the **Appt List** sheet.
- Clicks **Arrived**, **Roomed**, **Running late** or **No-show** next to a visit to notify that visit's clinician. Roomed fills in the clinician's usual room.
- Sends anything else (Cancelled, Checked out, a short message) to one clinician or **All clinicians**.
- Gets clinician requests in an inbox and clicks **Got it** to confirm.

**Clinicians**
- Clicks **Join Zoom** or **Open Nudge AI** to open their own saved link in a new tab.
- Tells the front desk **Ready for next patient**, **Send patient back** or **Ready for checkout** in one click, or sends Need help, Running behind or a short message.
- Sees front-desk notices in an inbox and clicks **Got it**.

Everyone sees a **Sent today** list showing whether each message has been seen, and by whom.

## Features

| Feature | Detail |
|---|---|
| Role picker | First run asks "Front desk or which clinician?" and remembers the answer on that computer. Click your name at the top to switch. |
| Self-service clinicians | A clinician who isn't on the list types their name and is added to the **Clinicians** sheet. |
| Personal launch links | Each clinician stores a Zoom room link and a Nudge AI link once. The buttons stay disabled until a link is saved. |
| Today's appointments | Read from the pasted NextGen report. Headers are found by name, so column order doesn't matter. Cancelled visits and other days are left out, and the rest are sorted by time. |
| Last status per visit | Each appointment shows the most recent status the front desk sent for it. |
| Acknowledgement | Every message records who clicked **Got it** and when. |
| One-click setup | **Set up workbook** creates the **Kinnect Log** and **Clinicians** sheets and their tables. |
| Demo mode | `taskpane.html?demo=1` runs on fictional data outside Excel, for training and testing. |

## Automations

These run without anyone clicking:

- **Live updates.** The pane listens for changes to the message table, including edits from other people in the shared workbook, and re-reads within about half a second. A refresh every 10 seconds catches anything the change events miss. Polling pauses while the pane is hidden.
- **Alerts.** A new, unacknowledged message to you plays a short chime, shows a red "N new messages" banner and puts the count in the browser tab title, e.g. "(2) Kinnect". These clear when you click **Got it**.
- **LASFIR conversion.** Names in the NextGen report, in "Last, First" or "First Last" form, are turned into LASFIR codes on screen. Names are never written to the log.
- **Appointment filtering.** Only today's visits that aren't cancelled are listed, sorted by time.
- **Room prefill.** **Roomed** fills in the clinician's default room from the Clinicians sheet.
- **Typing is kept.** The background refresh doesn't clear a half-typed message or close an open section.

**Not built yet** (candidates, not promises):
- a daily archive of older log rows
- a Power Automate flow that refreshes Appt List from NextGen automatically
- Teams or Outlook alerts when the pane is closed
- per-clinician acknowledgement for **All clinicians** messages

## Safety measures

- **PHI stays in Microsoft 365.** Messages live only in the shared workbook. The add-in code is hosted on GitHub Pages and contains no patient data; it only reads and writes the open workbook.
- **LASFIR only.** The patient field accepts letters only (2 to 6, for short names) and rejects anything else. The pane reminds users not to put names or other details in the note, and notes are capped at 140 characters.
- **No secrets.** There are no API keys, tokens or outside services.
- **Safe links.** Only `https://` links can be saved or opened, which blocks `javascript:` and other unsafe schemes. Links open in a new tab with `noopener`.
- **Escaped output.** Everything typed into the workbook is HTML-escaped before display, so a cell can't inject script into the pane.
- **Spreadsheets can't be committed.** `.gitignore` blocks `.xlsx`, `.xls`, `.xlsm` and `.csv`, so a real workbook or NextGen export can't be committed by accident.
- **Audit trail.** Every message keeps who sent it, to whom, when, who acknowledged it and when.
- **Access follows the workbook.** Anyone who can open the workbook can read the log, so share it only with clinic staff, invite-only.

## Known limits

- Alerts only appear while the pane is open in Excel on the web. Kinnect does not send operating-system pop-ups.
- "Who am I" is set per computer and is not an identity check. Anyone at a computer can choose any role. The workbook's own sharing is the real access control.
- One **Got it** marks a message seen for everyone it was sent to.
- The log grows without limit; the pane shows only today's rows.
- The appointment header names are guesses until the NextGen report is built.
- **Not yet run in Excel on the web.** Tested only in a headless browser with fictional data.

## Workbook layout

| Sheet | Table | Columns |
|---|---|---|
| Kinnect Log | `KinnectLog` | ID, Sent, From, To, Patient, Status, Room, Note, Acked by, Acked at |
| Clinicians | `KinnectClinicians` | Clinician, Zoom URL, Nudge URL, Room |
| Appt List | (none) | The NextGen report, pasted as-is |

## Rollout

1. Build the NextGen report with columns for appointment date, appointment time, patient, provider/resource and status. Match `APPT_HDR` in `taskpane.js` to its headers.
2. Upload `manifest.xml` in Excel on the web: Insert > Add-ins > Upload My Add-in.
3. Open a copy of the clinic workbook, then click Home > Kinnect > **Set up workbook**.
4. Pilot at one clinic with fictional patients first, then real visits.
5. Changes go live when they merge to `main`. Re-upload the manifest only when `manifest.xml` changes.
