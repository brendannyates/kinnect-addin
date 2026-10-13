# Kinnect: project brief

Version 0.2, October 2026. Owner: Brendan Yates.

## Why it exists

A patient who arrives at the front desk shouldn't have to wait because their clinician didn't hear about it. A clinician running a telehealth visit shouldn't be hunting through bookmarks for a Zoom room or the Nudge AI scribe. Kinnect is a single pane where the front desk and each **resource** (a clinician or a device such as MagVenture or BrainsWay) tell each other what's happening with each visit, and where clinicians launch the two tools they open every visit.

## Where it came from

Kinnect began as **kin-nect** in Floot: one-click "patient arrived" alerts per resource, a per-device "notify me" setting, four notification sounds (bing, boom, bop, bow) and browser pop-ups. This add-in carries those features into Excel on the web, in a workbook the clinic shares. That keeps every message inside Microsoft 365, which MHS already covers for patient information, and avoids a separate vendor, login or install.

kin-nect's Teams delivery was never wired up in Floot, and it isn't in this version either. **IT Support: Weekly Connect**, the other app in the same Floot account, is a separate product and is deliberately not part of Kinnect.

## Who uses it and how

**Front desk**
- Sees a card for every resource with **Patient arrived** and **Telehealth patient arrived**. One click notifies that resource, with no patient details.
- Works down today's appointment list, generated from the NextGen report pasted on **Appt List**. Each visit has **Arrived**, **Telehealth**, **Roomed**, **Running late** and **No-show** buttons. Roomed fills in the resource's usual room.
- Sends anything else (Cancelled, Checked out, a short message) to one resource or to **All resources**, adding a LASFIR code only when it's needed.
- Gets requests from resources in an inbox and clicks **Got it** to confirm.

**Clinicians and devices**
- Click **Join Zoom** or **Open Nudge AI** to open their own saved link in a new tab.
- Tell the front desk **Ready for next patient**, **Send patient back** or **Ready for checkout** in one click, or send Need help, Running behind or a short message.
- See "Your patient has arrived." or "Your telehealth patient has arrived." in their inbox and click **Got it**.

Everyone sees a **Sent today** list showing whether each message has been seen, and by whom.

## Features

| Feature | Detail |
|---|---|
| Resource cards | One-click Patient arrived and Telehealth patient arrived per clinician or device, carried over from kin-nect. |
| Role picker | First run asks "Front desk, or which clinician or device?" and remembers the answer on that computer. Click the name at the top to switch. |
| Self-service resources | A clinician or device that isn't listed can be added from the pane, and goes onto the **Resources** sheet. |
| Personal launch links | Each resource stores a Zoom room link and a Nudge AI link once. The buttons stay disabled until a link is saved. |
| Today's appointments | Read from the pasted NextGen report. Headers are found by name, so column order doesn't matter. Cancelled visits and other days are left out, and the rest are sorted by time. |
| Notification sounds | Each computer picks bing, boom, bop, bow, a built-in chime or no sound, sets the volume, and can play a test. |
| Desktop pop-ups | Optional, per computer, wherever the browser allows them. They never show the patient code. |
| Per-computer settings | Notify me on/off, sound, volume, pop-ups and (front desk only) whether one-click appointment alerts include the patient code. |
| Acknowledgement | Every message records who clicked **Got it** and when. |
| One-click setup | **Set up workbook** creates the **Kinnect Log** and **Resources** sheets and their tables. |
| Demo mode | `taskpane.html?demo=1` runs on fictional data outside Excel, for training and testing. |

## Automations

These run without anyone clicking:

- **Live updates.** The pane listens for changes to the message table, including edits from other people in the shared workbook, and re-reads within about half a second. A refresh every 10 seconds catches anything the change events miss. Polling pauses while the pane is hidden.
- **Alerts.** A new, unacknowledged message to you plays your chosen sound, shows a red "N new messages" banner, puts the count in the browser tab title, e.g. "(2) Kinnect", and, if turned on, shows a desktop pop-up. These clear when you click **Got it**.
- **Sound fallback.** If a chosen sound file is missing or slow to load, the built-in chime plays instead.
- **Duplicate guard.** The same alert to the same resource within 5 seconds, from a double click or two people at once, is sent once ("Already notified moments ago"). This is kin-nect's debounce.
- **LASFIR conversion.** Names in the NextGen report, in "Last, First" or "First Last" form, are turned into LASFIR codes on screen. Names are never written to the log.
- **Appointment filtering.** Only today's visits that aren't cancelled are listed, sorted by time.
- **Room prefill.** **Roomed** fills in the resource's default room from the Resources sheet.
- **Typing is kept.** The background refresh doesn't clear a half-typed message or close an open section.

**Not built yet** (candidates, not promises):
- Teams delivery through Power Automate, which kin-nect also had planned
- a daily archive of older log rows
- a Power Automate flow that refreshes Appt List from NextGen automatically
- per-resource acknowledgement for **All resources** messages

## Safety measures

- **PHI stays in Microsoft 365.** Messages live only in the shared workbook. The add-in code is hosted on GitHub Pages and contains no patient data; it only reads and writes the open workbook.
- **No identifier by default.** As in kin-nect, one-click alerts say only "Your patient has arrived." Adding a patient code is a deliberate choice: typed into the message form, or switched on per front-desk computer for appointment-list buttons.
- **LASFIR only.** The patient field accepts letters only (2 to 6, for short names) and rejects anything else. The pane reminds users not to put names or other details in the note, and notes are capped at 140 characters.
- **Pop-ups stay generic.** Desktop pop-ups can be seen by anyone walking past, so they never include the patient code.
- **No secrets.** There are no API keys, tokens or outside services.
- **Safe links.** Only `https://` links can be saved or opened, which blocks `javascript:` and other unsafe schemes. Links open in a new tab with `noopener`.
- **Escaped output.** Everything typed into the workbook is HTML-escaped before display, so a cell can't inject script into the pane.
- **Spreadsheets can't be committed.** `.gitignore` blocks `.xlsx`, `.xls`, `.xlsm` and `.csv`, so a real workbook or NextGen export can't be committed by accident.
- **Licensed sounds only.** Sound files go in a public repo, so only add sounds you have the right to publish.
- **Audit trail.** Every message keeps who sent it, to whom, when, who acknowledged it and when.
- **Access follows the workbook.** Anyone who can open the workbook can read the log, so share it only with clinic staff, invite-only.

## Known limits

- Alerts only appear while the pane is open in Excel on the web.
- Browsers play sound only after the pane has been clicked once, so click the pane after opening it.
- Excel may block the desktop pop-up permission prompt inside the pane. When it does, the banner and sound still work.
- "Who am I" is set per computer and is not an identity check. Anyone at a computer can choose any role. The workbook's own sharing is the real access control.
- One **Got it** marks a message seen for everyone it was sent to.
- The log grows without limit; the pane shows only today's rows.
- The appointment header names are guesses until the NextGen report is built.
- **Not yet run in Excel on the web.** Tested only in a headless browser with fictional data.

## Workbook layout

| Sheet | Table | Columns |
|---|---|---|
| Kinnect Log | `KinnectLog` | ID, Sent, From, To, Patient, Status, Room, Note, Acked by, Acked at |
| Resources | `KinnectResources` | Resource, Zoom URL, Nudge URL, Room |
| Appt List | (none) | The NextGen report, pasted as-is |

## Rollout

1. Download the four sounds from kin-nect's Settings page and save them in `assets/sounds/` as `bing`, `boom`, `bop` and `bow`, each as `.mp3`, `.wav`, `.ogg` or `.m4a`.
2. Build the NextGen report with columns for appointment date, appointment time, patient, provider/resource and status. Match `APPT_HDR` in `taskpane.js` to its headers.
3. Upload `manifest.xml` in Excel on the web: Insert > Add-ins > Upload My Add-in.
4. Open a copy of the clinic workbook, then click Home > Kinnect > **Set up workbook**.
5. Pilot at one clinic with fictional patients first, then real visits.
6. Changes go live when they merge to `main`. Re-upload the manifest only when `manifest.xml` changes.
