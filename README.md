# Kinnect (Excel on the web add-in)

Full brief (uses, features, automations, safety): [BRIEF.md](BRIEF.md).

Front desk and resource notifications inside one shared workbook, plus one-click **Join Zoom** and **Open Nudge AI** for clinicians. A resource is a clinician or a device (MagVenture, BrainsWay). Carries over the Floot kin-nect app. Version 0.2.

## How it works

- **Messages** are rows in the `KinnectLog` table on the **Kinnect Log** sheet. Anyone with the pane open sees new ones within seconds, with their chosen sound, a red banner, a count in the tab title and an optional desktop pop-up, until they click **Got it**.
- **Resources** (name, Zoom room, Nudge AI, default room) are rows in `KinnectResources` on the **Resources** sheet. Clinicians can edit their own links from the pane. Only `https://` links open.
- **Today's appointments** come from the NextGen report pasted on the **Appt List** sheet. The pane finds the header row by name (`APPT_HDR` in `taskpane.js`).
- **Sounds** are `assets/sounds/bing|boom|bop|bow` (`.mp3`, `.wav`, `.ogg` or `.m4a`). A missing file falls back to a built-in chime.
- **Per-computer settings** (who I am, notify on/off, sound, volume, pop-ups, patient code on one-click alerts) are saved in the browser, not the workbook.

One-click alerts carry no patient identifier. A LASFIR code (last three + first three letters) is optional. Names from the report are converted in the pane and never written to the log. The hosted files contain no PHI; `.gitignore` blocks spreadsheets and CSVs.

## Limits

- Alerts only show while the pane is open in Excel on the web. Sound needs one click in the pane first.
- One "Got it" marks a message seen for everyone it was sent to (including **All resources**).
- The log keeps growing; the pane only shows today's rows.
- `APPT_HDR` is a guess until the NextGen report exists.

## Try it

- **Demo, no Excel:** open `taskpane.html?demo=1` (fictional data).
- **In Excel on the web:** Insert > Add-ins > Upload My Add-in > `manifest.xml`, then Home > Kinnect > **Set up workbook**.
