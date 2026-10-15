# Builds "Kinnect Relay DEMO (fake data).xlsx" in the OneDrive Add-In ManifestsKinnect folder with Excel desktop (COM). Every name, patient, link and SOP is fake.
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File scriptsmake-demo-workbook.ps1
$ErrorActionPreference = "Stop"
$out = "C:\Users\byates\OneDrive - Mindful Health Solutions\Add-In Manifests\Kinnect\Kinnect Relay DEMO (fake data).xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 6) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $names = "Start Here", "Kinnect Log", "Resources", "Kinnect Tools", "Kinnect Guide", "Appt List"
  for ($i = 0; $i -lt 6; $i++) { $wb.Worksheets.Item($i + 1).Name = $names[$i] }

  function Put($ws, $rows, $table) {
    $r = $rows.Count; $c = $rows[0].Count
    for ($i = 0; $i -lt $r; $i++) { for ($j = 0; $j -lt $c; $j++) { $v = $rows[$i][$j]; $cell = $ws.Cells.Item($i + 1, $j + 1)
      # PowerShell's COM binder can lock Value2 to string after the first write, so numbers and formulas both go through Formula
      if ($v -is [double]) { $cell.Formula = $v.ToString([Globalization.CultureInfo]::InvariantCulture) } elseif ($v -is [string] -and $v.StartsWith("=")) { $cell.Formula = $v } else { $cell.Value2 = [string]$v } } }
    $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($r, $c))
    if ($table) { $lo = $ws.ListObjects.Add(1, $rng, [Type]::Missing, 1); $lo.Name = $table }
    [void]$rng.Columns.AutoFit()
  }
  function T($hms) { return "=TEXT(TODAY(),""yyyy-mm-dd"")&"" $hms""" }   # always today, as text, like Relay writes

  # Start Here
  $s = $wb.Worksheets.Item("Start Here")
  $lines = @(
    "Kinnect Relay - DEMO WORKBOOK (all names, patients, links and SOPs here are fake)",
    "",
    "1. Open this file in Excel on the web (from OneDrive), not desktop Excel.",
    "2. Insert > Add-ins > Upload My Add-in, and pick kinnect-manifest.xml from the Add-In Manifests folder (one level up from Kinnect).",
    "3. Home > Kinnect > Relay opens the pane.",
    "4. Pick Front desk, or a clinician or device. Click your name at the top of the pane to switch roles.",
    "",
    "Things to try:",
    "- Front desk: click Patient arrived on a resource card, or Arrived next to a demo appointment. Click twice to see the duplicate guard.",
    "- Switch to Dr. Avery Rivera: two front-desk messages are waiting. Click Got it. Join Zoom and Open Nudge AI open placeholder links.",
    "- Ops Guide tab: 16 fake SOPs with search and category filters. Links go to example.com.",
    "- Settings: sound, volume, test, pop-ups, and the Kinnect tool links (fake example.com links).",
    "- Open this workbook in two browser tabs as two roles to watch messages arrive live.",
    "",
    "Messages and appointments are dated with TODAY(), so the demo looks current any day you open it.",
    "Rows you add on Kinnect Log stay; delete them to reset. Never put real patient information in this file."
  )
  for ($i = 0; $i -lt $lines.Count; $i++) { $s.Cells.Item($i + 1, 1).Value2 = $lines[$i] }
  $s.Cells.Item(1, 1).Font.Bold = $true; $s.Cells.Item(1, 1).Font.Size = 14; $s.Cells.Item(8, 1).Font.Bold = $true
  $s.Columns.Item(1).ColumnWidth = 120

  # Kinnect Log (Sent / Acked at are TODAY() text formulas)
  $log = $wb.Worksheets.Item("Kinnect Log")
  $log.Range("A:J").NumberFormat = "@"
  $log.Range("B2:B9,J2:J4").NumberFormat = "General"   # before writing, so the TODAY() formulas calculate instead of staying text
  $rows = @(
    @("ID", "Sent", "From", "To", "Patient", "Status", "Room", "Note", "Acked by", "Acked at"),
    @("demo01", (T "08:41:10"), "Front desk", "Dr. Avery Rivera", "", "Patient arrived", "", "", "Dr. Avery Rivera", (T "08:42:03")),
    @("demo02", (T "08:44:30"), "Dr. Avery Rivera", "Front desk", "", "Ready for next patient", "", "", "Front desk", (T "08:45:12")),
    @("demo03", (T "08:58:05"), "Front desk", "MagVenture", "", "Patient arrived", "", "", "MagVenture", (T "08:58:40")),
    @("demo04", (T "09:10:22"), "Front desk", "Dr. Jordan Okafor", "", "Telehealth patient arrived", "", "", "", ""),
    @("demo05", (T "09:14:47"), "Front desk", "Dr. Avery Rivera", "DOEJAN", "Roomed", "2", "", "", ""),
    @("demo06", (T "09:20:13"), "BrainsWay", "Front desk", "", "Need help", "", "Chair 2 cap fit", "", ""),
    @("demo07", (T "09:31:55"), "Front desk", "All resources", "", "Message", "", "Lunch delivery at 12:15", "", ""),
    @("demo08", (T "09:36:02"), "Front desk", "Dr. Avery Rivera", "", "Running late", "", "About 10 minutes", "", "")
  )
  Put $log $rows "KinnectLog"

  Put $wb.Worksheets.Item("Resources") @(
    @("Resource", "Zoom URL", "Nudge URL", "Room"),
    @("Dr. Avery Rivera", "https://zoom.us/j/0000000001", "https://example.com/nudge-demo-rivera", "2"),
    @("Dr. Jordan Okafor", "https://zoom.us/j/0000000002", "https://example.com/nudge-demo-okafor", "3"),
    @("Sam Lee, PMHNP", "https://zoom.us/j/0000000003", "", "4"),
    @("MagVenture", "", "", "TMS 1"),
    @("BrainsWay", "", "", "TMS 2")
  ) "KinnectResources"

  Put $wb.Worksheets.Item("Kinnect Tools") @(
    @("Tool", "Workbook link"),
    @("TMS", "https://example.com/demo-tms-tracker"),
    @("ESK", "https://example.com/demo-esk-tracker"),
    @("Billing", "https://example.com/demo-billing-tickets"),
    @("IR Portal", "")
  ) "KinnectTools"

  $u = "https://example.com/demo-sop/"
  Put $wb.Worksheets.Item("Kinnect Guide") @(
    @("Category", "Section", "Title", "Audience", "Updated", "Link", "Purpose", "Notes"),
    @("Front Office", "", "Demo: Patient check-in", "Front Office", "2026-06-17", "${u}check-in", "Steps for checking a patient in and alerting their clinician.", ""),
    @("Front Office", "", "Demo: Patient check-out", "Front Office", "2026-06-25", "${u}check-out", "Closing out a visit and booking the next one.", ""),
    @("Front Office", "", "Demo: No-show and late cancel", "Care Navigator, Front Office", "2026-10-06", "${u}no-show", "How to document a no-show or late cancel and when a fee applies.", ""),
    @("Front Office", "Phones", "Demo: Answering the clinic line", "Front Office", "2026-05-02", "${u}phones", "Greeting, routing and message taking.", ""),
    @("Care Navigation", "", "Demo: Incoming referrals", "Care Navigator", "2026-04-18", "${u}referrals", "Logging a new referral and first contact within one business day.", ""),
    @("Care Navigation", "", "Demo: Insurance eligibility check", "Care Navigator", "2026-07-09", "${u}eligibility", "Verifying coverage before a first visit.", ""),
    @("Care Navigation", "Waitlist", "Demo: Waitlist and offers", "Care Navigator, Front Office", "2026-05-27", "${u}waitlist", "Adding patients to the waitlist and offering open slots.", ""),
    @("Interventional", "TMS", "Demo: TMS intake", "Care Navigator", "2026-08-28", "${u}tms-intake", "Screening and scheduling a TMS consult.", ""),
    @("Interventional", "TMS", "Demo: TMS daily treatment checklist", "Technician", "2026-03-11", "${u}tms-checklist", "Chair setup, motor threshold record and end-of-day log.", ""),
    @("Interventional", "Esketamine", "Demo: Esketamine session monitoring", "Technician", "2026-02-20", "${u}esk-monitoring", "Vitals schedule and observation period.", "Older copy kept for reference; see 2026 revision."),
    @("Safety", "", "Demo: Patient in crisis", "All staff", "2025-05-21", "${u}crisis", "What to do, who to call and how to document a crisis.", ""),
    @("Safety", "", "Demo: Suicidal ideation response", "Technician, Clinician", "2024-05-14", "${u}si", "Immediate steps when a patient reports suicidal thoughts.", ""),
    @("Medical Records", "", "Demo: Records request", "Medical Records", "2026-06-24", "${u}records-request", "Receiving, reviewing and sending a records request.", ""),
    @("Medical Records", "", "Demo: Chart amendment request", "Medical Records", "2026-08-27", "${u}amendment", "Handling a patient request to amend their chart.", ""),
    @("Patient Portal", "", "Demo: Replying to portal messages", "Front Office, Clinic Help Desk", "2026-05-16", "${u}portal-messages", "Reviewing and answering patient portal messages.", ""),
    @("Patient Portal", "", "Demo: Portal password reset", "Front Office", "2026-05-16", "${u}portal-reset", "Resetting a patient's portal password or username.", "")
  ) "KinnectGuide"

  # Appt List: plain report paste (no table), dated today
  $a = $wb.Worksheets.Item("Appt List")
  $appts = @(
    @("Appt Date", "Appt Time", "Patient", "Provider", "Status"),
    @("=TODAY()", (8.5/24), "Doe, Jane", "Dr. Avery Rivera", "Scheduled"),
    @("=TODAY()", (9/24), "Roe, Rich", "Dr. Jordan Okafor", "Scheduled"),
    @("=TODAY()", (9/24), "Poe, Alex", "MagVenture", "Scheduled"),
    @("=TODAY()", (9.5/24), "Moe, Casey", "BrainsWay", "Scheduled"),
    @("=TODAY()", (10/24), "Loe, Dana", "Dr. Avery Rivera", "Cancelled"),
    @("=TODAY()", (10.5/24), "Ng, Al", "Sam Lee, PMHNP", "Scheduled"),
    @("=TODAY()", (11/24), "Kay, Morgan", "Dr. Jordan Okafor", "Scheduled"),
    @("=TODAY()+1", (9/24), "Fay, Robin", "Dr. Avery Rivera", "Scheduled")
  )
  Put $a $appts $null
  $a.Range("A2:A9").NumberFormat = "m/d/yyyy"; $a.Range("B2:B9").NumberFormat = "h:mm AM/PM"; [void]$a.Range("A:E").Columns.AutoFit()

  $wb.Worksheets.Item("Start Here").Activate()
  $wb.SaveAs($out, 51)
  $wb.Close($false)
  "saved: $out"
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
