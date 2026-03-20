# Endpoint Advisor v7.0.0

Zero infrastructure. No Node.js. No database. No server to maintain.

## How it works

```
EA_Admin.html  →  commits ContentData.json to GitHub
                        ↓
              raw URL (no auth required)
                        ↓
         EA.ps1 polls on every endpoint
```

## Files

| File | Purpose |
|------|---------|
| `ContentData.json` | Announcements + support info. Host anywhere reachable by endpoints. |
| `admin/EA_Admin.html` | Admin dashboard. Open in any browser — no server needed. |
| `agent/EA.ps1` | PowerShell tray agent. Deploy via BigFix or GPO. |
| `agent/EA.config.json` | Agent config — URL and refresh interval. |
| `agent/EA_LOGO.ico` | Tray icon (normal state). |
| `agent/EA_LOGO_MSG.ico` | Tray icon (unread announcement state). |

## Setup

### 1. Host ContentData.json

Recommended: internal GitHub repo set to **Internal** visibility.

Use the raw URL — no token required:
```
https://github.yourcompany.com/raw/IT/EndpointAdvisor/main/ContentData.json
```

### 2. Configure the agent

Edit `agent/EA.config.json`:
```json
{
  "ContentDataUrl": "https://github.yourcompany.com/raw/IT/EndpointAdvisor/main/ContentData.json",
  "GitHubToken": "",
  "RefreshInterval": 900,
  "LogPath": "C:\\ProgramData\\EndpointAdvisor\\EA.log",
  "LogMaxSizeMB": 2
}
```

### 3. Deploy to endpoints

Copy all 4 files to a folder on the endpoint (e.g. `C:\Program Files\EndpointAdvisor\`) and run at logon via BigFix or GPO:

```
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Program Files\EndpointAdvisor\EA.ps1"
```

### 4. Manage announcements

Open `admin/EA_Admin.html` in a browser. You need a GitHub PAT with write access to the repo.

1. Enter your token + repo details in the sidebar
2. Click **Load from GitHub**
3. Add or edit announcements
4. Click **Commit to GitHub** — endpoints pick it up on next poll

## Dashboard Sections

| Section | Source | Description |
|---------|--------|-------------|
| Announcements | ContentData.json (GitHub) | Admin-managed announcements with priority, nag settings, targeting |
| BigFix Software Updates | BigFix REST API | Pending software offers from BigFix Self Service |
| System Patch Updates | ECM/SCCM (WMI) | Pending Windows updates with reboot detection |
| Account Info | Active Directory | Password expiration, account status |
| YubiKey Info | ykman CLI | PIV certificate details for slots 9a, 9c, 9e |
| Support | ContentData.json | Support contact info and links |

## Automatic Toast Notifications

### System Patch Restart Required Toast

When the agent detects pending Windows updates that require a restart, it proactively notifies the user via Windows toast notification.

**Behavior:**
- On app startup (after ~8 second delay), a background runspace queries `CCM_SoftwareUpdate` via WMI
- If any updates require a restart, a toast fires: *"X update(s) require a restart. Please open Software Center to apply the update."*
- Every **4 hours**, the check repeats and the toast fires again if the update is still pending
- Toast uses the Windows Toast Notification API (`ToastNotificationManager`)

**Acknowledgement:**
- When the user **opens the dashboard**, the `Acknowledged` flag is set to `true` — toasts stop for the current cycle
- On the next 4-hour cycle, the flag resets to `false` and the toast will fire again if updates are still pending
- This continues until the user actually installs the update and reboots

**Technical Details:**
- Toast state is managed via a synchronized hashtable (`$Script:ToastFlags`) shared between background runspaces and the main UI thread
- Flags: `RestartCount` (int), `PatchCount` (int), `Acknowledged` (bool)
- Line 901: Initialized at startup → `@{ RestartCount = 0; PatchCount = 0; Acknowledged = $false }`
- Line 837: Set `Acknowledged = $true` when dashboard window opens
- Line 906: Toast timer checks `Acknowledged` flag — skips if `$true`
- Line 926: 4-hour periodic timer resets `Acknowledged = $false` and re-runs WU check
- Background WU check uses a PowerShell runspace with `BeginInvoke()` (non-blocking)
- WMI query uses `Get-CimInstance` with `-OperationTimeoutSec 30` to prevent hangs

**Reboot Detection Logic:**
Updates are flagged as "restart required" if any of these conditions are true:
- `IsRebootPending` property is `$true`
- `EvaluationState` is 8, 9, or 10 (restart pending/required states)
- `RebootOutsideServiceWindow` is `$true`
- Update name matches: `Cumulative Update`, `Security Update`, or `Servicing Stack`

### Pending Patch Toast (No Reboot)

For updates that do NOT require a reboot, a gentler toast fires **once per day**:
- *"X update(s) available in Software Center. Please install at your earliest convenience."*
- Uses `Info` priority (not `Critical`)
- Tracked via user environment variable `EA_LAST_PATCH_TOAST` (stores today's date)
- Will not re-fire until the next calendar day

## Admin Panel (EA_Admin.html)

### Critical Announcement Nag Settings
- **Enabled toggle**: Enable/disable the announcement
- **Nag until acknowledged**: When enabled, the announcement will re-toast at a configurable interval
- **Nag interval (minutes)**: How often to re-display (default: 30 min)

### Future: Registry-Based Targeting
Planned feature to target announcements based on registry keys (e.g., department, team, location). Not yet implemented — waiting on registry key definitions.

## Version History

### v7.0.0 (March 2026)
- Renamed sections: "BigFix Software Updates" and "System Patch Updates"
- Automatic restart-required toast notifications (startup + every 4 hours)
- Daily pending patch toast notifications
- Acknowledged flag — toasts stop when dashboard is opened, resume on next cycle
- Reboot detection: checks `IsRebootPending`, `EvaluationState`, `RebootOutsideServiceWindow`, and update name patterns
- Removed YubiKey slot 9d from display
- Admin panel: increased spacing between enabled toggle and nag settings
- Version footer: displays EA version and Windows build info
- Software Center button fix (scope issue)
- WMI query timeout (30s) to prevent hangs

### v6.4.7 (Previous)
- Legacy version — see previous repository

## Scale

- Endpoints poll a static file — no database, no concurrency issues
- 9,000 endpoints at 15-min intervals = ~10 req/sec — trivial for any web server or GitHub CDN
- Agents read only — nothing on the endpoint can modify the content
