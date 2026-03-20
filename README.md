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

| Setting | Description | Default |
|---------|-------------|---------|
| `ContentDataUrl` | URL to the hosted ContentData.json file | Required |
| `GitHubToken` | Optional GitHub PAT if repo is private | Empty |
| `RefreshInterval` | How often to poll for changes (seconds) | 900 (15 min) |
| `LogPath` | Where to write the agent log | `C:\ProgramData\EndpointAdvisor\EA.log` |
| `LogMaxSizeMB` | Max log file size before rotation | 2 MB |

### 3. Deploy to endpoints

Copy all agent files to a folder on the endpoint (e.g. `C:\Program Files\EndpointAdvisor\`) and run at logon via BigFix or GPO:

```
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Program Files\EndpointAdvisor\EA.ps1"
```

### 4. Manage announcements

Open `admin/EA_Admin.html` in a browser. You need a GitHub PAT with write access to the repo.

1. Enter your token + repo details in the sidebar
2. Click **Load from GitHub**
3. Add or edit announcements
4. Click **Commit to GitHub** — endpoints pick it up on next poll

---

## Dashboard Sections

The agent displays a system tray dashboard with the following sections. Each section loads asynchronously in its own runspace so the UI never blocks.

### 1. Announcements

| Item | Detail |
|------|--------|
| **Source** | `ContentData.json` hosted on GitHub (or any web server) |
| **How it works** | Agent polls the JSON URL on a configurable interval. Parses announcements and evaluates client-side targeting conditions. |
| **Admin control** | Priority (info/warning/critical), enabled toggle, nag settings, links, expiration dates |
| **Nag behavior** | Critical announcements can be set to re-toast every X minutes until the user acknowledges them |
| **Toast** | Uses Windows Toast Notification API (`ToastNotificationManager`). Falls back to `NotifyIcon.ShowBalloonTip` if unavailable. |
| **Tray icon** | Switches to alert icon (`EA_LOGO_MSG.ico`) when unread announcements exist |

### 2. BigFix Software Updates

| Item | Detail |
|------|--------|
| **Source** | Local text file `C:\temp\X-Fixlet-Source_Count.txt` populated by a BigFix Fixlet |
| **How it works** | A BigFix Fixlet runs periodically on the endpoint and writes pending software offer names (one per line) to the text file. The agent reads this file and displays each update as a card. |
| **Data format** | Plain text, one application name per line (e.g. `Google Chrome 134.0.6998.89`) |
| **Staleness check** | Displays file timestamp and age in hours. Flags red if > 24 hours old. |
| **Action button** | "Open Self Service — Install Updates" launches BigFix Self Service Application (`BigFixSSA.exe`). Falls back to `BESClientUI.exe` if SSA is not installed. |
| **Button paths** | SSA: `C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe`<br>Client UI: `C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClientUI.exe` |

### 3. System Patch Updates

| Item | Detail |
|------|--------|
| **Source** | ECM/SCCM client via WMI — `CCM_SoftwareUpdate` class in `ROOT\ccm\ClientSDK` namespace |
| **How it works** | Queries `Get-CimInstance` with a 30-second timeout. Filters for `ComplianceState -eq 0` (not compliant = update pending). |
| **Reboot detection** | Each update is checked for restart requirement using multiple indicators (see below) |
| **Display** | Each pending update shown as an amber-bordered card with title and deadline |
| **Deadline coloring** | Red if deadline < 3 days away, grey otherwise |
| **Action button** | "Open Software Center" launches `SCClient.exe` or falls back to `softwarecenter:` URI |

**Reboot Detection Logic:**

An update is flagged as "(restart required)" if ANY of these conditions are true:

| Check | Description |
|-------|-------------|
| `IsRebootPending` | Direct boolean property on the WMI object |
| `EvaluationState` in (8, 9, 10) | Restart pending / restart required / reboot required states |
| `RebootOutsideServiceWindow` = true | Reboot might be required outside maintenance windows |
| Name matches pattern | `Cumulative Update`, `Security Update`, or `Servicing Stack` in the update title |

### 4. Account Info (Active Directory)

| Item | Detail |
|------|--------|
| **Source** | Active Directory via ADSI (no AD PowerShell module required) |
| **How it works** | Uses `[adsisearcher]` to query the current user's AD object by `sAMAccountName`. No domain controller binding issues — uses the default ADSI provider. |
| **Properties queried** | `displayName`, `pwdLastSet`, `userAccountControl`, `msDS-UserPasswordExpiryTimeComputed` |
| **Password expiry** | Calculated from `msDS-UserPasswordExpiryTimeComputed` (FileTime format). This property accounts for fine-grained password policies. |
| **Display** | Shows username, display name, password expiry date, and days remaining |
| **Color coding** | Green (> 14 days), Amber (≤ 14 days), Red (expired) |
| **Never expires** | Detected when the FileTime value is 0 or `Int64.MaxValue` |

### 5. YubiKey / Certificate Info

| Item | Detail |
|------|--------|
| **Source** | YubiKey Manager CLI (`ykman.exe`) + Windows Certificate Store |
| **ykman path** | `C:\Program Files\Yubico\Yubikey Manager\ykman.exe` |
| **How it works** | Runs `ykman info` to detect YubiKey presence, then exports PIV certificates from each slot using `ykman piv certificates export <slot> -` |
| **Slots checked** | 9a (Authentication), 9c (Digital Signature), 9e (Card Authentication). Slot 9d (Key Management) is excluded. |
| **Certificate parsing** | Exports PEM to temp file → loads as `X509Certificate2` → reads `NotAfter` date → calculates days remaining |
| **Virtual Smart Card** | Also checks `Cert:\CurrentUser\My` for certificates with "Virtual" or "Smart Card" in Subject/Issuer |
| **Color coding** | Green (> 14 days), Amber (≤ 14 days), Red (expired) |
| **Display** | Slot label, expiry date, days remaining |

### 6. Support Info

| Item | Detail |
|------|--------|
| **Source** | `ContentData.json` → `Data.Support` array |
| **How it works** | Same JSON file as announcements. Support entries have `Text`, `Details`, and `Links` fields. |
| **Display** | Purple-bordered cards with title, details text, and clickable hyperlinks |
| **Admin control** | Managed via the same `EA_Admin.html` interface |

### 7. Footer

| Item | Detail |
|------|--------|
| **OS Info** | Windows edition + display version (e.g. "Windows 11 Enterprise Build: 25H2"). Queried via `Get-CimInstance Win32_OperatingSystem` + registry `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion`. |
| **EA Version** | "Endpoint Advisor v7.0.0" displayed below OS info |

---

## Automatic Toast Notifications

### System Patch Restart Required Toast

When the agent detects pending Windows updates that require a restart, it proactively notifies the user via Windows toast notification.

**Behavior:**
1. On app startup (after ~8 second delay), a background PowerShell runspace queries `CCM_SoftwareUpdate` via WMI
2. If any updates require a restart, a toast fires: *"X update(s) require a restart. Please open Software Center to apply the update."*
3. Every **4 hours**, the check repeats and the toast fires again if the update is still pending
4. When the user **opens the dashboard**, toasts stop (acknowledged)
5. On the next 4-hour cycle, the acknowledged flag resets — if updates are still pending, toasting resumes

**Purpose:** Users complain about being rebooted at inconvenient times. This proactively lets them know they have a pending reboot so they can handle it when convenient.

**Technical Implementation:**

Toast state is managed via a synchronized hashtable shared between background runspaces and the main UI thread:

```powershell
$Script:ToastFlags = [hashtable]::Synchronized(@{
    RestartCount  = 0        # Number of updates requiring restart
    PatchCount    = 0        # Number of updates not requiring restart
    Acknowledged  = $false   # Has user seen the dashboard?
})
```

| Location | What happens |
|----------|-------------|
| Line 901 | Initialized at startup: `Acknowledged = $false` |
| Line 837 | Set `Acknowledged = $true` when dashboard window opens |
| Line 906 | Toast timer checks flag — skips if acknowledged |
| Line 926 | 4-hour periodic timer resets `Acknowledged = $false` and re-runs WU check |

**Startup WU Check Flow:**
1. 5-second timer fires after app launch
2. Creates a new PowerShell runspace (isolated, non-blocking)
3. Runspace queries `Get-CimInstance -Namespace ROOT\ccm\ClientSDK -ClassName CCM_SoftwareUpdate -OperationTimeoutSec 30`
4. Counts restart-required updates using the reboot detection logic
5. Sets `ToastFlags.RestartCount` on the synchronized hashtable
6. Main thread toast timer (polling every 5s) picks up the flag and calls `Show-Toast`

### Pending Patch Toast (No Reboot Required)

For updates that do NOT require a reboot:
- Fires **once per day** with Info priority
- *"X update(s) available in Software Center. Please install at your earliest convenience."*
- Tracked via user environment variable `EA_LAST_PATCH_TOAST` (stores today's date `yyyy-MM-dd`)
- Will not re-fire until the next calendar day

---

## Admin Panel (EA_Admin.html)

Single-file HTML admin interface — no server required. Opens in any browser.

### Features
- **GitHub integration**: Load/save `ContentData.json` directly to a GitHub repo
- **Announcement management**: Add, edit, delete announcements with priority levels
- **Support info management**: Add help desk contacts, links, documentation
- **Live preview**: See how announcements will render before publishing

### Announcement Settings

| Setting | Description |
|---------|-------------|
| Title | Announcement headline |
| Message | Body text (supports newlines) |
| Priority | Info (blue), Warning (amber), Critical (red) |
| Enabled | Toggle on/off without deleting |
| Nag until acknowledged | For critical items — re-toasts at a configurable interval |
| Nag interval | Minutes between re-notifications (default: 30) |
| Links | Optional action buttons with URLs |

### Future: Registry-Based Targeting

Planned feature to target announcements based on registry keys (e.g., department, team, location). Example: show announcement only to machines where `HKLM\SOFTWARE\Company\Department` = `Engineering`. Not yet implemented.

---

## Version History

### v7.0.0 (March 2026)
- Renamed sections: "BigFix Software Updates" and "System Patch Updates"
- Automatic restart-required toast notifications (startup + every 4 hours)
- Daily pending patch toast notifications (once per day, no reboot needed)
- Acknowledged flag — toasts stop when dashboard is opened, resume on next 4-hour cycle
- Reboot detection: checks `IsRebootPending`, `EvaluationState`, `RebootOutsideServiceWindow`, and update name patterns
- Background WU check runs independently of dashboard (no need to open UI to get toasted)
- Removed YubiKey slot 9d from display
- Admin panel: nag controls moved to separate styled row below enabled toggle
- Version footer: displays "Endpoint Advisor v7.0.0" and Windows build info
- Software Center button scope fix
- WMI query switched to `Get-CimInstance` with 30-second timeout to prevent hangs

### v6.4.7 (Previous)
- Legacy version — see previous repository

---

## Scale

- Endpoints poll a static file — no database, no concurrency issues
- 9,000 endpoints at 15-min intervals = ~10 req/sec — trivial for any web server or GitHub CDN
- Agents read only — nothing on the endpoint can modify the content
- All WMI/AD/certificate queries run in isolated PowerShell runspaces — UI thread never blocks
