# Endpoint Advisor v7.2.1

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

### 3. Configure company registry path

Edit line 27 of `EA.ps1`:
```powershell
$Script:CompanyRegPath = "YourCompanyName"
```

This is used for group-targeted announcements. The agent reads the device group from:
```
HKLM:\SOFTWARE\<CompanyRegPath>\targeting\GROUP
```

Set this registry key on endpoints via GPO or deployment script:
```powershell
New-Item -Path "HKLM:\SOFTWARE\YourCompanyName\targeting" -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\YourCompanyName\targeting" -Name "GROUP" -Value "ENG"
```

### 4. Deploy to endpoints

Copy all agent files to a folder on the endpoint (e.g. `C:\Program Files\EndpointAdvisor\`) and run at logon via BigFix or GPO:

```
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Program Files\EndpointAdvisor\EA.ps1"
```

### 5. Manage announcements

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
| **Admin control** | Priority (info/warning/critical), enabled toggle, nag settings, links, expiration dates, target group, conditional rules |
| **Nag behavior** | Critical announcements can be set to re-toast every X minutes until the user acknowledges them |
| **Toast behavior** | Only critical announcements trigger toasts — persistent with Dismiss + Snooze buttons. Info/Warning are dashboard-only (no toast). System Update toasts include Update Now + Snooze + Dismiss. |
| **Snooze** | Windows built-in snooze — defaults to ~10 minutes, then re-fires. Toast persists even after app exits. |
| **Tray icon** | Switches to alert icon (`EA_LOGO_MSG.ico`) when unread announcements exist |

#### Group-Targeted Announcements

Announcements can be targeted to specific device groups. The admin sets a `TargetGroup` value in the announcement (e.g., "ENG", "HR", "FINANCE"). Only endpoints whose registry `GROUP` value matches will see the announcement.

| Scenario | Behavior |
|----------|----------|
| Announcement has no TargetGroup | Shown to **all** devices |
| Announcement has TargetGroup = "All" | Shown to **all** devices |
| Announcement has TargetGroup = "ENG" | Only shown to devices where `GROUP = "ENG"` |
| Device has no GROUP registry key | **Never** sees targeted announcements (only untargeted ones) |

#### Inactive Announcement Visibility

In the admin panel, announcements that are not currently active are greyed out (50% opacity) for easy identification:

| State | Visual |
|-------|--------|
| **Disabled** | Greyed out + "Disabled" pill |
| **Expired** (EndDate in the past) | Greyed out + "Expired" pill |
| **Scheduled** (StartDate in the future) | Greyed out + "Scheduled" pill |
| **Active** | Full opacity, no status pill |

#### Conditional Announcements: Password Expiry

Announcements can be conditionally displayed based on the user's password status:

| Setting | Description |
|---------|-------------|
| Condition | `password_expiry` |
| ConditionThresholdDays | Number of days before expiry to start showing (default: 14) |
| ConditionKbUrl | Optional link to "how to change your password" KB article |

**How it works:**
1. Admin creates an announcement with Condition = `password_expiry` and threshold = 14 days
2. Agent queries AD for `msDS-UserPasswordExpiryTimeComputed`
3. If password expires within threshold → announcement is shown
4. If password is fine (>14 days) → announcement is hidden
5. When user changes password → announcement auto-hides on next check
6. If AD is unreachable → fails safe (announcement hidden)

#### Conditional Announcements: Certificate Expiry

Announcements can be conditionally displayed based on certificate expiration:

| Setting | Description |
|---------|-------------|
| Condition | `cert_expiry` |
| ConditionThresholdDays | Number of days before expiry to start showing (default: 14) |
| ConditionKbUrl | Optional link to certificate renewal instructions |

**How it works:**
1. Admin creates an announcement with Condition = `cert_expiry` and threshold = 14 days
2. Agent checks all certificates in `Cert:\CurrentUser\My` and YubiKey PIV slots (9a, 9c, 9e)
3. If any certificate expires within threshold → announcement is shown
4. If all certs are healthy → announcement is hidden

#### Registry Key Targeting

Announcements can be targeted based on registry key existence and value matching. This is useful for targeting specific OS versions, installed software, or configuration states.

| Setting | Description |
|---------|-------------|
| TargetRegKey | Registry key path (e.g., `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion`) |
| TargetRegValue | Optional value name to check under the key (e.g., `DisplayVersion`) |
| TargetRegData | Optional data match — announcement shows only if value data contains this text (e.g., `24H2`) |

**Matching behavior:**

| Configuration | Behavior |
|---------------|----------|
| Key only | Show if registry key exists |
| Key + Value | Show if the value name exists under the key |
| Key + Value + Data | Show if value data contains the specified text (case-insensitive) |

**Example — Target machines still on Windows 24H2:**
- Registry Key: `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion`
- Value Name: `DisplayVersion`
- Value Data Match: `24H2`

### 2. Application Updates (formerly "BigFix Software Updates")

| Item | Detail |
|------|--------|
| **Source** | Local text file `C:\temp\X-Fixlet-Source_Count.txt` populated by a BigFix Fixlet |
| **How it works** | A BigFix Fixlet runs periodically on the endpoint and writes pending software offer names (one per line) to the text file. The agent reads this file and displays each update as a card. |
| **Data format** | Plain text, one application name per line (e.g. `Google Chrome 134.0.6998.89`) |
| **Staleness check** | Displays file timestamp and age in hours. Flags red if > 24 hours old. |
| **Action button** | "Open Self Service — Install Updates" launches BigFix Self Service Application (`BigFixSSA.exe`). Falls back to `BESClientUI.exe` if SSA is not installed. |

### 3. Microsoft Updates (formerly "System Patch Updates")

| Item | Detail |
|------|--------|
| **Source** | ECM/SCCM client via WMI — `CCM_SoftwareUpdate` class in `ROOT\ccm\ClientSDK` namespace |
| **How it works** | Queries `Get-CimInstance` with a 30-second timeout. Filters for `ComplianceState -eq 0` (not compliant = update pending). |
| **Reboot detection** | Each update is checked for restart requirement using multiple indicators (see below) |
| **Display** | Each pending update shown as an amber-bordered card with title and deadline |
| **Deadline coloring** | Red if deadline < 3 days away, grey otherwise |
| **Action button** | "Open Software Center" launches `SCClient.exe` or falls back to `softwarecenter:` URI |
| **Toast: Update Now** | System update toast includes an "Update Now" button that opens Software Center via `softwarecenter:` protocol |

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

### 5. Certificate Monitoring

The agent monitors multiple certificate types from the Windows Certificate Store and YubiKey.

#### YubiKey PIV Certificates

| Item | Detail |
|------|--------|
| **Source** | YubiKey Manager CLI (`ykman.exe`) |
| **ykman path** | `C:\Program Files\Yubico\Yubikey Manager\ykman.exe` |
| **How it works** | Runs `ykman info` to detect YubiKey presence, then exports PIV certificates from each slot using `ykman piv certificates export <slot> -` |
| **Slots checked** | 9a (Authentication), 9c (Digital Signature), 9e (Card Authentication). Slot 9d (Key Management) is excluded. |
| **Certificate parsing** | Exports PEM to temp file → loads as `X509Certificate2` → reads `NotAfter` date → calculates days remaining |
| **Color coding** | Green (> 14 days), Amber (≤ 14 days), Red (expired) |
| **Display format** | Friendly name with subtitle (e.g., "Hardware token authentication") and status: `[OK] VALID`, `[WARNING] EXPIRING SOON`, or `[EXPIRED]` |

#### Smart Card Certificates (Physical, Virtual, YubiKey)

| Item | Detail |
|------|--------|
| **Source** | `Cert:\CurrentUser\My` — Windows Certificate Store |
| **Detection** | Finds certs with Smart Card Logon OID (`1.3.6.1.4.1.311.20.2.2`) and `HasPrivateKey = true` |
| **Classification** | Uses **certificate template name** (OID `1.3.6.1.4.1.311.21.7`) as the primary identifier, then falls back to Subject/Issuer text matching |
| **Display** | Separate rows for each type detected: |

| Label | How identified |
|-------|---------------|
| **Smart Card Cert** | Physical smart card — no "Virtual" or "Yubi" in template/subject |
| **Virtual Smart Card** | Template name contains "Virtual" (e.g., "Virtual Smart Card Logon") |
| **YubiKey Cert** | Template name contains "Yubi" or "PIV", or Subject/Issuer contains "Yubi" (unless template says "Virtual") |

#### Email Signing Certificate (S/MIME)

| Item | Detail |
|------|--------|
| **Source** | `Cert:\CurrentUser\My` |
| **Detection** | Enhanced Key Usage contains Secure Email OID (`1.3.6.1.5.5.7.3.4`) |
| **Display** | Shows expiry date and days remaining |

#### Email Encryption Certificate

| Item | Detail |
|------|--------|
| **Source** | `Cert:\CurrentUser\My` |
| **Detection** | Secure Email OID (`1.3.6.1.5.5.7.3.4`) AND KeyEncipherment key usage flag AND `HasPrivateKey = true` |
| **Display** | Shows as separate row from signing cert, even if same certificate serves both purposes |

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
| **EA Version** | "Endpoint Advisor v7.2.0" displayed below OS info |

---

## Toast Notification System

### Announcement Toast Behavior by Priority

| Priority | Toast? | Behavior |
|----------|--------|----------|
| **Info** | ❌ No | Dashboard only — no toast notification |
| **Warning** | ❌ No | Dashboard only — no toast notification |
| **Critical** | ✅ Yes | Persistent toast with **Dismiss** + **Snooze** buttons |

### Critical Announcement Toasts

Critical announcements use Windows `scenario="reminder"` toasts:
- **Stays on screen** until the user takes action (no auto-dismiss)
- **Dismiss** button — marks as seen, won't show again (unless Nag is enabled)
- **Snooze** button — dismisses temporarily, comes back after ~10 minutes (Windows default)
- **Toast persists even after the EA app is closed** — Windows owns the snooze timer
- **Nag behavior** — if NagEnabled is checked in the admin panel, the toast re-fires at the configured interval (default: 30 min) even after Dismiss, until the user opens the dashboard
- Info and Warning announcements **never** trigger a toast — they are visible only in the dashboard

### System Update Toast

When pending Windows updates require a restart:
- **Update Now** button — opens Software Center (`softwarecenter:` protocol)
- **Snooze** button — temporarily dismisses
- **Dismiss** button — marks as seen

### Restart Required Toast (Automatic)

| Item | Detail |
|------|--------|
| **Trigger** | On app startup (after ~8 second delay) + every 4 hours |
| **Condition** | Pending Windows updates that require a restart |
| **Message** | *"X update(s) require a restart. Please open Software Center to apply the update."* |
| **Acknowledged** | When user opens the dashboard, toasts stop. On next 4-hour cycle, if updates are still pending, toasting resumes. |

### Pending Patch Toast (No Reboot Required)

| Item | Detail |
|------|--------|
| **Trigger** | Once per day |
| **Condition** | Pending updates that do NOT require a reboot |
| **Message** | *"X update(s) available in Software Center. Please install at your earliest convenience."* |
| **Tracking** | User environment variable `EA_LAST_PATCH_TOAST` stores today's date. Won't re-fire until the next calendar day. |

---

## Admin Panel (EA_Admin.html)

Single-file HTML admin interface — no server required. Opens in any browser.

### Features
- **GitHub integration**: Load/save `ContentData.json` directly to a GitHub repo
- **Announcement management**: Add, edit, delete announcements with priority levels
- **Group targeting**: Target announcements to specific device groups
- **Conditional announcements**: Password expiry-based announcements with configurable threshold
- **Support info management**: Add help desk contacts, links, documentation
- **Live preview**: See how announcements will render before publishing

### Announcement Settings

| Setting | Description |
|---------|-------------|
| Title | Announcement headline |
| Message | Body text (supports newlines) |
| Priority | Info (blue), Warning (amber), Critical (red) |
| Enabled | Toggle on/off without deleting |
| Target Group | Leave blank for all devices. Enter group name (e.g., "ENG") to target specific devices. |
| Registry Key | Optional registry key targeting — only show if key/value/data matches on endpoint |
| Condition | Optional conditional display (`password_expiry` or `cert_expiry`) |
| Condition Threshold | Days threshold for conditional announcements (e.g., 14 days for password expiry) |
| KB Link | Optional knowledge base URL for conditional announcements |
| Nag until acknowledged | For critical items — re-toasts at a configurable interval |
| Nag interval | Minutes between re-notifications (default: 30) |
| Links | Optional action buttons with URLs |
| Start/End Date | Schedule when the announcement appears and expires |

---

## Architecture

### Threading Model

The dashboard uses **multi-threaded runspaces** — each major section runs in its own isolated PowerShell runspace:

| Thread | Section |
|--------|---------|
| Main thread | WPF UI (dashboard window) |
| Runspace 1 | Announcements (fetch + filter + display) |
| Runspace 2 | BigFix / Application Updates |
| Runspace 3 | Windows Update / Microsoft Updates |
| Runspace 4 | Support Info |
| Runspace 5 | Account Information + Certificates |

UI updates are pushed back to the main thread via `$Dispatcher.Invoke()`. Nothing blocks the UI — if Windows Update takes 30 seconds to query, the rest of the dashboard is already loaded and responsive.

### Toast State Management

Toast state is managed via a synchronized hashtable shared between background runspaces and the main UI thread:

```powershell
$Script:ToastFlags = [hashtable]::Synchronized(@{
    RestartCount  = 0        # Number of updates requiring restart
    PatchCount    = 0        # Number of updates not requiring restart
    Acknowledged  = $false   # Has user seen the dashboard?
})
```

### Scale

- Endpoints poll a static file — no database, no concurrency issues
- 9,000 endpoints at 15-min intervals = ~10 req/sec — trivial for any web server or GitHub CDN
- Agents read only — nothing on the endpoint can modify the content
- All WMI/AD/certificate queries run in isolated PowerShell runspaces — UI thread never blocks

---

## Version History

### v7.2.1 (April 2026)
- **Reboot toast dedupe fix** — persistent restart-required toast now shows only once while reboot state remains pending, instead of stacking duplicate toasts overnight or across repeated background checks
- **Toast state reset** — reboot toast visibility flag clears automatically once reboot-required updates are no longer detected, so future reboot events still notify correctly

### v7.2.0 (April 2026)
- **Inactive announcement visibility** — admin panel greys out disabled, expired, and scheduled announcements with status pills
- **Certificate expiry condition** — conditional announcements triggered by cert expiry within configurable threshold (checks user store + YubiKey)
- **Registry key targeting** — target announcements by registry key existence, value name, or value data match (e.g., target 24H2 machines)
- **Enriched certificate display** — friendly names, subtitles (e.g., "Used for email signing"), and `[OK] VALID` / `[WARNING] EXPIRING SOON` / `[EXPIRED]` status

### v7.1.0 (March 2026)
- **Persistent toast notifications** — critical/high priority announcements stay on screen with Dismiss + Snooze buttons
- **System update toast** — "Update Now" button opens Software Center
- **Group-targeted announcements** — admin targets announcements to specific device groups via registry-based targeting
- **Conditional announcements** — password expiry-based announcements with configurable threshold and KB link
- **Email encryption cert monitoring** — separate from email signing cert
- **Virtual smart card detection** — identified by certificate template name, displayed as separate row
- **Smart card cert classification** — distinguishes Physical, Virtual, and YubiKey certs using template names
- **Configurable company registry path** — `$Script:CompanyRegPath` variable at top of script
- **Renamed sections** — "BigFix Software Updates" → "Application Updates", "System Patch Updates" → "Microsoft Updates"

### v7.0.0 (March 2026)
- Automatic restart-required toast notifications (startup + every 4 hours)
- Daily pending patch toast notifications (once per day, no reboot needed)
- Acknowledged flag — toasts stop when dashboard is opened, resume on next 4-hour cycle
- Reboot detection: checks `IsRebootPending`, `EvaluationState`, `RebootOutsideServiceWindow`, and update name patterns
- Background WU check runs independently of dashboard
- Removed YubiKey slot 9d from display
- Admin panel: nag controls moved to separate styled row below enabled toggle
- Version footer
- Software Center button scope fix
- WMI query switched to `Get-CimInstance` with 30-second timeout

### v6.4.7 (Previous)
- Legacy version — see previous repository
