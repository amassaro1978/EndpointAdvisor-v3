# Endpoint Advisor v3 — Backendless Edition

Zero infrastructure. No Node.js. No database. No server to maintain.

## How it works

```
ContentData.json  ←  EA_Admin.html writes via GitHub API
      ↓
   (any URL — GitHub raw, IIS, file share)
      ↓
   EA.ps1 on 9,000 endpoints polls on a timer
```

## Files

| File | Purpose |
|------|---------|
| `ContentData.json` | The content file. Host this anywhere — GitHub, IIS, internal CDN |
| `admin/EA_Admin.html` | Admin dashboard. Open in any browser, no server needed |
| `agent/EA.ps1` | PowerShell tray agent — deploy via BigFix/GPO |
| `agent/EA.config.json` | Agent config — set the URL, refresh interval, platform |

## Setup

### 1. Host ContentData.json
Put it anywhere reachable by endpoints:
- Internal GitHub (raw URL)
- IIS static file
- Network share served via HTTP
- Any internal web server

### 2. Configure the agent
Edit `EA.config.json`:
```json
{
  "ContentDataUrl":  "https://raw.githubusercontent.com/YOUR_ORG/YOUR_REPO/main/ContentData.json",
  "RefreshInterval": 900,
  "Platform":        "Windows"
}
```

### 3. Deploy the agent
Via BigFix or GPO — copy the agent folder to `C:\ProgramData\EndpointAdvisor\`
and run at logon:
```
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\ProgramData\EndpointAdvisor\EA.ps1"
```

### 4. Use the admin dashboard
Open `EA_Admin.html` in a browser.
- Enter your GitHub token + repo details in the sidebar
- Click **Load from GitHub** to pull the current JSON
- Add/edit/delete announcements visually
- Click **Commit to GitHub** — agents pick it up on next poll

No token? Use **Download JSON**, update the file manually, re-upload to your host.

## Targeting

Targeted announcements are evaluated **client-side** by the agent.
The JSON condition is checked against the local machine (registry, etc.) — no server logic needed.

```json
{
  "Condition": {
    "Type":  "Registry",
    "Path":  "HKLM:\\SOFTWARE\\YourOrg\\Targeting",
    "Name":  "Group",
    "Value": "IT"
  }
}
```

## Scale

- 9,000 endpoints polling a static file every 15 min = trivial load on any web server
- No concurrent writes — agents only read
- No database — nothing to maintain or patch beyond the JSON file itself
