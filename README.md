# Endpoint Advisor v3

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

## Scale

- Endpoints poll a static file — no database, no concurrency issues
- 9,000 endpoints at 15-min intervals = ~10 req/sec — trivial for any web server or GitHub CDN
- Agents read only — nothing on the endpoint can modify the content
