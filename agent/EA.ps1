#Requires -Version 5.1
<#
.SYNOPSIS
    Endpoint Advisor v3 - Backendless edition
.DESCRIPTION
    Polls a JSON file (hosted on GitHub or any internal web server) for
    announcements, evaluates client-side targeting conditions, and displays
    a WPF system-tray dashboard. No backend server required.
#>

$Script:EAVersion = "7.0.0"

# ---- Config ------------------------------------------------------------------
$ScriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ConfigPath = Join-Path $ScriptDir "EA.config.json"

$DefaultConfig = @{
    ContentDataUrl  = "https://raw.githubusercontent.com/YOUR_ORG/YOUR_REPO/main/ContentData.json"
    RefreshInterval = 900
    LogPath         = "C:\ProgramData\EndpointAdvisor\EA.log"
    LogMaxSizeMB    = 2
}

function Get-Config {
    if (Test-Path $ConfigPath) {
        try { return Get-Content $ConfigPath -Raw | ConvertFrom-Json }
        catch { return $DefaultConfig }
    }
    return $DefaultConfig
}

$Config   = Get-Config
$Hostname = $env:COMPUTERNAME
$Username = $env:USERNAME

# ---- Logging -----------------------------------------------------------------
function Write-Log($msg) {
    try {
        $logDir = Split-Path $Config.LogPath
        if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
        Add-Content -Path $Config.LogPath -Value $entry -Encoding UTF8
        # Rotate at max size
        if ((Get-Item $Config.LogPath -ErrorAction SilentlyContinue).Length -gt ($Config.LogMaxSizeMB * 1MB)) {
            $bak = $Config.LogPath -replace '\.log$', '.bak.log'
            Move-Item $Config.LogPath $bak -Force
        }
    } catch {}
}

# ---- Content fetching --------------------------------------------------------
$Script:CachedContent    = $null
$Script:ContentVersion   = $null
$Script:SeenIds          = [System.Collections.Generic.HashSet[string]]::new()
$Script:LastNagTime      = @{}
$Script:ActiveRunspaces  = [System.Collections.Generic.List[object]]::new()

function Get-ContentData {
    try {
        $headers = @{ 'User-Agent' = 'EndpointAdvisor' }
        if ($Config.GitHubToken) { $headers['Authorization'] = "token $($Config.GitHubToken)" }
        $params = @{ Uri = $Config.ContentDataUrl; UseBasicParsing = $true; TimeoutSec = 20; Headers = $headers; ErrorAction = 'Stop' }
        $raw    = Invoke-WebRequest @params
        try {
            $data = $raw.Content | ConvertFrom-Json
            Write-Log "Content fetched OK (v$($data.contentVersion))"
            return $data
        } catch {
            $preview = $raw.Content.Substring(0, [Math]::Min(300, $raw.Content.Length))
            Write-Log "JSON parse failed. HTTP $($raw.StatusCode). Body: $preview"
            return $null
        }
    } catch {
        Write-Log "Content fetch failed: $_"
        return $null
    }
}

function Get-RelevantAnnouncements($data) {
    if (-not $data -or -not $data.Data -or -not $data.Data.Announcements) { return @() }
    $now     = Get-Date
    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($item in $data.Data.Announcements.Default) {
        if (-not $item) { continue }
        if ($item.Enabled -eq $false) { continue }
        if ($item.StartDate -and ([datetime]$item.StartDate) -gt $now) { continue }
        if ($item.EndDate   -and ([datetime]$item.EndDate)   -lt $now) { continue }
        $results.Add($item) | Out-Null
    }
    return $results.ToArray()
}

# ---- UI assemblies + WPF pre-warm --------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Pre-warm WPF so first dashboard open is instant
try {
    $pw  = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="0" Height="0" WindowStyle="None" ShowInTaskbar="False" Opacity="0"/>'
    $pwR = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($pw))
    $pwW = [System.Windows.Markup.XamlReader]::Load($pwR)
    $pwW.Show()
    $pwW.Close()
} catch {}

# ---- Icon helpers ------------------------------------------------------------
$Script:IconNormal = Join-Path $ScriptDir "EA_LOGO.ico"
$Script:IconAlert  = Join-Path $ScriptDir "EA_LOGO_MSG.ico"

Write-Log "ScriptDir=$ScriptDir IconNormal=$($Script:IconNormal) Exists=$(Test-Path $Script:IconNormal)"

function Get-TrayIcon([bool]$hasUnread = $false) {
    $ico = if ($hasUnread -and (Test-Path $Script:IconAlert)) { $Script:IconAlert }
           elseif (Test-Path $Script:IconNormal) { $Script:IconNormal }
           else { $null }
    if ($ico) {
        try { return New-Object System.Drawing.Icon($ico) }
        catch { return [System.Drawing.SystemIcons]::Information }
    }
    return [System.Drawing.SystemIcons]::Information
}

# ---- Tray icon ---------------------------------------------------------------
function New-TrayIcon {
    $tray = New-Object System.Windows.Forms.NotifyIcon
    $tray.Icon    = Get-TrayIcon $false
    $tray.Text    = "Endpoint Advisor"
    $tray.Visible = $true

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $mi   = { param($t,$c) $i = [System.Windows.Forms.ToolStripMenuItem]::new($t); $i.Add_Click($c); $i }

    $null = $menu.Items.Add((& $mi "Open Dashboard" { Show-Dashboard }))
    $null = $menu.Items.Add((& $mi "Refresh Now"    { Start-ContentRefresh }))
    $null = $menu.Items.Add([System.Windows.Forms.ToolStripSeparator]::new())
    $null = $menu.Items.Add((& $mi "Exit" {
        # Kill all background runspaces
        foreach ($rs in $Script:ActiveRunspaces) {
            try { $rs.Runspace.Close(); $rs.Runspace.Dispose(); $rs.Dispose() } catch {}
        }
        $Script:ActiveRunspaces.Clear()
        if ($Script:DashboardWindow) {
            try { $Script:DashboardWindow.Close() } catch {}
        }
        $Script:TrayIcon.Visible = $false
        $Script:TrayIcon.Dispose()
        [System.Windows.Forms.Application]::Exit()
    }))

    $tray.ContextMenuStrip = $menu
    $tray.Add_Click({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Show-Dashboard }
    })
    $tray.Add_DoubleClick({ Show-Dashboard })
    return $tray
}

# ---- AppId for toast ---------------------------------------------------------
function Register-AppId {
    $appId   = "EndpointAdvisor.Notifications"
    $regPath = "HKCU:\Software\Classes\AppUserModelId\$appId"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
        New-ItemProperty -Path $regPath -Name "DisplayName" -Value "Endpoint Advisor" -PropertyType String -Force | Out-Null
    }
    return $appId
}
$Script:ToastAppId = Register-AppId

function Show-Toast($title, $message, $priority) {
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]                       | Out-Null
        $xml = [Windows.Data.Xml.Dom.XmlDocument]::new()
        $xml.LoadXml("<toast><visual><binding template='ToastGeneric'><text>$([System.Security.SecurityElement]::Escape($title))</text><text>$([System.Security.SecurityElement]::Escape($message))</text></binding></visual></toast>")
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($Script:ToastAppId).Show([Windows.UI.Notifications.ToastNotification]::new($xml))
    } catch {
        $ico = switch ($priority) { 'critical'{'Error'} 'warning'{'Warning'} default{'Info'} }
        $Script:TrayIcon.ShowBalloonTip(5000, $title, $message, [System.Windows.Forms.ToolTipIcon]$ico)
    }
}

# ---- Helper: start a tracked runspace ----------------------------------------
function Start-TrackedRunspace {
    param([scriptblock]$Script, [hashtable]$Parameters)
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()
    $ps = [System.Management.Automation.PowerShell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript($Script)
    [void]$ps.AddParameters($Parameters)
    [void]$ps.BeginInvoke()
    $Script:ActiveRunspaces.Add($ps) | Out-Null
    return $ps
}

# ---- WPF helpers (UI thread only) --------------------------------------------
function New-SectionHeader($text) {
    $b = New-Object System.Windows.Controls.Border
    $b.Background   = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#334155")
    $b.CornerRadius = "4"; $b.Padding = "10,6"; $b.Margin = "0,12,0,6"
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $text; $tb.FontSize = 14; $tb.FontWeight = "SemiBold"
    $tb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#FFFFFF")
    $b.Child = $tb; $b
}
function New-InfoText($text) {
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $text; $tb.Foreground = "#94A3B8"; $tb.FontSize = 12; $tb.Margin = "4,4,0,4"
    $tb
}

# ---- Async section: Announcements --------------------------------------------
function Start-AnnouncementsLoad {
    param($Dispatcher, $Container, $ContentDataUrl, $GitHubToken, $TrayIcon, $IconAlert, $IconNormal)

    Start-TrackedRunspace -Script {
        param($Dispatcher, $Container, $ContentDataUrl, $GitHubToken, $TrayIcon, $IconAlert, $IconNormal)

        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

        # Fetch JSON
        $items     = @()
        $hasUnread = $false
        $fetchErr  = $null
        try {
            $headers = @{ 'User-Agent' = 'EndpointAdvisor' }
            if ($GitHubToken) { $headers['Authorization'] = "token $GitHubToken" }
            $params  = @{ Uri = $ContentDataUrl; UseBasicParsing = $true; TimeoutSec = 20; Headers = $headers; ErrorAction = 'Stop' }
            $raw     = Invoke-WebRequest @params
            $data    = $raw.Content | ConvertFrom-Json
            $now     = Get-Date

            foreach ($item in $data.Data.Announcements.Default) {
                if (-not $item -or $item.Enabled -eq $false) { continue }
                if ($item.StartDate -and ([datetime]$item.StartDate) -gt $now) { continue }
                if ($item.EndDate   -and ([datetime]$item.EndDate)   -lt $now) { continue }
                $items += $item
                $hasUnread = $true
            }
        } catch {
            $fetchErr = $_.ToString()
        }

        $Dispatcher.Invoke([Action]{
            $Container.Children.Clear()

            $mkBrush = { param($c) [System.Windows.Media.BrushConverter]::new().ConvertFrom($c) }
            $mkTb    = {
                param($text, $color = '#94A3B8', $size = 12, $wrap = $true)
                $tb = New-Object System.Windows.Controls.TextBlock
                $tb.Text = $text; $tb.FontSize = $size; $tb.Margin = "4,2,0,2"
                $tb.Foreground = & $mkBrush $color
                if ($wrap) { $tb.TextWrapping = "Wrap" }
                $tb
            }

            if ($fetchErr) {
                $Container.Children.Add((& $mkTb "Failed to load announcements." '#EF4444')) | Out-Null
            } elseif ($items.Count -eq 0) {
                $Container.Children.Add((& $mkTb "No announcements at this time.")) | Out-Null
            } else {
                foreach ($a in $items) {
                    $borderColor = switch ($a.Priority) {
                        'critical' { '#EF4444' }
                        'warning'  { '#F59E0B' }
                        default    { '#3B82F6' }
                    }
                    $bd = New-Object System.Windows.Controls.Border
                    $bd.Background      = & $mkBrush "#FFFFFF"
                    $bd.BorderBrush     = & $mkBrush $borderColor
                    $bd.BorderThickness = "3,0,0,0"
                    $bd.CornerRadius    = "6"
                    $bd.Margin          = "0,0,0,8"
                    $bd.Padding         = "14,10"

                    $sp = New-Object System.Windows.Controls.StackPanel
                    $title = if ($a.Title) { $a.Title } else { "" }
                    $sp.Children.Add((& $mkTb $title '#1E293B' 14)) | Out-Null

                    if ($a.Text) {
                        $clean = $a.Text -replace '\*\*|##|#|\*|`', ''
                        $sp.Children.Add((& $mkTb $clean '#94A3B8' 12)) | Out-Null
                    }

                    if ($a.Details) {
                        $det = New-Object System.Windows.Controls.Expander
                        $det.Header    = "Details"
                        $det.Margin    = "0,6,0,0"
                        $det.Foreground = & $mkBrush "#64748B"
                        $detTb = & $mkTb ($a.Details -replace '\*\*|##|#|\*|`','') '#94A3B8' 11
                        $detTb.Margin  = "8,4,0,4"
                        $det.Content   = $detTb
                        $sp.Children.Add($det) | Out-Null
                    }

                    if ($a.Links -and $a.Links.Count -gt 0) {
                        $linkPanel = New-Object System.Windows.Controls.WrapPanel
                        $linkPanel.Margin = "0,6,0,0"
                        foreach ($lnk in $a.Links) {
                            $hl = New-Object System.Windows.Documents.Hyperlink
                            $hl.NavigateUri = [uri]$lnk.Url
                            $hl.Add_RequestNavigate({ Start-Process $_.Uri.AbsoluteUri })
                            $hl.Inlines.Add($lnk.Name) | Out-Null
                            $rt = New-Object System.Windows.Controls.TextBlock
                            $rt.Inlines.Add($hl) | Out-Null
                            $rt.Margin    = "0,0,12,0"
                            $rt.FontSize  = 11
                            $linkPanel.Children.Add($rt) | Out-Null
                        }
                        $sp.Children.Add($linkPanel) | Out-Null
                    }

                    $bd.Child = $sp
                    $Container.Children.Add($bd) | Out-Null
                }
            }

            # Update tray icon
            try {
                $icoPath = if ($hasUnread -and (Test-Path $IconAlert)) { $IconAlert } elseif (Test-Path $IconNormal) { $IconNormal } else { $null }
                if ($icoPath) { $TrayIcon.Icon = New-Object System.Drawing.Icon($icoPath) }
            } catch {}
        })
    } -Parameters @{
        Dispatcher     = $Dispatcher;     Container  = $Container
        ContentDataUrl = $ContentDataUrl; GitHubToken = $GitHubToken
        TrayIcon       = $TrayIcon
        IconAlert      = $IconAlert;      IconNormal = $IconNormal
    }
}

# ---- Async section: BigFix ---------------------------------------------------
function Start-BigFixLoad {
    param($Dispatcher, $Container, $ScriptDir)

    Start-TrackedRunspace -Script {
        param($Dispatcher, $Container, $ScriptDir)
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        # Read BigFix update data from text file (populated by BigFix Fixlet)
        $fixletPath = "C:\temp\X-Fixlet-Source_Count.txt"
        $updates = @()
        $fileFound = $false
        $fileTime = $null
        try {
            if (Test-Path $fixletPath) {
                $fileFound = $true
                $fileTime = (Get-Item $fixletPath).LastWriteTime
                $lines = Get-Content -Path $fixletPath -ErrorAction Stop
                foreach ($line in $lines) {
                    $trimmed = $line.Trim()
                    if ($trimmed.Length -gt 0) { $updates += $trimmed }
                }
            }
        } catch {}

        # Check for BigFix Self Service Application
        $ssaPath = "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe"
        $besUIPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClientUI.exe"
        $hasSSA = Test-Path $ssaPath
        $hasBESUI = Test-Path $besUIPath

        $Dispatcher.Invoke([Action]{
            $Container.Children.Clear()
            $mkB = { param($c) [System.Windows.Media.BrushConverter]::new().ConvertFrom($c) }
            $mkT = { param($t,$c='#64748B',$s=12) $tb=New-Object System.Windows.Controls.TextBlock; $tb.Text=$t; $tb.Foreground=& $mkB $c; $tb.FontSize=$s; $tb.Margin="4,4,0,4"; $tb.TextWrapping="Wrap"; $tb }

            if ($updates.Count -gt 0) {
                # Show update count header
                $Container.Children.Add((& $mkT "$($updates.Count) application update(s) available" '#D97706' 12)) | Out-Null

                # Show each update as a card
                foreach ($u in $updates) {
                    $bd = New-Object System.Windows.Controls.Border
                    $bd.Background = & $mkB "#FFFBEB"; $bd.BorderBrush = & $mkB "#D97706"
                    $bd.BorderThickness = "3,0,0,0"; $bd.CornerRadius = "6"; $bd.Margin = "0,0,0,6"; $bd.Padding = "12,8"
                    $bd.Child = (& $mkT $u '#1E293B' 12)
                    $Container.Children.Add($bd) | Out-Null
                }

                # Show last updated timestamp
                if ($null -ne $fileTime) {
                    $age = [math]::Round(([datetime]::Now - $fileTime).TotalHours, 1)
                    $tsColor = if ($age -gt 24) { '#DC2626' } else { '#64748B' }
                    $Container.Children.Add((& $mkT "Last refreshed: $($fileTime.ToString('M/d/yyyy h:mm tt')) ($age hrs ago)" $tsColor 10)) | Out-Null
                }

                # Open Self Service button
                if ($hasSSA) {
                    $btn = New-Object System.Windows.Controls.Button
                    $btn.Content = "Open Self Service  —  Install Updates"; $btn.Background = & $mkB "#2563EB"; $btn.Foreground = & $mkB "#FFF"
                    $btn.BorderThickness = "0"; $btn.Padding = "14,6"; $btn.Margin = "4,8,0,4"; $btn.Cursor = [System.Windows.Input.Cursors]::Hand
                    $btn.HorizontalAlignment = "Left"; $btn.FontWeight = "SemiBold"
                    $btn.Add_Click({ try { Start-Process "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe" } catch { [System.Windows.MessageBox]::Show("Could not launch Self Service: $_","Endpoint Advisor","OK","Warning") } })
                    $Container.Children.Add($btn) | Out-Null
                } elseif ($hasBESUI) {
                    $btn = New-Object System.Windows.Controls.Button
                    $btn.Content = "Open BigFix Client UI"; $btn.Background = & $mkB "#2563EB"; $btn.Foreground = & $mkB "#FFF"
                    $btn.BorderThickness = "0"; $btn.Padding = "14,6"; $btn.Margin = "4,8,0,4"; $btn.Cursor = [System.Windows.Input.Cursors]::Hand
                    $btn.HorizontalAlignment = "Left"; $btn.FontWeight = "SemiBold"
                    $btn.Add_Click({ try { Start-Process "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClientUI.exe" } catch { [System.Windows.MessageBox]::Show("Could not launch BigFix Client UI: $_","Endpoint Advisor","OK","Warning") } })
                    $Container.Children.Add($btn) | Out-Null
                }
            } else {
                # No updates or file not found — show nothing
            }
        })
    } -Parameters @{ Dispatcher=$Dispatcher; Container=$Container; ScriptDir=$ScriptDir }
}

# ---- Async section: Windows Update -------------------------------------------
function Start-WULoad {
    param($Dispatcher, $Container, $ToastFlags)

    Start-TrackedRunspace -Script {
        param($Dispatcher, $Container, $ToastFlags)
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        $updates    = @()
        $ccmMissing = $false
        try {
            $raw = Get-CimInstance -Namespace ROOT\ccm\ClientSDK -ClassName CCM_SoftwareUpdate -OperationTimeoutSec 30 -ErrorAction Stop |
                   Where-Object { $_.ComplianceState -eq 0 }
            if ($raw) {
                foreach ($u in $raw) {
                    $rebootNeeded = $false
                    try {
                        # Check multiple reboot indicators from CCM
                        $rebootNeeded = [bool]$u.IsRebootPending -or
                            $u.EvaluationState -in @(8,9,10) -or
                            $u.RebootOutsideServiceWindow -eq $true -or
                            ($u.OverrideServiceWindows -ne $null) -or
                            ($u.Name -match 'Cumulative Update|Security Update|Servicing Stack')
                    } catch {
                        # If properties don't exist, flag cumulative/security updates as likely needing reboot
                        try { $rebootNeeded = $u.Name -match 'Cumulative Update|Security Update|Servicing Stack' } catch {}
                    }
                    $title = $u.Name
                    if ($rebootNeeded) { $title += " (restart required)" }
                    $updates += [PSCustomObject]@{
                        Title    = $title
                        Deadline = if ($u.Deadline) { [System.Management.ManagementDateTimeConverter]::ToDateTime($u.Deadline) } else { $null }
                    }
                }
            }
        } catch {
            $ccmMissing = $true
        }

        $Dispatcher.Invoke([Action]{
            $Container.Children.Clear()
            $mkB = { param($c) [System.Windows.Media.BrushConverter]::new().ConvertFrom($c) }
            $mkT = { param($t,$c='#94A3B8',$s=12) $tb=New-Object System.Windows.Controls.TextBlock; $tb.Text=$t; $tb.Foreground=& $mkB $c; $tb.FontSize=$s; $tb.Margin="4,4,0,4"; $tb.TextWrapping="Wrap"; $tb }

            if ($ccmMissing) {
                $Container.Children.Add((& $mkT "ECM client not found on this device.")) | Out-Null
            } elseif ($updates.Count -eq 0) {
                $Container.Children.Add((& $mkT "All OS patches are up to date." '#059669')) | Out-Null
            } else {
                $Container.Children.Add((& $mkT "$($updates.Count) patch(es) pending — please install via Software Center." '#D97706' 12)) | Out-Null

                foreach ($u in $updates) {
                    $bd = New-Object System.Windows.Controls.Border
                    $bd.BorderBrush=& $mkB "#D97706"; $bd.BorderThickness="3,0,0,0"; $bd.CornerRadius="6"; $bd.Margin="0,0,0,6"; $bd.Padding="10,8"
                    $bd.Background=& $mkB "#FFFBEB"
                    $sp = New-Object System.Windows.Controls.StackPanel
                    $sp.Children.Add((& $mkT $u.Title '#1E293B' 12)) | Out-Null
                    if ($u.Deadline) {
                        $deadlineColor = if ($u.Deadline -lt [datetime]::Now.AddDays(3)) { '#DC2626' } else { '#64748B' }
                        $sp.Children.Add((& $mkT "Deadline: $($u.Deadline.ToString('MMM d, yyyy h:mm tt'))" $deadlineColor 11)) | Out-Null
                    }
                    $bd.Child = $sp
                    $Container.Children.Add($bd) | Out-Null
                }
                $btn = New-Object System.Windows.Controls.Button
                $btn.Content="Open Software Center"; $btn.Background=& $mkB "#D97706"; $btn.Foreground=& $mkB "#FFFFFF"
                $btn.BorderThickness="0"; $btn.Padding="14,6"; $btn.Margin="0,8,0,0"; $btn.Cursor=[System.Windows.Input.Cursors]::Hand; $btn.HorizontalAlignment="Left"; $btn.FontWeight="SemiBold"
                $btn.Add_Click({
                    try {
                        $p = "$env:WinDir\CCM\SCClient.exe"
                        if (Test-Path $p) { Start-Process $p } else { Start-Process "softwarecenter:" }
                    } catch { Start-Process "softwarecenter:" }
                })
                $Container.Children.Add($btn) | Out-Null
            }
        })

        # Set toast flags for the main thread to pick up via shared hashtable
        $restartUpdates = @($updates | Where-Object { $_.Title -match 'restart required' })
        $allPending = @($updates)

        # Debug: log what we found
        [System.IO.File]::AppendAllText("C:\temp\toast-debug.log",
            "$(Get-Date) | Updates: $($allPending.Count) | Restart: $($restartUpdates.Count) | ToastFlags ref: $($null -ne $ToastFlags)`r`n")

        if ($restartUpdates.Count -gt 0) {
            $ToastFlags.RestartCount = $restartUpdates.Count
        }
        # For non-reboot updates, toast once per day
        $noRebootUpdates = @($updates | Where-Object { $_.Title -notmatch 'restart required' })
        if ($noRebootUpdates.Count -gt 0) {
            $todayKey = [datetime]::Now.ToString('yyyy-MM-dd')
            $lastPatchToast = [Environment]::GetEnvironmentVariable('EA_LAST_PATCH_TOAST', 'User')
            if ($lastPatchToast -ne $todayKey) {
                $ToastFlags.PatchCount = $noRebootUpdates.Count
                [Environment]::SetEnvironmentVariable('EA_LAST_PATCH_TOAST', $todayKey, 'User')
            }
        }
    } -Parameters @{ Dispatcher=$Dispatcher; Container=$Container; ToastFlags=$Script:ToastFlags }
}

# ---- Async section: Support --------------------------------------------------
function Start-SupportLoad {
    param($Dispatcher, $Container, $ContentDataUrl, $GitHubToken)

    Start-TrackedRunspace -Script {
        param($Dispatcher, $Container, $ContentDataUrl, $GitHubToken)
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        $support = @()
        try {
            $headers = @{ 'User-Agent' = 'EndpointAdvisor' }
            if ($GitHubToken) { $headers['Authorization'] = "token $GitHubToken" }
            $params = @{ Uri = $ContentDataUrl; UseBasicParsing = $true; TimeoutSec = 20; Headers = $headers; ErrorAction = 'Stop' }
            $raw    = Invoke-WebRequest @params
            $data   = $raw.Content | ConvertFrom-Json
            if ($data.Data.Support) { $support = @($data.Data.Support) }
        } catch {}

        $Dispatcher.Invoke([Action]{
            $Container.Children.Clear()
            $mkB = { param($c) [System.Windows.Media.BrushConverter]::new().ConvertFrom($c) }
            $mkT = { param($t,$c='#94A3B8',$s=12) $tb=New-Object System.Windows.Controls.TextBlock; $tb.Text=$t; $tb.Foreground=& $mkB $c; $tb.FontSize=$s; $tb.Margin="4,4,0,4"; $tb.TextWrapping="Wrap"; $tb }

            if ($support.Count -eq 0) {
                $Container.Children.Add((& $mkT "No support info configured.")) | Out-Null
            } else {
                foreach ($s in $support) {
                    $bd = New-Object System.Windows.Controls.Border
                    $bd.Background=& $mkB "#F5F3FF"; $bd.BorderBrush=& $mkB "#8B5CF6"
                    $bd.BorderThickness="3,0,0,0"; $bd.CornerRadius="6"; $bd.Margin="0,0,0,8"; $bd.Padding="14,10"
                    $sp = New-Object System.Windows.Controls.StackPanel
                    if ($s.Text) { $sp.Children.Add((& $mkT $s.Text '#1E293B' 13)) | Out-Null }
                    if ($s.Details) {
                        $lines = $s.Details -replace '\\n',"`n"
                        $sp.Children.Add((& $mkT $lines '#94A3B8' 12)) | Out-Null
                    }
                    if ($s.Links -and @($s.Links).Count -gt 0) {
                        $lp = New-Object System.Windows.Controls.WrapPanel; $lp.Margin="0,6,0,0"
                        foreach ($lnk in @($s.Links)) {
                            try {
                                $hl = New-Object System.Windows.Documents.Hyperlink
                                $hl.NavigateUri = [uri]$lnk.Url
                                $hl.Add_RequestNavigate({ Start-Process $_.Uri.AbsoluteUri })
                                $hl.Inlines.Add($lnk.Name) | Out-Null
                                $rt = New-Object System.Windows.Controls.TextBlock
                                $rt.Inlines.Add($hl) | Out-Null
                                $rt.Margin="0,0,12,0"; $rt.FontSize=11
                                $lp.Children.Add($rt) | Out-Null
                            } catch {}
                        }
                        $sp.Children.Add($lp) | Out-Null
                    }
                    $bd.Child = $sp
                    $Container.Children.Add($bd) | Out-Null
                }
            }
        })
    } -Parameters @{ Dispatcher=$Dispatcher; Container=$Container; ContentDataUrl=$ContentDataUrl; GitHubToken=$GitHubToken }
}

# ---- Async section: Account Information --------------------------------------
function Start-AccountLoad {
    param($Dispatcher, $Container)

    Start-TrackedRunspace -Script {
        param($Dispatcher, $Container)
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        $accountName  = $env:USERNAME
        $displayName  = ""
        $pwdExpiryStr = "Unknown"
        $pwdDaysLeft  = $null
        $pwdExpired   = $false

        try {
            # Pull AD info via ADSI (no AD module required)
            $searcher = [adsisearcher]"(samaccountname=$accountName)"
            $searcher.PropertiesToLoad.AddRange(@("displayName","pwdLastSet","userAccountControl","msDS-UserPasswordExpiryTimeComputed"))
            $result = $searcher.FindOne()
            if ($result) {
                if ($result.Properties["displayName"].Count -gt 0) {
                    $displayName = $result.Properties["displayName"][0]
                }
                # msDS-UserPasswordExpiryTimeComputed is most reliable — it accounts for fine-grained policies
                $expiryProp = $result.Properties["msds-userpasswordexpirytimecomputed"]
                if ($expiryProp.Count -gt 0) {
                    $ft = $expiryProp[0]
                    if ($ft -gt 0 -and $ft -ne [Int64]::MaxValue) {
                        $expiryDate  = [datetime]::FromFileTime($ft)
                        $pwdDaysLeft = [math]::Ceiling(($expiryDate - [datetime]::Now).TotalDays)
                        $pwdExpired  = $pwdDaysLeft -lt 0
                        $pwdExpiryStr = $expiryDate.ToString("MMMM d, yyyy")
                    } else {
                        $pwdExpiryStr = "Never expires"
                    }
                }
            }
        } catch {}

        # --- Certificate checks ---
        $certRows = @()
        $alertDays = 14

        # YubiKey PIV certificates
        $ykmanPath = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
        try {
            if (Test-Path $ykmanPath) {
                $ykInfo = & $ykmanPath info 2>$null
                if ($ykInfo) {
                    foreach ($slot in @("9a", "9c", "9e")) {
                        $certPem = & $ykmanPath "piv" "certificates" "export" $slot "-" 2>$null
                        if ($certPem -and ($certPem -join "`n") -match "-----BEGIN CERTIFICATE-----") {
                            $tempFile = [System.IO.Path]::GetTempFileName()
                            ($certPem -join "`n") | Out-File $tempFile -Encoding ASCII
                            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tempFile)
                            Remove-Item $tempFile -Force
                            $daysLeft = [math]::Ceiling(($cert.NotAfter - [datetime]::Now).TotalDays)
                            $color = if ($daysLeft -lt 0) { "#DC2626" } elseif ($daysLeft -le $alertDays) { "#D97706" } else { "#059669" }
                            $label = if ($daysLeft -lt 0) { "EXPIRED" } else { "$($cert.NotAfter.ToString('MMM d, yyyy')) ($daysLeft days)" }
                            $certRows += ,@("YubiKey (Slot $slot):", $label, $color)
                        }
                    }
                }
            }
        } catch {}

        # Virtual Smart Card certificate (avoid accessing PrivateKey to prevent PIN prompt)
        try {
            $vscCert = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object {
                $_.HasPrivateKey -and ($_.Subject -match "Virtual" -or $_.Issuer -match "Smart Card" -or $_.Issuer -match "Virtual")
            } | Sort-Object NotAfter -Descending | Select-Object -First 1
            if ($vscCert) {
                $daysLeft = [math]::Ceiling(($vscCert.NotAfter - [datetime]::Now).TotalDays)
                $color = if ($daysLeft -lt 0) { "#DC2626" } elseif ($daysLeft -le $alertDays) { "#D97706" } else { "#059669" }
                $label = if ($daysLeft -lt 0) { "EXPIRED" } else { "$($vscCert.NotAfter.ToString('MMM d, yyyy')) ($daysLeft days)" }
                $certRows += ,@("Smart Card Cert:", $label, $color)
            }
        } catch {}

        # Email signing (S/MIME) certificate
        try {
            $emailCert = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object {
                $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.4"  # Secure Email OID
            } | Sort-Object NotAfter -Descending | Select-Object -First 1
            if ($emailCert) {
                $daysLeft = [math]::Ceiling(($emailCert.NotAfter - [datetime]::Now).TotalDays)
                $color = if ($daysLeft -lt 0) { "#DC2626" } elseif ($daysLeft -le $alertDays) { "#D97706" } else { "#059669" }
                $label = if ($daysLeft -lt 0) { "EXPIRED" } else { "$($emailCert.NotAfter.ToString('MMM d, yyyy')) ($daysLeft days)" }
                $certRows += ,@("Email Signing Cert:", $label, $color)
            }
        } catch {}

        $Dispatcher.Invoke([Action]{
            $Container.Children.Clear()
            $mkB = { param($c) [System.Windows.Media.BrushConverter]::new().ConvertFrom($c) }

            $bd = New-Object System.Windows.Controls.Border
            $bd.Background    = & $mkB "#F8FAFC"
            $bd.BorderBrush   = & $mkB "#E2E8F0"
            $bd.BorderThickness = "1"
            $bd.CornerRadius  = "6"
            $bd.Padding       = "14,12"
            $bd.Margin        = "0,0,0,8"

            $grid = New-Object System.Windows.Controls.Grid
            $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = "Auto"
            $col2 = New-Object System.Windows.Controls.ColumnDefinition; $col2.Width = "*"
            $grid.ColumnDefinitions.Add($col1); $grid.ColumnDefinitions.Add($col2)

            # Build account rows as an array of [label, value, color]
            $nameVal = if ($displayName) { "$displayName ($accountName)" } else { $accountName }
            $rows = @(
                ,@("Account:", $nameVal, "#1E293B")
            )
            if ($pwdExpired) {
                $rows += ,@("Password Expires:", "EXPIRED — please change your password", "#DC2626")
            } elseif ($pwdExpiryStr -eq "Never expires") {
                $rows += ,@("Password Expires:", "Never expires", "#059669")
            } elseif ($null -ne $pwdDaysLeft) {
                $expiryColor = if ($pwdDaysLeft -le 7) { "#DC2626" } elseif ($pwdDaysLeft -le 14) { "#D97706" } else { "#1E293B" }
                $expiryLabel = if ($pwdDaysLeft -le 14) { "$pwdExpiryStr ($pwdDaysLeft days)" } else { $pwdExpiryStr }
                $rows += ,@("Password Expires:", $expiryLabel, $expiryColor)
            } else {
                $rows += ,@("Password Expires:", $pwdExpiryStr, "#1E293B")
            }

            # Add certificate rows
            foreach ($cr in $certRows) { $rows += ,$cr }

            foreach ($r in $rows) {
                $row = New-Object System.Windows.Controls.RowDefinition; $row.Height = "Auto"
                $grid.RowDefinitions.Add($row)
                $rowIdx = $grid.RowDefinitions.Count - 1

                $lbl = New-Object System.Windows.Controls.TextBlock
                $lbl.Text = $r[0]; $lbl.Foreground = & $mkB "#64748B"
                $lbl.FontSize = 12; $lbl.Margin = "0,4,16,4"; $lbl.FontWeight = "SemiBold"
                [System.Windows.Controls.Grid]::SetRow($lbl, $rowIdx)
                [System.Windows.Controls.Grid]::SetColumn($lbl, 0)
                $grid.Children.Add($lbl) | Out-Null

                $val = New-Object System.Windows.Controls.TextBlock
                $val.Text = $r[1]; $val.Foreground = & $mkB $r[2]
                $val.FontSize = 12; $val.Margin = "0,4,0,4"
                [System.Windows.Controls.Grid]::SetRow($val, $rowIdx)
                [System.Windows.Controls.Grid]::SetColumn($val, 1)
                $grid.Children.Add($val) | Out-Null
            }

            $bd.Child = $grid
            $Container.Children.Add($bd) | Out-Null
        })
    } -Parameters @{ Dispatcher=$Dispatcher; Container=$Container }
}

# ---- Dashboard ---------------------------------------------------------------
function Show-Dashboard {
    $Script:TrayIcon.Icon = Get-TrayIcon $false

    if ($Script:DashboardWindow -and $Script:DashboardWindow.IsVisible) {
        $Script:DashboardWindow.Activate(); return
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Endpoint Advisor" Width="620" Height="740"
        WindowStartupLocation="CenterScreen" Background="#F1F5F9" ShowInTaskbar="False">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="Expander">
            <Setter Property="Foreground" Value="#64748B"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize"   Value="11"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Grid.Row="0" Background="#4A90D9" Padding="16,12">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Endpoint Advisor" FontSize="18" FontWeight="Bold" Foreground="#FFFFFF" VerticalAlignment="Center"/>
                <TextBlock x:Name="HostLabel" FontSize="11" Foreground="#DBEAFE" Margin="12,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
        </Border>
        <ScrollViewer Grid.Row="1" Margin="12" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="MainPanel"/>
        </ScrollViewer>
    </Grid>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    if (Test-Path $Script:IconNormal) {
        try {
            $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
            $bmp.BeginInit(); $bmp.UriSource = New-Object System.Uri($Script:IconNormal,[System.UriKind]::Absolute); $bmp.EndInit()
            $window.Icon = $bmp
        } catch {}
    }

    $window.FindName("HostLabel").Text = $env:COMPUTERNAME

    $panel = $window.FindName("MainPanel")

    # Announcements
    $panel.Children.Add((New-SectionHeader "Announcements")) | Out-Null
    $annPanel = New-Object System.Windows.Controls.StackPanel
    $annPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($annPanel) | Out-Null

    # Software Updates (BigFix)
    $panel.Children.Add((New-SectionHeader "BigFix Software Updates")) | Out-Null
    $swPanel = New-Object System.Windows.Controls.StackPanel
    $swPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($swPanel) | Out-Null

    # Pending Windows Updates
    $panel.Children.Add((New-SectionHeader "System Patch Updates")) | Out-Null
    $wuPanel = New-Object System.Windows.Controls.StackPanel
    $wuPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($wuPanel) | Out-Null

    # Support
    $panel.Children.Add((New-SectionHeader "Support")) | Out-Null
    $supPanel = New-Object System.Windows.Controls.StackPanel
    $supPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($supPanel) | Out-Null

    # Account Information
    $panel.Children.Add((New-SectionHeader "Account Information")) | Out-Null
    $acctPanel = New-Object System.Windows.Controls.StackPanel
    $acctPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($acctPanel) | Out-Null

    # Windows Version footer
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $build = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue)
    $displayVersion = $build.DisplayVersion  # e.g. "25H2"
    $productName = $osInfo.Caption -replace 'Microsoft ', ''  # e.g. "Windows 11 Enterprise"
    $versionText = "$productName Build: $displayVersion"
    $verBlock = New-Object System.Windows.Controls.TextBlock
    $verBlock.Text = $versionText
    $verBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#94A3B8")
    $verBlock.FontSize = 10; $verBlock.Margin = "0,16,0,2"; $verBlock.HorizontalAlignment = "Center"
    $panel.Children.Add($verBlock) | Out-Null

    $eaVerBlock = New-Object System.Windows.Controls.TextBlock
    $eaVerBlock.Text = "Endpoint Advisor v$($Script:EAVersion)"
    $eaVerBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#64748B")
    $eaVerBlock.FontSize = 9; $eaVerBlock.Margin = "0,0,0,4"; $eaVerBlock.HorizontalAlignment = "Center"
    $panel.Children.Add($eaVerBlock) | Out-Null

    $Script:DashboardWindow = $window
    $window.Show()

    $dispatcher = $window.Dispatcher
    Start-AnnouncementsLoad -Dispatcher $dispatcher -Container $annPanel `
        -ContentDataUrl $Config.ContentDataUrl -GitHubToken $Config.GitHubToken `
        -TrayIcon $Script:TrayIcon -IconAlert $Script:IconAlert -IconNormal $Script:IconNormal
    Start-BigFixLoad   -Dispatcher $dispatcher -Container $swPanel -ScriptDir $ScriptDir
    Start-WULoad       -Dispatcher $dispatcher -Container $wuPanel -ToastFlags $Script:ToastFlags
    Start-SupportLoad  -Dispatcher $dispatcher -Container $supPanel `
        -ContentDataUrl $Config.ContentDataUrl -GitHubToken $Config.GitHubToken
    Start-AccountLoad  -Dispatcher $dispatcher -Container $acctPanel
}

# ---- Polling / refresh -------------------------------------------------------
function Start-ContentRefresh {
    $data = Get-ContentData
    if (-not $data) { return }
    $newVersion = $data.contentVersion

    if ($newVersion -ne $Script:ContentVersion) {
        $Script:ContentVersion = $newVersion
        $relevant = Get-RelevantAnnouncements $data

        $now = [datetime]::UtcNow
        foreach ($a in $relevant) {
            $shouldNotify = $false
            if (-not $Script:SeenIds.Contains($a.id)) {
                $shouldNotify = $true
                [void]$Script:SeenIds.Add($a.id)
            } elseif ($a.Priority -eq 'critical' -and $a.NagEnabled -and $a.PSObject.Properties['NagEnabled']) {
                $interval = if ($a.NagIntervalMinutes) { $a.NagIntervalMinutes } else { 30 }
                $last = $Script:LastNagTime[$a.id]
                if (-not $last -or ($now - $last).TotalMinutes -ge $interval) { $shouldNotify = $true }
            }
            if ($shouldNotify) {
                $Script:LastNagTime[$a.id] = $now
                # Only toast/nag for critical announcements — info/warning show in dashboard only
                if ($a.Priority -eq 'critical') {
                    $snippet = if ($a.Text) { $a.Text.Substring(0, [Math]::Min(200, $a.Text.Length)) } else { "" }
                    Show-Toast $a.Title $snippet $a.Priority
                }
            }
        }

        $hasUnread = $relevant | Where-Object { -not $Script:SeenIds.Contains($_.id) }
        if ($hasUnread) { $Script:TrayIcon.Icon = Get-TrayIcon $true }
    }

    $Script:CachedContent = $data
    Write-Log "Refresh complete (v$newVersion)"
}

# ---- Main --------------------------------------------------------------------
$Script:DashboardWindow = $null
$Script:TrayIcon        = New-TrayIcon

# Periodic refresh timer
$timer          = New-Object System.Windows.Forms.Timer
$timer.Interval = $Config.RefreshInterval * 1000
$timer.Add_Tick({ Start-ContentRefresh })
$timer.Start()

# Shared toast flags — synchronized hashtable accessible from runspaces
$Script:ToastFlags = [hashtable]::Synchronized(@{ RestartCount = 0; PatchCount = 0 })
$toastTimer = New-Object System.Windows.Forms.Timer
$toastTimer.Interval = 5000
$toastTimer.Add_Tick({
    if ($Script:ToastFlags.RestartCount -gt 0) {
        $count = $Script:ToastFlags.RestartCount
        $Script:ToastFlags.RestartCount = 0
        Show-Toast "System Restart Required" "$count update(s) require a restart. Please open Software Center to apply the update." "critical"
    }
    if ($Script:ToastFlags.PatchCount -gt 0) {
        $count = $Script:ToastFlags.PatchCount
        $Script:ToastFlags.PatchCount = 0
        Show-Toast "Software Updates Available" "$count update(s) available in Software Center. Please install at your earliest convenience." "info"
    }
})
$toastTimer.Start()

# Deferred startup refresh - fires 1s after message pump starts so UI never blocks
$startupTimer          = New-Object System.Windows.Forms.Timer
$startupTimer.Interval = 1000
$startupTimer.Add_Tick({
    $startupTimer.Stop()
    $startupTimer.Dispose()
    Start-ContentRefresh
})
$startupTimer.Start()

Write-Log "Endpoint Advisor started on $Hostname (ScriptDir=$ScriptDir)"

[System.Windows.Forms.Application]::Run()
