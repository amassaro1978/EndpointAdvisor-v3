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

# ---- CHANGE THIS to your company's registry path ----------------------------
# Group targeting reads from: HKLM:\SOFTWARE\<CompanyRegPath>\targeting\GROUP
# Set this to match your organization's registry namespace (e.g. "Contoso", "AcmeCorp")
$Script:CompanyRegPath = "CompanyName"

# ---- CHANGE THIS to your company's display name for branding -----------------
# This appears in the window title, footer, toasts, tray tooltip, etc.
# Example: "Acme Corp" → shows as "Acme Corp Endpoint Advisor"
$Script:CompanyBrand = ""   # Leave empty for just "Endpoint Advisor"
$Script:BrandedName  = if ($Script:CompanyBrand) { "$($Script:CompanyBrand) Endpoint Advisor" } else { "Endpoint Advisor" }

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

# Read targeting group from registry (set by IT via GPO or deployment)
$Script:DeviceGroup = ""
try {
    $regGroup = Get-ItemProperty -Path "HKLM:\SOFTWARE\$($Script:CompanyRegPath)\targeting" -Name "GROUP" -ErrorAction SilentlyContinue
    if ($regGroup -and $regGroup.GROUP) { $Script:DeviceGroup = $regGroup.GROUP }
} catch {}

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

        # Group targeting — skip if announcement targets a specific group and we're not in it
        if ($item.TargetGroup -and $item.TargetGroup.Trim() -ne "" -and $item.TargetGroup -ne "All") {
            if ($Script:DeviceGroup -ne $item.TargetGroup) { continue }
        }

        # Registry key targeting — skip if key/value doesn't match on this machine
        if ($item.TargetRegKey -and $item.TargetRegKey.Trim() -ne "") {
            $regMatch = $false
            try {
                $regPath = $item.TargetRegKey.Trim()
                # Normalize: accept both HKLM\... and Registry::HKLM\... formats
                if ($regPath -notmatch '^Registry::') {
                    $regPath = "Registry::$regPath"
                }
                if (Test-Path $regPath) {
                    if ($item.TargetRegValue -and $item.TargetRegValue.Trim() -ne "") {
                        # Check specific value
                        $val = Get-ItemProperty -Path $regPath -Name $item.TargetRegValue.Trim() -ErrorAction SilentlyContinue
                        if ($val) {
                            $actualData = $val.($item.TargetRegValue.Trim())
                            if ($item.TargetRegData -and $item.TargetRegData.Trim() -ne "") {
                                # Match value data (contains)
                                $regMatch = "$actualData" -like "*$($item.TargetRegData.Trim())*"
                            } else {
                                # Value exists, no data filter
                                $regMatch = $true
                            }
                        }
                    } else {
                        # Key exists, no value check needed
                        $regMatch = $true
                    }
                }
            } catch {}
            if (-not $regMatch) { continue }
        }

        # Conditional: password_expiry — only show if user password expires within threshold
        if ($item.Condition -eq "password_expiry") {
            $thresholdDays = if ($item.ConditionThresholdDays) { [int]$item.ConditionThresholdDays } else { 14 }
            $daysLeft = $null
            try {
                $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
                $searcher.PropertiesToLoad.Add("msDS-UserPasswordExpiryTimeComputed") | Out-Null
                $adResult = $searcher.FindOne()
                if ($adResult) {
                    $expiryProp = $adResult.Properties["msds-userpasswordexpirytimecomputed"]
                    if ($expiryProp.Count -gt 0) {
                        $ft = $expiryProp[0]
                        if ($ft -gt 0 -and $ft -ne [Int64]::MaxValue) {
                            $expiryDate = [datetime]::FromFileTime($ft)
                            $daysLeft   = [math]::Ceiling(($expiryDate - [datetime]::Now).TotalDays)
                        }
                    }
                }
            } catch {}
            # Only include if daysLeft is known and within threshold
            if ($null -eq $daysLeft -or $daysLeft -gt $thresholdDays) { continue }
        }

        # Conditional: cert_expiry — only show if any certificate expires within threshold
        if ($item.Condition -eq "cert_expiry") {
            $thresholdDays = if ($item.ConditionThresholdDays) { [int]$item.ConditionThresholdDays } else { 14 }
            $certExpiringSoon = $false
            try {
                $allCerts = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object { $_.HasPrivateKey }
                foreach ($c in $allCerts) {
                    $d = [math]::Ceiling(($c.NotAfter - [datetime]::Now).TotalDays)
                    if ($d -ge 0 -and $d -le $thresholdDays) { $certExpiringSoon = $true; break }
                }
            } catch {}
            # Also check YubiKey certs if ykman is available
            if (-not $certExpiringSoon) {
                $ykPath = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
                if (Test-Path $ykPath) {
                    try {
                        foreach ($slot in @("9a", "9c", "9e")) {
                            $pem = & $ykPath "piv" "certificates" "export" $slot "-" 2>$null
                            if ($pem -and ($pem -join "`n") -match "-----BEGIN CERTIFICATE-----") {
                                $tmp = [System.IO.Path]::GetTempFileName()
                                ($pem -join "`n") | Out-File $tmp -Encoding ASCII
                                $ykCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tmp)
                                Remove-Item $tmp -Force
                                $d = [math]::Ceiling(($ykCert.NotAfter - [datetime]::Now).TotalDays)
                                if ($d -ge 0 -and $d -le $thresholdDays) { $certExpiringSoon = $true; break }
                            }
                        }
                    } catch {}
                }
            }
            if (-not $certExpiringSoon) { continue }
        }

        # Conditional: cert_expiry_email — only show if an email/signing cert expires within threshold
        if ($item.Condition -eq "cert_expiry_email") {
            $thresholdDays = if ($item.ConditionThresholdDays) { [int]$item.ConditionThresholdDays } else { 14 }
            $certExpiringSoon = $false
            $emailEkus = @('Secure Email','Email Protection')
            try {
                $allCerts = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object { $_.HasPrivateKey }
                foreach ($c in $allCerts) {
                    $ekus = $c.EnhancedKeyUsageList.FriendlyName
                    $isEmail = $emailEkus | Where-Object { $ekus -contains $_ }
                    if (-not $isEmail) { continue }
                    $d = [math]::Ceiling(($c.NotAfter - [datetime]::Now).TotalDays)
                    if ($d -ge 0 -and $d -le $thresholdDays) { $certExpiringSoon = $true; break }
                }
            } catch {}
            if (-not $certExpiringSoon) { continue }
        }

        # Conditional: cert_expiry_auth — only show if a client auth/smart card cert expires within threshold
        if ($item.Condition -eq "cert_expiry_auth") {
            $thresholdDays = if ($item.ConditionThresholdDays) { [int]$item.ConditionThresholdDays } else { 14 }
            $certExpiringSoon = $false
            $authEkus = @('Client Authentication','Smart Card Logon')
            try {
                $allCerts = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object { $_.HasPrivateKey }
                foreach ($c in $allCerts) {
                    $ekus = $c.EnhancedKeyUsageList.FriendlyName
                    $isAuth = $authEkus | Where-Object { $ekus -contains $_ }
                    if (-not $isAuth) { continue }
                    $d = [math]::Ceiling(($c.NotAfter - [datetime]::Now).TotalDays)
                    if ($d -ge 0 -and $d -le $thresholdDays) { $certExpiringSoon = $true; break }
                }
            } catch {}
            # Also check YubiKey PIV auth slot (9a)
            if (-not $certExpiringSoon) {
                $ykPath = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
                if (Test-Path $ykPath) {
                    try {
                        $pem = & $ykPath "piv" "certificates" "export" "9a" "-" 2>$null
                        if ($pem -and ($pem -join "`n") -match "-----BEGIN CERTIFICATE-----") {
                            $tmp = [System.IO.Path]::GetTempFileName()
                            ($pem -join "`n") | Out-File $tmp -Encoding ASCII
                            $ykCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tmp)
                            Remove-Item $tmp -Force
                            $d = [math]::Ceiling(($ykCert.NotAfter - [datetime]::Now).TotalDays)
                            if ($d -ge 0 -and $d -le $thresholdDays) { $certExpiringSoon = $true }
                        }
                    } catch {}
                }
            }
            if (-not $certExpiringSoon) { continue }
        }

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
    $tray.Text    = $Script:BrandedName
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
        New-ItemProperty -Path $regPath -Name "DisplayName" -Value $Script:BrandedName -PropertyType String -Force | Out-Null
    }
    return $appId
}
$Script:ToastAppId = Register-AppId

function Show-Toast($title, $message, $priority, [string]$toastType = "announcement") {
    # toastType: "announcement" | "update"
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]                       | Out-Null

        $escapedTitle   = [System.Security.SecurityElement]::Escape($title)
        $escapedMessage = [System.Security.SecurityElement]::Escape($message)

        if ($toastType -eq "update") {
            # System update toast — "Update Now", "Snooze", "Dismiss"
            $toastXml = @"
<toast scenario='reminder' duration='long'>
  <visual>
    <binding template='ToastGeneric'>
      <text>$escapedTitle</text>
      <text>$escapedMessage</text>
    </binding>
  </visual>
  <actions>
    <action content='Update Now' activationType='protocol' arguments='softwarecenter:'/>
    <action content='Snooze' activationType='system' arguments='snooze' hint-inputId='snoozeTime'/>
    <action content='Dismiss' activationType='system' arguments='dismiss'/>
  </actions>
</toast>
"@
        } elseif ($priority -eq 'critical' -or $priority -eq 'high') {
            # Persistent reminder-style toast for critical/high announcements
            $toastXml = @"
<toast scenario='reminder' duration='long'>
  <visual>
    <binding template='ToastGeneric'>
      <text>$escapedTitle</text>
      <text>$escapedMessage</text>
    </binding>
  </visual>
  <actions>
    <action content='Dismiss' activationType='background' arguments='acknowledge'/>
    <action content='Snooze' activationType='system' arguments='snooze' hint-inputId='snoozeTime'/>
  </actions>
</toast>
"@
        } else {
            # Normal (info/warning) toast — standard style, no persistent scenario
            $toastXml = @"
<toast>
  <visual>
    <binding template='ToastGeneric'>
      <text>$escapedTitle</text>
      <text>$escapedMessage</text>
    </binding>
  </visual>
</toast>
"@
        }

        $xml = [Windows.Data.Xml.Dom.XmlDocument]::new()
        $xml.LoadXml($toastXml)
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
    param($Dispatcher, $Container, $ContentDataUrl, $GitHubToken, $TrayIcon, $IconAlert, $IconNormal, $CompanyRegPath)

    Start-TrackedRunspace -Script {
        param($Dispatcher, $Container, $ContentDataUrl, $GitHubToken, $TrayIcon, $IconAlert, $IconNormal, $CompanyRegPath)

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

            # Read device group inside runspace (Script: scope not shared across runspaces)
            $deviceGroup = ""
            try {
                $rg = Get-ItemProperty -Path "HKLM:\SOFTWARE\$CompanyRegPath\targeting" -Name "GROUP" -ErrorAction SilentlyContinue
                if ($rg -and $rg.GROUP) { $deviceGroup = $rg.GROUP }
            } catch {}

            foreach ($item in $data.Data.Announcements.Default) {
                if (-not $item -or $item.Enabled -eq $false) { continue }
                if ($item.StartDate -and ([datetime]$item.StartDate) -gt $now) { continue }
                if ($item.EndDate   -and ([datetime]$item.EndDate)   -lt $now) { continue }

                # Group targeting — skip if announcement targets a specific group and we're not in it
                if ($item.TargetGroup -and $item.TargetGroup.Trim() -ne "" -and $item.TargetGroup -ne "All") {
                    if ($deviceGroup -ne $item.TargetGroup) { continue }
                }

                # Registry key targeting — skip if key/value doesn't match on this machine
                if ($item.TargetRegKey -and $item.TargetRegKey.Trim() -ne "") {
                    $regMatch = $false
                    try {
                        $regPath = $item.TargetRegKey.Trim()
                        if ($regPath -notmatch '^Registry::') { $regPath = "Registry::$regPath" }
                        if (Test-Path $regPath) {
                            if ($item.TargetRegValue -and $item.TargetRegValue.Trim() -ne "") {
                                $val = Get-ItemProperty -Path $regPath -Name $item.TargetRegValue.Trim() -ErrorAction SilentlyContinue
                                if ($val) {
                                    $actualData = $val.($item.TargetRegValue.Trim())
                                    if ($item.TargetRegData -and $item.TargetRegData.Trim() -ne "") {
                                        $regMatch = "$actualData" -like "*$($item.TargetRegData.Trim())*"
                                    } else { $regMatch = $true }
                                }
                            } else { $regMatch = $true }
                        }
                    } catch {}
                    if (-not $regMatch) { continue }
                }

                # Conditional: cert_expiry — any cert with private key expiring within threshold
                if ($item.Condition -eq "cert_expiry") {
                    $thresh = if ($item.ConditionThresholdDays) { [int]$item.ConditionThresholdDays } else { 14 }
                    $firing = $false
                    try {
                        foreach ($c in (Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object { $_.HasPrivateKey })) {
                            $d = [math]::Ceiling(($c.NotAfter - [datetime]::Now).TotalDays)
                            if ($d -ge 0 -and $d -le $thresh) { $firing = $true; break }
                        }
                    } catch {}
                    if (-not $firing) { continue }
                }

                # Conditional: cert_expiry_email — email/signing cert expiring within threshold
                if ($item.Condition -eq "cert_expiry_email") {
                    $thresh = if ($item.ConditionThresholdDays) { [int]$item.ConditionThresholdDays } else { 14 }
                    $firing = $false
                    $emailEkus = @('Secure Email','Email Protection')
                    try {
                        foreach ($c in (Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object { $_.HasPrivateKey })) {
                            $ekus = $c.EnhancedKeyUsageList.FriendlyName
                            if (-not ($emailEkus | Where-Object { $ekus -contains $_ })) { continue }
                            $d = [math]::Ceiling(($c.NotAfter - [datetime]::Now).TotalDays)
                            if ($d -ge 0 -and $d -le $thresh) { $firing = $true; break }
                        }
                    } catch {}
                    if (-not $firing) { continue }
                }

                # Conditional: cert_expiry_auth — client auth/smart card cert expiring within threshold
                if ($item.Condition -eq "cert_expiry_auth") {
                    $thresh = if ($item.ConditionThresholdDays) { [int]$item.ConditionThresholdDays } else { 14 }
                    $firing = $false
                    $authEkus = @('Client Authentication','Smart Card Logon')
                    try {
                        foreach ($c in (Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object { $_.HasPrivateKey })) {
                            $ekus = $c.EnhancedKeyUsageList.FriendlyName
                            if (-not ($authEkus | Where-Object { $ekus -contains $_ })) { continue }
                            $d = [math]::Ceiling(($c.NotAfter - [datetime]::Now).TotalDays)
                            if ($d -ge 0 -and $d -le $thresh) { $firing = $true; break }
                        }
                    } catch {}
                    if (-not $firing) { continue }
                }

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

                    # KB article link (from ConditionKbUrl field)
                    $kbUrl = $null
                    if ($a.PSObject.Properties['ConditionKbUrl'] -and $a.ConditionKbUrl) {
                        $kbUrl = $a.ConditionKbUrl
                    }
                    if ($kbUrl) {
                        try {
                            $hl = New-Object System.Windows.Documents.Hyperlink
                            $hl.NavigateUri = [uri]$kbUrl
                            $hl.Add_RequestNavigate({ Start-Process $_.Uri.AbsoluteUri })
                            $hl.Inlines.Add("Knowledge Base Article") | Out-Null
                            $hlTb = New-Object System.Windows.Controls.TextBlock
                            $hlTb.Inlines.Add($hl) | Out-Null
                            $hlTb.Margin = "4,4,0,4"; $hlTb.FontSize = 11
                            $sp.Children.Add($hlTb) | Out-Null
                        } catch {}
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
        CompanyRegPath = $CompanyRegPath
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
                    $btn.Add_Click({ try { Start-Process "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe" } catch { [System.Windows.MessageBox]::Show("Could not launch Self Service: $_",$Script:BrandedName,"OK","Warning") } })
                    $Container.Children.Add($btn) | Out-Null
                } elseif ($hasBESUI) {
                    $btn = New-Object System.Windows.Controls.Button
                    $btn.Content = "Open BigFix Client UI"; $btn.Background = & $mkB "#2563EB"; $btn.Foreground = & $mkB "#FFF"
                    $btn.BorderThickness = "0"; $btn.Padding = "14,6"; $btn.Margin = "4,8,0,4"; $btn.Cursor = [System.Windows.Input.Cursors]::Hand
                    $btn.HorizontalAlignment = "Left"; $btn.FontWeight = "SemiBold"
                    $btn.Add_Click({ try { Start-Process "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClientUI.exe" } catch { [System.Windows.MessageBox]::Show("Could not launch BigFix Client UI: $_",$Script:BrandedName,"OK","Warning") } })
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

        # Smart card cert detection — enumerate by reader to distinguish physical, virtual, and YubiKey
        try {
            $allUserCerts = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue
            $scCerts = $allUserCerts | Where-Object {
                $_.HasPrivateKey -and
                $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.4.1.311.20.2.2"  # Smart Card Logon only
            } | Sort-Object NotAfter -Descending

            # Detect reader types via WMI (no PIN prompt)
            $readers = @{}
            try {
                $scReaders = Get-WmiObject -Query "SELECT * FROM Win32_PnPEntity WHERE Caption LIKE '%smart card%' OR Caption LIKE '%Yubi%' OR Caption LIKE '%Virtual%'" -ErrorAction SilentlyContinue
                foreach ($r in $scReaders) {
                    if ($r.Caption -match "Virtual") { $readers["virtual"] = $true }
                    if ($r.Caption -match "Yubi") { $readers["yubikey"] = $true }
                }
            } catch {}

            # Track which cert types we've shown
            $shownPhysical = $false; $shownVirtual = $false; $shownYubiKey = $false

            foreach ($cert in $scCerts) {
                $daysLeft = [math]::Ceiling(($cert.NotAfter - [datetime]::Now).TotalDays)
                $color = if ($daysLeft -lt 0) { "#DC2626" } elseif ($daysLeft -le $alertDays) { "#D97706" } else { "#059669" }
                $label = if ($daysLeft -lt 0) { "EXPIRED" } else { "$($cert.NotAfter.ToString('MMM d, yyyy')) ($daysLeft days)" }

                # Determine cert type — check template name first (most reliable), then subject/issuer
                $templateName = ""
                try {
                    $tmplExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq "1.3.6.1.4.1.311.21.7" }
                    if ($tmplExt) { $templateName = $tmplExt.Format($false) }
                } catch {}
                $certInfo = "$($cert.Subject) $($cert.Issuer)"

                if ($templateName -match "(?i)virtual" -and -not $shownVirtual) {
                    $certRows += ,@("Virtual Smart Card:", $label, $color)
                    $shownVirtual = $true
                } elseif ($templateName -match "(?i)yubi|(?i)piv" -and -not $shownYubiKey) {
                    $certRows += ,@("YubiKey Cert:", $label, $color)
                    $shownYubiKey = $true
                } elseif ($certInfo -match "Yubi" -and $templateName -notmatch "(?i)virtual" -and -not $shownYubiKey) {
                    $certRows += ,@("YubiKey Cert:", $label, $color)
                    $shownYubiKey = $true
                } elseif ($certInfo -match "Virtual" -and -not $shownVirtual) {
                    $certRows += ,@("Virtual Smart Card:", $label, $color)
                    $shownVirtual = $true
                } elseif (-not $shownPhysical) {
                    $certRows += ,@("Smart Card Cert:", $label, $color)
                    $shownPhysical = $true
                }
            }

            # If we detected a virtual reader but no cert matched by name, show it separately
            if ($readers["virtual"] -and -not $shownVirtual -and $scCerts.Count -gt 1) {
                # The cert that isn't YubiKey and isn't the first physical one is likely virtual
                $remaining = $scCerts | Select-Object -Skip ([int]$shownPhysical + [int]$shownYubiKey) | Select-Object -First 1
                if ($remaining) {
                    $daysLeft = [math]::Ceiling(($remaining.NotAfter - [datetime]::Now).TotalDays)
                    $color = if ($daysLeft -lt 0) { "#DC2626" } elseif ($daysLeft -le $alertDays) { "#D97706" } else { "#059669" }
                    $label = if ($daysLeft -lt 0) { "EXPIRED" } else { "$($remaining.NotAfter.ToString('MMM d, yyyy')) ($daysLeft days)" }
                    $certRows += ,@("Virtual Smart Card:", $label, $color)
                }
            }
        } catch {}

        # Email signing (S/MIME) certificate — Digital Signature OID 1.3.6.1.5.5.7.3.4 (Secure Email)
        try {
            $emailSignCert = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object {
                $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.4"  # Secure Email / S/MIME
            } | Sort-Object NotAfter -Descending | Select-Object -First 1
            if ($emailSignCert) {
                $daysLeft = [math]::Ceiling(($emailSignCert.NotAfter - [datetime]::Now).TotalDays)
                $color = if ($daysLeft -lt 0) { "#DC2626" } elseif ($daysLeft -le $alertDays) { "#D97706" } else { "#059669" }
                $label = if ($daysLeft -lt 0) { "EXPIRED" } else { "$($emailSignCert.NotAfter.ToString('MMM d, yyyy')) ($daysLeft days)" }
                $certRows += ,@("Email Signing Cert:", $label, $color)
            }
        } catch {}

        # Email encryption certificate — Key Encipherment usage + Secure Email OID (1.3.6.1.5.5.7.3.4)
        # Separate from signing: look for Key Encipherment key usage flag (0x20 = 32)
        try {
            $emailEncCert = Get-ChildItem "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object {
                $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.4" -and
                $_.HasPrivateKey -and
                (($_.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension] } |
                    ForEach-Object { ($_.KeyUsages -band [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::KeyEncipherment) -ne 0 }) -contains $true)
            } | Sort-Object NotAfter -Descending | Select-Object -First 1

            # Show encryption cert even if same as signing cert (multi-purpose certs are common)
            if ($emailEncCert) {
                $daysLeft = [math]::Ceiling(($emailEncCert.NotAfter - [datetime]::Now).TotalDays)
                $color = if ($daysLeft -lt 0) { "#DC2626" } elseif ($daysLeft -le $alertDays) { "#D97706" } else { "#059669" }
                $label = if ($daysLeft -lt 0) { "EXPIRED" } else { "$($emailEncCert.NotAfter.ToString('MMM d, yyyy')) ($daysLeft days)" }
                $certRows += ,@("Email Encryption Cert:", $label, $color)
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

            # ---- Certificate rows (enriched with friendly names + status) ----
            $certMeta = @{
                "YubiKey (Slot 9a):"     = @("YubiKey (Slot 9A):", "Hardware token authentication")
                "YubiKey (Slot 9c):"     = @("YubiKey (Slot 9C):", "Hardware token digital signature")
                "YubiKey (Slot 9e):"     = @("YubiKey (Slot 9E):", "Hardware token card authentication")
                "YubiKey Cert:"          = @("YubiKey Certificate:", "Hardware token authentication")
                "Smart Card Cert:"       = @("Smart Card:", "Physical smart card login")
                "Virtual Smart Card:"    = @("Virtual Smart Card:", "Virtual smart card login")
                "Email Signing Cert:"    = @("Signing Certificate:", "Used for email signing")
                "Email Encryption Cert:" = @("Encryption Certificate:", "Used for email encryption")
            }

            foreach ($cr in $certRows) {
                $meta = $certMeta[$cr[0]]
                $friendlyLabel = if ($meta) { $meta[0] } else { $cr[0] }
                $subtitle      = if ($meta) { $meta[1] } else { $null }
                $isExpired = $cr[1] -eq "EXPIRED"
                $isWarning = $cr[2] -eq "#D97706"
                $statusTag = if ($isExpired) { "[EXPIRED]" } elseif ($isWarning) { "[WARNING]" } else { "[OK]" }
                $statusWord = if ($isExpired) { "EXPIRED" } elseif ($isWarning) { "EXPIRING SOON" } else { "VALID" }
                $expiryDetail = if ($isExpired) { "" } else { " — Expires $($cr[1])" }
                $enrichedValue = "$statusTag $statusWord$expiryDetail"

                # Label row (friendly name + subtitle)
                $row = New-Object System.Windows.Controls.RowDefinition; $row.Height = "Auto"
                $grid.RowDefinitions.Add($row)
                $rowIdx = $grid.RowDefinitions.Count - 1

                $lblStack = New-Object System.Windows.Controls.StackPanel
                $lbl = New-Object System.Windows.Controls.TextBlock
                $lbl.Text = $friendlyLabel; $lbl.Foreground = & $mkB "#64748B"
                $lbl.FontSize = 12; $lbl.FontWeight = "SemiBold"
                $lblStack.Children.Add($lbl) | Out-Null
                if ($subtitle) {
                    $sub = New-Object System.Windows.Controls.TextBlock
                    $sub.Text = $subtitle; $sub.Foreground = & $mkB "#94A3B8"
                    $sub.FontSize = 10
                    $lblStack.Children.Add($sub) | Out-Null
                }
                $lblStack.Margin = "0,4,16,4"
                [System.Windows.Controls.Grid]::SetRow($lblStack, $rowIdx)
                [System.Windows.Controls.Grid]::SetColumn($lblStack, 0)
                $grid.Children.Add($lblStack) | Out-Null

                $val = New-Object System.Windows.Controls.TextBlock
                $val.Text = $enrichedValue; $val.Foreground = & $mkB $cr[2]
                $val.FontSize = 12; $val.Margin = "0,4,0,4"; $val.VerticalAlignment = "Center"
                [System.Windows.Controls.Grid]::SetRow($val, $rowIdx)
                [System.Windows.Controls.Grid]::SetColumn($val, 1)
                $grid.Children.Add($val) | Out-Null
            }
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
        Title="$($Script:BrandedName)" Width="620" Height="740"
        WindowStartupLocation="CenterScreen" Background="#F1F5F9" ShowInTaskbar="False" ResizeMode="NoResize">
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
                <TextBlock Text="$($Script:BrandedName)" FontSize="18" FontWeight="Bold" Foreground="#FFFFFF" VerticalAlignment="Center"/>
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

    # Application Updates (BigFix)
    $panel.Children.Add((New-SectionHeader "Application Updates")) | Out-Null
    $swPanel = New-Object System.Windows.Controls.StackPanel
    $swPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($swPanel) | Out-Null

    # Microsoft Updates (Windows/CCM patches)
    $panel.Children.Add((New-SectionHeader "Microsoft Updates")) | Out-Null
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
    $eaVerBlock.Text = "$($Script:BrandedName) v$($Script:EAVersion)"
    $eaVerBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#64748B")
    $eaVerBlock.FontSize = 9; $eaVerBlock.Margin = "0,0,0,4"; $eaVerBlock.HorizontalAlignment = "Center"
    $panel.Children.Add($eaVerBlock) | Out-Null

    $Script:DashboardWindow = $window
    $Script:ToastFlags.Acknowledged = $true  # User opened dashboard — stop toasting
    $window.Show()

    $dispatcher = $window.Dispatcher
    Start-AnnouncementsLoad -Dispatcher $dispatcher -Container $annPanel `
        -ContentDataUrl $Config.ContentDataUrl -GitHubToken $Config.GitHubToken `
        -TrayIcon $Script:TrayIcon -IconAlert $Script:IconAlert -IconNormal $Script:IconNormal `
        -CompanyRegPath $Script:CompanyRegPath
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
                    Show-Toast $a.Title $snippet $a.Priority "announcement"
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
$Script:ToastFlags = [hashtable]::Synchronized(@{ RestartCount = 0; PatchCount = 0; Acknowledged = $false; RestartToastVisible = $false })
$toastTimer = New-Object System.Windows.Forms.Timer
$toastTimer.Interval = 5000
$toastTimer.Add_Tick({
    # Don't toast if user has acknowledged (opened dashboard or clicked toast)
    if ($Script:ToastFlags.Acknowledged) { return }

    if ($Script:ToastFlags.RestartCount -gt 0 -and -not $Script:ToastFlags.RestartToastVisible) {
        $count = $Script:ToastFlags.RestartCount
        $Script:ToastFlags.RestartCount = 0
        $Script:ToastFlags.RestartToastVisible = $true
        Show-Toast "System Restart Required" "$count update(s) require a restart. Open Software Center to apply and restart." "critical" "update"
    }
    if ($Script:ToastFlags.PatchCount -gt 0) {
        $count = $Script:ToastFlags.PatchCount
        $Script:ToastFlags.PatchCount = 0
        Show-Toast "Microsoft Updates Available" "$count update(s) available in Software Center. Please install at your earliest convenience." "info" "update"
    }
})
$toastTimer.Start()

# Periodic WU re-check timer — every 4 hours, re-checks and re-toasts if not acknowledged
$wuPeriodicTimer = New-Object System.Windows.Forms.Timer
$wuPeriodicTimer.Interval = 4 * 60 * 60 * 1000  # 4 hours in ms
$wuPeriodicTimer.Add_Tick({
    # Reset acknowledged flag so toast can fire again only if a new pending state exists
    $Script:ToastFlags.Acknowledged = $false
    # Re-run background WU check
    $tf = $Script:ToastFlags
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps = [powershell]::Create().AddScript({
        param($ToastFlags)
        try {
            $raw = Get-CimInstance -Namespace ROOT\ccm\ClientSDK -ClassName CCM_SoftwareUpdate -OperationTimeoutSec 30 -ErrorAction Stop |
                   Where-Object { $_.ComplianceState -eq 0 }
            if ($raw) {
                $restartCount = 0
                foreach ($u in $raw) {
                    $isRestart = $false
                    try {
                        $isRestart = [bool]$u.IsRebootPending -or
                            $u.EvaluationState -in @(8,9,10) -or
                            $u.RebootOutsideServiceWindow -eq $true -or
                            ($u.Name -match 'Cumulative Update|Security Update|Servicing Stack')
                    } catch {
                        try { $isRestart = $u.Name -match 'Cumulative Update|Security Update|Servicing Stack' } catch {}
                    }
                    if ($isRestart) { $restartCount++ }
                }
                if ($restartCount -gt 0) {
                    if (-not $ToastFlags.RestartToastVisible) { $ToastFlags.RestartCount = $restartCount }
                } else {
                    $ToastFlags.RestartToastVisible = $false
                }
            }
        } catch {}
    }).AddArgument($tf)
    $ps.Runspace = $rs
    $ps.BeginInvoke() | Out-Null
})
$wuPeriodicTimer.Start()

# Deferred startup refresh - fires 1s after message pump starts so UI never blocks
$startupTimer          = New-Object System.Windows.Forms.Timer
$startupTimer.Interval = 1000
$startupTimer.Add_Tick({
    $startupTimer.Stop()
    $startupTimer.Dispose()
    Start-ContentRefresh
})
$startupTimer.Start()

# Background WU check on startup — runs independently of dashboard for toast notifications
try {
    $wuStartupTimer = New-Object System.Windows.Forms.Timer
    $wuStartupTimer.Interval = 5000
    $wuStartupTimer.Add_Tick({
        try {
            $wuStartupTimer.Stop()
            $wuStartupTimer.Dispose()
            [System.IO.File]::AppendAllText("C:\temp\toast-debug.log", "$(Get-Date) | WU startup timer fired`r`n")
            $tf = $Script:ToastFlags
            $rs = [runspacefactory]::CreateRunspace()
            $rs.Open()
            $ps = [powershell]::Create().AddScript({
                param($ToastFlags)
                try {
                    [System.IO.File]::AppendAllText("C:\temp\toast-debug.log", "$(Get-Date) | WU runspace started`r`n")
                    $raw = Get-CimInstance -Namespace ROOT\ccm\ClientSDK -ClassName CCM_SoftwareUpdate -OperationTimeoutSec 30 -ErrorAction Stop |
                           Where-Object { $_.ComplianceState -eq 0 }
                    [System.IO.File]::AppendAllText("C:\temp\toast-debug.log", "$(Get-Date) | WU query done, found $(@($raw).Count) updates`r`n")
                    if ($raw) {
                        $restartCount = 0
                        $patchCount = 0
                        foreach ($u in $raw) {
                            $isRestart = $false
                            try {
                                $isRestart = [bool]$u.IsRebootPending -or
                                    $u.EvaluationState -in @(8,9,10) -or
                                    $u.RebootOutsideServiceWindow -eq $true -or
                                    ($u.Name -match 'Cumulative Update|Security Update|Servicing Stack')
                            } catch {
                                try { $isRestart = $u.Name -match 'Cumulative Update|Security Update|Servicing Stack' } catch {}
                            }
                            if ($isRestart) { $restartCount++ } else { $patchCount++ }
                        }
                        [System.IO.File]::AppendAllText("C:\temp\toast-debug.log", "$(Get-Date) | Restart: $restartCount Patch: $patchCount`r`n")
                        if ($restartCount -gt 0) {
                    if (-not $ToastFlags.RestartToastVisible) { $ToastFlags.RestartCount = $restartCount }
                } else {
                    $ToastFlags.RestartToastVisible = $false
                }
                        if ($patchCount -gt 0) {
                            $todayKey = [datetime]::Now.ToString('yyyy-MM-dd')
                            $lastToast = [Environment]::GetEnvironmentVariable('EA_LAST_PATCH_TOAST', 'User')
                            if ($lastToast -ne $todayKey) {
                                $ToastFlags.PatchCount = $patchCount
                                [Environment]::SetEnvironmentVariable('EA_LAST_PATCH_TOAST', $todayKey, 'User')
                            }
                        }
                    }
                } catch {
                    [System.IO.File]::AppendAllText("C:\temp\toast-debug.log", "$(Get-Date) | WU runspace ERROR: $($_.Exception.Message)`r`n")
                }
            }).AddArgument($tf)
            $ps.Runspace = $rs
            $ps.BeginInvoke() | Out-Null
        } catch {
            [System.IO.File]::AppendAllText("C:\temp\toast-debug.log", "$(Get-Date) | Timer tick ERROR: $($_.Exception.Message)`r`n")
        }
    })
    $wuStartupTimer.Start()
} catch {
    [System.IO.File]::AppendAllText("C:\temp\toast-debug.log", "$(Get-Date) | Timer setup ERROR: $($_.Exception.Message)`r`n")
}

Write-Log "$($Script:BrandedName) started on $Hostname (ScriptDir=$ScriptDir)"

[System.Windows.Forms.Application]::Run()
