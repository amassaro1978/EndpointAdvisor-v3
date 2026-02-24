#Requires -Version 5.1
<#
.SYNOPSIS
    Endpoint Advisor v3 - Backendless edition
.DESCRIPTION
    Polls a JSON file (hosted on GitHub or any internal web server) for
    announcements, evaluates client-side targeting conditions, and displays
    a WPF system-tray dashboard. No backend server required.
#>

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
    $tray.Add_Click({ Show-Dashboard })
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
    $tb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#E2E8F0")
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
                    $bd.Background      = & $mkBrush "#1E293B"
                    $bd.BorderBrush     = & $mkBrush $borderColor
                    $bd.BorderThickness = "3,0,0,0"
                    $bd.CornerRadius    = "6"
                    $bd.Margin          = "0,0,0,8"
                    $bd.Padding         = "14,10"

                    $sp = New-Object System.Windows.Controls.StackPanel
                    $title = if ($a.Title) { $a.Title } else { "" }
                    $sp.Children.Add((& $mkTb $title '#E2E8F0' 14)) | Out-Null

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

        $available = $false; $offers = @()
        try {
            $reg = "HKLM:\SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\_BESClient_LocalAPI_Enable"
            if (Test-Path $reg) {
                $v = (Get-ItemProperty $reg -ErrorAction Stop).'value'
                if ($v -eq '1') {
                    $resp = Invoke-RestMethod "http://127.0.0.1:52311/api/offers" -TimeoutSec 5 -ErrorAction Stop
                    $available = $true
                    if ($resp -and $resp.offers) {
                        foreach ($o in $resp.offers) { $offers += [PSCustomObject]@{ Id=$o.id; Name=$o.name; Version=$o.version } }
                    }
                }
            }
        } catch {}

        $Dispatcher.Invoke([Action]{
            $Container.Children.Clear()
            $mkB = { param($c) [System.Windows.Media.BrushConverter]::new().ConvertFrom($c) }
            $mkT = { param($t,$c='#94A3B8',$s=12) $tb=New-Object System.Windows.Controls.TextBlock; $tb.Text=$t; $tb.Foreground=& $mkB $c; $tb.FontSize=$s; $tb.Margin="4,4,0,4"; $tb.TextWrapping="Wrap"; $tb }

            if ($available -and $offers.Count -gt 0) {
                foreach ($o in $offers) {
                    $bd = New-Object System.Windows.Controls.Border
                    $bd.Background=& $mkB "#1E293B"; $bd.BorderBrush=& $mkB "#10B981"; $bd.BorderThickness="3,0,0,0"; $bd.CornerRadius="6"; $bd.Margin="0,0,0,8"; $bd.Padding="14,10"
                    $dk = New-Object System.Windows.Controls.DockPanel
                    $btn = New-Object System.Windows.Controls.Button
                    $btn.Content="Install"; $btn.Background=& $mkB "#10B981"; $btn.Foreground=& $mkB "#FFF"
                    $btn.BorderThickness="0"; $btn.Padding="12,4"; $btn.Cursor=[System.Windows.Input.Cursors]::Hand; $btn.Tag=$o.Id
                    $btn.Add_Click({ try { Invoke-RestMethod "http://127.0.0.1:52311/api/offers/$($this.Tag)/accept" -Method Post -TimeoutSec 15 | Out-Null; [System.Windows.MessageBox]::Show("Submitted.","Endpoint Advisor","OK","Information") } catch { [System.Windows.MessageBox]::Show("Failed: $_","Endpoint Advisor","OK","Error") } })
                    [System.Windows.Controls.DockPanel]::SetDock($btn,[System.Windows.Controls.Dock]::Right)
                    $dk.Children.Add($btn) | Out-Null
                    $inf = New-Object System.Windows.Controls.StackPanel
                    $inf.Children.Add((& $mkT $o.Name '#E2E8F0' 13)) | Out-Null
                    if ($o.Version) { $inf.Children.Add((& $mkT "Version: $($o.Version)")) | Out-Null }
                    $dk.Children.Add($inf) | Out-Null
                    $bd.Child = $dk; $Container.Children.Add($bd) | Out-Null
                }
            } elseif ($available) {
                $Container.Children.Add((& $mkT "No software offers available.")) | Out-Null
            } else {
                $Container.Children.Add((& $mkT "BigFix Local API not available.")) | Out-Null
                $btn = New-Object System.Windows.Controls.Button
                $btn.Content="Open BigFix Self Service"; $btn.Background=& $mkB "#6366F1"; $btn.Foreground=& $mkB "#FFF"
                $btn.BorderThickness="0"; $btn.Padding="14,6"; $btn.Margin="4,4,0,4"; $btn.Cursor=[System.Windows.Input.Cursors]::Hand; $btn.HorizontalAlignment="Left"
                $btn.Add_Click({ $p="C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClientUI.exe"; if(Test-Path $p){Start-Process $p}else{[System.Windows.MessageBox]::Show("Not installed.","Endpoint Advisor","OK","Warning")} })
                $Container.Children.Add($btn) | Out-Null
            }
        })
    } -Parameters @{ Dispatcher=$Dispatcher; Container=$Container; ScriptDir=$ScriptDir }
}

# ---- Async section: Windows Update -------------------------------------------
function Start-WULoad {
    param($Dispatcher, $Container)

    Start-TrackedRunspace -Script {
        param($Dispatcher, $Container)
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        $updates = @()
        $timedOut = $false
        try {
            # Run WU query in a background job with 15s timeout to prevent hang
            $job = Start-Job -ScriptBlock {
                $sess = New-Object -ComObject Microsoft.Update.Session
                $res  = $sess.CreateUpdateSearcher().Search("IsInstalled=0 AND IsHidden=0")
                $out  = @()
                foreach ($u in $res.Updates) {
                    $kb = if ($u.KBArticleIDs.Count -gt 0) { "KB$($u.KBArticleIDs.Item(0))" } else { "" }
                    $out += @{ Title=$u.Title; KB=$kb }
                }
                return $out
            }
            $completed = $job | Wait-Job -Timeout 15
            if ($completed) {
                $result = Receive-Job $job
                foreach ($r in $result) { $updates += [PSCustomObject]@{ Title=$r.Title; KB=$r.KB } }
            } else {
                $timedOut = $true
                Stop-Job $job -ErrorAction SilentlyContinue
            }
            Remove-Job $job -Force -ErrorAction SilentlyContinue
        } catch {}

        $Dispatcher.Invoke([Action]{
            $Container.Children.Clear()
            $mkB = { param($c) [System.Windows.Media.BrushConverter]::new().ConvertFrom($c) }
            $mkT = { param($t,$c='#94A3B8',$s=12) $tb=New-Object System.Windows.Controls.TextBlock; $tb.Text=$t; $tb.Foreground=& $mkB $c; $tb.FontSize=$s; $tb.Margin="4,4,0,4"; $tb.TextWrapping="Wrap"; $tb }

            if ($timedOut) {
                $Container.Children.Add((& $mkT "Update check timed out. Open Windows Update to check manually.")) | Out-Null
                $btn = New-Object System.Windows.Controls.Button
                $btn.Content="Open Windows Update"; $btn.Background=& $mkB "#F59E0B"; $btn.Foreground=& $mkB "#0F172A"
                $btn.BorderThickness="0"; $btn.Padding="14,6"; $btn.Margin="0,6,0,0"; $btn.Cursor=[System.Windows.Input.Cursors]::Hand; $btn.HorizontalAlignment="Left"; $btn.FontWeight="SemiBold"
                $btn.Add_Click({ Start-Process "softwarecenter:" })
                $Container.Children.Add($btn) | Out-Null
            } elseif ($updates.Count -eq 0) {
                $Container.Children.Add((& $mkT "System is up to date.")) | Out-Null
            } else {
                foreach ($u in $updates) {
                    $bd = New-Object System.Windows.Controls.Border
                    $bd.Background=& $mkB "#1E293B"; $bd.BorderBrush=& $mkB "#F59E0B"; $bd.BorderThickness="3,0,0,0"; $bd.CornerRadius="6"; $bd.Margin="0,0,0,6"; $bd.Padding="10,8"
                    $label = if ($u.KB) { "$($u.KB) - $($u.Title)" } else { $u.Title }
                    $bd.Child = (& $mkT $label '#E2E8F0' 12); $Container.Children.Add($bd) | Out-Null
                }
                $btn = New-Object System.Windows.Controls.Button
                $btn.Content="Install Patches"; $btn.Background=& $mkB "#F59E0B"; $btn.Foreground=& $mkB "#0F172A"
                $btn.BorderThickness="0"; $btn.Padding="14,6"; $btn.Margin="0,6,0,0"; $btn.Cursor=[System.Windows.Input.Cursors]::Hand; $btn.HorizontalAlignment="Left"; $btn.FontWeight="SemiBold"
                $btn.Add_Click({ Start-Process "softwarecenter:" })
                $Container.Children.Add($btn) | Out-Null
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
        Title="Endpoint Advisor" Width="620" Height="740"
        WindowStartupLocation="CenterScreen" Background="#0F172A" ShowInTaskbar="False">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#E2E8F0"/>
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
        <Border Grid.Row="0" Background="#1E293B" Padding="16,12">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Endpoint Advisor" FontSize="18" FontWeight="Bold" VerticalAlignment="Center"/>
                <TextBlock x:Name="HostLabel" FontSize="11" Foreground="#475569" Margin="12,0,0,0" VerticalAlignment="Center"/>
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
    $panel.Children.Add((New-SectionHeader "Software Updates")) | Out-Null
    $swPanel = New-Object System.Windows.Controls.StackPanel
    $swPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($swPanel) | Out-Null

    # Pending Windows Updates
    $panel.Children.Add((New-SectionHeader "Pending Updates")) | Out-Null
    $wuPanel = New-Object System.Windows.Controls.StackPanel
    $wuPanel.Children.Add((New-InfoText "Loading...")) | Out-Null
    $panel.Children.Add($wuPanel) | Out-Null

    $Script:DashboardWindow = $window
    $window.Show()

    $dispatcher = $window.Dispatcher
    Start-AnnouncementsLoad -Dispatcher $dispatcher -Container $annPanel `
        -ContentDataUrl $Config.ContentDataUrl -GitHubToken $Config.GitHubToken `
        -TrayIcon $Script:TrayIcon -IconAlert $Script:IconAlert -IconNormal $Script:IconNormal
    Start-BigFixLoad -Dispatcher $dispatcher -Container $swPanel -ScriptDir $ScriptDir
    Start-WULoad     -Dispatcher $dispatcher -Container $wuPanel
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
            } elseif ($a.NagEnabled -and $a.PSObject.Properties['NagEnabled']) {
                $interval = if ($a.NagIntervalMinutes) { $a.NagIntervalMinutes } else { 30 }
                $last = $Script:LastNagTime[$a.id]
                if (-not $last -or ($now - $last).TotalMinutes -ge $interval) { $shouldNotify = $true }
            }
            if ($shouldNotify) {
                $Script:LastNagTime[$a.id] = $now
                $snippet = if ($a.Text) { $a.Text.Substring(0, [Math]::Min(200, $a.Text.Length)) } else { "" }
                Show-Toast $a.Title $snippet $a.Priority
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
