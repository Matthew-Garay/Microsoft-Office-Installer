# Microsoft Office Installer - Graphical Edition
# Downloads the official Office Deployment Tool (ODT) from Microsoft and installs
# the selected Office edition. It only installs Office; it does NOT activate it.
#
# Usage:  irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-gui.ps1 | iex
#    or:  .\install-gui.ps1
#
# Everything runs inside a script block so that nothing leaks into the caller's
# session and no 'exit' closes the user's PowerShell window when run via iex.

& {
$ErrorActionPreference = 'Stop'
$ScriptUrl = 'https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-gui.ps1'

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {}

# ------------------------------------------------------ elevation + STA ------
# WinForms needs an STA thread; Windows PowerShell 5.1 (powershell.exe) is STA.
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$isSta   = [Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA'
if (-not $isAdmin -or -not $isSta) {
    Write-Host ''
    Write-Host '  Opening the installer in an Administrator window...' -ForegroundColor Yellow
    if ($PSCommandPath) {
        $argList = "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    } else {
        $argList = "-NoProfile -STA -ExecutionPolicy Bypass -Command `"[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; irm '$ScriptUrl' | iex`""
    }
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argList | Out-Null
    } catch {
        Write-Host '  Elevation was cancelled. Nothing was changed.' -ForegroundColor Red
    }
    Write-Host ''
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Add-Type fails if the type already exists (script run twice in one window).
if (-not ('OfficeDpiHelper' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class OfficeDpiHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int a);
}
"@
}
try { [void][OfficeDpiHelper]::SetProcessDpiAwareness(1) } catch { try { [void][OfficeDpiHelper]::SetProcessDPIAware() } catch {} }

# ------------------------------------------------------------------- config --
$logFile     = Join-Path $env:TEMP 'OfficeInstallerGUI_Install.log'
$odtTemp     = Join-Path $env:TEMP 'OfficeInstaller_ODT'
$odtExe      = Join-Path $env:TEMP 'OfficeInstaller_ODT.exe'
$odtPage     = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
$odtFallback = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe'

# Shared mutable state (event handlers write here instead of to loose variables).
$state = @{ OdtReady = $false; Lang = 'en-US'; Updating = $false; Busy = $false }

# C = channel, P/V/J = product IDs for Office / Visio / Project.
# Office 2016 has no volume Click-to-Run edition, so the retail C2R product is used.
# Office 2013 is not offered: it needs the retired 2013 ODT and is out of support.
$versions = @(
    @{ Label = 'Office LTSC Professional Plus 2024'; C = 'PerpetualVL2024'; P = 'ProPlus2024Volume'; V = 'VisioPro2024Volume'; J = 'ProjectPro2024Volume' }
    @{ Label = 'Office LTSC Professional Plus 2021'; C = 'PerpetualVL2021'; P = 'ProPlus2021Volume'; V = 'VisioPro2021Volume'; J = 'ProjectPro2021Volume' }
    @{ Label = 'Office Professional Plus 2019';      C = 'PerpetualVL2019'; P = 'ProPlus2019Volume'; V = 'VisioPro2019Volume'; J = 'ProjectPro2019Volume' }
    @{ Label = 'Office Professional Plus 2016';      C = 'Current';         P = 'ProPlusRetail';     V = 'VisioProRetail';     J = 'ProjectProRetail' }
)

$langs = @(
    @{ Code = 'en-US'; Name = 'English';    Flag = 'US' }
    @{ Code = 'es-ES'; Name = 'Spanish';    Flag = 'ES' }
    @{ Code = 'fr-FR'; Name = 'French';     Flag = 'FR' }
    @{ Code = 'de-DE'; Name = 'German';     Flag = 'DE' }
    @{ Code = 'pt-BR'; Name = 'Portuguese'; Flag = 'BR' }
    @{ Code = 'it-IT'; Name = 'Italian';    Flag = 'IT' }
    @{ Code = 'nl-NL'; Name = 'Dutch';      Flag = 'NL' }
    @{ Code = 'pl-PL'; Name = 'Polish';     Flag = 'PL' }
    @{ Code = 'ru-RU'; Name = 'Russian';    Flag = 'RU' }
    @{ Code = 'ja-JP'; Name = 'Japanese';   Flag = 'JP' }
)

# I = tag, D = display name, X = official ODT ExcludeApp ID, C = checked by default.
$apps = @(
    @{ I = 'Word';             D = 'Word';           X = 'Word';       C = $true;  Addon = $false }
    @{ I = 'Excel';            D = 'Excel';          X = 'Excel';      C = $true;  Addon = $false }
    @{ I = 'PowerPoint';       D = 'PowerPoint';     X = 'PowerPoint'; C = $true;  Addon = $false }
    @{ I = 'Outlook';          D = 'Outlook';        X = 'Outlook';    C = $false; Addon = $false }
    @{ I = 'Access';           D = 'Access';         X = 'Access';     C = $false; Addon = $false }
    @{ I = 'Publisher';        D = 'Publisher';      X = 'Publisher';  C = $false; Addon = $false }
    @{ I = 'OneNote';          D = 'OneNote';        X = 'OneNote';    C = $false; Addon = $false }
    @{ I = 'SkypeForBusiness'; D = 'Skype for Bus.'; X = 'Lync';       C = $false; Addon = $false }
    @{ I = 'Project';          D = 'Project';        X = $null;        C = $false; Addon = $true }
    @{ I = 'Visio';            D = 'Visio';          X = $null;        C = $false; Addon = $true }
)
$alwaysExcluded = @('Groove', 'OneDrive', 'Teams')

$th = @{
    Bg       = [System.Drawing.Color]::FromArgb(255, 255, 255)
    Panel    = [System.Drawing.Color]::FromArgb(248, 250, 252)
    Text     = [System.Drawing.Color]::FromArgb(15, 23, 42)
    Dim      = [System.Drawing.Color]::FromArgb(100, 116, 139)
    Accent   = [System.Drawing.Color]::FromArgb(220, 38, 38)
    AccentH  = [System.Drawing.Color]::FromArgb(185, 28, 28)
    Ok       = [System.Drawing.Color]::FromArgb(22, 163, 74)
    Bad      = [System.Drawing.Color]::FromArgb(220, 38, 38)
    Warn     = [System.Drawing.Color]::FromArgb(202, 138, 4)
    Track    = [System.Drawing.Color]::FromArgb(226, 232, 240)
    Border   = [System.Drawing.Color]::FromArgb(203, 213, 225)
    CardBd   = [System.Drawing.Color]::FromArgb(203, 213, 225)
    CardOn   = [System.Drawing.Color]::FromArgb(254, 242, 242)
    CardOnBd = [System.Drawing.Color]::FromArgb(220, 38, 38)
}

$appColors = @{
    Office           = [System.Drawing.Color]::FromArgb(220, 62, 21)
    Word             = [System.Drawing.Color]::FromArgb(43, 87, 151)
    Excel            = [System.Drawing.Color]::FromArgb(33, 115, 70)
    PowerPoint       = [System.Drawing.Color]::FromArgb(183, 71, 42)
    Outlook          = [System.Drawing.Color]::FromArgb(0, 120, 212)
    Access           = [System.Drawing.Color]::FromArgb(165, 51, 84)
    Publisher        = [System.Drawing.Color]::FromArgb(0, 122, 128)
    OneNote          = [System.Drawing.Color]::FromArgb(115, 69, 178)
    SkypeForBusiness = [System.Drawing.Color]::FromArgb(0, 120, 215)
    Project          = [System.Drawing.Color]::FromArgb(49, 114, 59)
    Visio            = [System.Drawing.Color]::FromArgb(65, 84, 178)
}

# ------------------------------------------------------------------ helpers --
function Write-Log {
    param([string]$M, [string]$C = 'Gray')
    try { Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $M" } catch {}
    Write-Host $M -ForegroundColor $C
}

function Remove-Temp {
    if (Test-Path $odtTemp) { Remove-Item $odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $odtExe)  { Remove-Item $odtExe -Force -ErrorAction SilentlyContinue }
}

# Keeps the window responsive while waiting for background work.
function Wait-UI {
    param([scriptblock]$Until)
    while (-not (& $Until)) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
}

function Get-ExistingOffice {
    $key = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (-not (Test-Path $key)) { return $null }
    $p = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
    if (-not $p) { return $null }
    return @{ Products = [string]$p.ProductReleaseIds; Platform = [string]$p.Platform }
}

function Get-OdtUrls {
    $urls = @()
    try {
        $page = Invoke-WebRequest -Uri $odtPage -UseBasicParsing -TimeoutSec 30
        $rx = 'https://download\.microsoft\.com/download/[^"''\s<>]+?officedeploymenttool[^"''\s<>]*?\.exe'
        foreach ($m in [regex]::Matches($page.Content, $rx)) { $urls += $m.Value }
    } catch {
        Write-Log "Could not read the ODT download page: $_" 'DarkYellow'
    }
    $urls += $odtFallback
    return @($urls | Select-Object -Unique)
}

function New-RoundedPath {
    param([float]$W, [float]$H, [float]$R)
    $d = $R * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($W - $d, 0, $d, $d, 270, 90)
    $path.AddArc($W - $d, $H - $d, $d, $d, 0, 90)
    $path.AddArc(0, $H - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-AppLogo {
    param([string]$Tag, [int]$Size = 34)
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $bg = $appColors[$Tag]; if (-not $bg) { $bg = $th.Accent }
        $path = New-RoundedPath $Size $Size ([float]($Size * 0.22))
        $br = New-Object System.Drawing.SolidBrush($bg)
        $g.FillPath($br, $path); $br.Dispose(); $path.Dispose()

        $letter = $Tag.Substring(0, 1)
        if ($Tag -eq 'OneNote') { $letter = 'N' }
        $font = New-Object System.Drawing.Font('Segoe UI', [float]($Size * 0.5), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
        $tb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $g.DrawString($letter, $font, $tb, (New-Object System.Drawing.RectangleF(0, 1, $Size, $Size)), $sf)
        $tb.Dispose(); $font.Dispose(); $sf.Dispose()
    } finally { $g.Dispose() }
    return $bmp
}

function New-FlagImage {
    param([string]$Code, [int]$W = 32, [int]$H = 22)
    $bmp = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $path = New-RoundedPath $W $H 4
        $g.SetClip($path)
        $fill = {
            param([float]$x, [float]$y, [float]$w, [float]$h, [int]$r, [int]$gg, [int]$b)
            $br = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($r, $gg, $b))
            $g.FillRectangle($br, $x, $y, $w, $h); $br.Dispose()
        }
        switch ($Code) {
            'US' {
                & $fill 0 0 $W $H 178 34 52
                for ($i = 0; $i -lt 7; $i++) { & $fill 0 ($i * $H / 13 * 2) $W ($H / 13) 255 255 255 }
                & $fill 0 0 ($W * 0.42) ($H * 7 / 13) 60 59 110
            }
            'ES' {
                & $fill 0 0 $W $H 170 21 27
                & $fill 0 ($H * 0.25) $W ($H * 0.5) 252 209 22
            }
            'FR' { & $fill 0 0 ($W / 3) $H 0 85 164; & $fill ($W / 3) 0 ($W / 3) $H 255 255 255; & $fill (2 * $W / 3) 0 ($W / 3 + 1) $H 239 65 53 }
            'DE' { & $fill 0 0 $W ($H / 3) 0 0 0; & $fill 0 ($H / 3) $W ($H / 3) 221 0 0; & $fill 0 (2 * $H / 3) $W ($H / 3 + 1) 255 206 0 }
            'BR' {
                & $fill 0 0 $W $H 0 156 59
                $pts = [System.Drawing.PointF[]]@(
                    (New-Object System.Drawing.PointF(($W / 2), 2)),
                    (New-Object System.Drawing.PointF(($W - 2), ($H / 2))),
                    (New-Object System.Drawing.PointF(($W / 2), ($H - 2))),
                    (New-Object System.Drawing.PointF(2, ($H / 2)))
                )
                $yb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 223, 0))
                $g.FillPolygon($yb, $pts); $yb.Dispose()
                $bb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 39, 118))
                $g.FillEllipse($bb, [float]($W * 0.32), [float]($H * 0.22), [float]($W * 0.36), [float]($H * 0.56)); $bb.Dispose()
            }
            'IT' { & $fill 0 0 ($W / 3) $H 0 146 70; & $fill ($W / 3) 0 ($W / 3) $H 255 255 255; & $fill (2 * $W / 3) 0 ($W / 3 + 1) $H 206 43 55 }
            'NL' { & $fill 0 0 $W ($H / 3) 174 28 40; & $fill 0 ($H / 3) $W ($H / 3) 255 255 255; & $fill 0 (2 * $H / 3) $W ($H / 3 + 1) 33 70 135 }
            'PL' { & $fill 0 0 $W ($H / 2) 255 255 255; & $fill 0 ($H / 2) $W ($H / 2) 220 20 60 }
            'RU' { & $fill 0 0 $W ($H / 3) 255 255 255; & $fill 0 ($H / 3) $W ($H / 3) 0 57 166; & $fill 0 (2 * $H / 3) $W ($H / 3 + 1) 213 43 30 }
            'JP' {
                & $fill 0 0 $W $H 255 255 255
                $rb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(188, 0, 45))
                $g.FillEllipse($rb, [float]($W * 0.32), [float]($H * 0.2), [float]($W * 0.36), [float]($H * 0.6)); $rb.Dispose()
            }
            default { & $fill 0 0 $W $H 120 130 140 }
        }
        $g.ResetClip()
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(140, 150, 160))
        $g.DrawPath($pen, $path); $pen.Dispose(); $path.Dispose()
    } finally { $g.Dispose() }
    return $bmp
}

function New-Label {
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Text, [float]$Size = 9, [string]$Style = 'Regular', $Color = $th.Text)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, $H)
    $l.Font = New-Object System.Drawing.Font('Segoe UI', $Size, [System.Drawing.FontStyle]$Style)
    $l.ForeColor = $Color
    $l.BackColor = [System.Drawing.Color]::Transparent
    $Parent.Controls.Add($l)
    return $l
}

# ------------------------------------------------------- preparation dialog --
function Show-PrepDialog {
    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = 'Dpi'
    $f.Text = 'Microsoft Office Installer - Preparation'
    $f.ClientSize = New-Object System.Drawing.Size(460, 150)
    $f.FormBorderStyle = 'FixedDialog'
    $f.ControlBox = $false
    $f.StartPosition = 'CenterScreen'
    $f.BackColor = $th.Bg
    $f.TopMost = $true
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    [void](New-Label $f 24 24 410 24 "$([char]0x25B8) Preparing Office Deployment Tool..." 10 'Bold')
    $status = New-Label $f 24 54 410 20 'Starting...' 9 'Regular' $th.Dim
    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Location = New-Object System.Drawing.Point(24, 88)
    $pb.Size = New-Object System.Drawing.Size(410, 22)
    $pb.Style = 'Marquee'
    $f.Controls.Add($pb)

    $f.Show()
    $setStatus = { param($t) $status.Text = $t; [System.Windows.Forms.Application]::DoEvents() }

    try {
        & $setStatus 'Cleaning up previous files...'
        Remove-Temp
        New-Item -ItemType Directory -Path $odtTemp -Force | Out-Null

        & $setStatus 'Looking up the latest Office Deployment Tool...'
        $urls = Get-OdtUrls
        $ok = $false
        foreach ($u in $urls) {
            & $setStatus 'Downloading from download.microsoft.com...'
            Write-Log "Downloading ODT from $u"
            $wc = New-Object System.Net.WebClient
            $task = $wc.DownloadFileTaskAsync($u, $odtExe)
            Wait-UI { $task.IsCompleted }
            $wc.Dispose()
            if (-not $task.IsFaulted -and -not $task.IsCanceled) { $ok = $true; break }
            Write-Log "Source failed: $u ($($task.Exception.InnerException.Message))" 'DarkYellow'
        }
        if (-not $ok) { throw 'Could not download the Office Deployment Tool. Check your internet connection.' }

        & $setStatus 'Verifying digital signature...'
        $sig = Get-AuthenticodeSignature -FilePath $odtExe
        if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
            throw "The downloaded file is not signed by Microsoft (signature: $($sig.Status))."
        }

        & $setStatus 'Extracting Office Deployment Tool...'
        $p = Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:`"$odtTemp`"" -PassThru
        $null = $p.Handle   # caches the handle so ExitCode is available later
        Wait-UI { $p.HasExited }
        if ($p.ExitCode -ne 0) { throw "Extraction failed (exit code $($p.ExitCode))." }
        if (-not (Test-Path (Join-Path $odtTemp 'setup.exe'))) { throw 'setup.exe not found after extraction.' }

        $state.OdtReady = $true
        Write-Log 'ODT ready' 'Green'
    } catch {
        Write-Log "Prep error: $_" 'Red'
        $f.TopMost = $false
        [void][System.Windows.Forms.MessageBox]::Show("Failed to prepare the Office Deployment Tool.`n`n$($_.Exception.Message)", 'Microsoft Office Installer - Error', 'OK', 'Error')
    } finally {
        $f.Close(); $f.Dispose()
    }
}

# ---------------------------------------------------------------- main form --
function Show-MainForm {
    $M = 24; $FW = 920; $GW = $FW - 2 * $M; $fnt = 'Segoe UI'

    $existing = Get-ExistingOffice
    $arch = if ([Environment]::Is64BitOperatingSystem) { '64' } else { '32' }
    if ($existing -and $existing.Platform -eq 'x86') { $arch = '32' }

    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = 'Dpi'
    $f.Text = 'Microsoft Office Installer'
    $f.ClientSize = New-Object System.Drawing.Size($FW, 700)
    $f.StartPosition = 'CenterScreen'
    $f.FormBorderStyle = 'FixedSingle'
    $f.MaximizeBox = $false
    $f.BackColor = $th.Bg
    $f.Font = New-Object System.Drawing.Font($fnt, 9)

    # Header
    $hd = New-Object System.Windows.Forms.Panel
    $hd.Location = New-Object System.Drawing.Point(0, 0)
    $hd.Size = New-Object System.Drawing.Size($FW, 85)
    $hd.BackColor = $th.Panel
    $f.Controls.Add($hd)

    $logo = New-Object System.Windows.Forms.PictureBox
    $logo.Image = New-AppLogo 'Office' 50
    $logo.Location = New-Object System.Drawing.Point($M, 18)
    $logo.Size = New-Object System.Drawing.Size(50, 50)
    $hd.Controls.Add($logo)

    [void](New-Label $hd ($M + 62) 14 500 30 'Microsoft Office Installer' 15 'Bold')
    [void](New-Label $hd ($M + 62) 46 600 20 'Install Microsoft Office LTSC & Professional Plus on Windows (no activation)' 9 'Regular' $th.Dim)
    $badge = New-Label $hd ($FW - 110) 28 86 26 "$arch-bit" 8.5 'Bold' $th.Ok
    $badge.TextAlign = 'MiddleCenter'; $badge.BackColor = $th.Bg

    $bar = New-Object System.Windows.Forms.Panel
    $bar.Location = New-Object System.Drawing.Point(0, 85)
    $bar.Size = New-Object System.Drawing.Size($FW, 3)
    $bar.BackColor = $th.Accent
    $f.Controls.Add($bar)

    $y = 104
    if ($existing) {
        $warn = New-Label $f $M $y $GW 18 "!  Office is already installed ($($existing.Products)). Installing another version next to it may fail." 9 'Bold' $th.Warn
        $y += 24
    }

    # 1. Version
    [void](New-Label $f $M $y 800 20 '1    VERSION' 9.5 'Bold' $th.Accent)
    $y += 22
    $cv = New-Object System.Windows.Forms.ComboBox
    $cv.Location = New-Object System.Drawing.Point($M, $y)
    $cv.Size = New-Object System.Drawing.Size($GW, 28)
    $cv.DropDownStyle = 'DropDownList'
    $cv.Font = New-Object System.Drawing.Font($fnt, 10)
    $cv.FlatStyle = 'Flat'
    foreach ($v in $versions) { [void]$cv.Items.Add($v.Label) }
    $cv.SelectedIndex = 0
    $f.Controls.Add($cv)
    $y += 42

    # 2. Language
    [void](New-Label $f $M $y 800 20 '2    LANGUAGE' 9.5 'Bold' $th.Accent)
    $y += 22
    $langPanel = New-Object System.Windows.Forms.Panel
    $langPanel.Location = New-Object System.Drawing.Point($M, $y)
    $langPanel.Size = New-Object System.Drawing.Size($GW, 80)
    $langPanel.BackColor = $th.Panel
    $f.Controls.Add($langPanel)

    $lCols = 5; $lGap = 10; $lPad = 12; $lBh = 30
    $lBw = [Math]::Floor(($GW - 2 * $lPad - ($lCols - 1) * $lGap) / $lCols)
    $langButtons = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $langs.Count; $i++) {
        $lg = $langs[$i]; $col = $i % $lCols; $row = [Math]::Floor($i / $lCols)
        $b = New-Object System.Windows.Forms.Button
        $b.Tag = $lg.Code
        $b.Text = "  $($lg.Name)"
        $b.TextAlign = 'MiddleLeft'
        $b.Image = New-FlagImage $lg.Flag
        $b.ImageAlign = 'MiddleLeft'
        $b.TextImageRelation = 'ImageBeforeText'
        $b.Location = New-Object System.Drawing.Point(($lPad + $col * ($lBw + $lGap)), (8 + $row * ($lBh + 6)))
        $b.Size = New-Object System.Drawing.Size($lBw, $lBh)
        $b.FlatStyle = 'Flat'
        $b.FlatAppearance.BorderSize = 1
        $b.Cursor = 'Hand'
        $b.Add_Click({ $state.Lang = [string]$this.Tag; Paint-LangButtons; Update-Summary })
        $langPanel.Controls.Add($b)
        [void]$langButtons.Add($b)
    }
    $fontNormal = New-Object System.Drawing.Font($fnt, 8.5)
    $fontBold   = New-Object System.Drawing.Font($fnt, 8.5, [System.Drawing.FontStyle]::Bold)
    function Paint-LangButtons {
        foreach ($b in $langButtons) {
            $on = ([string]$b.Tag -eq $state.Lang)
            $b.BackColor = if ($on) { $th.CardOn } else { $th.Panel }
            $b.ForeColor = if ($on) { $th.Text } else { $th.Dim }
            $b.FlatAppearance.BorderColor = if ($on) { $th.CardOnBd } else { $th.CardBd }
            $b.Font = if ($on) { $fontBold } else { $fontNormal }
        }
    }
    Paint-LangButtons
    $y += 80 + 20

    # 3. Applications
    [void](New-Label $f $M $y 800 20 '3    APPLICATIONS' 9.5 'Bold' $th.Accent)
    $y += 22
    $ap = New-Object System.Windows.Forms.Panel
    $ap.Location = New-Object System.Drawing.Point($M, $y)
    $ap.Size = New-Object System.Drawing.Size($GW, 154)
    $ap.BackColor = $th.Panel
    $f.Controls.Add($ap)

    $csa = New-Object System.Windows.Forms.CheckBox
    $csa.Text = 'Select all applications'
    $csa.Location = New-Object System.Drawing.Point(12, 8)
    $csa.Size = New-Object System.Drawing.Size(200, 20)
    $csa.Font = $fontBold
    $csa.BackColor = [System.Drawing.Color]::Transparent
    $ap.Controls.Add($csa)

    $pd = 12; $gap = 10; $cols = 5; $cardH = 54
    $spc = [Math]::Floor(($GW - 2 * $pd - ($cols - 1) * $gap) / $cols) + $gap
    $bw = $spc - $gap
    $checks = New-Object System.Collections.ArrayList
    $cards = @{}

    for ($i = 0; $i -lt $apps.Count; $i++) {
        $a = $apps[$i]; $col = $i % $cols; $row = [Math]::Floor($i / $cols)
        $card = New-Object System.Windows.Forms.Panel
        $card.Tag = $a.I
        $card.Location = New-Object System.Drawing.Point(($pd + $col * $spc), (34 + $row * ($cardH + 6)))
        $card.Size = New-Object System.Drawing.Size($bw, $cardH)
        $card.BorderStyle = 'FixedSingle'
        $card.Cursor = 'Hand'
        $ap.Controls.Add($card)

        $pic = New-Object System.Windows.Forms.PictureBox
        $pic.Image = New-AppLogo $a.I 34
        $pic.Location = New-Object System.Drawing.Point(8, 9)
        $pic.Size = New-Object System.Drawing.Size(34, 34)
        $pic.Tag = $a.I
        $card.Controls.Add($pic)

        $lbl = New-Label $card 48 7 ($bw - 52) 20 $a.D 8.5 'Bold'
        $lbl.Tag = $a.I

        $c = New-Object System.Windows.Forms.CheckBox
        $c.Tag = $a.I
        $c.Location = New-Object System.Drawing.Point(48, 27)
        $c.Size = New-Object System.Drawing.Size(20, 20)
        $c.Checked = $a.C
        $c.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($c)
        [void]$checks.Add($c)
        $cards[$a.I] = @{ Check = $c; Card = $card }

        $toggle = { $chk = $cards[[string]$this.Tag].Check; $chk.Checked = -not $chk.Checked }
        $card.Add_Click($toggle); $pic.Add_Click($toggle); $lbl.Add_Click($toggle)
        $c.Add_CheckedChanged({
            if (-not $state.Updating) {
                $all = $true; foreach ($x in $checks) { if (-not $x.Checked) { $all = $false; break } }
                $state.Updating = $true; $csa.Checked = $all; $state.Updating = $false
            }
            Paint-Cards; Update-Summary
        })
    }
    function Paint-Cards {
        foreach ($k in $cards.Keys) {
            $cards[$k].Card.BackColor = if ($cards[$k].Check.Checked) { $th.CardOn } else { $th.Panel }
        }
    }
    $csa.Add_CheckedChanged({
        if ($state.Updating) { return }
        $state.Updating = $true
        foreach ($x in $checks) { $x.Checked = $csa.Checked }
        $state.Updating = $false
        Paint-Cards; Update-Summary
    })
    Paint-Cards
    $y += 154 + 18

    # Summary
    $sum = New-Object System.Windows.Forms.Panel
    $sum.Location = New-Object System.Drawing.Point($M, $y)
    $sum.Size = New-Object System.Drawing.Size($GW, 76)
    $sum.BackColor = $th.Panel
    $sum.BorderStyle = 'FixedSingle'
    $f.Controls.Add($sum)
    function Add-SummaryField {
        param([int]$X, [int]$Y, [string]$Caption)
        [void](New-Label $sum $X $Y 410 14 $Caption 7.5 'Bold' $th.Dim)
        $v = New-Label $sum $X ($Y + 14) 410 18 '' 9
        $v.AutoEllipsis = $true
        return $v
    }
    $sumEd   = Add-SummaryField 14 8 'EDITION'
    $sumLang = Add-SummaryField 14 40 'LANGUAGE'
    $sumArch = Add-SummaryField 440 8 'ARCHITECTURE'
    $sumApps = Add-SummaryField 440 40 'APPLICATIONS'
    $y += 86

    # Status + progress
    $sb = New-Label $f $M $y $GW 18 'Ready to install' 9 'Regular' $th.Dim
    $y += 22
    $track = New-Object System.Windows.Forms.Panel
    $track.Location = New-Object System.Drawing.Point($M, $y)
    $track.Size = New-Object System.Drawing.Size($GW, 6)
    $track.BackColor = $th.Track
    $f.Controls.Add($track)
    $fill = New-Object System.Windows.Forms.Panel
    $fill.Location = New-Object System.Drawing.Point(0, 0)
    $fill.Size = New-Object System.Drawing.Size(0, 6)
    $fill.BackColor = $th.Accent
    $track.Controls.Add($fill)

    function Set-Progress {
        param([int]$Pct, [string]$Text, [string]$State = 'run')
        $Pct = [Math]::Max(0, [Math]::Min(100, $Pct))
        $fill.Width = [int][Math]::Floor($GW * $Pct / 100)
        $sb.Text = $Text
        switch ($State) {
            'ok'    { $fill.BackColor = $th.Ok;     $sb.ForeColor = $th.Ok }
            'bad'   { $fill.BackColor = $th.Bad;    $sb.ForeColor = $th.Bad }
            'warn'  { $fill.BackColor = $th.Warn;   $sb.ForeColor = $th.Warn }
            'idle'  { $fill.BackColor = $th.Accent; $sb.ForeColor = $th.Dim }
            default { $fill.BackColor = $th.Accent; $sb.ForeColor = $th.Text }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Get-SelectedTags { @($checks | Where-Object { $_.Checked } | ForEach-Object { [string]$_.Tag }) }

    function Update-Summary {
        $sumEd.Text = [string]$cv.SelectedItem
        $sumLang.Text = "$(($langs | Where-Object { $_.Code -eq $state.Lang }).Name) ($($state.Lang))"
        $sumArch.Text = "$arch-bit"
        $picked = Get-SelectedTags
        if ($picked.Count -eq 0) {
            $sumApps.Text = 'Nothing selected'
            $sumApps.ForeColor = $th.Warn
        } else {
            $sumApps.Text = (@($picked | ForEach-Object { $t = $_; ($apps | Where-Object { $_.I -eq $t }).D }) -join ', ')
            $sumApps.ForeColor = $th.Text
        }
    }
    $cv.Add_SelectedIndexChanged({ Update-Summary })
    Update-Summary
    $y += 24

    # Buttons
    $biW = 200; $bcW = 140; $btnGap = 14
    $btnStart = [Math]::Floor(($FW - $biW - $bcW - $btnGap) / 2)
    $bi = New-Object System.Windows.Forms.Button
    $bi.Text = "$([char]0x25B6)  Install Office"
    $bi.Location = New-Object System.Drawing.Point($btnStart, $y)
    $bi.Size = New-Object System.Drawing.Size($biW, 40)
    $bi.Font = New-Object System.Drawing.Font($fnt, 9.5, [System.Drawing.FontStyle]::Bold)
    $bi.FlatStyle = 'Flat'
    $bi.FlatAppearance.BorderSize = 0
    $bi.FlatAppearance.MouseOverBackColor = $th.AccentH
    $bi.BackColor = $th.Accent
    $bi.ForeColor = [System.Drawing.Color]::White
    $bi.Cursor = 'Hand'
    $f.Controls.Add($bi)

    $bc = New-Object System.Windows.Forms.Button
    $bc.Text = "$([char]0x2715)  Close"
    $bc.Location = New-Object System.Drawing.Point(($btnStart + $biW + $btnGap), $y)
    $bc.Size = New-Object System.Drawing.Size($bcW, 40)
    $bc.Font = New-Object System.Drawing.Font($fnt, 9.5)
    $bc.FlatStyle = 'Flat'
    $bc.FlatAppearance.BorderColor = $th.Border
    $bc.FlatAppearance.MouseOverBackColor = $th.Track
    $bc.BackColor = $th.Panel
    $bc.ForeColor = $th.Dim
    $bc.Cursor = 'Hand'
    $f.Controls.Add($bc)

    $by = New-Label $f $M ($y + 50) $GW 18 "Microsoft Office Installer  $([char]0x00B7)  by Matthew Garay" 8.5 'Regular' $th.Dim
    $by.TextAlign = 'MiddleCenter'
    $f.ClientSize = New-Object System.Drawing.Size($FW, ($y + 80))

    $bc.Add_Click({ $f.Close() })
    $f.Add_FormClosing({
        param($s, $e)
        if ($state.Busy) {
            $e.Cancel = $true
            [void][System.Windows.Forms.MessageBox]::Show('Office is being installed. Please wait until it finishes.', 'Microsoft Office Installer', 'OK', 'Information')
        }
    })

    $bi.Add_Click({
        $ver = $versions[$cv.SelectedIndex]
        $lang = $state.Lang
        $sel = Get-SelectedTags

        if ($sel.Count -eq 0) {
            Set-Progress 0 'Nothing selected. Pick at least one application.' 'warn'
            [void][System.Windows.Forms.MessageBox]::Show('Select at least one application before installing.', 'Microsoft Office Installer', 'OK', 'Warning')
            return
        }

        $note = ''
        if ($ver.C -eq 'PerpetualVL2024' -and $sel -contains 'Publisher') { $note = "`n`nNote: Publisher is not part of Office LTSC 2024 and will be skipped." }
        $ask = "About to install:`n`n  Edition  : $($ver.Label)`n  Language : $lang`n  System   : $arch-bit`n  Apps     : $($sel -join ', ')$note`n`nOffice will be installed without activation. Continue?"
        if ([System.Windows.Forms.MessageBox]::Show($ask, 'Microsoft Office Installer - Confirm', 'YesNo', 'Question') -ne 'Yes') {
            Set-Progress 0 'Cancelled. Nothing was changed.' 'idle'
            return
        }

        $state.Busy = $true
        $bi.Enabled = $false; $bc.Enabled = $false; $cv.Enabled = $false; $langPanel.Enabled = $false; $ap.Enabled = $false
        try {
            Set-Progress 15 'Generating configuration...'
            $x = New-Object System.Text.StringBuilder
            [void]$x.AppendLine('<Configuration>')
            [void]$x.AppendLine("  <Add OfficeClientEdition=`"$arch`" Channel=`"$($ver.C)`">")
            $baseSel = @($apps | Where-Object { -not $_.Addon -and $sel -contains $_.I })
            if ($baseSel.Count -gt 0) {
                [void]$x.AppendLine("    <Product ID=`"$($ver.P)`">")
                [void]$x.AppendLine("      <Language ID=`"$lang`" />")
                foreach ($a in $apps) { if (-not $a.Addon -and $sel -notcontains $a.I) { [void]$x.AppendLine("      <ExcludeApp ID=`"$($a.X)`" />") } }
                foreach ($ex in $alwaysExcluded) { [void]$x.AppendLine("      <ExcludeApp ID=`"$ex`" />") }
                [void]$x.AppendLine('    </Product>')
            }
            if ($sel -contains 'Project') { [void]$x.AppendLine("    <Product ID=`"$($ver.J)`"><Language ID=`"$lang`" /></Product>") }
            if ($sel -contains 'Visio')   { [void]$x.AppendLine("    <Product ID=`"$($ver.V)`"><Language ID=`"$lang`" /></Product>") }
            [void]$x.AppendLine('  </Add>')
            [void]$x.AppendLine('  <Display Level="Full" AcceptEULA="TRUE" />')
            [void]$x.AppendLine('</Configuration>')
            $cp = Join-Path $odtTemp 'configuration.xml'
            Set-Content -Path $cp -Value $x.ToString() -Encoding UTF8
            Write-Log "Configuration written to $cp"

            Set-Progress 40 'Downloading and installing Office (this can take 10-30 minutes)...'
            Write-Log "Running setup.exe /configure $cp"
            $proc = Start-Process -FilePath (Join-Path $odtTemp 'setup.exe') -ArgumentList "/configure `"$cp`"" -PassThru
            $null = $proc.Handle
            Wait-UI { $proc.HasExited }

            if ($proc.ExitCode -eq 0) {
                Set-Progress 100 'Done. Office installed successfully.' 'ok'
                Write-Log 'Office installed successfully' 'Green'
                [void][System.Windows.Forms.MessageBox]::Show('Office installed successfully.', 'Microsoft Office Installer - Success', 'OK', 'Information')
            } else {
                Set-Progress 100 "Installation failed (code $($proc.ExitCode))." 'bad'
                Write-Log "setup.exe failed with exit code $($proc.ExitCode)" 'Red'
                [void][System.Windows.Forms.MessageBox]::Show("Installation failed (code $($proc.ExitCode)).`n`nOffice setup logs are in:`n$env:TEMP", 'Microsoft Office Installer - Error', 'OK', 'Error')
            }
        } catch {
            Set-Progress 0 "Error: $($_.Exception.Message)" 'bad'
            Write-Log "Error: $_" 'Red'
            [void][System.Windows.Forms.MessageBox]::Show("$($_.Exception.Message)", 'Microsoft Office Installer - Error', 'OK', 'Error')
        } finally {
            $state.Busy = $false
            $bi.Enabled = $true; $bc.Enabled = $true; $cv.Enabled = $true; $langPanel.Enabled = $true; $ap.Enabled = $true
        }
    })

    [void]$f.ShowDialog()
    $f.Dispose()
}

# --------------------------------------------------------------------- main --
try {
    Write-Log '=== Microsoft Office Installer (GUI) started ===' 'Cyan'
    Show-PrepDialog
    if ($state.OdtReady) { Show-MainForm }
} catch {
    Write-Log "Error: $_" 'Red'
    [void][System.Windows.Forms.MessageBox]::Show("$($_.Exception.Message)", 'Microsoft Office Installer - Error', 'OK', 'Error')
} finally {
    Remove-Temp
    Write-Log '=== Microsoft Office Installer (GUI) finished ===' 'Cyan'
}
}
