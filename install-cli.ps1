# Microsoft Office Installer - Command Line Edition
# Downloads the official Office Deployment Tool (ODT) from Microsoft and installs
# the selected Office edition. It only installs Office; it does NOT activate it.
#
# Usage:  irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-cli.ps1 | iex
#    or:  .\install-cli.ps1
#
# Everything runs inside a script block so that nothing leaks into the caller's
# session and no 'exit' closes the user's PowerShell window when run via iex.

& {
$ErrorActionPreference = 'Stop'
$ScriptUrl = 'https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/install-cli.ps1'

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {}

# ---------------------------------------------------------------- elevation --
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host ""
    Write-Host "  Administrator rights are required. Opening an elevated window..." -ForegroundColor Yellow
    if ($PSCommandPath) {
        $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    } else {
        $argList = "-NoProfile -ExecutionPolicy Bypass -Command `"[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; irm '$ScriptUrl' | iex`""
    }
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argList | Out-Null
    } catch {
        Write-Host "  Elevation was cancelled. Nothing was changed." -ForegroundColor Red
    }
    Write-Host ""
    return
}

# ------------------------------------------------------------------- config --
$logFile    = Join-Path $env:TEMP 'OfficeInstallerCLI_Install.log'
$odtTemp    = Join-Path $env:TEMP 'OfficeInstaller_ODT'
$odtExe     = Join-Path $env:TEMP 'OfficeInstaller_ODT.exe'
$odtPage    = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
$odtFallback = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe'

# C = channel, P/V/J = product IDs for Office / Visio / Project.
# Office 2016 has no volume Click-to-Run edition, so the retail C2R product is used.
# Office 2013 is not offered: it needs the retired 2013 ODT and is out of support.
$versions = @(
    @{ Label = 'Office LTSC Professional Plus 2024'; C = 'PerpetualVL2024'; P = 'ProPlus2024Volume'; V = 'VisioPro2024Volume'; J = 'ProjectPro2024Volume' }
    @{ Label = 'Office LTSC Professional Plus 2021'; C = 'PerpetualVL2021'; P = 'ProPlus2021Volume'; V = 'VisioPro2021Volume'; J = 'ProjectPro2021Volume' }
    @{ Label = 'Office Professional Plus 2019';      C = 'PerpetualVL2019'; P = 'ProPlus2019Volume'; V = 'VisioPro2019Volume'; J = 'ProjectPro2019Volume' }
    @{ Label = 'Office Professional Plus 2016';      C = 'Current';         P = 'ProPlusRetail';     V = 'VisioProRetail';     J = 'ProjectProRetail' }
)

$languages = @(
    @{ Code = 'en-US'; Name = 'English' }
    @{ Code = 'es-ES'; Name = 'Spanish' }
    @{ Code = 'fr-FR'; Name = 'French' }
    @{ Code = 'de-DE'; Name = 'German' }
    @{ Code = 'pt-BR'; Name = 'Brazilian Portuguese' }
    @{ Code = 'it-IT'; Name = 'Italian' }
    @{ Code = 'nl-NL'; Name = 'Dutch' }
    @{ Code = 'pl-PL'; Name = 'Polish' }
    @{ Code = 'ru-RU'; Name = 'Russian' }
    @{ Code = 'ja-JP'; Name = 'Japanese' }
)

# Tag = what the user picks; Exclude = the official ODT ExcludeApp ID.
$appCatalog = @(
    @{ Tag = 'Word';             Label = 'Word';               Exclude = 'Word';       Addon = $false }
    @{ Tag = 'Excel';            Label = 'Excel';              Exclude = 'Excel';      Addon = $false }
    @{ Tag = 'PowerPoint';       Label = 'PowerPoint';         Exclude = 'PowerPoint'; Addon = $false }
    @{ Tag = 'Outlook';          Label = 'Outlook';            Exclude = 'Outlook';    Addon = $false }
    @{ Tag = 'Access';           Label = 'Access';             Exclude = 'Access';     Addon = $false }
    @{ Tag = 'Publisher';        Label = 'Publisher';          Exclude = 'Publisher';  Addon = $false }
    @{ Tag = 'OneNote';          Label = 'OneNote';            Exclude = 'OneNote';    Addon = $false }
    @{ Tag = 'SkypeForBusiness'; Label = 'Skype for Business'; Exclude = 'Lync';       Addon = $false }
    @{ Tag = 'Project';          Label = 'Project';            Exclude = $null;        Addon = $true }
    @{ Tag = 'Visio';            Label = 'Visio';              Exclude = $null;        Addon = $true }
)
$alwaysExcluded = @('Groove', 'OneDrive', 'Teams')

# ------------------------------------------------------------------ helpers --
function Color {
    param([string]$T, [string]$C = 'White')
    Write-Host $T -ForegroundColor $C -NoNewline
}
function Rule { Color ('  ' + ([string][char]0x2500) * 74) 'DarkGray'; Write-Host '' }
function Ok   { param([string]$M) Color "  $([char]0x2714) " 'Green';  Color $M 'White'; Write-Host '' }
function Fail { param([string]$M) Color "  $([char]0x2718) " 'Red';    Color $M 'Red';   Write-Host '' }
function Info { param([string]$M) Color "  $([char]0x25B8) " 'Cyan';   Color $M 'White'; Write-Host '' }
function Warn { param([string]$M) Color "  !  " 'Yellow';              Color $M 'Yellow'; Write-Host '' }
function Dim  { param([string]$M) Color "    $M" 'DarkGray'; Write-Host '' }

function Write-Log {
    param([string]$M)
    try { Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $M" } catch {}
}

function Remove-Temp {
    if (Test-Path $odtTemp) { Remove-Item $odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $odtExe)  { Remove-Item $odtExe -Force -ErrorAction SilentlyContinue }
}

function Wait-Exit {
    Write-Host ''
    Color '  Press Enter to close ' 'DarkGray'
    try { [void](Read-Host) } catch {}
}

function Show-Banner {
    Clear-Host
    $W = 66
    $h = [string][char]0x2550
    function Box([string]$txt, [string]$col) {
        $pad = $W - $txt.Length
        $l = [Math]::Floor($pad / 2); $r = $pad - $l
        Color ("  $([char]0x2551)" + (' ' * $l)) 'Cyan'
        Color $txt $col
        Color ((' ' * $r) + "$([char]0x2551)") 'Cyan'
        Write-Host ''
    }
    Write-Host ''
    Color ("  $([char]0x2554)" + ($h * $W) + "$([char]0x2557)") 'Cyan'; Write-Host ''
    Box 'MICROSOFT OFFICE INSTALLER' 'White'
    Box 'Command Line Edition' 'Cyan'
    Color ("  $([char]0x255A)" + ($h * $W) + "$([char]0x255D)") 'Cyan'; Write-Host ''
    $by = 'by Matthew Garay'
    Color ((' ' * [Math]::Floor((($W + 4) - $by.Length) / 2)) + $by) 'DarkGray'; Write-Host ''
    Write-Host ''
    Rule
}

function Show-Step {
    param([int]$N, [int]$Total, [string]$Title)
    Write-Host ''
    Color "  $([char]0x25CF) STEP $N/$Total " 'Cyan'
    Color "$([char]0x2500) " 'DarkGray'
    Color $Title 'White'
    Write-Host ''
    Write-Host ''
}

function Ask-Choice {
    param([string[]]$Options, [int]$Default = 1)
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Color ('    {0,2}. ' -f ($i + 1)) 'DarkGray'
        Color $Options[$i] 'White'
        Write-Host ''
    }
    Write-Host ''
    Dim "Enter a number 1-$($Options.Count), or press Enter for $Default"
    while ($true) {
        Color '  > ' 'Cyan'
        $raw = Read-Host
        if ($null -eq $raw -or -not $raw.Trim()) { return $Default - 1 }
        $n = 0
        if ([int]::TryParse($raw.Trim(), [ref]$n) -and $n -ge 1 -and $n -le $Options.Count) { return $n - 1 }
        Fail "Not one of 1-$($Options.Count): $raw"
    }
}

function Ask-Apps {
    for ($i = 0; $i -lt $appCatalog.Count; $i++) {
        $a = $appCatalog[$i]
        Color ('    {0,2}. ' -f ($i + 1)) 'DarkGray'
        Color $a.Label 'White'
        if ($a.Addon) { Color '  (separate product)' 'DarkGray' }
        Write-Host ''
    }
    Write-Host ''
    Dim 'Numbers separated by commas, e.g. 1,2,3 = Word, Excel, PowerPoint.'
    Dim 'Press Enter for 1,2,3. Type "all" for everything.'
    while ($true) {
        Color '  > ' 'Cyan'
        $raw = Read-Host
        if ($null -eq $raw -or -not $raw.Trim()) { $raw = '1,2,3' }
        $raw = $raw.Trim()
        if ($raw -eq 'all') { return @($appCatalog | ForEach-Object { $_.Tag }) }
        $picked = @(); $bad = $false
        foreach ($part in ($raw -split '[,\s]+')) {
            if (-not $part) { continue }
            $n = 0
            if (-not [int]::TryParse($part, [ref]$n) -or $n -lt 1 -or $n -gt $appCatalog.Count) {
                Fail "Not a valid option: $part"; $bad = $true; continue
            }
            $tag = $appCatalog[$n - 1].Tag
            if ($picked -notcontains $tag) { $picked += $tag }
        }
        if ($bad) { continue }
        if ($picked.Count -eq 0) { Fail 'Nothing selected.'; continue }
        return $picked
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
        Write-Log "Could not read the ODT download page: $_"
    }
    $urls += $odtFallback
    return @($urls | Select-Object -Unique)
}

function Initialize-Odt {
    Remove-Temp
    New-Item -ItemType Directory -Path $odtTemp -Force | Out-Null

    Info 'Looking up the latest Office Deployment Tool...'
    $downloaded = $false
    foreach ($u in (Get-OdtUrls)) {
        try {
            Info 'Downloading from download.microsoft.com...'
            Write-Log "Downloading ODT from $u"
            (New-Object System.Net.WebClient).DownloadFile($u, $odtExe)
            $downloaded = $true
            break
        } catch {
            Write-Log "Source failed: $u ($_)"
            Warn 'Source unavailable, trying the next one.'
        }
    }
    if (-not $downloaded) { throw 'Could not download the Office Deployment Tool.' }

    $sig = Get-AuthenticodeSignature -FilePath $odtExe
    if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        throw "The downloaded file is not signed by Microsoft (signature: $($sig.Status))."
    }
    Ok 'Download verified (signed by Microsoft)'

    Info 'Extracting...'
    $p = Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:`"$odtTemp`"" -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw "Extraction failed (exit code $($p.ExitCode))." }
    if (-not (Test-Path (Join-Path $odtTemp 'setup.exe'))) { throw 'setup.exe not found after extraction.' }
    Write-Log "ODT ready at $odtTemp"
}

function New-ConfigurationXml {
    param($Ver, [string]$Lang, [string[]]$Selected, [string]$Arch)
    $x = New-Object System.Text.StringBuilder
    [void]$x.AppendLine('<Configuration>')
    [void]$x.AppendLine("  <Add OfficeClientEdition=`"$Arch`" Channel=`"$($Ver.C)`">")
    $baseSelected = @($appCatalog | Where-Object { -not $_.Addon -and $Selected -contains $_.Tag })
    if ($baseSelected.Count -gt 0) {
        [void]$x.AppendLine("    <Product ID=`"$($Ver.P)`">")
        [void]$x.AppendLine("      <Language ID=`"$Lang`" />")
        foreach ($a in $appCatalog) {
            if (-not $a.Addon -and $Selected -notcontains $a.Tag) { [void]$x.AppendLine("      <ExcludeApp ID=`"$($a.Exclude)`" />") }
        }
        foreach ($ex in $alwaysExcluded) { [void]$x.AppendLine("      <ExcludeApp ID=`"$ex`" />") }
        [void]$x.AppendLine('    </Product>')
    }
    if ($Selected -contains 'Project') { [void]$x.AppendLine("    <Product ID=`"$($Ver.J)`"><Language ID=`"$Lang`" /></Product>") }
    if ($Selected -contains 'Visio')   { [void]$x.AppendLine("    <Product ID=`"$($Ver.V)`"><Language ID=`"$Lang`" /></Product>") }
    [void]$x.AppendLine('  </Add>')
    [void]$x.AppendLine('  <Display Level="Full" AcceptEULA="TRUE" />')
    [void]$x.AppendLine('</Configuration>')
    $path = Join-Path $odtTemp 'configuration.xml'
    Set-Content -Path $path -Value $x.ToString() -Encoding UTF8
    Write-Log "Configuration written to $path"
    return $path
}

# --------------------------------------------------------------------- main --
try {
    Write-Log '=== Microsoft Office Installer (CLI) started ==='
    Write-Log "PowerShell: $($PSVersionTable.PSVersion)"

    $existing = Get-ExistingOffice
    $arch = if ([Environment]::Is64BitOperatingSystem) { '64' } else { '32' }
    if ($existing -and $existing.Platform -eq 'x86') { $arch = '32' }

    Show-Banner

    if ($existing) {
        Warn "Office is already installed: $($existing.Products)"
        Dim 'Installing a different version or edition next to it may fail.'
        Dim "Using $arch-bit to match the installed Office."
        Write-Host ''
    }

    Show-Step 1 4 'Select the Office version'
    $ver = $versions[(Ask-Choice -Options @($versions | ForEach-Object { $_.Label }))]
    Ok $ver.Label

    Show-Step 2 4 'Select the language'
    $langIndex = Ask-Choice -Options @($languages | ForEach-Object { "$($_.Name) ($($_.Code))" })
    $lang = $languages[$langIndex].Code
    Ok $lang

    Show-Step 3 4 'Select the applications'
    $selected = Ask-Apps
    foreach ($t in $selected) { Dim "+ $(($appCatalog | Where-Object { $_.Tag -eq $t }).Label)" }
    if ($ver.C -eq 'PerpetualVL2024' -and $selected -contains 'Publisher') {
        Warn 'Publisher is not part of Office LTSC 2024 and will be skipped.'
    }

    Write-Host ''
    Color '  Summary' 'White'; Write-Host ''
    Dim "Edition  : $($ver.Label)"
    Dim "Language : $lang"
    Dim "Apps     : $(($selected | ForEach-Object { $t = $_; ($appCatalog | Where-Object { $_.Tag -eq $t }).Label }) -join ', ')"
    Dim "System   : $arch-bit"
    Dim 'Office is installed only. Activate it with your own license key.'
    Write-Host ''
    Color '  Press Enter to start, or type n to cancel: ' 'White'
    $ans = Read-Host
    if ($ans -and $ans.Trim().ToLower() -in @('n', 'no')) {
        Info 'Cancelled. Nothing was changed.'
        Write-Log 'Cancelled by the user'
        Wait-Exit
        return
    }

    Show-Step 4 4 'Download and install'
    Initialize-Odt
    Ok 'Office Deployment Tool ready'

    $configPath = New-ConfigurationXml -Ver $ver -Lang $lang -Selected $selected -Arch $arch
    Info 'Installing Office. The Office setup window will show the progress...'
    Dim 'This can take 10-30 minutes depending on your connection.'
    Write-Log "Running setup.exe /configure $configPath"

    $proc = Start-Process -FilePath (Join-Path $odtTemp 'setup.exe') -ArgumentList "/configure `"$configPath`"" -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        Fail "Installation failed (exit code $($proc.ExitCode))."
        Dim "Office setup logs: $env:TEMP (files named <PC name>-<date>.log)"
        Write-Log "setup.exe failed with exit code $($proc.ExitCode)"
    } else {
        Ok 'Office installed successfully.'
        Write-Log 'Office installed successfully'
    }
} catch {
    Write-Host ''
    Fail "$($_.Exception.Message)"
    Write-Log "Error: $_"
    Dim "Log file: $logFile"
} finally {
    Remove-Temp
    Write-Log '=== Microsoft Office Installer (CLI) finished ==='
}
Wait-Exit
}
