<#
.SYNOPSIS
    VSFI (Very Simple First Installation) by Artus Poulain
    Le script d'installation ultime pour makers et devs

.DESCRIPTION
    Installation personnalisable de l'environnement dev / maker / creatif
    - Interface graphique pour choisir les applications
    - Support Winget, Microsoft Store et Chocolatey
    - Fallback automatique sur Chocolatey si Winget echoue
    - 60+ applications pre-configurees
    
.PARAMETER NoPrompt
    Installe tous les paquets marques "par defaut" sans demander

.PARAMETER SelectAll
    Pre-selectionne tous les paquets dans l'interface
    
.PARAMETER SkipReboot
    Ne propose pas le redemarrage a la fin

.EXAMPLE
    .\VSFI.ps1
    .\VSFI.ps1 -NoPrompt
    .\VSFI.ps1 -SelectAll

.NOTES
    Auteur  : Artus Poulain
    Version : 3.0
#>

param(
    [switch]$NoPrompt,
    [switch]$SelectAll,
    [switch]$SkipReboot
)

# ============================================================================
# AUTO-ELEVATION ADMIN + MODE STA
# ============================================================================
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    $argList = @('-ExecutionPolicy', 'Bypass', '-NoProfile', '-STA', '-File', "`"$PSCommandPath`"")
    if ($NoPrompt)   { $argList += '-NoPrompt' }
    if ($SelectAll)  { $argList += '-SelectAll' }
    if ($SkipReboot) { $argList += '-SkipReboot' }
    Start-Process powershell.exe -Verb RunAs -ArgumentList $argList
    exit
}

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $argList = @('-ExecutionPolicy', 'Bypass', '-NoProfile', '-STA', '-File', "`"$PSCommandPath`"")
    if ($NoPrompt)   { $argList += '-NoPrompt' }
    if ($SelectAll)  { $argList += '-SelectAll' }
    if ($SkipReboot) { $argList += '-SkipReboot' }
    Start-Process powershell.exe -ArgumentList $argList -Wait -NoNewWindow
    exit
}

# ============================================================================
# CONFIG
# ============================================================================
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
$ProgressPreference = "SilentlyContinue"
$ErrorActionPreference = "Continue"
try {
    $Host.UI.RawUI.WindowTitle = "VSFI by Artus Poulain"
    $Host.UI.RawUI.BackgroundColor = 'Black'
    $Host.UI.RawUI.ForegroundColor = 'White'
    Clear-Host
} catch { }
try { Add-Type -AssemblyName System.Windows.Forms } catch { }

$script:LogFile = Join-Path $env:TEMP ("VSFI-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")

# ============================================================================
# FONCTIONS D'AFFICHAGE
# ============================================================================

function Show-Banner {
    try {
        $raw = $Host.UI.RawUI
        $minWidth = 100
        $win = $raw.WindowSize
        $buf = $raw.BufferSize
        $newWidth = [math]::Max($minWidth, $win.Width)
        if ($buf.Width -lt $newWidth) {
            $raw.BufferSize = New-Object Management.Automation.Host.Size($newWidth, $buf.Height)
        }
        if ($win.Width -lt $newWidth) {
            $raw.WindowSize = New-Object Management.Automation.Host.Size($newWidth, $win.Height)
        }
    } catch { }
    Clear-Host
    $bar = "=" * 60
    Write-Host ""
    Write-Host "  $bar" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "   VSFI - Very Simple First Installation" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   by Artus Poulain" -ForegroundColor Magenta
    Write-Host "   youtube.com/LesFreresPoulain" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  $bar" -ForegroundColor DarkCyan
    Write-Host ""
    [Console]::Out.Flush()
}

function Write-Step {
    param([string]$Step, [string]$Title)
    Write-Host ""
    Write-Host "  $("=" * 60)" -ForegroundColor DarkCyan
    Write-Host "  >> " -NoNewline -ForegroundColor Cyan
    Write-Host "ETAPE $Step " -NoNewline -ForegroundColor White
    Write-Host "- $Title" -ForegroundColor Yellow
    Write-Host "  $("=" * 60)" -ForegroundColor DarkCyan
    Write-Host ""
    [Console]::Out.Flush()
}

function Write-SectionHeader {
    param([string]$Title, [int]$Count)
    Write-Host ""
    Write-Host "    --- " -NoNewline -ForegroundColor DarkYellow
    Write-Host "$Title" -NoNewline -ForegroundColor Yellow
    Write-Host " ($Count apps) " -NoNewline -ForegroundColor DarkGray
    Write-Host "$("-" * 20)" -ForegroundColor DarkYellow
    Write-Host ""
    [Console]::Out.Flush()
}

function Write-AppResult {
    param([string]$Name, [string]$Status)
    
    switch ($Status) {
        "OK"       { $icon = "[+]"; $color = "Green" }
        "SKIP"     { $icon = "[-]"; $color = "DarkGray" }
        "FAIL"     { $icon = "[X]"; $color = "Red" }
        "FALLBACK" { $icon = "[~]"; $color = "Magenta" }
        default    { $icon = "[?]"; $color = "Gray" }
    }
    
    $padding = 45 - $Name.Length
    if ($padding -lt 1) { $padding = 1 }
    $dots = "." * $padding
    
    Write-Host "    " -NoNewline
    Write-Host "$icon " -NoNewline -ForegroundColor $color
    Write-Host "$Name" -NoNewline -ForegroundColor White
    Write-Host " $dots " -NoNewline -ForegroundColor DarkGray
    Write-Host "$Status" -ForegroundColor $color
    [Console]::Out.Flush()
}

function Write-Installing {
    param([string]$Name, [string]$Source, [int]$Index, [int]$Total)
    Write-Host "    >>> ($Index/$Total) $Name via $Source ..." -ForegroundColor Yellow
    [Console]::Out.Flush()
}

function Write-Log {
    param([string]$Message, [string]$Color = "Gray")
    Write-Host "    " -NoNewline
    Write-Host "* " -NoNewline -ForegroundColor DarkGray
    Write-Host "$Message" -ForegroundColor $Color
    [Console]::Out.Flush()
}

function Write-Success { 
    param([string]$Message) 
    Write-Host "    " -NoNewline
    Write-Host "[OK] " -NoNewline -ForegroundColor Green
    Write-Host "$Message" -ForegroundColor Green 
    [Console]::Out.Flush()
}

function Write-Warning2 { 
    param([string]$Message) 
    Write-Host "    " -NoNewline
    Write-Host "[!] " -NoNewline -ForegroundColor Yellow
    Write-Host "$Message" -ForegroundColor Yellow 
    [Console]::Out.Flush()
}

function Write-Error2 { 
    param([string]$Message) 
    Write-Host "    " -NoNewline
    Write-Host "[X] " -NoNewline -ForegroundColor Red
    Write-Host "$Message" -ForegroundColor Red 
    [Console]::Out.Flush()
}

function Show-Summary {
    param($Stats)
    $total = $Stats.Success + $Stats.Fallback + $Stats.Skip + $Stats.Fail
    Write-Host ""
    Write-Host "  $("=" * 60)" -ForegroundColor Green
    Write-Host "  " -NoNewline
    Write-Host "[OK] INSTALLATION TERMINEE" -NoNewline -ForegroundColor Green
    if ($total -gt 0) { Write-Host "  ($total applications traitees)" -ForegroundColor DarkGray } else { Write-Host "" }
    Write-Host "  $("=" * 60)" -ForegroundColor Green
    Write-Host ""
    Write-Host "    Installes      : " -NoNewline -ForegroundColor Green
    Write-Host "$($Stats.Success + $Stats.Fallback)" -NoNewline -ForegroundColor White
    if ($Stats.Fallback -gt 0) {
        Write-Host " (dont $($Stats.Fallback) via fallback Choco)" -ForegroundColor Magenta
    } else {
        Write-Host ""
    }
    Write-Host "    Deja presents  : " -NoNewline -ForegroundColor DarkGray
    Write-Host "$($Stats.Skip)" -ForegroundColor White
    $failColor = if ($Stats.Fail -gt 0) { "Red" } else { "DarkGray" }
    Write-Host "    Echecs         : " -NoNewline -ForegroundColor $failColor
    Write-Host "$($Stats.Fail)" -ForegroundColor White
    Write-Host ""
    [Console]::Out.Flush()
}

function Write-FileLog {
    param([Parameter(Mandatory=$true)][string]$Message)
    try {
        $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8
    } catch { }
}

# ============================================================================
# FONCTIONS UTILITAIRES
# ============================================================================

function Test-Command {
    param([Parameter(Mandatory=$true)][string]$cmd)
    return [bool](Get-Command $cmd -ErrorAction SilentlyContinue)
}

function Get-WingetExePath {
    $candidates = @()
    try { $candidates += Get-Command winget.exe -All -ErrorAction SilentlyContinue } catch { }
    try { $candidates += Get-Command winget -All -ErrorAction SilentlyContinue } catch { }

    foreach ($c in $candidates) {
        if ($null -ne $c -and $c.CommandType -eq 'Application' -and $c.Source -match '(?i)winget\.exe$') {
            return $c.Source
        }
    }

    $fallback = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
    if (Test-Path -LiteralPath $fallback) { return $fallback }

    return $null
}

function Invoke-Winget {
    param([Parameter(ValueFromRemainingArguments=$true)][string[]]$WingetArgs)

    if (-not $script:WingetExe) {
        $script:WingetExe = Get-WingetExePath
    }

    if (-not $script:WingetExe) {
        throw "Winget introuvable"
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $script:WingetExe
    $psi.Arguments = ($WingetArgs | ForEach-Object {
        if ($_ -match '\s') { "`"$_`"" } else { $_ }
    }) -join ' '
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.StandardInput.Close()

    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
    $stderrTask = $proc.StandardError.ReadToEndAsync()

    $isSTA = ([System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA')
    while (-not $proc.HasExited) {
        if ($isSTA) {
            try { [System.Windows.Forms.Application]::DoEvents() } catch { }
        }
        Start-Sleep -Milliseconds 50
    }
    $proc.WaitForExit()

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $null = $stderrTask.GetAwaiter().GetResult()

    $global:LASTEXITCODE = $proc.ExitCode
    $proc.Dispose()

    if ($stdout) { $stdout.TrimEnd() }
}

function Test-WingetInstalled($packageId) {
    $result = Invoke-Winget -WingetArgs @('list','--id',$packageId,'--exact')
    if ($LASTEXITCODE -ne 0) {
        $result = Invoke-Winget -WingetArgs @('list','--source','msstore','--id',$packageId,'--exact')
        if ($LASTEXITCODE -ne 0) { return $false }
    }
    $text = if ($result) { $result } else { "" }
    if ($text -match "No installed package found" -or $text -match "Aucun paquet install" -or $text -match "Aucun package install") { return $false }
    return ($text -match [regex]::Escape($packageId))
}

function Test-ChocoInstalled($packageName) {
    $result = choco list --exact $packageName 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $text = ($result | Out-String)
    return ($text -match "(?m)^$([regex]::Escape($packageName))[\s|]")
}

function Initialize-WingetAgreements {
    try {
        $null = Invoke-Winget -WingetArgs @('search','--accept-source-agreements','--disable-interactivity','yourphoneupdater')
        $null = Invoke-Winget -WingetArgs @('source','update','--disable-interactivity')
    } catch { }
}

function Export-InstalledApps {
    param([string]$OutputPath)

    Show-Banner
    Write-Step "BACKUP" "Analyse des logiciels installes"

    $allApps = @()

    # Winget (via winget export -> JSON)
    Write-Log "Scan des applications Winget..." Cyan
    try {
        $tmpExport = Join-Path $env:TEMP "vsfi-export-$([guid]::NewGuid().ToString('N')).json"
        $null = Invoke-Winget -WingetArgs @('export','-o',$tmpExport,'--disable-interactivity','--accept-source-agreements')
        if (Test-Path $tmpExport) {
            $exportData = Get-Content -Path $tmpExport -Raw | ConvertFrom-Json
            foreach ($source in $exportData.Sources) {
                $srcName = if ($source.SourceDetails.Name) { $source.SourceDetails.Name.ToLower() } else { 'winget' }
                foreach ($pkg in $source.Packages) {
                    $pkgId = $pkg.PackageIdentifier
                    if ($pkgId) {
                        $allApps += [PSCustomObject]@{
                            Name    = $pkgId
                            Id      = $pkgId
                            Version = ''
                            Source  = $srcName
                        }
                    }
                }
            }
            Remove-Item $tmpExport -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Warning2 "Erreur lors du scan Winget: $_"
    }
    Write-Success "$($allApps.Count) applications detectees via Winget"

    # Chocolatey
    $chocoCount = 0
    if (Test-Command choco) {
        Write-Log "Scan des applications Chocolatey..." Cyan
        try {
            $chocoRaw = (choco list 2>&1 | Out-String)
            $chocoLines = $chocoRaw -split "`n"
            foreach ($line in $chocoLines) {
                if ($line -match '^([a-zA-Z0-9._-]+)\s+([\d.]+)') {
                    $chocoId = $Matches[1]
                    $chocoVer = $Matches[2]
                    if ($chocoId -eq 'chocolatey') { continue }
                    $alreadyFound = $false
                    foreach ($existing in $allApps) {
                        if ($existing.Id -eq $chocoId) { $alreadyFound = $true; break }
                        $mapped = $WingetToChocoMap[$existing.Id]
                        if ($mapped -and $mapped -eq $chocoId) { $alreadyFound = $true; break }
                    }
                    if (-not $alreadyFound) {
                        $allApps += [PSCustomObject]@{
                            Name    = $chocoId
                            Id      = $chocoId
                            Version = $chocoVer
                            Source  = 'choco'
                        }
                        $chocoCount++
                    }
                }
            }
        } catch {
            Write-Warning2 "Erreur lors du scan Chocolatey: $_"
        }
        Write-Success "$chocoCount applications supplementaires via Chocolatey"
    }

    # Sauvegarder
    $export = [PSCustomObject]@{
        ExportDate   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ComputerName = $env:COMPUTERNAME
        UserName     = $env:USERNAME
        TotalApps    = $allApps.Count
        Applications = $allApps
    }

    if (-not $OutputPath) {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = "VSFI - Sauvegarder le backup"
        $sfd.Filter = "JSON (*.json)|*.json"
        $sfd.FileName = "vsfi-backup-$(Get-Date -Format 'yyyyMMdd').json"
        $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if ($sfd.ShowDialog() -ne 'OK') {
            Write-Warning2 "Backup annule."
            return
        }
        $OutputPath = $sfd.FileName
    }

    $export | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host ""
    Write-Success "Backup sauvegarde : $OutputPath"
    Write-Log "$($allApps.Count) applications exportees" Cyan
    Write-Host ""
    [Console]::Out.Flush()
}

function Show-WelcomeDialog {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    } catch { return 'install' }

    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { return 'install' }

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $bgDark   = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $fgWhite  = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $fgGray   = [System.Drawing.Color]::FromArgb(170, 170, 170)
    $cyan     = [System.Drawing.Color]::FromArgb(0, 180, 216)
    $green    = [System.Drawing.Color]::FromArgb(80, 200, 120)
    $orange   = [System.Drawing.Color]::FromArgb(230, 160, 50)
    $red      = [System.Drawing.Color]::FromArgb(220, 80, 80)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "VSFI by Artus Poulain"
    $form.StartPosition = 'CenterScreen'
    $form.Width = 500
    $form.Height = 380
    $form.BackColor = $bgDark
    $form.ForeColor = $fgWhite
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Definition)

    $script:welcomeResult = $null

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "VSFI"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $cyan
    $lblTitle.AutoSize = $false
    $lblTitle.TextAlign = 'MiddleCenter'
    $lblTitle.Dock = 'None'
    $lblTitle.Size = New-Object System.Drawing.Size(460, 50)
    $lblTitle.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($lblTitle)

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = "by Artus Poulain"
    $lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
    $lblSub.ForeColor = $fgGray
    $lblSub.AutoSize = $false
    $lblSub.TextAlign = 'MiddleCenter'
    $lblSub.Size = New-Object System.Drawing.Size(460, 25)
    $lblSub.Location = New-Object System.Drawing.Point(20, 70)
    $form.Controls.Add($lblSub)

    $makeBtn = {
        param($text, $desc, $bgColor, $y)
        $btn = New-Object System.Windows.Forms.Button
        $btn.FlatStyle = 'Flat'
        $btn.FlatAppearance.BorderSize = 0
        $btn.BackColor = $bgColor
        $btn.ForeColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
        $btn.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $btn.Text = $text
        $btn.Size = New-Object System.Drawing.Size(400, 48)
        $btn.Location = New-Object System.Drawing.Point(50, $y)
        $btn.Cursor = 'Hand'
        $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
            [math]::Min($bgColor.R + 30, 255),
            [math]::Min($bgColor.G + 30, 255),
            [math]::Min($bgColor.B + 30, 255))
        $btn
    }

    $btnInstall = & $makeBtn "Installer des applications" "" $green 120
    $btnInstall.Add_Click({ $script:welcomeResult = 'install'; $form.Close() })
    $form.Controls.Add($btnInstall)

    $btnBackup = & $makeBtn "Sauvegarder mes logiciels (Backup)" "" $orange 180
    $btnBackup.Add_Click({ $script:welcomeResult = 'backup'; $form.Close() })
    $form.Controls.Add($btnBackup)

    $btnQuit = & $makeBtn "Quitter" "" $red 240
    $btnQuit.ForeColor = [System.Drawing.Color]::White
    $btnQuit.Add_Click({ $script:welcomeResult = $null; $form.Close() })
    $form.Controls.Add($btnQuit)

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "v3.0 - youtube.com/LesFreresPoulain"
    $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $lblVersion.AutoSize = $false
    $lblVersion.TextAlign = 'MiddleCenter'
    $lblVersion.Size = New-Object System.Drawing.Size(460, 20)
    $lblVersion.Location = New-Object System.Drawing.Point(20, 305)
    $form.Controls.Add($lblVersion)

    $form.ShowDialog() | Out-Null
    return $script:welcomeResult
}

function Show-AppSelectionDialog {
    param([array]$Apps, [switch]$SelectAll, [string]$ImportFile)

    $importIds = @{}
    if ($ImportFile -and (Test-Path $ImportFile)) {
        try {
            $importData = Get-Content -Path $ImportFile -Raw | ConvertFrom-Json
            foreach ($app in $importData.Applications) { $importIds[$app.Id] = $true }
            $catalogIds = @{}
            foreach ($a in $Apps) { $catalogIds[$a.Id] = $true }
            $extraApps = @()
            foreach ($app in $importData.Applications) {
                if (-not $catalogIds.ContainsKey($app.Id)) {
                    $extraApps += @{
                        Category    = "99-Import"
                        Name        = if ($app.Name) { $app.Name } else { $app.Id }
                        Id          = $app.Id
                        Source      = if ($app.Source) { $app.Source } else { "winget" }
                        Default     = $false
                        Description = "Importe depuis backup"
                    }
                }
            }
            if ($extraApps.Count -gt 0) {
                $Apps = @($Apps) + $extraApps
                Write-Warning2 "$($extraApps.Count) applications hors catalogue ajoutees depuis l'import"
            }
            Write-Success "$($importIds.Count) applications chargees depuis le fichier d'import"
        } catch {
            Write-Warning2 "Erreur lecture fichier import: $_"
        }
    }

    $defaultSelection = $Apps | Where-Object { if ($SelectAll) { $true } else { $_.Default -eq $true } }

    $canShowWinForms = $true
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    } catch {
        $canShowWinForms = $false
    }

    if ($canShowWinForms -and ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA')) {
        $canShowWinForms = $false
    }

    if (-not $canShowWinForms) {
        if (Test-Command Out-GridView) {
            $gridData = $Apps | ForEach-Object {
                [PSCustomObject]@{
                    "Categorie"   = $_.Category
                    "Application" = $_.Name
                    "Description" = $_.Description
                    "Source"      = $_.Source
                    "Id"          = $_.Id
                }
            } | Sort-Object Categorie, Application
            return $gridData | Out-GridView -Title "VSFI - Selection des applications by Artus Poulain" -OutputMode Multiple
        }

        Write-Warning2 "Interface graphique indisponible : installation sur selection par defaut"
        return $defaultSelection | ForEach-Object { [PSCustomObject]@{ Id = $_.Id } }
    }

    [System.Windows.Forms.Application]::EnableVisualStyles()

    # -- Couleurs du theme --
    $bgDark      = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $bgMedium    = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $bgLight     = [System.Drawing.Color]::FromArgb(55, 55, 60)
    $bgRow1      = [System.Drawing.Color]::FromArgb(37, 37, 40)
    $bgRow2      = [System.Drawing.Color]::FromArgb(47, 47, 52)
    $fgWhite     = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $fgGray      = [System.Drawing.Color]::FromArgb(170, 170, 170)
    $accentCyan  = [System.Drawing.Color]::FromArgb(0, 180, 216)
    $accentGreen = [System.Drawing.Color]::FromArgb(80, 200, 120)
    $accentRed   = [System.Drawing.Color]::FromArgb(220, 80, 80)
    $fontUI      = New-Object System.Drawing.Font("Segoe UI", 9)
    $fontTitle   = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $fontSub     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $fontBtn     = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    # -- Formulaire --
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "VSFI by Artus Poulain"
    $form.StartPosition = 'CenterScreen'
    $form.Width = 1150
    $form.Height = 750
    $form.BackColor = $bgDark
    $form.ForeColor = $fgWhite
    $form.Font = $fontUI
    $form.FormBorderStyle = 'Sizable'
    $form.MinimumSize = New-Object System.Drawing.Size(900, 500)
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Definition)

    # -- Header panel --
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = 'Top'
    $headerPanel.Height = 100
    $headerPanel.BackColor = $bgMedium
    $headerPanel.Padding = New-Object System.Windows.Forms.Padding(20, 10, 20, 10)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "VSFI"
    $lblTitle.Font = $fontTitle
    $lblTitle.ForeColor = $accentCyan
    $lblTitle.AutoSize = $true
    $lblTitle.Location = New-Object System.Drawing.Point(20, 10)
    $headerPanel.Controls.Add($lblTitle)

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = "Selectionne les applications a installer puis clique Installer"
    $lblSub.Font = $fontSub
    $lblSub.ForeColor = $fgGray
    $lblSub.AutoSize = $true
    $lblSub.Location = New-Object System.Drawing.Point(22, 42)
    $headerPanel.Controls.Add($lblSub)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "0 / $($Apps.Count) selectionnees"
    $lblCount.Font = $fontBtn
    $lblCount.ForeColor = $accentGreen
    $lblCount.AutoSize = $true
    $lblCount.Anchor = 'Top, Right'
    $lblCount.Location = New-Object System.Drawing.Point(($form.Width - 420), 10)
    $headerPanel.Controls.Add($lblCount)

    # -- Barre de recherche --
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Rechercher :"
    $lblSearch.Font = $fontUI
    $lblSearch.ForeColor = $fgGray
    $lblSearch.AutoSize = $true
    $lblSearch.Location = New-Object System.Drawing.Point(22, 68)
    $headerPanel.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Font = $fontUI
    $txtSearch.BackColor = $bgLight
    $txtSearch.ForeColor = $fgWhite
    $txtSearch.BorderStyle = 'FixedSingle'
    $txtSearch.Width = 350
    $txtSearch.Height = 24
    $txtSearch.Location = New-Object System.Drawing.Point(110, 65)
    $txtSearch.Anchor = 'Top, Left, Right'
    $headerPanel.Controls.Add($txtSearch)

    # -- DataGridView --
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.BorderStyle = 'None'
    $grid.CellBorderStyle = 'SingleHorizontal'
    $grid.BackgroundColor = $bgDark
    $grid.GridColor = [System.Drawing.Color]::FromArgb(60, 60, 65)
    $grid.DefaultCellStyle.BackColor = $bgRow1
    $grid.DefaultCellStyle.ForeColor = $fgWhite
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 120, 150)
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $grid.DefaultCellStyle.Font = $fontUI
    $grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)
    $grid.AlternatingRowsDefaultCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $bgRow2
    $grid.AlternatingRowsDefaultCellStyle.ForeColor = $fgWhite
    $grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 120, 150)
    $grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $bgMedium
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $accentCyan
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle.Alignment = 'MiddleLeft'
    $grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(6, 4, 4, 4)
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersHeight = 36
    $grid.RowTemplate.Height = 30

    $colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colSelect.Name = 'Selection'
    $colSelect.HeaderText = ''
    $colSelect.AutoSizeMode = 'None'
    $colSelect.Width = 40
    $grid.Columns.Add($colSelect) | Out-Null

    $colCat = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colCat.Name = 'Categorie'
    $colCat.HeaderText = 'Categorie'
    $colCat.ReadOnly = $true
    $colCat.FillWeight = 18
    $grid.Columns.Add($colCat) | Out-Null

    $colApp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colApp.Name = 'Application'
    $colApp.HeaderText = 'Application'
    $colApp.ReadOnly = $true
    $colApp.FillWeight = 22
    $grid.Columns.Add($colApp) | Out-Null

    $colDesc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDesc.Name = 'Description'
    $colDesc.HeaderText = 'Description'
    $colDesc.ReadOnly = $true
    $colDesc.FillWeight = 40
    $grid.Columns.Add($colDesc) | Out-Null

    $colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSource.Name = 'Source'
    $colSource.HeaderText = 'Source'
    $colSource.ReadOnly = $true
    $colSource.FillWeight = 10
    $grid.Columns.Add($colSource) | Out-Null

    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.Name = 'Statut'
    $colStatus.HeaderText = 'Statut'
    $colStatus.ReadOnly = $true
    $colStatus.FillWeight = 10
    $colStatus.DefaultCellStyle.Alignment = 'MiddleCenter'
    $grid.Columns.Add($colStatus) | Out-Null

    $colId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colId.Name = 'Id'
    $colId.HeaderText = 'Id'
    $colId.ReadOnly = $true
    $colId.Visible = $false
    $grid.Columns.Add($colId) | Out-Null

    # -- Detection des apps deja installees (batch) --
    Write-Log "Detection des applications deja installees..." Cyan
    $installedIds = @{}
    try {
        $wingetOutput = Invoke-Winget -WingetArgs @('list','--disable-interactivity')
        foreach ($a in $Apps) {
            if ($a.Source -eq 'winget' -or $a.Source -eq 'msstore') {
                if ($wingetOutput -match [regex]::Escape($a.Id)) {
                    $installedIds[$a.Id] = $a.Source
                }
            }
        }
    } catch { }
    try {
        if (Test-Command choco) {
            $chocoOutput = (choco list 2>&1 | Out-String)
            foreach ($a in $Apps) {
                if ($installedIds.ContainsKey($a.Id)) { continue }
                if ($a.Source -eq 'choco') {
                    if ($chocoOutput -match "(?m)^$([regex]::Escape($a.Id))[\s|]") {
                        $installedIds[$a.Id] = 'choco'
                    }
                } else {
                    $chocoId = $WingetToChocoMap[$a.Id]
                    if ($chocoId -and ($chocoOutput -match "(?m)^$([regex]::Escape($chocoId))[\s|]")) {
                        $installedIds[$a.Id] = 'choco'
                    }
                }
            }
        }
    } catch { }

    $installedColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $installedGreen = [System.Drawing.Color]::FromArgb(80, 200, 120)
    $sortedApps = $Apps | Sort-Object Category, Name
    foreach ($app in $sortedApps) {
        $checked = if ($importIds.Count -gt 0) { $importIds.ContainsKey($app.Id) } elseif ($SelectAll) { $true } else { [bool]$app.Default }
        $isInstalled = $installedIds.ContainsKey($app.Id)
        $statusText = if ($isInstalled) { "Installe" } else { "" }
        $realSource = if ($isInstalled) { $installedIds[$app.Id] } else { $app.Source }
        $rowIdx = $grid.Rows.Add($checked, $app.Category, $app.Name, $app.Description, $realSource, $statusText, $app.Id)
        if ($isInstalled) {
            $grid.Rows[$rowIdx].Cells['Statut'].Style.ForeColor = $installedGreen
            $grid.Rows[$rowIdx].Cells['Statut'].Style.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
            $grid.Rows[$rowIdx].Cells['Categorie'].Style.ForeColor = $installedColor
            $grid.Rows[$rowIdx].Cells['Application'].Style.ForeColor = $installedColor
            $grid.Rows[$rowIdx].Cells['Description'].Style.ForeColor = $installedColor
            $grid.Rows[$rowIdx].Cells['Source'].Style.ForeColor = $installedColor
        }
    }

    # -- Compteur dynamique --
    $grid.Add_CellValueChanged({
        $count = 0
        foreach ($r in $grid.Rows) { if ($r.Cells['Selection'].Value -eq $true) { $count++ } }
        $lblCount.Text = "$count / $($Apps.Count) selectionnees  |  $($installedIds.Count) deja installees"
    })
    $grid.Add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) { $grid.CommitEdit('CurrentCellChange') }
    })

    # -- Compteur initial --
    $initCount = 0
    foreach ($r in $grid.Rows) { if ($r.Cells['Selection'].Value -eq $true) { $initCount++ } }
    $installedTotal = $installedIds.Count
    $lblCount.Text = "$initCount / $($Apps.Count) selectionnees  |  $installedTotal deja installees"

    # -- Recherche temps reel --
    $txtSearch.Add_TextChanged({
        $filter = $txtSearch.Text.Trim()
        foreach ($row in $grid.Rows) {
            if ($filter -eq '') {
                $row.Visible = $true
            } else {
                $match = $false
                foreach ($col in @('Application','Categorie','Description','Source')) {
                    $val = $row.Cells[$col].Value
                    if ($val -and $val.ToString() -match [regex]::Escape($filter)) { $match = $true; break }
                }
                $row.Visible = $match
            }
        }
    })

    # -- Bottom layout (TableLayoutPanel pour positionnement fiable) --
    $btnStyle = {
        param($btn, $bgColor, $fgColor, $w)
        $btn.FlatStyle = 'Flat'
        $btn.FlatAppearance.BorderSize = 0
        $btn.BackColor = $bgColor
        $btn.ForeColor = $fgColor
        $btn.Font = $fontBtn
        $btn.Width = $w
        $btn.Height = 36
        $btn.Cursor = 'Hand'
        $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
            [math]::Min($bgColor.R + 25, 255),
            [math]::Min($bgColor.G + 25, 255),
            [math]::Min($bgColor.B + 25, 255))
    }

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = 'Tout cocher'
    & $btnStyle $btnAll $bgLight $fgWhite 120
    $btnAll.Add_Click({
        foreach ($row in $grid.Rows) { $row.Cells['Selection'].Value = $true }
    })

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Tout decocher'
    & $btnStyle $btnNone $bgLight $fgWhite 130
    $btnNone.Add_Click({
        foreach ($row in $grid.Rows) { $row.Cells['Selection'].Value = $false }
    })

    $accentOrange = [System.Drawing.Color]::FromArgb(230, 160, 50)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Exporter'
    & $btnStyle $btnExport $accentOrange ([System.Drawing.Color]::FromArgb(20, 20, 20)) 110
    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = "VSFI - Exporter la selection"
        $sfd.Filter = "JSON (*.json)|*.json"
        $sfd.FileName = "vsfi-selection-$(Get-Date -Format 'yyyyMMdd').json"
        $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if ($sfd.ShowDialog() -eq 'OK') {
            $exportList = @()
            foreach ($row in $grid.Rows) {
                if ($row.Cells['Selection'].Value -eq $true) {
                    $exportList += [PSCustomObject]@{
                        Name   = $row.Cells['Application'].Value
                        Id     = $row.Cells['Id'].Value
                        Source = $row.Cells['Source'].Value
                    }
                }
            }
            $exportData = [PSCustomObject]@{
                ExportDate   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                ComputerName = $env:COMPUTERNAME
                TotalApps    = $exportList.Count
                Applications = $exportList
            }
            $exportData | ConvertTo-Json -Depth 5 | Out-File -FilePath $sfd.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("$($exportList.Count) applications exportees vers :`n$($sfd.FileName)", "Export reussi", 'OK', 'Information')
        }
    })

    $btnImport = New-Object System.Windows.Forms.Button
    $btnImport.Text = 'Importer'
    & $btnStyle $btnImport $accentOrange ([System.Drawing.Color]::FromArgb(20, 20, 20)) 110
    $btnImport.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = "VSFI - Importer une selection"
        $ofd.Filter = "JSON (*.json)|*.json"
        $ofd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if ($ofd.ShowDialog() -eq 'OK') {
            try {
                $importData = Get-Content -Path $ofd.FileName -Raw | ConvertFrom-Json
                $importIds = @{}
                foreach ($app in $importData.Applications) { $importIds[$app.Id] = $true }
                $existingIds = @{}
                foreach ($row in $grid.Rows) {
                    $rowId = $row.Cells['Id'].Value
                    $existingIds[$rowId] = $true
                    $row.Cells['Selection'].Value = $importIds.ContainsKey($rowId)
                }
                $extraCount = 0
                foreach ($app in $importData.Applications) {
                    if (-not $existingIds.ContainsKey($app.Id)) {
                        $appName = if ($app.Name) { $app.Name } else { $app.Id }
                        $appSource = if ($app.Source) { $app.Source } else { 'winget' }
                        $grid.Rows.Add($true, '99-Import', $appName, 'Importe depuis backup', $appSource, '', $app.Id) | Out-Null
                        $extraCount++
                    }
                }
                $msg = "$($importIds.Count) applications importees"
                if ($extraCount -gt 0) { $msg += "`n$extraCount applications hors catalogue ajoutees" }
                [System.Windows.Forms.MessageBox]::Show("$msg`n`nDepuis : $($ofd.FileName)", "Import reussi", 'OK', 'Information')
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'import :`n$_", "Erreur", 'OK', 'Error')
            }
        }
    })

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Installer'
    & $btnStyle $btnOk $accentGreen ([System.Drawing.Color]::FromArgb(20, 20, 20)) 140
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Annuler'
    & $btnStyle $btnCancel $accentRed ([System.Drawing.Color]::White) 120
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $bottomLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $bottomLayout.Dock = 'Bottom'
    $bottomLayout.Height = 56
    $bottomLayout.BackColor = $bgMedium
    $bottomLayout.ColumnCount = 2
    $bottomLayout.RowCount = 1
    $null = $bottomLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $null = $bottomLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    $leftFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $leftFlow.Dock = 'Fill'
    $leftFlow.FlowDirection = 'LeftToRight'
    $leftFlow.WrapContents = $false
    $leftFlow.BackColor = $bgMedium
    $leftFlow.Padding = New-Object System.Windows.Forms.Padding(10, 8, 0, 0)
    $leftFlow.Controls.AddRange(@($btnAll, $btnNone, $btnExport, $btnImport))

    $rightFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $rightFlow.Dock = 'Fill'
    $rightFlow.FlowDirection = 'RightToLeft'
    $rightFlow.WrapContents = $false
    $rightFlow.BackColor = $bgMedium
    $rightFlow.Padding = New-Object System.Windows.Forms.Padding(0, 8, 10, 0)
    $rightFlow.Controls.AddRange(@($btnOk, $btnCancel))

    $bottomLayout.Controls.Add($leftFlow, 0, 0)
    $bottomLayout.Controls.Add($rightFlow, 1, 0)

    $form.Controls.Add($grid)
    $form.Controls.Add($headerPanel)
    $form.Controls.Add($bottomLayout)
    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel

    $dialogResult = $form.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $selected = @()
    foreach ($row in $grid.Rows) {
        if ($row.Cells['Selection'].Value -eq $true) {
            $idVal = $row.Cells['Id'].Value
            if ($idVal) { $selected += [PSCustomObject]@{ Id = $idVal } }
        }
    }

    return $selected
}

# ============================================================================
# CATALOGUE DES APPLICATIONS (60+ apps)
# ============================================================================

$AppCatalog = @(
    # ==================== UTILITAIRES ====================
    @{ Category="01-Utilitaires"; Name="NanaZip"; Id="M2Team.NanaZip"; Source="winget"; Default=$true; Description="Archiveur 7-Zip moderne" },
    @{ Category="01-Utilitaires"; Name="Notepad++"; Id="Notepad++.Notepad++"; Source="winget"; Default=$true; Description="Editeur texte avance" },
    @{ Category="01-Utilitaires"; Name="Everything"; Id="voidtools.Everything"; Source="winget"; Default=$true; Description="Recherche fichiers instantanee" },
    @{ Category="01-Utilitaires"; Name="PowerToys"; Id="Microsoft.PowerToys"; Source="winget"; Default=$true; Description="Outils Microsoft (FancyZones...)" },
    @{ Category="01-Utilitaires"; Name="TreeSize Free"; Id="JAMSoftware.TreeSize.Free"; Source="winget"; Default=$false; Description="Analyse espace disque" },
    @{ Category="01-Utilitaires"; Name="Rufus"; Id="Rufus.Rufus"; Source="winget"; Default=$true; Description="Creation cles USB bootables" },
    @{ Category="01-Utilitaires"; Name="HWiNFO"; Id="REALiX.HWiNFO"; Source="winget"; Default=$false; Description="Monitoring hardware detaille" },
    @{ Category="01-Utilitaires"; Name="CrystalDiskInfo"; Id="CrystalDewWorld.CrystalDiskInfo"; Source="winget"; Default=$false; Description="Sante des disques" },
    @{ Category="01-Utilitaires"; Name="UniGetUI"; Id="MartiCliment.UniGetUI"; Source="winget"; Default=$true; Description="Interface graphique gestionnaires paquets" },

    # ==================== NAVIGATEUR ====================
    @{ Category="02-Navigateur"; Name="Brave"; Id="Brave.Brave"; Source="winget"; Default=$true; Description="Navigateur prive base Chromium" },
    @{ Category="02-Navigateur"; Name="Firefox"; Id="Mozilla.Firefox"; Source="winget"; Default=$false; Description="Navigateur Mozilla" },

    # ==================== COMMUNICATION ====================
    @{ Category="03-Communication"; Name="Thunderbird"; Id="Mozilla.Thunderbird"; Source="winget"; Default=$true; Description="Client email" },
    @{ Category="03-Communication"; Name="Discord"; Id="discord.install"; Source="choco"; Default=$true; Description="Chat gaming/communaute" },
    @{ Category="03-Communication"; Name="Slack"; Id="SlackTechnologies.Slack"; Source="winget"; Default=$false; Description="Communication pro" },
    @{ Category="03-Communication"; Name="Zoom"; Id="Zoom.Zoom"; Source="winget"; Default=$false; Description="Visioconference" },
    @{ Category="03-Communication"; Name="Microsoft Teams"; Id="Microsoft.Teams"; Source="winget"; Default=$false; Description="Visio/chat Microsoft" },

    # ==================== DEV ====================
    @{ Category="04-Dev"; Name="Git"; Id="Git.Git"; Source="winget"; Default=$true; Description="Controle de version" },
    @{ Category="04-Dev"; Name="GitHub CLI"; Id="GitHub.cli"; Source="winget"; Default=$true; Description="GitHub en ligne de commande" },
    @{ Category="04-Dev"; Name="Docker Desktop"; Id="Docker.DockerDesktop"; Source="winget"; Default=$true; Description="Conteneurs Docker" },
    @{ Category="04-Dev"; Name="Windsurf (Codeium)"; Id="Codeium.Windsurf"; Source="winget"; Default=$true; Description="IDE avec IA integree" },
    @{ Category="04-Dev"; Name="VS Code"; Id="Microsoft.VisualStudioCode"; Source="winget"; Default=$false; Description="Editeur code Microsoft" },
    @{ Category="04-Dev"; Name="Windows Terminal"; Id="Microsoft.WindowsTerminal"; Source="winget"; Default=$true; Description="Terminal moderne Microsoft" },
    @{ Category="04-Dev"; Name="Node.js LTS"; Id="OpenJS.NodeJS.LTS"; Source="winget"; Default=$true; Description="Runtime JavaScript" },
    @{ Category="04-Dev"; Name="Python 3.9"; Id="Python.Python.3.9"; Source="winget"; Default=$false; Description="Python 3.9 (compatibilite)" },
    @{ Category="04-Dev"; Name="Python 3.10"; Id="Python.Python.3.10"; Source="winget"; Default=$false; Description="Python 3.10" },
    @{ Category="04-Dev"; Name="Python 3.12"; Id="Python.Python.3.12"; Source="winget"; Default=$true; Description="Python 3.12 (recent)" },
    @{ Category="04-Dev"; Name="Postman"; Id="Postman.Postman"; Source="winget"; Default=$false; Description="Test API REST" },
    @{ Category="04-Dev"; Name="HeidiSQL"; Id="HeidiSQL.HeidiSQL"; Source="winget"; Default=$false; Description="Client MySQL/MariaDB leger" },
    @{ Category="04-Dev"; Name="DBeaver"; Id="dbeaver.dbeaver"; Source="winget"; Default=$false; Description="Client BDD universel" },

    # ==================== TERMINAL ====================
    @{ Category="05-Terminal"; Name="Tabby"; Id="Eugeny.Tabby"; Source="winget"; Default=$true; Description="Terminal moderne multi-protocole" },
    @{ Category="05-Terminal"; Name="Termius"; Id="Termius.Termius"; Source="winget"; Default=$true; Description="Client SSH/SFTP moderne et cloud sync" },
    @{ Category="05-Terminal"; Name="PuTTY"; Id="PuTTY.PuTTY"; Source="winget"; Default=$false; Description="Client SSH/Telnet classique" },

    # ==================== RESEAU ====================
    @{ Category="06-Reseau"; Name="FileZilla"; Id="TimKosse.FileZilla.Client"; Source="winget"; Default=$true; Description="Client FTP/SFTP" },
    @{ Category="06-Reseau"; Name="RustDesk"; Id="RustDesk.RustDesk"; Source="winget"; Default=$true; Description="Bureau distant open-source" },
    @{ Category="06-Reseau"; Name="Tailscale"; Id="tailscale.tailscale"; Source="winget"; Default=$true; Description="VPN mesh simple" },
    @{ Category="06-Reseau"; Name="WinSCP"; Id="WinSCP.WinSCP"; Source="winget"; Default=$false; Description="Transfert SFTP/SCP avec GUI" },
    @{ Category="06-Reseau"; Name="Wireshark"; Id="WiresharkFoundation.Wireshark"; Source="winget"; Default=$false; Description="Analyse reseau" },
    @{ Category="06-Reseau"; Name="Advanced IP Scanner"; Id="Famatech.AdvancedIPScanner"; Source="winget"; Default=$false; Description="Scan reseau local" },

    # ==================== MULTIMEDIA ====================
    @{ Category="07-Multimedia"; Name="VLC"; Id="VideoLAN.VLC"; Source="winget"; Default=$true; Description="Lecteur multimedia universel" },
    @{ Category="07-Multimedia"; Name="OBS Studio"; Id="OBSProject.OBSStudio"; Source="winget"; Default=$true; Description="Streaming/enregistrement" },
    @{ Category="07-Multimedia"; Name="Spotify"; Id="Spotify.Spotify"; Source="winget"; Default=$true; Description="Musique en streaming" },
    @{ Category="07-Multimedia"; Name="Handbrake"; Id="HandBrake.HandBrake"; Source="winget"; Default=$true; Description="Encodage/conversion video" },
    @{ Category="07-Multimedia"; Name="FFmpeg"; Id="Gyan.FFmpeg"; Source="winget"; Default=$true; Description="Outils video en CLI" },
    @{ Category="07-Multimedia"; Name="Kdenlive"; Id="KDE.Kdenlive"; Source="winget"; Default=$false; Description="Montage video open-source" },
    @{ Category="07-Multimedia"; Name="Streamlabs"; Id="Streamlabs.Streamlabs"; Source="winget"; Default=$false; Description="OBS ameliore pour streaming" },

    # ==================== AUDIO ====================
    @{ Category="08-Audio"; Name="Audacity"; Id="Audacity.Audacity"; Source="winget"; Default=$true; Description="Edition audio" },
    @{ Category="08-Audio"; Name="VoiceMeeter Banana"; Id="VB-Audio.Voicemeeter.Banana"; Source="winget"; Default=$false; Description="Mixeur audio virtuel" },
    @{ Category="08-Audio"; Name="VB-Cable"; Id="VB-Audio.Cable"; Source="winget"; Default=$false; Description="Cable audio virtuel" },

    # ==================== IMAGES ====================
    @{ Category="09-Images"; Name="Caesium"; Id="SaeraSoft.CaesiumImageCompressor"; Source="winget"; Default=$true; Description="Compression images" },
    @{ Category="09-Images"; Name="GIMP"; Id="GIMP.GIMP"; Source="winget"; Default=$false; Description="Edition photo avancee" },
    @{ Category="09-Images"; Name="Inkscape"; Id="Inkscape.Inkscape"; Source="winget"; Default=$false; Description="Dessin vectoriel (logos, laser)" },
    @{ Category="09-Images"; Name="ShareX"; Id="ShareX.ShareX"; Source="winget"; Default=$true; Description="Screenshots + annotations" },
    @{ Category="09-Images"; Name="ImageGlass"; Id="DuongDieuPhap.ImageGlass"; Source="winget"; Default=$false; Description="Visionneuse images legere" },

    # ==================== 3D / CAD / MAKER ====================
    @{ Category="10-Maker-3D"; Name="PrusaSlicer"; Id="Prusa3D.PrusaSlicer"; Source="winget"; Default=$true; Description="Slicer impression 3D" },
    @{ Category="10-Maker-3D"; Name="Bambu Studio"; Id="Bambulab.Bambustudio"; Source="winget"; Default=$true; Description="Slicer Bambu Lab" },
    @{ Category="10-Maker-3D"; Name="UltiMaker Cura"; Id="Ultimaker.Cura"; Source="winget"; Default=$false; Description="Slicer alternatif" },
    @{ Category="10-Maker-3D"; Name="Fusion 360"; Id="autodesk-fusion360"; Source="choco"; Default=$true; Description="CAD/CAM Autodesk" },
    @{ Category="10-Maker-3D"; Name="Blender"; Id="BlenderFoundation.Blender"; Source="winget"; Default=$false; Description="3D, animation, rendu" },
    @{ Category="10-Maker-3D"; Name="FreeCAD"; Id="FreeCAD.FreeCAD"; Source="winget"; Default=$false; Description="CAD open-source" },
    @{ Category="10-Maker-3D"; Name="OpenSCAD"; Id="OpenSCAD.OpenSCAD"; Source="winget"; Default=$false; Description="CAD parametrique (code)" },
    @{ Category="10-Maker-3D"; Name="KiCad"; Id="KiCad.KiCad"; Source="winget"; Default=$false; Description="Design PCB/circuits" },
    @{ Category="10-Maker-3D"; Name="Arduino IDE"; Id="ArduinoSA.IDE.stable"; Source="winget"; Default=$true; Description="Programmation Arduino" },

    # ==================== DOMOTIQUE / IOT ====================
    @{ Category="11-Domotique"; Name="MQTT Explorer"; Id="ThomasNordworger.MQTTExplorer"; Source="winget"; Default=$true; Description="Client MQTT debug" },

    # ==================== PRODUCTIVITE ====================
    @{ Category="12-Productivite"; Name="LibreOffice"; Id="TheDocumentFoundation.LibreOffice"; Source="winget"; Default=$true; Description="Suite bureautique" },
    @{ Category="12-Productivite"; Name="iCloud"; Id="9PKTQ5699M62"; Source="msstore"; Default=$true; Description="Sync Apple" },
    @{ Category="12-Productivite"; Name="Notion"; Id="Notion.Notion"; Source="winget"; Default=$false; Description="Notes/wiki/gestion projet" },
    @{ Category="12-Productivite"; Name="Obsidian"; Id="Obsidian.Obsidian"; Source="winget"; Default=$false; Description="Notes Markdown linkees" },

    # ==================== IA ====================
    @{ Category="13-IA"; Name="LM Studio"; Id="ElementLabs.LMStudio"; Source="winget"; Default=$true; Description="LLM locaux" },
    @{ Category="13-IA"; Name="Claude"; Id="Anthropic.Claude"; Source="winget"; Default=$true; Description="Assistant IA Anthropic" },
    @{ Category="13-IA"; Name="ChatGPT"; Id="9NT1R1C2HH7J"; Source="msstore"; Default=$true; Description="Assistant IA OpenAI" },

    # ==================== GAMING ====================
    @{ Category="14-Gaming"; Name="Steam"; Id="Valve.Steam"; Source="winget"; Default=$true; Description="Plateforme jeux Valve" },
    @{ Category="14-Gaming"; Name="Epic Games Launcher"; Id="EpicGames.EpicGamesLauncher"; Source="winget"; Default=$false; Description="Jeux gratuits chaque semaine" },
    @{ Category="14-Gaming"; Name="GOG Galaxy"; Id="GOG.Galaxy"; Source="winget"; Default=$false; Description="Launcher unifie" },

    # ==================== PERIPHERIQUES ====================
    @{ Category="15-Peripheriques"; Name="Logitech Options+"; Id="Logitech.OptionsPlus"; Source="winget"; Default=$true; Description="Config souris/clavier Logi" }
)

# ============================================================================
# MAPPING WINGET -> CHOCOLATEY (pour fallback)
# ============================================================================

$WingetToChocoMap = @{
    "M2Team.NanaZip" = "nanazip"
    "Notepad++.Notepad++" = "notepadplusplus"
    "voidtools.Everything" = "everything"
    "Microsoft.PowerToys" = "powertoys"
    "JAMSoftware.TreeSize.Free" = "treesizefree"
    "Rufus.Rufus" = "rufus"
    "REALiX.HWiNFO" = "hwinfo"
    "CrystalDewWorld.CrystalDiskInfo" = "crystaldiskinfo"
    "Brave.Brave" = "brave"
    "Mozilla.Firefox" = "firefox"
    "Mozilla.Thunderbird" = "thunderbird"
    "SlackTechnologies.Slack" = "slack"
    "Zoom.Zoom" = "zoom"
    "Microsoft.Teams" = "microsoft-teams"
    "Git.Git" = "git"
    "GitHub.cli" = "gh"
    "Docker.DockerDesktop" = "docker-desktop"
    "Microsoft.VisualStudioCode" = "vscode"
    "Microsoft.WindowsTerminal" = "microsoft-windows-terminal"
    "OpenJS.NodeJS.LTS" = "nodejs-lts"
    "Python.Python.3.9" = "python39"
    "Python.Python.3.10" = "python310"
    "Python.Python.3.12" = "python312"
    "Postman.Postman" = "postman"
    "HeidiSQL.HeidiSQL" = "heidisql"
    "dbeaver.dbeaver" = "dbeaver"
    "Eugeny.Tabby" = "tabby"
    "Termius.Termius" = "termius"
    "PuTTY.PuTTY" = "putty"
    "TimKosse.FileZilla.Client" = "filezilla"
    "WinSCP.WinSCP" = "winscp"
    "WiresharkFoundation.Wireshark" = "wireshark"
    "VideoLAN.VLC" = "vlc"
    "OBSProject.OBSStudio" = "obs-studio"
    "Spotify.Spotify" = "spotify"
    "HandBrake.HandBrake" = "handbrake"
    "Gyan.FFmpeg" = "ffmpeg"
    "KDE.Kdenlive" = "kdenlive"
    "Audacity.Audacity" = "audacity"
    "VB-Audio.Voicemeeter.Banana" = "voicemeeter-banana"
    "GIMP.GIMP" = "gimp"
    "Inkscape.Inkscape" = "inkscape"
    "ShareX.ShareX" = "sharex"
    "Prusa3D.PrusaSlicer" = "prusaslicer"
    "Bambulab.Bambustudio" = "bambu-studio"
    "Ultimaker.Cura" = "cura-new"
    "BlenderFoundation.Blender" = "blender"
    "FreeCAD.FreeCAD" = "freecad"
    "OpenSCAD.OpenSCAD" = "openscad"
    "KiCad.KiCad" = "kicad"
    "ArduinoSA.IDE.stable" = "arduino"
    "TheDocumentFoundation.LibreOffice" = "libreoffice-fresh"
    "Notion.Notion" = "notion"
    "Obsidian.Obsidian" = "obsidian"
    "ElementLabs.LMStudio" = "lm-studio"
    "Valve.Steam" = "steam"
    "EpicGames.EpicGamesLauncher" = "epicgameslauncher"
    "GOG.Galaxy" = "goggalaxy"
    "Logitech.OptionsPlus" = "logitech-options"
    "MartiCliment.UniGetUI" = "wingetui"
    "RustDesk.RustDesk" = "rustdesk"
    "tailscale.tailscale" = "tailscale"
    "Famatech.AdvancedIPScanner" = "advanced-ip-scanner"
    "Streamlabs.Streamlabs" = "streamlabs-obs"
    "VB-Audio.Cable" = "vb-cable"
    "SaeraSoft.CaesiumImageCompressor" = "caesium.install"
    "DuongDieuPhap.ImageGlass" = "imageglass"
    "ThomasNordworger.MQTTExplorer" = "mqtt-explorer"
    "Anthropic.Claude" = "claude"
    "Codeium.Windsurf" = "windsurf"
}

# ============================================================================
# DEMARRAGE
# ============================================================================

Show-Banner

Write-Success "Privileges administrateur OK (STA: $([System.Threading.Thread]::CurrentThread.ApartmentState))"

# Ecran d'accueil GUI
if (-not $NoPrompt) {
    $welcomeChoice = Show-WelcomeDialog
    if (-not $welcomeChoice) {
        Write-Warning2 "Abandon."
        exit 0
    }
    if ($welcomeChoice -eq 'backup') {
        Export-InstalledApps
        Write-Host ""
        $null = Read-Host "Appuie sur Entree pour quitter"
        exit 0
    }
}

# Verif internet
try {
    if (Test-Command Test-NetConnection) {
        $net = Test-NetConnection -ComputerName "www.microsoft.com" -Port 443 -WarningAction SilentlyContinue
        if (-not $net.TcpTestSucceeded) { throw "Echec connexion" }
    } else {
        $null = Invoke-WebRequest -Uri "https://www.msftconnecttest.com/connecttest.txt" -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    }
    Write-Success "Connexion internet OK"
} catch {
    Write-Error2 "Pas de connexion internet !"
    $null = Read-Host "Appuie sur Entree pour quitter"
    exit 1
}

# ============================================================================
# SELECTION DES APPLICATIONS
# ============================================================================

Write-Host ""
if ($NoPrompt) {
    Write-Warning2 "Mode automatique : installation des paquets par defaut"
    $selectedApps = $AppCatalog | Where-Object { $_.Default -eq $true }
    Write-Log "$($selectedApps.Count) applications selectionnees" Cyan
} else {
    Write-Log "Ouverture de la fenetre de selection..." Cyan
    Write-Log "Coche les applications a installer puis clique OK" DarkGray
    
    $selected = Show-AppSelectionDialog -Apps $AppCatalog -SelectAll:$SelectAll
    
    if (-not $selected -or $selected.Count -eq 0) {
        Write-Warning2 "Aucune application selectionnee. Abandon."
        exit 0
    }
    
    $catalogById = @{}
    foreach ($a in $AppCatalog) { $catalogById[$a.Id] = $a }
    $selectedIds = $selected | ForEach-Object { $_.Id }
    $selectedApps = @($selectedIds | ForEach-Object { if ($catalogById.ContainsKey($_)) { $catalogById[$_] } })
    
    Write-Success "$($selectedApps.Count) applications selectionnees"
}

# Separer par source
$wingetApps = $selectedApps | Where-Object { $_.Source -eq "winget" }
$msstoreApps = $selectedApps | Where-Object { $_.Source -eq "msstore" }
$chocoApps = $selectedApps | Where-Object { $_.Source -eq "choco" }

Write-Log "Winget: $($wingetApps.Count) | MS Store: $($msstoreApps.Count) | Chocolatey: $($chocoApps.Count)" DarkGray

# ============================================================================
# ETAPE 1 - WINGET
# ============================================================================
Write-Step "1/4" "Gestionnaires de paquets"

$wingetPath = Get-WingetExePath
if ($wingetPath) {
    Write-Success "Winget $(Invoke-Winget -WingetArgs @('--version'))"
} else {
    Write-Warning2 "Installation de Winget..."
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        $url = "https://github.com/asheroto/winget-install/releases/latest/download/winget-install.ps1"
        $tempScript = "$env:TEMP\winget-install.ps1"
        Invoke-WebRequest -Uri $url -OutFile $tempScript -UseBasicParsing
        & powershell.exe -ExecutionPolicy Bypass -File $tempScript
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        if (Get-WingetExePath) {
            Write-Success "Winget installe"
        } else {
            throw "Echec"
        }
    } catch {
        Write-Error2 "Impossible d'installer Winget"
        exit 1
    }
}

# Accepter les licences
Write-Log "Acceptation des licences et mise a jour des sources ..." Cyan
$laStart = Get-Date
Initialize-WingetAgreements
$laDuration = (Get-Date) - $laStart
Write-Success "Licences acceptees ($($laDuration.ToString("mm\:ss")))"

# Chocolatey (toujours installer pour le fallback)
if (Test-Command choco) {
    Write-Success "Chocolatey $(choco --version)"
} else {
    Write-Warning2 "Installation de Chocolatey..."
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        if (Test-Command choco) {
            choco feature enable -n allowGlobalConfirmation 2>&1 | Out-Null
            Write-Success "Chocolatey installe"
        } else {
            throw "Echec"
        }
    } catch {
        Write-Error2 "Impossible d'installer Chocolatey"
    }
}

# ============================================================================
# ETAPE 2 - INSTALLATION DES APPLICATIONS
# ============================================================================
Write-Step "2/4" "Installation des applications"

$stats = @{ Success = 0; Skip = 0; Fail = 0; Fallback = 0 }
$failedApps = @()

Write-Log "Log: $script:LogFile" DarkGray
Write-FileLog "Demarrage installation - Winget:$($wingetApps.Count) MSStore:$($msstoreApps.Count) Choco:$($chocoApps.Count)"

# --- WINGET ---
if ($wingetApps.Count -gt 0) {
    Write-SectionHeader "WINGET" $wingetApps.Count

    $i = 0
    $total = $wingetApps.Count
    foreach ($app in $wingetApps) {
        $i++
        Write-FileLog "[WINGET] START $i/$total - $($app.Name) ($($app.Id))"
        $appStart = Get-Date
        if (Test-WingetInstalled -packageId $app.Id) {
            Write-AppResult $app.Name "SKIP"
            $stats.Skip++
            Write-FileLog "[WINGET] SKIP  $i/$total - $($app.Name) ($($app.Id))"
            continue
        }
        
        Write-Installing -Name $app.Name -Source "winget" -Index $i -Total $total
        $result = Invoke-Winget -WingetArgs @('install','--id',$app.Id,'--exact','--silent','--accept-package-agreements','--accept-source-agreements','--disable-interactivity')
        $duration = (Get-Date) - $appStart
        Write-Log "Termine en $($duration.ToString("mm\:ss"))" DarkGray
        Write-FileLog ("[WINGET] EXIT {0} - {1} ({2}) - duration {3}" -f $LASTEXITCODE, $app.Name, $app.Id, $duration.ToString("hh\:mm\:ss"))
        
        if ($LASTEXITCODE -eq 0) {
            Write-AppResult $app.Name "OK"
            $stats.Success++
        } elseif ($LASTEXITCODE -eq -1978335189) {
            Write-AppResult $app.Name "SKIP"
            $stats.Skip++
        } else {
            Write-AppResult $app.Name "FAIL"
            $stats.Fail++
            if ($WingetToChocoMap.ContainsKey($app.Id)) {
                $failedApps += @{
                    Name = $app.Name
                    WingetId = $app.Id
                    ChocoId = $WingetToChocoMap[$app.Id]
                }
            }
        }
    }
}

# --- MICROSOFT STORE ---
if ($msstoreApps.Count -gt 0) {
    Write-SectionHeader "MICROSOFT STORE" $msstoreApps.Count

    $i = 0
    $total = $msstoreApps.Count
    foreach ($app in $msstoreApps) {
        $i++
        Write-FileLog "[MSSTORE] START $i/$total - $($app.Name) ($($app.Id))"
        $appStart = Get-Date
        if (Test-WingetInstalled -packageId $app.Id) {
            Write-AppResult $app.Name "SKIP"
            $stats.Skip++
            Write-FileLog "[MSSTORE] SKIP  $i/$total - $($app.Name) ($($app.Id))"
            continue
        }
        Write-Installing -Name $app.Name -Source "msstore" -Index $i -Total $total
        $result = Invoke-Winget -WingetArgs @('install','--id',$app.Id,'--source','msstore','--silent','--accept-package-agreements','--accept-source-agreements','--disable-interactivity')
        $duration = (Get-Date) - $appStart
        Write-Log "Termine en $($duration.ToString("mm\:ss"))" DarkGray
        Write-FileLog ("[MSSTORE] EXIT {0} - {1} ({2}) - duration {3}" -f $LASTEXITCODE, $app.Name, $app.Id, $duration.ToString("hh\:mm\:ss"))
        
        if ($LASTEXITCODE -eq 0) {
            Write-AppResult $app.Name "OK"
            $stats.Success++
        } elseif ($LASTEXITCODE -eq -1978335189) {
            Write-AppResult $app.Name "SKIP"
            $stats.Skip++
        } else {
            Write-AppResult $app.Name "FAIL"
            $stats.Fail++
        }
    }
}

# --- CHOCOLATEY ---
if ($chocoApps.Count -gt 0 -and (Test-Command choco)) {
    Write-SectionHeader "CHOCOLATEY" $chocoApps.Count

    $i = 0
    $total = $chocoApps.Count
    foreach ($app in $chocoApps) {
        $i++
        Write-FileLog "[CHOCO] START $i/$total - $($app.Name) ($($app.Id))"
        $appStart = Get-Date
        if (Test-ChocoInstalled -packageName $app.Id) {
            Write-AppResult $app.Name "SKIP"
            $stats.Skip++
            Write-FileLog "[CHOCO] SKIP  $i/$total - $($app.Name) ($($app.Id))"
            continue
        }
        
        Write-Installing -Name $app.Name -Source "choco" -Index $i -Total $total
        $result = choco install $app.Id -y --no-progress 2>&1
        $duration = (Get-Date) - $appStart
        Write-Log "Termine en $($duration.ToString("mm\:ss"))" DarkGray
        Write-FileLog ("[CHOCO] EXIT {0} - {1} ({2}) - duration {3}" -f $LASTEXITCODE, $app.Name, $app.Id, $duration.ToString("hh\:mm\:ss"))
        
        if ($LASTEXITCODE -eq 0) {
            Write-AppResult $app.Name "OK"
            $stats.Success++
        } else {
            Write-AppResult $app.Name "FAIL"
            $stats.Fail++
        }
    }
}

# ============================================================================
# ETAPE 3 - FALLBACK CHOCOLATEY
# ============================================================================
if ($failedApps.Count -gt 0 -and (Test-Command choco)) {
    Write-Step "3/4" "Recuperation des echecs via Chocolatey"
    
    Write-SectionHeader "FALLBACK" $failedApps.Count

    $i = 0
    $total = $failedApps.Count
    foreach ($app in $failedApps) {
        $i++
        Write-FileLog "[FALLBACK] START $i/$total - $($app.Name) ($($app.ChocoId))"
        $appStart = Get-Date
        if (Test-ChocoInstalled -packageName $app.ChocoId) {
            Write-AppResult "$($app.Name)" "SKIP"
            $stats.Skip++
            if ($stats.Fail -gt 0) { $stats.Fail-- }
            Write-FileLog "[FALLBACK] SKIP  $i/$total - $($app.Name) ($($app.ChocoId))"
            continue
        }
        
        Write-Installing -Name $app.Name -Source "choco (fallback)" -Index $i -Total $total
        $result = choco install $app.ChocoId -y --no-progress 2>&1
        $duration = (Get-Date) - $appStart
        Write-Log "Termine en $($duration.ToString("mm\:ss"))" DarkGray
        Write-FileLog ("[FALLBACK] EXIT {0} - {1} ({2}) - duration {3}" -f $LASTEXITCODE, $app.Name, $app.ChocoId, $duration.ToString("hh\:mm\:ss"))
        
        if ($LASTEXITCODE -eq 0) {
            Write-AppResult "$($app.Name) via $($app.ChocoId)" "FALLBACK"
            $stats.Fallback++
            if ($stats.Fail -gt 0) { $stats.Fail-- }
        } else {
            Write-AppResult "$($app.Name)" "FAIL"
        }
    }
} else {
    Write-Step "3/4" "Fallback Chocolatey"
    Write-Success "Aucun echec a recuperer"
}

# ============================================================================
# ETAPE 4 - POST-INSTALLATION
# ============================================================================
Write-Step "4/4" "Configuration finale"

Write-FileLog "Fin installation - Success:$($stats.Success) Skip:$($stats.Skip) Fail:$($stats.Fail) Fallback:$($stats.Fallback)"

# Rafraichir PATH
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
Write-Success "PATH rafraichi"

# Mise a jour pip
if (Test-Command python) {
    python -m pip install --upgrade pip --quiet 2>&1 | Out-Null
    Write-Success "pip mis a jour"
}

# ============================================================================
# RAPPORT FINAL
# ============================================================================
Show-Summary -Stats $stats

Write-Host ""
Write-Host "    VERSIONS DETECTEES :" -ForegroundColor Cyan
if (Get-WingetExePath) { Write-Host "      winget : " -NoNewline -ForegroundColor DarkGray; Write-Host "$(Invoke-Winget -WingetArgs @('--version'))".Trim() -ForegroundColor White }
if (Test-Command choco)  { Write-Host "      choco  : " -NoNewline -ForegroundColor DarkGray; Write-Host "$(choco --version)".Trim() -ForegroundColor White }
if (Test-Command git)    { Write-Host "      git    : " -NoNewline -ForegroundColor DarkGray; Write-Host "$((git --version 2>&1) -replace 'git version ','')".Trim() -ForegroundColor White }
if (Test-Command node)   { Write-Host "      node   : " -NoNewline -ForegroundColor DarkGray; Write-Host "$(node --version 2>&1)".Trim() -ForegroundColor White }
if (Test-Command python) { Write-Host "      python : " -NoNewline -ForegroundColor DarkGray; Write-Host "$((python --version 2>&1) -replace 'Python ','')".Trim() -ForegroundColor White }
Write-Host ""

# Reboot
if (-not $SkipReboot) {
    Write-Warning2 "Redemarrage recommande pour Docker et PATH"
    Write-Host ""
    $response = Read-Host "    Redemarrer maintenant ? (O/N)"
    if ($response -match "^[OoYy]") {
        Write-Host ""
        Write-Host "    Redemarrage dans 10 secondes..." -ForegroundColor Yellow
        Start-Sleep -Seconds 10
        Restart-Computer -Force
    }
}

Write-Host ""
Write-Host "  $("=" * 60)" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "    Merci d'avoir utilise " -NoNewline -ForegroundColor Gray
Write-Host "VSFI" -NoNewline -ForegroundColor Cyan
Write-Host " ! Lance UniGetUI pour gerer tes apps." -ForegroundColor Gray
Write-Host ""
Write-Host "    by Artus Poulain - youtube.com/LesFreresPoulain" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  $("=" * 60)" -ForegroundColor DarkCyan
Write-Host ""