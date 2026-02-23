# ==========================================================
# APP MANAGER - Windows Edition v1.4.8.0
# Copyright 2026 FarTekTV
#
# Licensed under the Apache License, Version 2.0
# You may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed
# on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS
# OF ANY KIND, either express or implied.
# ==========================================================

foreach ($alias in @('Copier','Deplacer','Supprimer','Renommer','Effacer')) {
    try { Remove-Alias $alias -Force -ErrorAction SilentlyContinue } catch {}
    if (-not (Get-Alias $alias -ErrorAction SilentlyContinue)) {
        try { Set-Alias $alias Write-Host -Scope Global -ErrorAction SilentlyContinue } catch {}
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$BG       = [System.Drawing.Color]::FromArgb(15, 15, 25)
$BG2      = [System.Drawing.Color]::FromArgb(22, 22, 38)
$BG3      = [System.Drawing.Color]::FromArgb(30, 30, 50)
$ACCENT   = [System.Drawing.Color]::FromArgb(80, 200, 255)
$ACCENT2  = [System.Drawing.Color]::FromArgb(160, 100, 255)
$SUCCESS  = [System.Drawing.Color]::FromArgb(80, 230, 150)
$DANGER   = [System.Drawing.Color]::FromArgb(255, 80, 100)
$WARN     = [System.Drawing.Color]::FromArgb(255, 200, 60)
$GREEN2   = [System.Drawing.Color]::FromArgb(40, 120, 90)
$DARK     = [System.Drawing.Color]::FromArgb(40, 40, 70)
$FG       = [System.Drawing.Color]::FromArgb(220, 220, 240)
$FGDIM    = [System.Drawing.Color]::FromArgb(120, 120, 150)
$ROW_SEL  = [System.Drawing.Color]::FromArgb(40, 80, 140)
$ROW_ALT  = [System.Drawing.Color]::FromArgb(18, 18, 30)
$WHITE    = [System.Drawing.Color]::White

$CAT_COLORS = @{
    "Games"          = [System.Drawing.Color]::FromArgb(255, 140, 60)
    "Development" = [System.Drawing.Color]::FromArgb(80, 200, 255)
    "Multimedia"    = [System.Drawing.Color]::FromArgb(200, 80, 255)
    "Office"   = [System.Drawing.Color]::FromArgb(80, 200, 120)
    "System"       = [System.Drawing.Color]::FromArgb(200, 180, 60)
    "Network"        = [System.Drawing.Color]::FromArgb(60, 180, 220)
    "Security"      = [System.Drawing.Color]::FromArgb(255, 80, 100)
    "Microsoft"     = [System.Drawing.Color]::FromArgb(50, 140, 240)
    "Other"         = [System.Drawing.Color]::FromArgb(130, 130, 160)
}

$FONT_MAIN  = New-Object System.Drawing.Font('Segoe UI', 10)
$FONT_BOLD  = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$FONT_TITLE = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$FONT_MONO  = New-Object System.Drawing.Font('Consolas', 9)
$FONT_SMALL = New-Object System.Drawing.Font('Segoe UI', 8)
$FONT_CAT   = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)

function Clean-Name {
    param([string]$s)
    $out = ''
    foreach ($ch in $s.ToCharArray()) {
        $code = [int]$ch
        if ($code -ge 32 -and $code -le 126) { $out += $ch }
        elseif ($code -ge 160 -and $code -le 255) {
            $out += $ch
        }
    }
    return $out.Trim() -replace '\s+', ' '
}

# --- Categorie ---
function Get-AppCategory {
    param([string]$Name, [string]$Publisher)
    $n = $Name.ToLower()
    $p = $Publisher.ToLower()
    if ($n -match 'epic|steam|game|gamer|gaming|ubisoft|ea app|origin|battle|blizzard|riot|valorant|fortnite|minecraft|gog|itch' -or
        $p -match 'epic games|valve|ubisoft|electronic arts|activision|blizzard|riot games|2k games|bethesda') { return 'Games' }
    if ($n -match 'visual studio|vscode|vs code|git |python|node|sdk|docker|postman|jetbrains|intellij|android studio|github|powershell|wsl|mingw|cmake|gcc|llvm|rust |ruby|php|mysql|postgres|mongodb|redis|dbeaver|notepad|sublime|vim|anaconda|jupyter' -or
        $p -match 'jetbrains|github') { return 'Development' }
    if ($n -match 'vlc|spotify|audacity|obs |blender|gimp|photoshop|premiere|after effect|lightroom|winamp|foobar|media player|codec|k-lite|handbrake|paint|inkscape|krita|davinci|vegas|kdenlive|discord|zoom|skype|twitch|plex|kodi|itunes|groove' -or
        $p -match 'adobe|blackmagic|videolan') { return 'Multimedia' }
    if ($n -match 'office|word|excel|powerpoint|outlook|onenote|libreoffice|openoffice|acrobat|pdf|foxit|writer|calc|impress|notion|obsidian|evernote|trello|slack|thunderbird' -or
        $p -match 'the document foundation') { return 'Office' }
    if ($n -match 'antivirus|malware|kaspersky|avast|avira|bitdefender|norton|mcafee|defender|firewall|vpn|keepass|bitwarden|lastpass|1password|authy') { return 'Security' }
    if ($n -match 'chrome|firefox|edge|opera|brave|vivaldi|browser|putty|teamviewer|anydesk|remote desktop|ssh|wireshark|nmap') { return 'Network' }
    if ($n -match 'driver|directx|redistribut|runtime|framework|vcredist|openal|physx|cuda|nvidia|amd |intel |realtek|logitech|razer|corsair|hwinfo|hwmonitor|cpu-z|gpu-z|speccy|ccleaner|7-zip|winrar|rufus|virtualbox|vmware|chocolatey|winget|autohotkey' -or
        $p -match 'nvidia|advanced micro|realtek|logitech|razer') { return 'System' }
    if ($n -match 'microsoft|windows |msix|xbox|cortana|onedrive' -or
        $p -match 'microsoft') { return 'Microsoft' }
    return 'Other'
}

# --- Chargement des applications ---
function Get-RegApps {
    param([string]$hive, [string]$subkey, [Microsoft.Win32.RegistryView]$view)
    $apps = @()
    try {
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
        $key  = $root.OpenSubKey($subkey)
        if (-not $key) { Write-Log "SKIP: $hive\$subkey ($view) - key not found"; return $apps }
        foreach ($subName in $key.GetSubKeyNames()) {
            try {
                $sub = $key.OpenSubKey($subName)
                if (-not $sub) { continue }
                $displayName = $sub.GetValue('DisplayName')
                if (-not $displayName -or $displayName.Trim() -eq '') { continue }
                $apps += [PSCustomObject]@{
                    DisplayName     = $displayName.Trim()
                    DisplayVersion  = $sub.GetValue('DisplayVersion')
                    Publisher       = $sub.GetValue('Publisher')
                    EstimatedSize   = $sub.GetValue('EstimatedSize')
                    UninstallString = $sub.GetValue('UninstallString')
                    InstallDate     = $sub.GetValue('InstallDate')
                    InstallLocation = $sub.GetValue('InstallLocation')
                    DisplayIcon     = $sub.GetValue('DisplayIcon')
                }
                $sub.Close()
            } catch {}
        }
        $key.Close()
        $root.Close()
    } catch {}
    return $apps
}

function Get-InstalledApps {
    $apps = @()
    $seen = @{}

    $sources = @(
        @{ Hive='LocalMachine'; Key='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';          View='Registry64' },
        @{ Hive='LocalMachine'; Key='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';          View='Registry32' },
        @{ Hive='LocalMachine'; Key='SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; View='Registry64' },
        @{ Hive='CurrentUser';  Key='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';          View='Registry64' }
    )

    foreach ($src in $sources) {
        $hiveEnum = [Microsoft.Win32.RegistryHive]::$($src.Hive)
        $viewEnum = [Microsoft.Win32.RegistryView]::$($src.View)
        $entries  = Get-RegApps $hiveEnum $src.Key $viewEnum
        Write-Log "SOURCE $($src.Hive)\$($src.Key) [$($src.View)] -> $($entries.Count) entrees"
        foreach ($entry in $entries) {
            try {
                $rawName = $entry.DisplayName
                if ($seen.ContainsKey($rawName)) { continue }
                $seen[$rawName] = $true

                $installPath = ''
                if ($entry.InstallLocation -and (Test-Path $entry.InstallLocation)) {
                    $installPath = $entry.InstallLocation.TrimEnd('\')
                }
                if ($installPath -eq '' -and $entry.DisplayIcon) {
                    try {
                        $iconPath = ($entry.DisplayIcon -replace ',\d+$','' -replace '"','').Trim()
                        if ($iconPath -and $iconPath -match '\.(exe|dll)$' -and (Test-Path $iconPath)) {
                            $p = Split-Path $iconPath -Parent
                            if ($p) { $installPath = $p }
                        }
                    } catch {}
                }
                if ($installPath -eq '' -and $entry.UninstallString) {
                    try {
                        $ustr = ($entry.UninstallString -replace '"','').Trim()
                        if ($ustr -match '\.exe') {
                            $ep = (($ustr -split '\.exe')[0] + '.exe').Trim()
                            if ($ep -and (Test-Path $ep)) {
                                $p = Split-Path $ep -Parent
                                if ($p) { $installPath = $p }
                            }
                        }
                    } catch {}
                }

                $pub  = if ($entry.Publisher)     { $entry.Publisher.Trim() } else { 'N/A' }
                $size = if ($entry.EstimatedSize) { "$([math]::Round([int]$entry.EstimatedSize/1024,1)) MB" } else { 'N/A' }
                $apps += [PSCustomObject]@{
                    Name         = $rawName
                    Version      = if ($entry.DisplayVersion) { $entry.DisplayVersion } else { 'N/A' }
                    Publisher    = $pub
                    Size         = $size
                    UninstallCmd = if ($entry.UninstallString) { $entry.UninstallString } else { '' }
                    InstallDate  = if ($entry.InstallDate)     { $entry.InstallDate }     else { 'N/A' }
                    InstallPath  = $installPath
                    Source       = 'WIN'
                    Category     = Get-AppCategory $rawName $pub
                }
            } catch {}
        }
    }

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        try {
            $tmpFile = [System.IO.Path]::GetTempFileName()
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName               = 'winget'
            $psi.Arguments              = 'list --accept-source-agreements --disable-interactivity'
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $psi.UseShellExecute        = $false
            $psi.CreateNoWindow         = $true
            $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $proc = [System.Diagnostics.Process]::Start($psi)
            $stdout = $proc.StandardOutput.ReadToEnd()
            $proc.WaitForExit()
            $wgLines = $stdout -split '
?
' | Select-Object -Skip 3
            Write-Log "WINGET: $($wgLines.Count) lignes recues"
            foreach ($line in $wgLines) {
                if ($line -match '\S') {
                    $line = $line -replace '[--]',''
                    $parts = $line -split '\s{2,}'
                    if ($parts.Count -ge 1 -and $parts[0].Trim() -ne '') {
                        $cleanName = Clean-Name $parts[0].Trim()
                        if ($cleanName -eq '' -or $cleanName.Length -lt 2) { continue }
                        if ($seen.ContainsKey($cleanName)) { continue }
                        $seen[$cleanName] = $true
                        $apps += [PSCustomObject]@{
                            Name         = $cleanName
                            Version      = if ($parts.Count -gt 1) { $parts[1].Trim() } else { 'N/A' }
                            Publisher    = 'N/A'
                            Size         = 'N/A'
                            UninstallCmd = ''
                            InstallDate  = 'N/A'
                            InstallPath  = ''
                            Source       = 'WINGET'
                            Category     = Get-AppCategory $cleanName 'N/A'
                        }
                    }
                }
            }
        } catch { Write-Log "WINGET ERROR: $_" }
    }

    Write-Log "TOTAL apres dedup: $($apps.Count) apps"
    return $apps | Sort-Object Name
}

function Run-WithConsole {
    param([string]$Title, [scriptblock]$Action)
    $console = New-Object System.Windows.Forms.Form
    $console.Text = $Title
    $console.Size = New-Object System.Drawing.Size(720, 450)
    $console.BackColor = $BG
    $console.Font = $FONT_MONO
    $console.StartPosition = 'CenterScreen'
    $console.FormBorderStyle = 'FixedDialog'
    $console.MaximizeBox = $false
    $output = New-Object System.Windows.Forms.RichTextBox
    $output.Dock = 'Fill'
    $output.BackColor = $BG
    $output.ForeColor = $SUCCESS
    $output.Font = $FONT_MONO
    $output.ReadOnly = $true
    $output.BorderStyle = 'None'
    $output.Padding = New-Object System.Windows.Forms.Padding(10)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Dock = 'Bottom'
    $btnClose.Height = 36
    $btnClose.BackColor = $BG3
    $btnClose.ForeColor = $FG
    $btnClose.FlatStyle = 'Flat'
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Add_Click({ $console.Close() })
    $console.Controls.Add($output)
    $console.Controls.Add($btnClose)
    $console.Show()
    $output.AppendText("=== $Title ===`n`n")
    $console.Refresh()
    try {
        $result = & $Action 2>&1
        foreach ($line in $result) { $output.AppendText("  $line`n"); $console.Refresh() }
    } catch {
        $output.SelectionColor = $DANGER
        $output.AppendText("`n  ERREUR: $_`n")
    }
    $output.AppendText("`n  [OK] Termine.`n")
    $console.Refresh()
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'APP MANAGER'
$form.Size = New-Object System.Drawing.Size(1280, 760)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 550)
$form.BackColor = $BG
$form.ForeColor = $FG
$form.Font = $FONT_MAIN
$form.StartPosition = 'CenterScreen'

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.SplitterWidth = 2
$split.Panel1MinSize = 140
$split.BackColor = $BG3
$form.Controls.Add($split)
$split.SplitterDistance = 175

$split.Panel1.BackColor = $BG2
$lblCatTitle = New-Object System.Windows.Forms.Label
$lblCatTitle.Text = 'CATEGORIES'
$lblCatTitle.Font = $FONT_CAT
$lblCatTitle.ForeColor = $FGDIM
$lblCatTitle.Location = New-Object System.Drawing.Point(12, 14)
$lblCatTitle.AutoSize = $true
$split.Panel1.Controls.Add($lblCatTitle)

$script:catButtons = @{}
$script:currentCat = 'All'

function Make-CatButton {
    param([string]$Label, [int]$Y, [System.Drawing.Color]$Color)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = '  ' + $Label
    $btn.Width = 165
    $btn.Height = 30
    $btn.Location = New-Object System.Drawing.Point(5, $Y)
    $btn.BackColor = $BG3
    $btn.ForeColor = $Color
    $btn.Font = $FONT_CAT
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.BorderColor = $BG3
    $btn.TextAlign = 'MiddleLeft'
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Tag = ($Label -split '  ')[0].Trim()
    $btn.Add_Click({
        $script:currentCat = $this.Tag
        Update-CatButtons
        Apply-Filter
    })
    $split.Panel1.Controls.Add($btn)
    return $btn
}

function Update-CatButtons {
    foreach ($key in $script:catButtons.Keys) {
        $b = $script:catButtons[$key]
        if ($b.Tag -eq $script:currentCat) {
            $b.BackColor = [System.Drawing.Color]::FromArgb(35, 60, 100)
            $b.FlatAppearance.BorderColor = $ACCENT
            $b.FlatAppearance.BorderSize = 1
        } else {
            $b.BackColor = $BG3
            $b.FlatAppearance.BorderColor = $BG3
            $b.FlatAppearance.BorderSize = 0
        }
    }
}

$split.Panel2.BackColor = $BG

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 68
$header.BackColor = $BG2

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'APP MANAGER'
$lblTitle.Font = $FONT_TITLE
$lblTitle.ForeColor = $ACCENT
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(18, 8)
$header.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = 'Manage all your installed applications'
$lblSub.Font = $FONT_SMALL
$lblSub.ForeColor = $FGDIM
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(20, 40)
$header.Controls.Add($lblSub)

$lblStats = New-Object System.Windows.Forms.Label
$lblStats.Text = 'Loading...'
$lblStats.Font = $FONT_SMALL
$lblStats.ForeColor = $FGDIM
$lblStats.AutoSize = $true
$lblStats.Location = New-Object System.Drawing.Point(600, 26)
$header.Controls.Add($lblStats)

$statusBar = New-Object System.Windows.Forms.Panel
$statusBar.Dock = 'Bottom'
$statusBar.Height = 28
$statusBar.BackColor = $BG2
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = '  Ready'
$lblStatus.ForeColor = $FGDIM
$lblStatus.Font = $FONT_SMALL
$lblStatus.Dock = 'Fill'
$lblStatus.TextAlign = 'MiddleLeft'
$lblStatus.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$statusBar.Controls.Add($lblStatus)

$toolbar = New-Object System.Windows.Forms.Panel
$toolbar.Dock = 'Top'
$toolbar.Height = 52
$toolbar.BackColor = $BG3

$btnLaunch = New-Object System.Windows.Forms.Button
$btnLaunch.Text = '>  Launch'
$btnLaunch.Width = 115
$btnLaunch.Height = 34
$btnLaunch.Location = New-Object System.Drawing.Point(8, 9)
$btnLaunch.BackColor = $ACCENT2
$btnLaunch.ForeColor = $WHITE
$btnLaunch.Font = $FONT_BOLD
$btnLaunch.FlatStyle = 'Flat'
$btnLaunch.FlatAppearance.BorderSize = 0
$btnLaunch.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolbar.Controls.Add($btnLaunch)

$btnFolder = New-Object System.Windows.Forms.Button
$btnFolder.Text = '[] Folder'
$btnFolder.Width = 115
$btnFolder.Height = 34
$btnFolder.Location = New-Object System.Drawing.Point(130, 9)
$btnFolder.BackColor = $GREEN2
$btnFolder.ForeColor = $WHITE
$btnFolder.Font = $FONT_BOLD
$btnFolder.FlatStyle = 'Flat'
$btnFolder.FlatAppearance.BorderSize = 0
$btnFolder.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolbar.Controls.Add($btnFolder)

$btnUpdate = New-Object System.Windows.Forms.Button
$btnUpdate.Text = '^ Update'
$btnUpdate.Width = 140
$btnUpdate.Height = 34
$btnUpdate.Location = New-Object System.Drawing.Point(253, 9)
$btnUpdate.BackColor = $WARN
$btnUpdate.ForeColor = $WHITE
$btnUpdate.Font = $FONT_BOLD
$btnUpdate.FlatStyle = 'Flat'
$btnUpdate.FlatAppearance.BorderSize = 0
$btnUpdate.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolbar.Controls.Add($btnUpdate)

$btnUninstall = New-Object System.Windows.Forms.Button
$btnUninstall.Text = 'X Uninstall'
$btnUninstall.Width = 140
$btnUninstall.Height = 34
$btnUninstall.Location = New-Object System.Drawing.Point(400, 9)
$btnUninstall.BackColor = $DANGER
$btnUninstall.ForeColor = $WHITE
$btnUninstall.Font = $FONT_BOLD
$btnUninstall.FlatStyle = 'Flat'
$btnUninstall.FlatAppearance.BorderSize = 0
$btnUninstall.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolbar.Controls.Add($btnUninstall)

$btnInfo = New-Object System.Windows.Forms.Button
$btnInfo.Text = 'i  Info'
$btnInfo.Width = 105
$btnInfo.Height = 34
$btnInfo.Location = New-Object System.Drawing.Point(547, 9)
$btnInfo.BackColor = $BG2
$btnInfo.ForeColor = $WHITE
$btnInfo.Font = $FONT_BOLD
$btnInfo.FlatStyle = 'Flat'
$btnInfo.FlatAppearance.BorderSize = 0
$btnInfo.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolbar.Controls.Add($btnInfo)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(665, 14)
$searchBox.Width = 210
$searchBox.Height = 28
$searchBox.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 42)
$searchBox.ForeColor = $FGDIM
$searchBox.Font = $FONT_MAIN
$searchBox.BorderStyle = 'FixedSingle'
$searchBox.Text = 'Search...'
$toolbar.Controls.Add($searchBox)

$btnClearSearch = New-Object System.Windows.Forms.Button
$btnClearSearch.Text = 'X'
$btnClearSearch.Width = 26
$btnClearSearch.Height = 28
$btnClearSearch.Location = New-Object System.Drawing.Point(877, 14)
$btnClearSearch.BackColor = [System.Drawing.Color]::FromArgb(60, 30, 35)
$btnClearSearch.ForeColor = $DANGER
$btnClearSearch.Font = $FONT_BOLD
$btnClearSearch.FlatStyle = 'Flat'
$btnClearSearch.FlatAppearance.BorderSize = 0
$btnClearSearch.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClearSearch.Visible = $false
$toolbar.Controls.Add($btnClearSearch)

$searchTimer = New-Object System.Windows.Forms.Timer
$searchTimer.Interval = 250

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'R  Refresh'
$btnRefresh.Width = 100
$btnRefresh.Height = 34
$btnRefresh.Location = New-Object System.Drawing.Point(892, 9)
$btnRefresh.BackColor = $DARK
$btnRefresh.ForeColor = $WHITE
$btnRefresh.Font = $FONT_BOLD
$btnRefresh.FlatStyle = 'Flat'
$btnRefresh.FlatAppearance.BorderSize = 0
$btnRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolbar.Controls.Add($btnRefresh)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.BackgroundColor = $BG
$grid.GridColor = $BG3
$grid.BorderStyle = 'None'
$grid.RowHeadersVisible = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.MultiSelect = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.ReadOnly = $true
$grid.Font = $FONT_MAIN
$grid.AutoSizeRowsMode = 'None'
$grid.RowTemplate.Height = 32
$grid.CellBorderStyle = 'SingleHorizontal'
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeight = 38
$grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
$grid.ColumnHeadersDefaultCellStyle.BackColor = $BG3
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $ACCENT
$grid.ColumnHeadersDefaultCellStyle.Font = $FONT_BOLD
$grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $BG3
$grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $ACCENT
$grid.DefaultCellStyle.BackColor = $BG
$grid.DefaultCellStyle.ForeColor = $FG
$grid.DefaultCellStyle.SelectionBackColor = $ROW_SEL
$grid.DefaultCellStyle.SelectionForeColor = $WHITE
$grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$grid.AlternatingRowsDefaultCellStyle.BackColor = $ROW_ALT
$grid.AlternatingRowsDefaultCellStyle.ForeColor = $FG
$grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = $ROW_SEL

$split.Panel2.Controls.Add($statusBar)
$split.Panel2.Controls.Add($header)
$split.Panel2.Controls.Add($toolbar)
$split.Panel2.Controls.Add($grid)

$colDefs = @(
    @{Name='Category';    Header='Category';   Fill=$false; Width=105},
    @{Name='Name';        Header='Application'; Fill=$true;  Width=0},
    @{Name='Version';     Header='Version';     Fill=$false; Width=110},
    @{Name='Publisher';   Header='Publisher';     Fill=$false; Width=185},
    @{Name='Size';        Header='Size';      Fill=$false; Width=80},
    @{Name='Source';      Header='Source';      Fill=$false; Width=70},
    @{Name='InstallDate'; Header='Install Date'; Fill=$false; Width=92},
    @{Name='InstallPath'; Header='Folder';     Fill=$false; Width=255}
)
foreach ($col in $colDefs) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $col.Name
    $c.HeaderText = $col.Header
    $c.DataPropertyName = $col.Name
    if ($col.Fill) { $c.AutoSizeMode = 'Fill' } else { $c.Width = $col.Width }
    $grid.Columns.Add($c) | Out-Null
}

$script:allApps      = @()
$script:filteredApps = @()

function Build-CategoryButtons {
    $toRemove = @($split.Panel1.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] })
    foreach ($b in $toRemove) { $split.Panel1.Controls.Remove($b) }
    $script:catButtons = @{}
    $catCounts = @{}
    $catCounts['All'] = $script:allApps.Count
    foreach ($app in $script:allApps) {
        if (-not $catCounts.ContainsKey($app.Category)) { $catCounts[$app.Category] = 0 }
        $catCounts[$app.Category]++
    }
    $y = 36
    $orderedCats = @('All') + ($catCounts.Keys | Where-Object { $_ -ne 'All' } | Sort-Object)
    foreach ($cat in $orderedCats) {
        if (-not $catCounts.ContainsKey($cat)) { continue }
        $color = if ($CAT_COLORS.ContainsKey($cat)) { $CAT_COLORS[$cat] } else { $FG }
        $label = "$cat  ($($catCounts[$cat]))"
        $btn = Make-CatButton $label $y $color
        $script:catButtons[$cat] = $btn
        $y += 33
    }
    Update-CatButtons
}


$script:basePath = try {
    if ($MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    } elseif ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) {
        Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    } else { $PWD.Path }
} catch { $PWD.Path }
$script:dbPath  = Join-Path $script:basePath 'apps_db.json'
$script:logPath = Join-Path $script:basePath 'scan_log.txt'

function Write-Log {
    param([string]$msg)
    try {
        $ts = Get-Date -Format 'HH:mm:ss'
        Add-Content -Path $script:logPath -Value "[$ts] $msg" -Encoding UTF8
    } catch {}
}

function Save-DB {
    $data = $script:allApps | ForEach-Object {
        @{
            Name         = $_.Name
            Version      = $_.Version
            Publisher    = $_.Publisher
            Size         = $_.Size
            UninstallCmd = $_.UninstallCmd
            InstallDate  = $_.InstallDate
            InstallPath  = $_.InstallPath
            Source       = $_.Source
            Category     = $_.Category
        }
    }
    $data | ConvertTo-Json -Depth 3 | Set-Content -Path $script:dbPath -Encoding UTF8
    $lblStatus.Text = "  DB saved : $($script:allApps.Count) apps -> $script:dbPath"
}

function Load-DB {
    if (-not (Test-Path $script:dbPath)) { return $null }
    try {
        $raw = Get-Content -Path $script:dbPath -Encoding UTF8 -Raw
        $parsed = $raw | ConvertFrom-Json
        $apps = @()
        foreach ($item in $parsed) {
            $apps += [PSCustomObject]@{
                Name         = $item.Name
                Version      = $item.Version
                Publisher    = $item.Publisher
                Size         = $item.Size
                UninstallCmd = $item.UninstallCmd
                InstallDate  = $item.InstallDate
                InstallPath  = $item.InstallPath
                Source       = $item.Source
                Category     = $item.Category
            }
        }
        return $apps | Sort-Object Name
    } catch { return $null }
}

function Sync-DB {
    $lblStatus.Text = '  Scan en cours...'
    $form.Refresh()
    $live = Get-InstalledApps

    $dbMap  = @{}
    foreach ($a in $script:allApps) { $dbMap[$a.Name] = $a }

    $liveMap = @{}
    foreach ($a in $live) { $liveMap[$a.Name] = $a }

    $added   = 0
    $removed = 0
    $updated = 0

    foreach ($name in $liveMap.Keys) {
        if (-not $dbMap.ContainsKey($name)) {
            $script:allApps += $liveMap[$name]
            $added++
        } else {
            $existing = $dbMap[$name]
            $fresh    = $liveMap[$name]
            if ($existing.Version -ne $fresh.Version -or $existing.InstallPath -ne $fresh.InstallPath) {
                $existing.Version      = $fresh.Version
                $existing.InstallPath  = $fresh.InstallPath
                $existing.Size         = $fresh.Size
                $updated++
            }
        }
    }

    $script:allApps = $script:allApps | Where-Object { $liveMap.ContainsKey($_.Name) }
    $removed = ($dbMap.Count + $added) - $script:allApps.Count

    $script:allApps = $script:allApps | Sort-Object Name
    Save-DB
    Build-CategoryButtons
    Apply-Filter

    $msg = "  Sync: +$added added"
    if ($updated -gt 0) { $msg += "  ~$updated updated" }
    if ($removed -gt 0) { $msg += "  -$removed removed" }
    $lblStatus.Text = $msg
}

function Load-Data {
    Write-Log "=== Load-Data started ==="
    Write-Log "OS: $([System.Environment]::Is64BitOperatingSystem) | Process64: $([System.Environment]::Is64BitProcess)"
    $cached = Load-DB
    if ($cached -and $cached.Count -gt 0) {
        $script:allApps = $cached
        $lblStats.Text = "$($script:allApps.Count) apps (cached)"
        Build-CategoryButtons
        Apply-Filter
        $lblStatus.Text = '  Cache charge, scan en cours...'
        $form.Refresh()
        Sync-DB
    } else {
        $lblStatus.Text = '  Premier scan...'
        $form.Refresh()
        $script:allApps = Get-InstalledApps
        Save-DB
        $lblStats.Text = "$($script:allApps.Count) apps found"
        Build-CategoryButtons
        Apply-Filter
    }
    $lblStats.Text = "$($script:allApps.Count) apps found"
}

function Apply-Filter {
    $search = $searchBox.Text.Trim()
    $sa = ($search -ne 'Search...' -and $search -ne '')
    $btnClearSearch.Visible = $sa
    $script:filteredApps = $script:allApps | Where-Object {
        $mc = ($script:currentCat -eq 'All' -or $_.Category -eq $script:currentCat)
        $ms = (-not $sa -or
               $_.Name        -like "*$search*" -or
               $_.Publisher   -like "*$search*" -or
               $_.Category    -like "*$search*" -or
               $_.Version     -like "*$search*" -or
               $_.InstallPath -like "*$search*")
        $mc -and $ms
    }
    $grid.Rows.Clear()
    foreach ($app in $script:filteredApps) {
        $grid.Rows.Add($app.Category, $app.Name, $app.Version, $app.Publisher, $app.Size, $app.Source, $app.InstallDate, $app.InstallPath) | Out-Null
    }
    foreach ($row in $grid.Rows) {
        $cat = $row.Cells['Category'].Value
        if ($CAT_COLORS.ContainsKey($cat)) {
            $row.Cells['Category'].Style.ForeColor = $CAT_COLORS[$cat]
            $row.Cells['Category'].Style.Font = $FONT_CAT
        }
        switch ($row.Cells['Source'].Value) {
            'WINGET' { $row.Cells['Source'].Style.ForeColor = $SUCCESS }
            'WIN'    { $row.Cells['Source'].Style.ForeColor = $ACCENT }
        }
        $row.Cells['InstallPath'].Style.ForeColor = $FGDIM
        $row.Cells['InstallPath'].Style.Font = $FONT_MONO
    }
    $lblStatus.Text = "  $($script:filteredApps.Count) of $($script:allApps.Count) applications"
}

function Get-SelectedApp {
    if ($grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Please select an application first.', 'No Selection', 'OK', 'Information') | Out-Null
        return $null
    }
    return $script:filteredApps[$grid.SelectedRows[0].Index]
}

$btnLaunch.Add_Click({
    $app = Get-SelectedApp
    if (-not $app) { return }
    $launched = $false
    if ($app.InstallPath -ne '' -and (Test-Path $app.InstallPath)) {
        $exes = Get-ChildItem -Path $app.InstallPath -Filter '*.exe' -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notmatch '(?i)(unins|setup|update|crash|repair|redist|vcredist)' } |
                Sort-Object Length -Descending
        if ($exes.Count -gt 0) {
            try { Start-Process $exes[0].FullName; $launched = $true; $lblStatus.Text = "  Launched: $($exes[0].Name)" } catch {}
        }
    }
    if (-not $launched) {
        $sn = ($app.Name -split ' ')[0]
        try { Start-Process $sn -ErrorAction Stop; $launched = $true; $lblStatus.Text = "  Launched: $sn" } catch {}
    }
    if (-not $launched) {
        if ($app.InstallPath -ne '' -and (Test-Path $app.InstallPath)) {
            Start-Process 'explorer.exe' $app.InstallPath
            $lblStatus.Text = '  Folder opened (no exe found)'
        } else {
            [System.Windows.Forms.MessageBox]::Show("Impossible de lancer '$($app.Name)'.`nUtilisez le menu Demarrer.", 'Launch Failed', 'OK', 'Warning') | Out-Null
        }
    }
})

$btnFolder.Add_Click({
    $app = Get-SelectedApp
    if (-not $app) { return }
    if ($app.InstallPath -ne '' -and (Test-Path $app.InstallPath)) {
        Start-Process 'explorer.exe' $app.InstallPath
        $lblStatus.Text = "  Folder: $($app.InstallPath)"
    } else {
        $res = [System.Windows.Forms.MessageBox]::Show('Chemin introuvable.`n`nOuvrir Program Files ?', 'Path Not Found', 'YesNo', 'Question')
        if ($res -eq 'Yes') { Start-Process 'explorer.exe' "$env:ProgramFiles" }
    }
})

$btnUpdate.Add_Click({
    $app = Get-SelectedApp
    if (-not $app) { return }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show('winget is not installed. Install App Installer from the Microsoft Store.', 'winget Not Found', 'OK', 'Warning') | Out-Null
        return
    }
    $res = [System.Windows.Forms.MessageBox]::Show("Update: $($app.Name) ?", 'Update', 'YesNo', 'Question')
    if ($res -eq 'Yes') {
        $appName = $app.Name
        Run-WithConsole "Update - $appName" {
            winget upgrade --name "$appName" --accept-source-agreements --accept-package-agreements 2>&1
        }
        $lblStatus.Text = "  Done: $appName"
    }
})

$btnUninstall.Add_Click({
    $app = Get-SelectedApp
    if (-not $app) { return }
    $res = [System.Windows.Forms.MessageBox]::Show("Uninstall permanently:`n`n$($app.Name)`n`nAction irreversible.", 'Confirm', 'YesNo', 'Warning')
    if ($res -eq 'Yes') {
        $appName = $app.Name
        $appCmd  = $app.UninstallCmd
        if ($appCmd -ne '') {
            Run-WithConsole "Uninstall - $appName" {
                if ($appCmd -match '(?i)^msiexec') {
                    $msiArgs = $appCmd -replace '(?i)msiexec\.exe\s*',''
                    Start-Process 'msiexec.exe' -ArgumentList $msiArgs -Wait 2>&1
                } else {
                    Start-Process 'cmd.exe' -ArgumentList "/c `"$appCmd`"" -Wait 2>&1
                }
                'Command executed.'
            }
        } elseif (Get-Command winget -ErrorAction SilentlyContinue) {
            Run-WithConsole "Uninstall - $appName" {
                winget uninstall --name "$appName" --accept-source-agreements 2>&1
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show('No uninstall command found. Use Settings > Apps.', 'Error', 'OK', 'Warning') | Out-Null
        }
        Load-Data
    }
})

$btnInfo.Add_Click({
    $app = Get-SelectedApp
    if (-not $app) { return }
    $info = New-Object System.Windows.Forms.Form
    $info.Text = 'Infos'
    $info.Size = New-Object System.Drawing.Size(580, 440)
    $info.BackColor = $BG
    $info.ForeColor = $FG
    $info.Font = $FONT_MAIN
    $info.StartPosition = 'CenterScreen'
    $info.FormBorderStyle = 'FixedDialog'
    $info.MaximizeBox = $false
    $txt = New-Object System.Windows.Forms.RichTextBox
    $txt.Dock = 'Fill'
    $txt.BackColor = $BG2
    $txt.ForeColor = $FG
    $txt.Font = $FONT_MONO
    $txt.ReadOnly = $true
    $txt.BorderStyle = 'None'
    $txt.Padding = New-Object System.Windows.Forms.Padding(20)
    $catColor = if ($CAT_COLORS.ContainsKey($app.Category)) { $CAT_COLORS[$app.Category] } else { $FG }
    $txt.SelectionColor = $catColor
    $txt.AppendText("[ $($app.Category) ]`n")
    $txt.SelectionColor = $ACCENT
    $txt.AppendText("APPLICATION`n")
    $txt.SelectionColor = $FG
    $txt.AppendText("  $($app.Name)`n`n")
    $pairs = @(
        @('VERSION',     $app.Version),
        @('PUBLISHER',     $app.Publisher),
        @('SIZE',      $app.Size),
        @('SOURCE',      $app.Source),
        @('INSTALL DATE', $app.InstallDate),
        @('FOLDER',     $app.InstallPath),
        @('UNINSTALL CMD', $app.UninstallCmd)
    )
    foreach ($pair in $pairs) {
        $txt.SelectionColor = $FGDIM
        $txt.AppendText("  $($pair[0].PadRight(14))")
        $txt.SelectionColor = $FG
        $txt.AppendText("$($pair[1])`n")
    }
    $pnl = New-Object System.Windows.Forms.Panel
    $pnl.Dock = 'Bottom'
    $pnl.Height = 44
    $pnl.BackColor = $BG3
    $btnOF = New-Object System.Windows.Forms.Button
    $btnOF.Text = 'Open Folder'
    $btnOF.Width = 150
    $btnOF.Height = 30
    $btnOF.Location = New-Object System.Drawing.Point(10, 7)
    $btnOF.BackColor = $GREEN2
    $btnOF.ForeColor = $WHITE
    $btnOF.FlatStyle = 'Flat'
    $btnOF.FlatAppearance.BorderSize = 0
    $btnOF.Add_Click({
        if ($app.InstallPath -ne '' -and (Test-Path $app.InstallPath)) {
            Start-Process 'explorer.exe' $app.InstallPath
        } else {
            [System.Windows.Forms.MessageBox]::Show('Path not found.', 'Erreur', 'OK', 'Warning') | Out-Null
        }
    })
    $btnCI = New-Object System.Windows.Forms.Button
    $btnCI.Text = 'Close'
    $btnCI.Width = 100
    $btnCI.Height = 30
    $btnCI.Location = New-Object System.Drawing.Point(170, 7)
    $btnCI.BackColor = $BG2
    $btnCI.ForeColor = $FG
    $btnCI.FlatStyle = 'Flat'
    $btnCI.FlatAppearance.BorderSize = 0
    $btnCI.Add_Click({ $info.Close() })
    $pnl.Controls.Add($btnOF)
    $pnl.Controls.Add($btnCI)
    $info.Controls.Add($txt)
    $info.Controls.Add($pnl)
    $info.ShowDialog() | Out-Null
})

$btnRefresh.Add_Click({
    $lblStatus.Text = '  Scan complet en cours...'
    $form.Refresh()
    $script:allApps = Get-InstalledApps
    Save-DB
    $lblStats.Text = "$($script:allApps.Count) apps found"
    Build-CategoryButtons
    Apply-Filter
})

$searchBox.Add_GotFocus({
    if ($searchBox.Text -eq 'Search...') {
        $searchBox.Text = ''
        $searchBox.ForeColor = $FG
    }
})
$searchBox.Add_LostFocus({
    if ($searchBox.Text -eq '') {
        $searchBox.Text = 'Search...'
        $searchBox.ForeColor = $FGDIM
        $btnClearSearch.Visible = $false
    }
})
$searchBox.Add_TextChanged({
    if ($searchBox.Text -ne 'Search...') {
        $searchBox.ForeColor = $FG
        $searchTimer.Stop()
        $searchTimer.Start()
    }
})
$searchBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $searchBox.Text = ''
        $searchBox.ForeColor = $FGDIM
        $btnClearSearch.Visible = $false
        Apply-Filter
        $grid.Focus()
    }
})
$searchTimer.Add_Tick({
    $searchTimer.Stop()
    Apply-Filter
})

$btnClearSearch.Add_Click({
    $searchBox.Text = ''
    $searchBox.ForeColor = $FGDIM
    $btnClearSearch.Visible = $false
    Apply-Filter
    $searchBox.Focus()
})
$grid.Add_CellDoubleClick({ $btnInfo.PerformClick() })

$form.Add_Shown({ Load-Data })
[void]$form.ShowDialog()