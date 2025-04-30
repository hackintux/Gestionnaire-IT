
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function New-ColoredTabPage($title, $color) {
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $title
    $tab.BackColor = $color
    return $tab
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Gestionnaire IT - Maintenance et Diagnostic ClicOnLine"
$form.Size = New-Object System.Drawing.Size(800, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(760, 250)
$tabControl.Location = New-Object System.Drawing.Point(10, 10)

$tabControl.DrawMode = 'OwnerDrawFixed'

$tabControl.Add_DrawItem({
    param($sender, $e)

    $brush = switch ($e.Index) {
        0 { [System.Drawing.Brushes]::Gold }
        1 { [System.Drawing.Brushes]::MediumPurple }
        2 { [System.Drawing.Brushes]::DarkSeaGreen }
        3 { [System.Drawing.Brushes]::LightCoral }
        default { [System.Drawing.Brushes]::LightGray }
    }

    $tabPage = $sender.TabPages[$e.Index]
    $e.Graphics.FillRectangle($brush, $e.Bounds)

    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'
    $sf.LineAlignment = 'Center'

    $e.Graphics.DrawString($tabPage.Text, $form.Font, [System.Drawing.Brushes]::Black, [System.Drawing.RectangleF]$e.Bounds, $sf)
})



$tabMaintenance = New-ColoredTabPage "Maintenance" ([System.Drawing.Color]::LightGoldenrodYellow)
$tabDiag = New-ColoredTabPage "Diagnostic" ([System.Drawing.Color]::Lavender)
$tabNettoyage = New-ColoredTabPage "Nettoyage / Boost" ([System.Drawing.Color]::Honeydew)
$tabRapport = New-ColoredTabPage "Rapports" ([System.Drawing.Color]::MistyRose)

$tabControl.TabPages.AddRange(@($tabMaintenance, $tabDiag, $tabNettoyage, $tabRapport))
$form.Controls.Add($tabControl)

$textBoxLogs = New-Object System.Windows.Forms.TextBox -Property @{
    Multiline = $true
    ScrollBars = 'Vertical'
    Size = New-Object System.Drawing.Size(760, 350)
    Location = New-Object System.Drawing.Point(10, 280)
    ReadOnly = $true
    BackColor = "Black"
    ForeColor = "Lime"
}
$form.Controls.Add($textBoxLogs)

function Write-Log {
    param([string]$message)
    $timestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $textBoxLogs.AppendText("[$timestamp] $message`r`n")
    $textBoxLogs.SelectionStart = $textBoxLogs.Text.Length
    $textBoxLogs.ScrollToCaret()
}

# Message de bienvenue
Write-Log "Bienvenue dans le Gestionnaire IT de ClicOnLine. Toutes les actions effectuées s'afficheront ici."

function New-TabButton($tab, $text, $x, $y, $action) {
    $btn = New-Object System.Windows.Forms.Button -Property @{
        Text = $text
        Size = New-Object System.Drawing.Size(330,40)
        Location = New-Object System.Drawing.Point($x,$y)
        BackColor = [System.Drawing.Color]::LightSteelBlue
    }
    $btn.Add_Click($action)
    $tab.Controls.Add($btn)
}

function Scan-WindowsUpdate {
    try {
        Write-Log "Scan Windows Update..."
        $serviceWU = Get-Service -Name wuauserv -ErrorAction Stop
        Write-Log "Service Windows Update : $($serviceWU.Status)"
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        if ($historyCount -gt 0) {
            $lastEntry = $searcher.QueryHistory($historyCount-1, 1)
            Write-Log "Dernière mise à jour : $($lastEntry.Date) - Statut : $($lastEntry.ResultCode)"
        } else {
            Write-Log "Aucune mise à jour trouvée."
        }
    } catch {
        Write-Log "Erreur scan Windows Update : $_"
    }
}

function Repair-WindowsUpdate {
    Write-Log "Réinitialisation des composants Windows Update..."
    Stop-Service wuauserv -Force
    Remove-Item -Path "C:\Windows\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Service wuauserv
    Write-Log "Réinitialisation terminée."
}

function Force-WindowsUpdateDetection {
    Write-Log "Détection des mises à jour forcée..."
    UsoClient StartScan
}

function Restart-PCCountdown {
    shutdown /r /t 60
    Write-Log "Redémarrage prévu dans 60 secondes..."
}

function Create-SystemRestorePoint {
    try {
        Write-Log "Création point de restauration système..."
        Checkpoint-Computer -Description "Point_Restauration_Outil_IT" -RestorePointType "MODIFY_SETTINGS"
        Write-Log "Point de restauration créé."
    } catch {
        Write-Log "Erreur création point de restauration : $_"
    }
}

function Scan-InstalledApps {
    Write-Log "Scan des logiciels installés..."
    $softwares = @("Google Chrome", "Adobe Acrobat Reader")
    foreach ($software in $softwares) {
        $found = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
                 Where-Object { $_.DisplayName -like "*$software*" }
        if (!$found) {
            $found = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
                     Where-Object { $_.DisplayName -like "*$software*" }
        }
        if ($found) {
            Write-Log "$software trouvé : version $($found.DisplayVersion)"
        } else {
            Write-Log "$software non trouvé."
        }
    }
}

function Check-ObsoleteDrivers {
    Write-Log "Scan des pilotes en cours..."
    $drivers = Get-WmiObject Win32_PnPSignedDriver -ErrorAction SilentlyContinue
    $limitDate = (Get-Date).AddYears(-2)
    foreach ($driver in $drivers) {
        if ($driver.DriverDate) {
            $driverDate = [datetime]::ParseExact($driver.DriverDate.Substring(0,8), 'yyyyMMdd', $null)
            $status = if ($driverDate -lt $limitDate) { "Obsolète" } else { "OK" }
            Write-Log "$($driver.DeviceName) - $driverDate - $($driver.DriverVersion) - $status"
        }
    }
}

function Show-NetworkConnections {
    Write-Log "Connexions réseau actives :"
    $netstat = netstat -anob 2>$null
    foreach ($line in $netstat) {
        if ($line -match "^ *(TCP|UDP)") {
            Write-Log $line.Trim()
        }
    }
}

function Show-SystemHealthDashboard {
    Write-Log "État santé du système :"
    $cpu = Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
    $ram = Get-CimInstance Win32_OperatingSystem
    $total = [math]::Round($ram.TotalVisibleMemorySize / 1MB, 2)
    $free = [math]::Round($ram.FreePhysicalMemory / 1MB, 2)
    $used = $total - $free
    Write-Log "CPU : $cpu %, RAM : $used / $total Go"
}

function Check-Antivirus {
    $antivirus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if ($antivirus) {
        foreach ($av in $antivirus) {
            Write-Log "Antivirus : $($av.displayName)"
        }
    } else {
        Write-Log "Aucun antivirus détecté."
    }
}

function Quick-SystemClean {
    $paths = @("$env:TEMP", "$env:WINDIR\Temp", "$env:WINDIR\Prefetch", "$env:WINDIR\SoftwareDistribution\Download")
    foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Write-Log "Nettoyé : $path"
        }
    }
}

function Boost-PCPerformance {
    Write-Log "Activation du boost PC..."
    powercfg -setactive SCHEME_MIN
    Write-Log "Plan Haute performance activé."
}

function Uninstall-TargetedApps {
    $appsToRemove = @("Microsoft Teams", "OneDrive", "Java")
    foreach ($app in $appsToRemove) {
        $products = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*$app*" }
        foreach ($prod in $products) {
            $prod.Uninstall() | Out-Null
            Write-Log "$($prod.Name) désinstallé."
        }
    }
}

function Uninstall-Bitdefender {
    Write-Log "Suppression de Bitdefender..."
    Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Bitdefender*" } | ForEach-Object {
        $_.Uninstall() | Out-Null
        Write-Log "$($_.Name) désinstallé."
    }
}

function Update-Apps {
    Write-Log "MAJ Chrome et Adobe via winget..."
    winget upgrade --id Google.Chrome --silent --accept-source-agreements --accept-package-agreements
    winget upgrade --id Adobe.Acrobat.Reader.64-bit --silent --accept-source-agreements --accept-package-agreements
}

function Export-LogHtml {
    $path = "$env:USERPROFILE\Desktop\rapport_maintenance.html"
    $html = $textBoxLogs.Lines -join "<br>"
    Set-Content -Path $path -Value "<html><body><pre>$html</pre></body></html>"
    Write-Log "Export HTML terminé : $path"
}

function Check-CriticalServices {
    $services = @("wuauserv", "bits", "WinDefend")
    foreach ($svc in $services) {
        $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($s.Status -ne 'Running') {
            Write-Log "$svc NON démarré"
        } else {
            Write-Log "$svc OK"
        }
    }
}

function Install-WingetIfMissing {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "Winget non installé. Téléchargez-le depuis le Microsoft Store."
    } else {
        Write-Log "Winget est déjà installé."
    }
}

# Boutons
New-TabButton $tabMaintenance "Scanner Windows Update" 20 20 { Scan-WindowsUpdate }
New-TabButton $tabMaintenance "Réparer Windows Update" 400 20 { Repair-WindowsUpdate }
New-TabButton $tabMaintenance "Forcer Windows Update" 20 80 { Force-WindowsUpdateDetection }
New-TabButton $tabMaintenance "Créer un point de restauration" 400 80 { Create-SystemRestorePoint }
New-TabButton $tabMaintenance "Redémarrer le PC (60s)" 20 140 { Restart-PCCountdown }

New-TabButton $tabDiag "Scanner Logiciels Installés" 20 20 { Scan-InstalledApps }
New-TabButton $tabDiag "Vérifier pilotes obsolètes" 400 20 { Check-ObsoleteDrivers }
New-TabButton $tabDiag "Voir connexions réseau" 20 80 { Show-NetworkConnections }
New-TabButton $tabDiag "Tableau de bord santé PC" 400 80 { Show-SystemHealthDashboard }
New-TabButton $tabDiag "Lister Antivirus installés" 20 140 { Check-Antivirus }

New-TabButton $tabNettoyage "Nettoyage rapide du système" 20 20 { Quick-SystemClean }
New-TabButton $tabNettoyage "Booster le PC" 400 20 { Boost-PCPerformance }
New-TabButton $tabNettoyage "Désinstaller applis ciblées" 20 80 { Uninstall-TargetedApps }
New-TabButton $tabNettoyage "Désinstaller Bitdefender" 400 80 { Uninstall-Bitdefender }
New-TabButton $tabNettoyage "Mettre à jour Chrome / Adobe" 20 140 { Update-Apps }

New-TabButton $tabRapport "Exporter le rapport (HTML)" 20 20 { Export-LogHtml }
New-TabButton $tabRapport "Vérifier services critiques" 400 20 { Check-CriticalServices }
New-TabButton $tabRapport "Installer Winget (si absent)" 20 80 { Install-WingetIfMissing }

$form.ShowDialog()
