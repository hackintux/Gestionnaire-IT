
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

$textBoxLogs = New-Object System.Windows.Forms.RichTextBox -Property @{
    Multiline = $true
    ScrollBars = 'Vertical'
    Size = New-Object System.Drawing.Size(760, 300)  # au lieu de 350
    Location = New-Object System.Drawing.Point(10, 280)
    ReadOnly = $true
    BackColor = "Black"
    ForeColor = "white"
}
$form.Controls.Add($textBoxLogs)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size(760, 15)
$progressBar.Location = New-Object System.Drawing.Point(10, 590)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

function Animate-ProgressBar {
    param (
        [System.Windows.Forms.ProgressBar]$progressBar,
        [int]$durationSeconds = 30
    )

    if (-not $progressBar) {
        Write-LogError "ProgressBar non définie."
        return
    }

    $progressBar.Visible = $true
    $progressBar.Style = 'Blocks'
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    $stepCount = $durationSeconds * 10
    for ($i = 1; $i -le $stepCount; $i++) {
        $progress = [math]::Round(($i / $stepCount) * 100)
        $progressBar.Value = [math]::Min($progress, 100)
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.Application]::DoEvents()
    }

    $progressBar.Value = 100
    Start-Sleep -Milliseconds 300
    $progressBar.Visible = $false
}


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

function Write-LogError {
    param([string]$message)
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $textBoxLogs.SelectionColor = [System.Drawing.Color]::Red
    $textBoxLogs.AppendText("[$timestamp] $message`r`n")
    $textBoxLogs.SelectionColor = $textBoxLogs.ForeColor
    $textBoxLogs.SelectionStart = $textBoxLogs.Text.Length
    $textBoxLogs.ScrollToCaret()
}

function Write-LogAvert {
    param([string]$message)
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $textBoxLogs.SelectionColor = [System.Drawing.Color]::yellow
    $textBoxLogs.AppendText("[$timestamp] $message`r`n")
    $textBoxLogs.SelectionColor = $textBoxLogs.ForeColor
    $textBoxLogs.SelectionStart = $textBoxLogs.Text.Length
    $textBoxLogs.ScrollToCaret()
}

function Write-LogOk {
    param([string]$message)
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $textBoxLogs.SelectionColor = [System.Drawing.Color]::Lime
    $textBoxLogs.AppendText("[$timestamp] $message`r`n")
    $textBoxLogs.SelectionColor = $textBoxLogs.ForeColor
    $textBoxLogs.SelectionStart = $textBoxLogs.Text.Length
    $textBoxLogs.ScrollToCaret()
}


function Scan-WindowsUpdate {
    try {
        Write-Log "Scan Windows Update..."
        $serviceWU = Get-Service -Name wuauserv -ErrorAction Stop
        Write-Log "Service Windows Update : $($serviceWU.Status)"
	Animate-ProgressBar -progressBar $progressBar -durationSeconds 7
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $results = $searcher.Search("IsInstalled=0 and Type='Software'").Updates

        if ($results.Count -eq 0) {
            Write-LogOk "Aucune mise à jour disponible."
            return
        }

        $updateList = ""
        for ($i = 0; $i -lt $results.Count; $i++) {
            $title = $results.Item($i).Title
            $updateList += "$($i+1). $title`n"
        }

        Write-LogAvert "Mises à jour disponibles :"
        Write-Log $updateList.Trim()

        $dialogResult = [System.Windows.Forms.MessageBox]::Show(
            "Les mises à jour suivantes sont disponibles :`n`n$updateList`nVoulez-vous les installer ?",
            "Mises à jour détectées",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $downloader = $session.CreateUpdateDownloader()
            $downloader.Updates = $results
            Write-Log "Téléchargement des mises à jour..."
            $downloader.Download()

            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $results
            Write-Log "Installation des mises à jour..."
            $result = $installer.Install()

            Write-LogOk "Résultat de l'installation : $($result.ResultCode)"
        } else {
            Write-LogAvert "Installation des mises à jour annulée par l'utilisateur."
        }
    } catch {
        Write-LogError "Erreur durant le scan Windows Update : $_"
    }
}

function Repair-WindowsUpdate {
    Write-Log "Réinitialisation des composants Windows Update..."
    Stop-Service wuauserv -Force
    Remove-Item -Path "C:\Windows\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Service wuauserv
    Write-LogOk "Réinitialisation terminée."
}

function Force-WindowsUpdateDetection {
    Write-Log "Détection des mises à jour forcée..."

    try {
        $process = Start-Process -FilePath "UsoClient.exe" -ArgumentList "StartScan" -NoNewWindow -PassThru -ErrorAction Stop
        $process.WaitForExit()

        if ($process.ExitCode -eq 0) {
            Write-LogOk "Détection des mises à jour lancée avec succès."
        } else {
            Write-LogAvert "Commande UsoClient terminée avec un code inattendu : $($process.ExitCode)"
        }
    } catch {
        Write-LogError "Erreur lors du lancement de la détection : $_"
    }
}


function Restart-PCCountdown {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous redémarrer l’ordinateur dans 60 secondes ?`nVous pouvez encore annuler avec shutdown /a.",
        "Confirmation de redémarrage",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        shutdown /r /t 60
        Write-Log "Redémarrage programmé dans 60 secondes."
        [System.Windows.Forms.MessageBox]::Show(
            "Le redémarrage est prévu dans 60 secondes.`nUtilisez shutdown /a pour l'annuler.",
            "Redémarrage en attente",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } else {
        Write-Log "Redémarrage annulé par l’utilisateur."
    }
}


function Create-SystemRestorePoint {
    try {
        Write-Log "Création point de restauration système..."
        Checkpoint-Computer -Description "Point_Restauration_Outil_IT" -RestorePointType "MODIFY_SETTINGS"
        Write-LogOk "Point de restauration créé."
    } catch {
        Write-LogError "Erreur création point de restauration : $_"
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
            Write-LogOk "$software trouvé : version $($found.DisplayVersion)"
        } else {
            Write-LogAvert "$software non trouvé."
        }
    }
}

function Check-ObsoleteDrivers {
    Write-Log "Scan des pilotes obsolètes..."

    $drivers = Get-WmiObject Win32_PnPSignedDriver -ErrorAction SilentlyContinue
    $limitDate = (Get-Date).AddYears(-2)
    $obsolete = @()

    foreach ($driver in $drivers) {
        if ($driver.DriverDate -and $driver.DeviceName -ne "") {
            $driverDate = try {
                [datetime]::ParseExact($driver.DriverDate.Substring(0,8), 'yyyyMMdd', $null)
            } catch {
                continue
            }

            if ($driverDate -lt $limitDate -and $driver.DeviceName -notmatch "PCI|USB|Audio|Graphics|LAN|Wireless|Bluetooth") {
                $obsolete += $driver
                Write-LogAvert "$($driver.DeviceName) - $driverDate - $($driver.DriverVersion) - POTENTIELLEMENT OBSOLÈTE"
            }
        }
    }

    if ($obsolete.Count -eq 0) {
        Write-LogOk "Aucun pilote obsolète détecté ou désactivable en toute sécurité."
        return
    }

    Write-LogAvert "Nombre total de pilotes potentiellement obsolètes : $($obsolete.Count)"

    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Il y a $($obsolete.Count) pilotes potentiellement obsolètes.`nSouhaitez-vous les désinstaller ?",
        "Suppression des pilotes obsolètes",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        foreach ($d in $obsolete) {
            try {
                $infName = $d.InfName
                Write-Log "Suppression du pilote : $($d.DeviceName) ($infName)..."
                pnputil /delete-driver "$infName" /uninstall /force /quiet
                Write-Log "Pilote supprimé : $infName"
            } catch {
                Write-LogError "Erreur suppression pilote $($d.DeviceName) : $_"
            }
        }
    } else {
        Write-Log "Suppression des pilotes annulée par l'utilisateur."
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
        Write-LogAvert "Aucun antivirus détecté."
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
    $appsToRemove = @("Microsof Teams", "OneDriv", "ava")
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
