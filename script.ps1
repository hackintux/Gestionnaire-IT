

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
        0 { [System.Drawing.Brushes]::SteelBlue }
        1 { [System.Drawing.Brushes]::MediumSeaGreen }
        2 { [System.Drawing.Brushes]::DarkOrange }
        3 { [System.Drawing.Brushes]::Gold }
        4 { [System.Drawing.Brushes]::Gainsboro}
        5 { [System.Drawing.Brushes]::MediumOrchid }
        6 { [System.Drawing.Brushes]::Turquoise }
        7 { [System.Drawing.Brushes]::CornflowerBlue }
        8 { [System.Drawing.Brushes]::IndianRed }
        default { [System.Drawing.Brushes]::DarkSeaGreen }
    }

    $tabPage = $sender.TabPages[$e.Index]
    $e.Graphics.FillRectangle($brush, $e.Bounds)

    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'
    $sf.LineAlignment = 'Center'

    $e.Graphics.DrawString($tabPage.Text, $form.Font, [System.Drawing.Brushes]::Black, [System.Drawing.RectangleF]$e.Bounds, $sf)
})



$tabMaj = New-ColoredTabPage "Mise a jour" ([System.Drawing.Color]::SteelBlue)
$tabDiag = New-ColoredTabPage "Diagnostic" ([System.Drawing.Color]::MediumSeaGreen)
$tabNettoyage = New-ColoredTabPage "Nettoyage" ([System.Drawing.Color]::DarkOrange)
$tabBoost = New-ColoredTabPage "Boost" ([System.Drawing.Color]::Gold)
$tabRapport = New-ColoredTabPage "Rapports" ([System.Drawing.Color]::Gainsboro)
$tabo365 = New-ColoredTabPage "Office 365" ([System.Drawing.Color]::MediumOrchid)
$tabWildix = New-ColoredTabPage "Wildix" ([System.Drawing.Color]::Turquoise)
$tabReseaux = New-ColoredTabPage "Reseaux" ([System.Drawing.Color]::CornflowerBlue)
$tabRouteur = New-ColoredTabPage "Routeur" ([System.Drawing.Color]::IndianRed)

$tabControl.TabPages.AddRange(@($tabMaj, $tabDiag, $tabNettoyage, $tabBoost, $tabRapport, $tabo365, $tabWildix, $tabReseaux, $tabRouteur))
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
        Write-LogError "ProgressBar non definie."
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
Write-Log "Bienvenue dans le Gestionnaire IT de ClicOnLine. Toutes les actions effectuees s'afficheront ici."

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

function Write-LogInfo {
    param([string]$message)
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $textBoxLogs.SelectionColor = [System.Drawing.Color]::Cyan
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
	    Animate-ProgressBar -progressBar $progressBar -durationSeconds 4
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $results = $searcher.Search("IsInstalled=0 and Type='Software'").Updates

        if ($results.Count -eq 0) {
            Write-LogOk "Aucune mise a jour disponible."
            return
        }

        $updateList = ""
        for ($i = 0; $i -lt $results.Count; $i++) {
            $title = $results.Item($i).Title
            $updateList += "$($i+1). $title`n"
        }

        Write-LogAvert "Mises a jour disponibles :"
        Write-Log $updateList.Trim()

        $dialogResult = [System.Windows.Forms.MessageBox]::Show(
            "Les mises a jour suivantes sont disponibles :`n`n$updateList`nVoulez-vous les installer ?",
            "Mises a jour detectees",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $downloader = $session.CreateUpdateDownloader()
            $downloader.Updates = $results
            Write-Log "Telechargement des mises a jour..."
            Animate-ProgressBar -progressBar $progressBar -durationSeconds 2
            $downloader.Download()
            
            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $results
            Write-Log "Installation des mises a jour..."
            Animate-ProgressBar -progressBar $progressBar -durationSeconds 2
            $result = $installer.Install()

            Write-LogOk "Resultat de l'installation : $($result.ResultCode)"
        } else {
            Write-LogAvert "Installation des mises a jour annulee par l'utilisateur."
        }
    } catch {
        Write-LogError "Erreur durant le scan Windows Update : $_"
    }
}

function Repair-WindowsUpdate {
    Write-Log "Reinitialisation complete des composants Windows Update..."

    try {
        # Arret des services necessaires
        Write-Log "Arret des services wuauserv et bits..."
        Animate-ProgressBar -progressBar $progressBar -durationSeconds 2
        Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service bits -Force -ErrorAction SilentlyContinue

        # Suppression du cache de MAJ
        Write-Log "Suppression du contenu de SoftwareDistribution..."
        Animate-ProgressBar -progressBar $progressBar -durationSeconds 3
        Remove-Item -Path "C:\Windows\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue

        # Suppression de catroot2 (base de validation cryptographique)
        Write-Log "Suppression du dossier catroot2..."
        Animate-ProgressBar -progressBar $progressBar -durationSeconds 3
        Remove-Item -Path "C:\Windows\System32\catroot2" -Recurse -Force -ErrorAction SilentlyContinue

        # Reenregistrement des composants Update (optionnel mais utile)
        Write-Log "Reenregistrement des DLLs Windows Update..."
        Animate-ProgressBar -progressBar $progressBar -durationSeconds 2
        $dlls = @("atl.dll","urlmon.dll","mshtml.dll","shdocvw.dll","browseui.dll","jscript.dll","vbscript.dll","scrrun.dll","msxml.dll","msxml3.dll","msxml6.dll","wuapi.dll","wuaueng.dll","wucltui.dll","wups.dll","wups2.dll","wuweb.dll","qmgr.dll","qmgrprxy.dll","wuaueng1.dll")
        foreach ($dll in $dlls) {
            regsvr32 /s $dll 2>$null
        }

        # Redemarrage des services
        Write-Log "Redemarrage des services..."
        Animate-ProgressBar -progressBar $progressBar -durationSeconds 4
        Start-Service bits -ErrorAction SilentlyContinue
        Start-Service wuauserv -ErrorAction SilentlyContinue

        Write-LogOk "Reinitialisation complete terminee."
    } catch {
        Write-LogError "Erreur lors de la reinitialisation : $_"
    }
}


function Force-WindowsUpdateDetection {
    Write-Log "Detection des mises a jour forcee..."

    try {
        $process = Start-Process -FilePath "UsoClient.exe" -ArgumentList "StartScan" -NoNewWindow -PassThru -ErrorAction Stop
        $process.WaitForExit()

        if ($process.ExitCode -eq 0) {
            Write-LogOk "Detection des mises a jour lancee avec succes."
        } else {
            Write-LogAvert "Commande UsoClient terminee avec un code inattendu : $($process.ExitCode)"
        }
    } catch {
        Write-LogError "Erreur lors du lancement de la detection : $_"
    }
}


function Restart-PCCountdown {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous redemarrer leordinateur dans 10 secondes ?`nVous pouvez encore annuler avec shutdown /a.",
        "Confirmation de redemarrage",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        shutdown /r /t 10
        Write-Log "Redemarrage programme dans 10 secondes."
        [System.Windows.Forms.MessageBox]::Show(
            "Le redemarrage est prevu dans 10 secondes.`nUtilisez shutdown /a pour l'annuler.",
            "Redemarrage en attente",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } else {
        Write-Log "Redemarrage annule par l'utilisateur."
    }
}


function Create-SystemRestorePoint {
    try {
        Write-Log "Creation point de restauration systeme..."
        Checkpoint-Computer -Description "Point_Restauration_Outil_IT_ClicOnLine" -RestorePointType "MODIFY_SETTINGS"
        Animate-ProgressBar -progressBar $progressBar -durationSeconds 2
        Write-LogOk "Point de restauration cree."
    } catch {
        Write-LogError "Erreur creation point de restauration : $_"
    }
}


function Scan-InstalledApps {
    Write-Log "Scan des logiciels installes..."
    $softwares = @("Google Chrome", "Adobe Acrobat Reader")
    foreach ($software in $softwares) {
        $found = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
                 Where-Object { $_.DisplayName -like "*$software*" }
        if (!$found) {
            $found = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
                     Where-Object { $_.DisplayName -like "*$software*" }
        }
        if ($found) {
            Write-LogOk "$software trouve : version $($found.DisplayVersion)"
        } else {
            Write-LogAvert "$software non trouve."
        }
    }
}

function Check-ObsoleteDrivers {
    Write-Log "Scan des pilotes obsoletes..."
    Animate-ProgressBar -progressBar $progressBar -durationSeconds 2

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
                Write-LogAvert "$($driver.DeviceName) - $driverDate - $($driver.DriverVersion) - POTENTIELLEMENT OBSOLeTE"
            }
        }
    }

    if ($obsolete.Count -eq 0) {
        Write-LogOk "Aucun pilote obsolete detecte ou desactivable en toute securite."
        return
    }

    Write-LogAvert "Nombre total de pilotes potentiellement obsoletes : $($obsolete.Count)"

    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Il y a $($obsolete.Count) pilotes potentiellement obsoletes.`nSouhaitez-vous les desinstaller ?",
        "Suppression des pilotes obsoletes",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        foreach ($d in $obsolete) {
            try {
                $infName = $d.InfName
                Write-Log "Suppression du pilote : $($d.DeviceName) ($infName)..."
                pnputil /delete-driver "$infName" /uninstall /force /quiet
                Write-Log "Pilote supprime : $infName"
            } catch {
                Write-LogError "Erreur suppression pilote $($d.DeviceName) : $_"
            }
        }
    } else {
        Write-Log "Suppression des pilotes annulee par l'utilisateur."
    }
}

function Show-NetworkConnections {
    Write-Log "Analyse des connexions reseau actives..."
    Animate-ProgressBar -progressBar $progressBar -durationSeconds 2

    try {
        $lines = netstat -anob 2>$null
        $connexions = @()
        $currentApp = ""

        foreach ($line in $lines) {
            $line = $line.Trim()

            # Nom de processus entre crochets
            if ($line -match "^\[(.+)\]$") {
                $currentApp = $matches[1]
                continue
            }

            # Ligne de connexion (TCP/UDP)
            if ($line -match "^(TCP|UDP)\s+(\S+)\s+(\S+)\s+(\S+)") {
                $proto  = $matches[1]
                $local  = $matches[2]
                $remote = $matches[3]
                $state  = $matches[4]

                $connexions += [PSCustomObject]@{
                    Protocole = $proto
                    Local     = $local
                    Distant   = $remote
                    Etat      = $state
                    Processus = $currentApp
                }
            }
        }

        if ($connexions.Count -eq 0) {
            Write-Log "Aucune connexion active detectee."
            return
        }

        foreach ($conn in $connexions) {
            $str = "$($conn.Protocole) $($conn.Local) -> $($conn.Distant) [$($conn.Etat)] ($($conn.Processus))"

            # Connexion potentiellement suspecte : IP externe
            if ($conn.Distant -notmatch '^192\.168\.|^10\.|^172\.(1[6-9]|2\d|3[01])|^127\.') {
                Write-LogError "? Connexion suspecte : $str"
            } else {
                Write-Log $str
            }
        }
    } catch {
        Write-LogError "Erreur durant l'analyse reseau : $_"
    }
}


function Show-SystemHealthDashboard {
    Write-Log "etat de sante du systeme :"

    if (-not (Test-IsAdmin)) {
        Write-LogAvert "ATTENTION : Script non lance en tant qu'administrateur. Certaines informations peuvent etre inaccessibles."
    }

    # Charge CPU
    try {
        $cpuLoad = Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
        Write-Log "Charge CPU : $cpuLoad %"
    } catch {
        Write-LogError "Erreur CPU : $_"
    }

    # RAM
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
        $usedRAM = $totalRAM - $freeRAM
        Write-Log "RAM : $usedRAM / $totalRAM Go utilises"
    } catch {
        Write-LogError "Erreur RAM : $_"
    }

    # Uptime
    try {
        $uptime = (Get-Date) - $os.LastBootUpTime
        Write-Log "Dernier demarrage : $([math]::Floor($uptime.TotalHours)) h $($uptime.Minutes) min"
    } catch {}

    # Batterie
    try {
        $batt = Get-CimInstance Win32_Battery
        if ($batt) {
            Write-Log "Batterie : $($batt.EstimatedChargeRemaining)%"
        }
    } catch {}

    # Sante disque (S.M.A.R.T.)
    try {
        $disks = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus -ErrorAction Stop
        $i = 0
        foreach ($disk in $disks) {
            $status = if ($disk.PredictFailure -eq $false) { "OK" } else { "? echec previsible" }
            Write-Log "Disque $i : S.M.A.R.T. => $status"
            $i++
        }
    } catch {
        Write-LogAvert "Impossible de lire l'etat S.M.A.R.T. des disques (acces refuse ?)"
    }

    # Temperature CPU (si dispo)
    try {
        $temps = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction SilentlyContinue
        foreach ($t in $temps) {
            $celsius = ($t.CurrentTemperature / 10) - 273.15
            $tempStr = [math]::Round($celsius, 1)
            Write-Log "Temperature CPU : $tempStr eC"
        }
    } catch {
        Write-LogAvert "Temperature CPU non accessible (capteur ou droits manquants)"
    }
}


function Check-Antivirus {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Fonction de telechargement robuste de NRnR
    function Download-NortonTool {
        $nortonToolUrl = "https://download.norton.com/nbr/NRnR.exe"
        $nortonToolPath = "$env:TEMP\NRnR.exe"

        Write-Log "Tentative de telechargement de l’outil Norton..."

        # Methode 1 : BITS
        try {
            Start-BitsTransfer -Source $nortonToolUrl -Destination $nortonToolPath -ErrorAction Stop
            Write-LogOk "Telechargement via BITS reussi."
            return $nortonToolPath
        } catch {
            Write-LogAvert "BITS echoue : $_"
        }

        # Methode 2 : curl.exe
        try {
            $curl = "$env:SystemRoot\System32\curl.exe"
            if (Test-Path $curl) {
                $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                $pinfo.FileName = $curl
                $pinfo.Arguments = "-L -o `"$nortonToolPath`" `"$nortonToolUrl`""
                $pinfo.RedirectStandardOutput = $true
                $pinfo.UseShellExecute = $false
                $pinfo.CreateNoWindow = $true
                $proc = [System.Diagnostics.Process]::Start($pinfo)
                $proc.WaitForExit()
                if (Test-Path $nortonToolPath) {
                    Write-LogOk "Telechargement via curl reussi."
                    return $nortonToolPath
                }
            }
        } catch {
            Write-LogAvert "curl echoue : $_"
        }

        # Methode 3 : Invoke-WebRequest
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11
            Invoke-WebRequest -Uri $nortonToolUrl -OutFile $nortonToolPath -UseBasicParsing
            Write-LogOk "Telechargement via Invoke-WebRequest reussi."
            return $nortonToolPath
        } catch {
            Write-LogError "Toutes les methodes de telechargement ont echoue : $_"
            return $null
        }
    }

    # Desinstallation Norton via NRnR
    function Force-Uninstall-Norton {
        $nortonToolPath = Download-NortonTool
        if (-not $nortonToolPath) {
            Write-LogError "echec complet du telechargement de l’outil Norton. Abandon."
            return
        }

        Write-Log "Execution de l’outil Norton en mode suppression silencieuse..."
        try {
            Start-Process -FilePath $nortonToolPath -ArgumentList "/Silent /RemoveOnly" -Wait -NoNewWindow
            Write-LogOk "Norton desinstalle avec succes."
        } catch {
            Write-LogError "Erreur pendant la desinstallation de Norton : $_"
        }

        Remove-Item -Path $nortonToolPath -Force -ErrorAction SilentlyContinue
    }

    # Detection des antivirus
    $antivirus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if (-not $antivirus) {
        Write-LogAvert "Aucun antivirus detecte."
        return
    }

    $uninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $avList = @()
    foreach ($av in $antivirus) {
        $name = $av.displayName
        $match = $null

        foreach ($regPath in $uninstallPaths) {
            $entries = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
            foreach ($entry in $entries) {
                if ($entry.DisplayName -and $entry.DisplayName -like "*$name*") {
                    $match = $entry
                    break
                }
            }
            if ($match) { break }
        }

        $avList += [PSCustomObject]@{
            Nom             = $name
            UninstallString = if ($match) { $match.UninstallString } else { $null }
        }
    }

    # Interface graphique
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Desinstallation des antivirus"
    $form.Size = New-Object System.Drawing.Size(480, 350)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Size = New-Object System.Drawing.Size(440, 220)
    $list.Location = New-Object System.Drawing.Point(10, 10)
    $list.CheckOnClick = $true

    foreach ($av in $avList) {
        $list.Items.Add($av.Nom)
    }

    $form.Controls.Add($list)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Desinstaller"
    $btnOK.Location = New-Object System.Drawing.Point(260, 250)
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(130, 250)
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Desinstallation annulee par l'utilisateur."
        return
    }

    $selected = $list.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucun antivirus selectionne pour suppression."
        return
    }

    foreach ($nom in $selected) {
        $av = $avList | Where-Object { $_.Nom -eq $nom } | Select-Object -First 1
        if ($av.UninstallString) {
            try {
                Write-Log "Desinstallation de $($av.Nom)..."
                $cmd = $av.UninstallString

                if ($cmd -match '^\"(.+?)\"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(" ")[0]
                    $args = $cmd.Substring($exe.Length)
                }

                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($av.Nom) desinstalle avec succes."
            } catch {
                Write-LogError "Erreur lors de la suppression de $($av.Nom) : $_"
            }
        } elseif ($av.Nom -like "*Norton*") {
            Write-LogAvert "Norton detecte sans commande de desinstallation. Utilisation de l’outil officiel..."
            Force-Uninstall-Norton
        } else {
            Write-LogAvert "Aucune commande de desinstallation trouvee pour $($av.Nom)"
        }
    }
}


function Quick-SystemClean {
    $paths = @("$env:TEMP", "$env:WINDIR\Temp", "$env:WINDIR\Prefetch", "$env:WINDIR\SoftwareDistribution\Download")
    foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Write-Log "Nettoye : $path"
        }
    }
}

function Boost-PCPerformance {
    # Boete de selection des actions
    $formBoost = New-Object System.Windows.Forms.Form
    $formBoost.Text = "Optimisation du PC"
    $formBoost.Size = New-Object System.Drawing.Size(420,330)
    $formBoost.StartPosition = "CenterParent"
    $formBoost.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Size = New-Object System.Drawing.Size(380,160)
    $checkList.Location = New-Object System.Drawing.Point(10,10)
    $checkList.CheckOnClick = $true
    $checkList.Items.AddRange(@(
        "Activer le plan Haute Performance",
        "Nettoyer le dossier Temp",
        "Vider la corbeille",
        "Arreter OneDrive",
        "Augmenter la RAM virtuelle (pagefile)",
        "Gerer les apps au demarrage"
    ))
    $formBoost.Controls.Add($checkList)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Lancer"
    $btnOK.Location = New-Object System.Drawing.Point(220,190)
    $btnOK.Add_Click({ $formBoost.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formBoost.Close() })
    $formBoost.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(100,190)
    $btnCancel.Add_Click({ $formBoost.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formBoost.Close() })
    $formBoost.Controls.Add($btnCancel)

    if ($formBoost.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Boost PC annule par leutilisateur."
        return
    }

    foreach ($item in $checkList.CheckedItems) {
        switch ($item) {
            "Activer le plan Haute Performance" {
                powercfg -setactive SCHEME_MIN
                Write-LogOk "Plan Haute performance active."
            }
            "Nettoyer le dossier Temp" {
                Remove-Item \"$env:TEMP\\*\" -Recurse -Force -ErrorAction SilentlyContinue
                Write-LogOk "Dossier TEMP nettoye."
            }
            "Vider la corbeille" {
                (New-Object -ComObject Shell.Application).NameSpace(0x0a).Items() | ForEach-Object {
                    try { Remove-Item $_.Path -Recurse -Force -ErrorAction SilentlyContinue } catch {}
                }
                Write-LogOk "Corbeille videe."
            }
            "Arreter OneDrive" {
                Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
                Write-LogOk "OneDrive arrete."
            }
            "Augmenter la RAM virtuelle" {
          	try {
    		$totalRAM_MB = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB
    		$pagefileSize = [math]::Round($totalRAM_MB * 1.5)
    		Write-Log "RAM installee : $([math]::Round($totalRAM_MB)) Mo => Fichier deechange : $pagefileSize Mo"

   	 	wmic computersystem where name="%computername%" set AutomaticManagedPagefile=False | Out-Null
    		wmic pagefileset where name="C:\\pagefile.sys" set InitialSize=$pagefileSize,MaximumSize=$pagefileSize | Out-Null

    		Write-LogOk "RAM virtuelle configuree e $pagefileSize Mo"
		} catch {
    		Write-LogError "Erreur configuration RAM virtuelle : $_"
		}
		}

            "Gerer les apps au demarrage" {
                try {
                    $startupApps = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location, User
                    $formStartup = New-Object System.Windows.Forms.Form
                    $formStartup.Text = "Applications au demarrage"
                    $formStartup.Size = New-Object System.Drawing.Size(550, 400)
                    $formStartup.StartPosition = "CenterParent"
                    $formStartup.Font = New-Object System.Drawing.Font("Segoe UI", 9)

                    $listView = New-Object System.Windows.Forms.ListView
                    $listView.View = 'Details'
                    $listView.CheckBoxes = $true
                    $listView.FullRowSelect = $true
                    $listView.Size = New-Object System.Drawing.Size(520,300)
                    $listView.Location = New-Object System.Drawing.Point(10,10)
                    $listView.Columns.Add("Nom", 150)
                    $listView.Columns.Add("Chemin", 300)

                    foreach ($app in $startupApps) {
                        $item = New-Object System.Windows.Forms.ListViewItem($app.Name)
                        $item.SubItems.Add($app.Command)
                        $item.Checked = $true  # On suppose qu'ils sont actifs
                        $listView.Items.Add($item)
                    }

                    $formStartup.Controls.Add($listView)

                    $btnDisable = New-Object System.Windows.Forms.Button
                    $btnDisable.Text = "Desactiver les coches"
                    $btnDisable.Location = New-Object System.Drawing.Point(300,320)
                    $btnDisable.Add_Click({
                        foreach ($entry in $listView.Items) {
                            if ($entry.Checked -eq $false) {
                                Write-LogAvert "Applications desactiver manuellement : $($entry.Text)"
                                # Pas d'API directe fiable pour les desactiver (demande teche planifiee ou registry selon contexte)
                            }
                        }
                        $formStartup.Close()
                    })
                    $formStartup.Controls.Add($btnDisable)

                    $formStartup.ShowDialog()
                } catch {
                    Write-LogError "Erreur affichage ou lecture des apps demarrage : $_"
                }
            }
        }
    }

    Write-LogOk "Optimisation terminee."
}

function Uninstall-TargetedApps {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $foundApps = @()

    foreach ($path in $registryPaths) {
        $entries = Get-ItemProperty $path -ErrorAction SilentlyContinue
        foreach ($entry in $entries) {
            if ($entry.DisplayName -and $entry.UninstallString) {
                if (-not ($foundApps | Where-Object { $_.DisplayName -eq $entry.DisplayName })) {
                    $foundApps += $entry
                }
            }
        }
    }

    if ($foundApps.Count -eq 0) {
        Write-Host "Aucune application trouvee pour desinstallation."
        return
    }

    # Interface graphique
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Selectionnez les applications a desinstaller"
    $form.Size = New-Object System.Drawing.Size(500, 450)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Size = New-Object System.Drawing.Size(460, 320)
    $list.Location = New-Object System.Drawing.Point(10, 10)
    $list.CheckOnClick = $true

    foreach ($app in $foundApps) {
        $display = $app.DisplayName
        if ($display -and !$list.Items.Contains($display)) {
            $list.Items.Add($display)
        }
    }

    $form.Controls.Add($list)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Desinstaller"
    $btnOK.Location = New-Object System.Drawing.Point(280, 350)
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(140, 350)
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Desinstallation annulee par l'utilisateur."
        return
    }

    $selected = $list.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucune application selectionnee pour desinstallation."
        return
    }

    foreach ($name in $selected) {
        $app = $foundApps | Where-Object { $_.DisplayName -eq $name } | Select-Object -First 1
        if ($app -and $app.UninstallString) {
            try {
                $cmd = $app.UninstallString
                Write-LogOk "Desinstallation de $($app.DisplayName)..."

                if ($cmd -match '^\"(.+?)\"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(" ")[0]
                    $args = $cmd.Substring($exe.Length)
                }

                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($app.DisplayName) desinstallee avec succes."
            } catch {
                Write-LogError "Erreur lors de la desinstallation de $($app.DisplayName) : $_"
            }
        } else {
            Write-LogInfo "Pas de commande de desinstallation trouvee pour $name"
        }
    }
}


function Update-Apps {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    Write-Log "Recherche des mises a jour disponibles..."

    try {
        $updates = winget upgrade --accept-source-agreements --accept-package-agreements | Select-String "^[^\s]+\s+[^\s]+\s+[^\s]+"
    } catch {
        Write-LogError "Erreur lors de la recuperation des mises a jour : $_"
        return
    }

    if (!$updates -or $updates.Count -eq 0) {
        Write-LogOk "Aucune mise a jour disponible detectee."
        return
    }

    $appList = @()
    foreach ($line in $updates) {
        $columns = ($line -replace '\s{2,}', '|').Split('|')
        if ($columns.Length -ge 3) {
            $appList += [PSCustomObject]@{
                Nom     = $columns[0].Trim()
                Id      = $columns[1].Trim()
                Version = $columns[2].Trim()
            }
        }
    }

    if ($appList.Count -eq 0) {
        Write-LogOk "Aucune mise a jour detectee apres traitement."
        return
    }

    # Interface graphique
    $formUpdate = New-Object System.Windows.Forms.Form
    $formUpdate.Text = "Mises a jour disponibles"
    $formUpdate.Size = New-Object System.Drawing.Size(500, 450)
    $formUpdate.StartPosition = "CenterScreen"
    $formUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Size = New-Object System.Drawing.Size(460, 320)
    $checkList.Location = New-Object System.Drawing.Point(10, 10)
    $checkList.CheckOnClick = $true

    foreach ($app in $appList) {
        $checkList.Items.Add("$($app.Nom) ($($app.Version))")
    }

    $formUpdate.Controls.Add($checkList)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Mettre a jour"
    $btnOK.Location = New-Object System.Drawing.Point(280, 350)
    $btnOK.Add_Click({ $formUpdate.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formUpdate.Close() })
    $formUpdate.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(140, 350)
    $btnCancel.Add_Click({ $formUpdate.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formUpdate.Close() })
    $formUpdate.Controls.Add($btnCancel)

    if ($formUpdate.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-LogAvert "Mise a jour annulee par l'utilisateur."
        return
    }

    $selected = $checkList.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucune application selectionnee pour mise a jour."
        return
    }

    foreach ($item in $selected) {
        $name = $item -replace '\s\([^)]+\)$', ''
        $targetApp = $appList | Where-Object { $_.Nom -eq $name }
        if ($targetApp) {
            try {
                Write-LogOk "Mise a jour de $($targetApp.Nom)..."
                $result = winget upgrade --id "$($targetApp.Id)" --silent --accept-source-agreements --accept-package-agreements 2>&1

                if ($result -match "No applicable update found" -or $result -match "Aucune mise") {
                    Write-LogInfo "$($targetApp.Nom) est deja a jour."
                } elseif ($result -match "error|echec|failed") {
                    Write-LogError "Erreur mise a jour $($targetApp.Nom) : $result"
                } else {
                    Write-LogOk "$($targetApp.Nom) mise a jour terminee."
                }
            } catch {
                Write-LogError "Exception mise a jour $($targetApp.Nom) : $_"
            }
        }
    }
}


function Export-LogHtml {
    $path = "$env:USERPROFILE\Desktop\rapport_maintenance.html"
    $html = $textBoxLogs.Lines -join "<br>"
    Set-Content -Path $path -Value "<html><body><pre>$html</pre></body></html>"
    Write-Log "Export HTML termine : $path"
}

function Check-CriticalServices {
    $services = @("wuauserv", "bits", "WinDefend")
    $nonRunning = @()

    foreach ($svc in $services) {
        $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($null -eq $s) {
            Write-LogError "Service $svc introuvable."
            continue
        }

        if ($s.Status -ne 'Running') {
            Write-LogAvert "$svc NON demarre"
            $nonRunning += $s
        } else {
            Write-LogOk "$svc OK"
        }
    }

    if ($nonRunning.Count -eq 0) {
        Write-Log "Tous les services critiques sont actifs."
        return
    }

    # Creer une interface pour choisir les services a redemarrer
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Demarrer les services arretes"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterParent"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Size = New-Object System.Drawing.Size(360, 180)
    $checkList.Location = New-Object System.Drawing.Point(10, 10)
    $checkList.CheckOnClick = $true

    foreach ($svc in $nonRunning) {
        $checkList.Items.Add($svc.Name)
    }

    $form.Controls.Add($checkList)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Demarrer selection"
    $btnOK.Location = New-Object System.Drawing.Point(220, 210)
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(80, 210)
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Demarrage des services annule par l'utilisateur."
        return
    }

    # Demarrage uniquement des services coches
    foreach ($selectedName in $checkList.CheckedItems) {
        try {
            Start-Service -Name $selectedName -ErrorAction Stop
            Write-LogOk "Service $selectedName demarre avec succes."
        } catch {
            Write-LogError "Erreur demarrage service $selectedName : $_"
        }
    }
}


function Install-WingetIfMissing {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "Winget non installe. Telechargez-le depuis le Microsoft Store."
    } else {
        Write-Log "Winget est deje installe."
    }
}

function Check-OfficeInstallation {
    Write-Log "Scan des versions Office..."
    $officePaths = @(
        "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\Software\Microsoft\Office\$($_)\Common\InstallRoot" # Pour anciens Office
    )

    $found = $false
    foreach ($path in $officePaths) {
        if (Test-Path $path) {
            Get-ItemProperty -Path $path | ForEach-Object {
                Write-LogOk "Office detecte"
                $found = $true
            }
        }
    }

    if (-not $found) {
        Write-LogAvert "Microsoft Office n'a pas ete detecte."
    }
}

function Repair-OfficeClickToRun {
    try {
        Write-Log "Tentative de reparation d'Office (Click-to-Run)..."
        Start-Process -FilePath "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" -ArgumentList "/repair user" -Wait -ErrorAction Stop
        Write-LogOk "Office lance en mode reparation."
    } catch {
        Write-LogError "Erreur lors de la reparation : $_"
    }
}

function Clear-OutlookCache {
    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Outlook",
        "$env:APPDATA\Microsoft\Outlook",
        "$env:USERPROFILE\AppData\Local\Microsoft\Outlook",
        "$env:USERPROFILE\AppData\Local\Temp\Outlook Logging"
    )

    foreach ($path in $cachePaths) {
        if (Test-Path $path) {
            Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Cache Outlook vide : $path"
        }
    }

    Write-LogOk "Nettoyage du cache Outlook termine."
}

function Clean-OutlookTempFolder {
    try {
        $tempPath = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\*\Outlook\Security")."OutlookSecureTempFolder"
        if (Test-Path $tempPath) {
            Remove-Item "$tempPath\*" -Force -Recurse -ErrorAction SilentlyContinue
            Write-LogOk "Dossier temporaire Outlook vide : $tempPath"
        } else {
            Write-LogAvert "Dossier temporaire Outlook non trouve."
        }
    } catch {
        Write-LogError "Erreur nettoyage dossier temp Outlook : $_"
    }
}

function Repair-OutlookPST {
    try {
        $scanpstPath = "$env:ProgramFiles\Common Files\System\MSMAPI\1036\SCANPST.EXE"
        if (-not (Test-Path $scanpstPath)) {
            $scanpstPath = "$env:ProgramFiles (x86)\Common Files\System\MSMAPI\1036\SCANPST.EXE"
        }

        if (Test-Path $scanpstPath) {
            Start-Process -FilePath $scanpstPath -Wait
            Write-Log "L'outil SCANPST a ete lance. L'utilisateur doit choisir le fichier PST/OST a reparer."
        } else {
            Write-LogError "SCANPST.EXE introuvable sur ce systeme."
        }
    } catch {
        Write-LogError "Erreur lors du lancement de SCANPST.EXE : $_"
    }
}

function Show-OutlookProfilesWithRepair {
    try {
        $officeKey = "HKCU:\Software\Microsoft\Office"
        $versions = Get-ChildItem -Path $officeKey | Where-Object { $_.Name -match 'Office\\\d+\.\d+' }

        $allProfiles = @()
        foreach ($v in $versions) {
            $versionPath = $v.PSPath
            $profileRoot = "$versionPath\Outlook\Profiles"
            if (Test-Path $profileRoot) {
                $profiles = Get-ChildItem -Path $profileRoot | Select-Object -ExpandProperty PSChildName
                foreach ($p in $profiles) {
                    $allProfiles += [PSCustomObject]@{
                        Version = $v.PSChildName
                        Profile = $p
                        RegPath = "$profileRoot\$p"
                    }
                }
            }
        }

        if ($allProfiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Aucun profil Outlook trouve.", "Information", 'OK', 'Information')
            return
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Profils Outlook - Selection pour reparation"
        $form.Size = New-Object System.Drawing.Size(500, 400)
        $form.StartPosition = "CenterScreen"

        $checkList = New-Object System.Windows.Forms.CheckedListBox
        $checkList.Size = New-Object System.Drawing.Size(460, 280)
        $checkList.Location = New-Object System.Drawing.Point(10, 10)
        $checkList.CheckOnClick = $true
        $checkList.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        foreach ($entry in $allProfiles) {
            $checkList.Items.Add("Office $($entry.Version) - $($entry.Profile)")
        }

        $form.Controls.Add($checkList)

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "Reparer"
        $btnOK.Location = New-Object System.Drawing.Point(280, 310)
        $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
        $form.Controls.Add($btnOK)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Annuler"
        $btnCancel.Location = New-Object System.Drawing.Point(100, 310)
        $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
        $form.Controls.Add($btnCancel)

        if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "Reparation annulee par l'utilisateur."
            return
        }

        $selected = $checkList.CheckedItems
        if ($selected.Count -eq 0) {
            Write-LogAvert "Aucun profil selectionne."
            return
        }

        # Reparation de base : suppression des fichiers .dat/.xml lies
        Write-LogAvert "Outlook arrete pour maintenance."
        Get-Process -Name outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

        $paths = @("$env:APPDATA\Microsoft\Outlook", "$env:LOCALAPPDATA\Microsoft\Outlook")
        foreach ($path in $paths) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -Include *.dat,*.xml -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Log "Nettoyage config dans : $path"
            }
        }

        foreach ($profileText in $selected) {
            Write-LogOk "Profil repare : $profileText"
        }

        Write-LogOk "Reparation des profils Outlook terminee."

    } catch {
        Write-LogError "Erreur durant la detection ou la reparation des profils Outlook : $_"
    }
}

function Get-SystemInfoPlus {
    $output = @{}

    # Infos systeme
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $computer = Get-CimInstance -ClassName Win32_ComputerSystem
    $bios = Get-CimInstance -ClassName Win32_BIOS
    $cpu = Get-CimInstance -ClassName Win32_Processor
    $gpu = Get-CimInstance -ClassName Win32_VideoController
    $mem = Get-CimInstance -ClassName Win32_PhysicalMemory
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
    $net = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

    # Resume
    $output["Nom de l'ordinateur"] = $env:COMPUTERNAME
    $output["Utilisateur actuel"] = $env:USERNAME
    $output["Administrateur ?"] = if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { "Oui" } else { "Non" }

    $output["Nom de l'OS"] = $os.Caption
    $output["Version OS"] = $os.Version
    $output["Architecture"] = $os.OSArchitecture
    $output["Fabricant"] = $computer.Manufacturer
    $output["Modele"] = $computer.Model
    $output["BIOS Version"] = $bios.SMBIOSBIOSVersion
    $output["CPU"] = $cpu.Name
    $output["Coeurs physiques"] = $cpu.NumberOfCores
    $output["Coeurs logiques"] = $cpu.NumberOfLogicalProcessors
    $output["RAM Totale (GB)"] = "{0:N2}" -f ($mem.Capacity | Measure-Object -Sum).Sum / 1GB

    $output["Disques"] = ($disk | ForEach-Object {
        "$($_.DeviceID) : $([math]::Round($_.Size / 1GB, 2)) GB - Libre: $([math]::Round($_.FreeSpace / 1GB, 2)) GB"
    }) -join " | "

    $output["Reseau"] = ($net | ForEach-Object {
        "$($_.Description): IP $($_.IPAddress -join ", "), MAC $($_.MACAddress), Passerelle $($_.DefaultIPGateway -join ', '), DNS $($_.DNSServerSearchOrder -join ', ')"
    }) -join "`n"

    # Affichage joli
    Write-Log "`n Informations systeme completes :`n" -ForegroundColor Cyan
    foreach ($key in $output.Keys) {
        Write-log ("{0,-25} : {1}" -f $key, $output[$key])
    }
}

function Redemarrer-BorneWildix {
    Add-Type -AssemblyName System.Windows.Forms

    # Identifiants codes en dur
    $username = "admin"
    $password = "monSuperMotDePasse123"

    # Fenêtre pour saisir uniquement l’IP
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Redemarrage Borne DECT Wildix"
    $form.Size = New-Object System.Drawing.Size(300,160)
    $form.StartPosition = "CenterScreen"

    $labelIP = New-Object System.Windows.Forms.Label
    $labelIP.Text = "Adresse IP de la borne :"
    $labelIP.Location = New-Object System.Drawing.Point(10,20)
    $labelIP.Size = New-Object System.Drawing.Size(250,20)

    $textBoxIP = New-Object System.Windows.Forms.TextBox
    $textBoxIP.Location = New-Object System.Drawing.Point(10,50)
    $textBoxIP.Size = New-Object System.Drawing.Size(260,20)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Redemarrer"
    $btnOK.Location = New-Object System.Drawing.Point(90,90)

    $btnOK.Add_Click({
        $ip = $textBoxIP.Text
        if (-not $ip) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez saisir une adresse IP.","Erreur",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $secpasswd = ConvertTo-SecureString $password -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential ($username, $secpasswd)

        try {
            $url = "http://$ip/cgi-bin/reboot"  # a adapter si necessaire
            $response = Invoke-WebRequest -Uri $url -Method POST -Credential $cred -TimeoutSec 10 -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("La borne a bien reçu la commande de redemarrage.","Succes",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("echec de la commande : $($_.Exception.Message)","Erreur",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }

        $form.Close()
    })

    $form.Controls.AddRange(@($labelIP, $textBoxIP, $btnOK))
    $form.ShowDialog()
}




function New-NetworkLocationWithAuth {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Creer un Emplacement Reseau"
    $form.Size = New-Object System.Drawing.Size(420, 320)
    $form.StartPosition = "CenterParent"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    # Chemin reseau
    $labelPath = New-Object System.Windows.Forms.Label
    $labelPath.Text = "Chemin reseau :"
    $labelPath.AutoSize = $true
    $labelPath.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($labelPath)

    $textPath = New-Object System.Windows.Forms.TextBox
    $textPath.Size = New-Object System.Drawing.Size(380, 20)
    $textPath.Location = New-Object System.Drawing.Point(10, 45)
    $form.Controls.Add($textPath)

    # Lettre de lecteur
    $labelLetter = New-Object System.Windows.Forms.Label
    $labelLetter.Text = "Lettre de lecteur :"
    $labelLetter.Location = New-Object System.Drawing.Point(10, 80)
    $labelLetter.AutoSize = $true
    $form.Controls.Add($labelLetter)

    $comboLetter = New-Object System.Windows.Forms.ComboBox
    $comboLetter.Location = New-Object System.Drawing.Point(150, 75)
    $comboLetter.Size = New-Object System.Drawing.Size(60, 20)
    $comboLetter.DropDownStyle = "DropDownList"
    65..90 | ForEach-Object {
    $char = [char]$_
    if (-not (Get-PSDrive -Name $char -ErrorAction SilentlyContinue)) {
        $comboLetter.Items.Add($char)
    }
    }
    if ($comboLetter.Items.Count -gt 0) {
    $comboLetter.SelectedIndex = 0
    } else {
        Write-LogAvert "Aucune lettre de lecteur disponible pour le mappage."
    }
    $form.Controls.Add($comboLetter)

    # Identifiant
    $labelUser = New-Object System.Windows.Forms.Label
    $labelUser.Text = "Nom d'utilisateur reseau :"
    $labelUser.Location = New-Object System.Drawing.Point(10, 110)
    $labelUser.AutoSize = $true
    $form.Controls.Add($labelUser)

    $textUser = New-Object System.Windows.Forms.TextBox
    $textUser.Size = New-Object System.Drawing.Size(300, 20)
    $textUser.Location = New-Object System.Drawing.Point(10, 135)
    $form.Controls.Add($textUser)

    # Mot de passe
    $labelPass = New-Object System.Windows.Forms.Label
    $labelPass.Text = "Mot de passe :"
    $labelPass.Location = New-Object System.Drawing.Point(10, 165)
    $labelPass.AutoSize = $true
    $form.Controls.Add($labelPass)

    $textPass = New-Object System.Windows.Forms.MaskedTextBox
    $textPass.Size = New-Object System.Drawing.Size(300, 20)
    $textPass.PasswordChar = "*"
    $textPass.Location = New-Object System.Drawing.Point(10, 190)
    $form.Controls.Add($textPass)

    # Boutons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Creer"
    $btnOK.Location = New-Object System.Drawing.Point(250, 230)
    $btnOK.Add_Click({
        $drive = $comboLetter.SelectedItem + ":"
        $path = $textPath.Text
        $user = $textUser.Text
        $pass = $textPass.Text

        if (-not ($path -like "\\*")) {
            [System.Windows.Forms.MessageBox]::Show("Chemin reseau invalide.", "Erreur", "OK", "Error")
            return
        }

        try {
            net use $drive /delete /yes | Out-Null
            $cmd = "net use $drive `"$path`" `"$pass`" /user:`"$user`" /persistent:yes"
            Invoke-Expression $cmd

            if (Test-Path "$drive\") {
                Write-LogOk "Lecteur $drive mappe vers $path avec succes."
            } else {
                Write-LogAvert "Le lecteur a ete mappe mais est inaccessible : $drive"
            }
            $form.Close()
        } catch {
            Write-LogError "Erreur de mappage reseau : $_"
        }
    })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(130, 230)
    $btnCancel.Add_Click({ $form.Close() })
    $form.Controls.Add($btnCancel)

    $form.ShowDialog() | Out-Null
}

function Create-LocalAdmin {
    param (
        [string]$username = "cliconline",
        [string]$password = "MotDePasseSuperSecurise123!"
    )

    try {
        if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) {
            Write-LogAvert "Le compte '$username' existe deja."
        } else {
            $securePass = ConvertTo-SecureString $password -AsPlainText -Force
            New-LocalUser -Name $username -Password $securePass -FullName "Compte Admin ClicOnLine" -Description "Compte IT local avec droits admin" -PasswordNeverExpires -UserMayNotChangePassword
            Add-LocalGroupMember -Group "Administrateurs" -Member $username
            Write-LogOk "Compte administrateur '$username' cree avec succes."
        }

        # Masquage du compte dans l'ecran de login
        $regPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }

        New-ItemProperty -Path $regPath -Name $username -PropertyType DWord -Value 0 -Force | Out-Null
        Write-LogOk "Compte '$username' masque de l’ecran de connexion."
    } catch {
        Write-LogError "Erreur creation ou masquage du compte '$username' : $_"
    }
}


# Boutons
New-TabButton $tabMaj "Scanner Windows Update" 20 20 { Scan-WindowsUpdate } "Information"
New-TabButton $tabMaj "Reinitialiser Windows Update" 20 80 { Repair-WindowsUpdate }
New-TabButton $tabMaj "Forcer Windows Update" 400 20 { Force-WindowsUpdateDetection }
New-TabButton $tabMaj "Installer Winget (si absent)" 400 80 { Install-WingetIfMissing }
New-TabButton $tabMaj "Mise a jour logiciels" 20 140 { Update-Apps }

New-TabButton $tabDiag "Scanner Logiciels Installes" 20 20 { Scan-InstalledApps } "Question"
New-TabButton $tabDiag "Verifier pilotes obsoletes" 400 20 { Check-ObsoleteDrivers } 
New-TabButton $tabDiag "Voir connexions reseau" 20 80 { Show-NetworkConnections }
New-TabButton $tabDiag "Tableau de bord sante PC" 400 80 { Show-SystemHealthDashboard } "Question"
New-TabButton $tabDiag "Lister Antivirus installes" 20 140 { Check-Antivirus } "Shield"

New-TabButton $tabNettoyage "Nettoyage rapide du systeme" 20 20 { Quick-SystemClean }
New-TabButton $tabNettoyage "Desinstallation logiciels" 20 80 { Uninstall-TargetedApps }
New-TabButton $tabNettoyage "Creer un point de restauration" 400 80 { Create-SystemRestorePoint }
New-TabButton $tabNettoyage "Redemarrer le PC (20s)" 20 140 { Restart-PCCountdown }

New-TabButton $tabBoost "Booster le PC" 20 20 { Boost-PCPerformance }

New-TabButton $tabRapport "Exporter le rapport (HTML)" 20 20 { Export-LogHtml }
New-TabButton $tabRapport "Verifier services critiques" 400 20 { Check-CriticalServices } "Warning"
New-TabButton $tabRapport "Informations de ce PC (test)" 20 80 { Get-SystemInfoPlus }
New-TabButton $tabRapport "Creer compte admin Cliconline" 400 80 { Get-SystemInfoPlus }

New-TabButton $tabo365 "Verification la presence Office" 20 20 { Check-OfficeInstallation } "Application"
New-TabButton $tabo365 "Reparation Office" 20 80 { Repair-OfficeClickToRun }
New-TabButton $tabo365 "Vider le cache d'Outlook" 20 140 { Clear-OutlookCache }
New-TabButton $tabo365 "Vider les dossiers temporaires d'Outlook" 400 20 { Clean-OutlookTempFolder }
New-TabButton $tabo365 "Reparation fichier PST/OST" 400 80 { Repair-OutlookPST }
New-TabButton $tabo365 "Reparation profil" 400 140 { Show-OutlookProfilesWithRepair }

New-TabButton $tabWildix "Redemarrage borne/telephone (test)" 20 20 { Redemarrer-BorneWildix }

New-TabButton $tabReseaux "Creation connecteur reseaux (test)" 20 20 { New-NetworkLocationWithAuth }

New-TabButton $tabRouteur "Redemarrage routeur DrayTech (test)" 20 20 {  }
New-TabButton $tabRouteur "Redemarrage routeur TP-Link (test)" 20 80 {  }
New-TabButton $tabRouteur "Redemarrage routeur Mikrotik (test)" 20 140 {  }
New-TabButton $tabRouteur "Creer NAT DrayTech (test)" 400 20 {  }
New-TabButton $tabRouteur "Creer NAT TP-Link (test)" 400 80 {  }
New-TabButton $tabRouteur "Creer NAT Mikrotik (test)" 400 140 {  }


$form.ShowDialog()
