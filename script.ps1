# GESTIONNAIRE IT - Maintenance et Diagnostic ClicOnLine

# ----------------------------
# VeRIFICATION DES PRIVILÈGES
# ----------------------------


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
        1 { [System.Drawing.Brushes]::Orange }
        2 { [System.Drawing.Brushes]::PaleVioletRed }
        3 { [System.Drawing.Brushes]::MediumPurple }
        4 { [System.Drawing.Brushes]::LightBlue }
        5 { [System.Drawing.Brushes]::DarkSeaGreen }
        default { [System.Drawing.Brushes]::DarkSeaGreen }
    }

    $tabPage = $sender.TabPages[$e.Index]
    $e.Graphics.FillRectangle($brush, $e.Bounds)

    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'
    $sf.LineAlignment = 'Center'

    $e.Graphics.DrawString($tabPage.Text, $form.Font, [System.Drawing.Brushes]::Black, [System.Drawing.RectangleF]$e.Bounds, $sf)
})



$tabMaj = New-ColoredTabPage "Mise a jour" ([System.Drawing.Color]::LightGoldenrodYellow)
$tabDiag = New-ColoredTabPage "Diagnostic" ([System.Drawing.Color]::LightSalmon)
$tabNettoyage = New-ColoredTabPage "Nettoyage" ([System.Drawing.Color]::LightCoral)
$tabBoost = New-ColoredTabPage "Boost" ([System.Drawing.Color]::Lavender)
$tabRapport = New-ColoredTabPage "Rapports" ([System.Drawing.Color]::LightBlue)
$tabo365 = New-ColoredTabPage "Office 365" ([System.Drawing.Color]::DarkSeaGreen)

$tabControl.TabPages.AddRange(@($tabMaj, $tabDiag, $tabNettoyage, $tabBoost, $tabRapport, $tabo365))
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

function Resize-Image($image, $width, $height) {
    $resized = New-Object System.Drawing.Bitmap $width, $height
    $graphics = [System.Drawing.Graphics]::FromImage($resized)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.DrawImage($image, 0, 0, $width, $height)
    $graphics.Dispose()
    return $resized
}

function New-TabButton {
    param(
        [System.Windows.Forms.TabPage]$tab,
        [string]$text,
        [int]$x,
        [int]$y,
        [ScriptBlock]$action,
        [string]$iconName = $null  # <- on change ici
    )

    $btn = New-Object System.Windows.Forms.Button -Property @{
        Text = $text
        Size = New-Object System.Drawing.Size(330, 40)
        Location = New-Object System.Drawing.Point($x, $y)
        BackColor = [System.Drawing.Color]::LightSteelBlue
        TextAlign = 'Middlecenter'
    }

    if ($iconName) {
        try {
            $icon = Resize-Image ([System.Drawing.SystemIcons]::$iconName.ToBitmap()) 24 24
            $btn.Image = $icon
            $btn.ImageAlign = 'MiddleLeft'
            $btn.TextAlign = 'Middlecenter'
        } catch {
            Write-Warning "Icône système '$iconName' introuvable."
        }
    }

    $btn.Add_Click({
        Set-Status "Action en cours : $text..."
        & $action
        Set-Status "Prêt."
    })

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
            $downloader.Download()

            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $results
            Write-Log "Installation des mises a jour..."
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
        Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service bits -Force -ErrorAction SilentlyContinue

        # Suppression du cache de MAJ
        Write-Log "Suppression du contenu de SoftwareDistribution..."
        Remove-Item -Path "C:\Windows\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue

        # Suppression de catroot2 (base de validation cryptographique)
        Write-Log "Suppression du dossier catroot2..."
        Remove-Item -Path "C:\Windows\System32\catroot2" -Recurse -Force -ErrorAction SilentlyContinue

        # Reenregistrement des composants Update (optionnel mais utile)
        Write-Log "Reenregistrement des DLLs Windows Update..."
        $dlls = @("atl.dll","urlmon.dll","mshtml.dll","shdocvw.dll","browseui.dll","jscript.dll","vbscript.dll","scrrun.dll","msxml.dll","msxml3.dll","msxml6.dll","wuapi.dll","wuaueng.dll","wucltui.dll","wups.dll","wups2.dll","wuweb.dll","qmgr.dll","qmgrprxy.dll","wuaueng1.dll")
        foreach ($dll in $dlls) {
            regsvr32 /s $dll 2>$null
        }

        # Redemarrage des services
        Write-Log "Redemarrage des services..."
        Start-Service bits -ErrorAction SilentlyContinue
        Start-Service wuauserv -ErrorAction SilentlyContinue

        Write-LogOk "Reinitialisation complete terminee."
    } catch {
        Write-LogError "Erreur lors de la reinitialisation : $_"
    }
}


function Force-WindowsUpdateDetection {
    Write-Log "Detection des misesajour forcee..."

    try {
        $process = Start-Process -FilePath "UsoClient.exe" -ArgumentList "StartScan" -NoNewWindow -PassThru -ErrorAction Stop
        $process.WaitForExit()

        if ($process.ExitCode -eq 0) {
            Write-LogOk "Detection des misesajour lancee avec succes."
        } else {
            Write-LogAvert "Commande UsoClient terminee avec un code inattendu : $($process.ExitCode)"
        }
    } catch {
        Write-LogError "Erreur lors du lancement de la detection : $_"
    }
}


function Restart-PCCountdown {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous redemarrer leordinateur dans 20 secondes ?`nVous pouvez encore annuler avec shutdown /a.",
        "Confirmation de redemarrage",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        shutdown /r /t 20
        Write-Log "Redemarrage programme dans 20 secondes."
        [System.Windows.Forms.MessageBox]::Show(
            "Le redemarrage est prevu dans 20 secondes.`nUtilisez shutdown /a pour l'annuler.",
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
        Checkpoint-Computer -Description "Point_Restauration_Outil_IT" -RestorePointType "MODIFY_SETTINGS"
        Write-LogOk "Point de restauration creer."
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
                Write-LogAvert "$($driver.DeviceName) - $driverDate - $($driver.DriverVersion) - POTENTIELLEMENT OBSOLETE"
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
    Animate-ProgressBar -progressBar $progressBar -durationSeconds 4

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


function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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
            Write-Log "Temperature CPU : $tempStr °C"
        }
    } catch {
        Write-LogAvert "Temperature CPU non accessible (capteur ou droits manquants)"
    }
}


function Check-Antivirus {
    $antivirus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if ($antivirus) {
        foreach ($av in $antivirus) {
            Write-LogOk "Antivirus : $($av.displayName)"
        }
    } else {
        Write-LogAvert "Aucun antivirus detecte."
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
        "Augmenter la RAM virtuelle",
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
        Write-Log "Boost PC annule par l'utilisateur."
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
    		Write-Log "RAM installee : $([math]::Round($totalRAM_MB)) Mo => Fichier d'echange : $pagefileSize Mo"

   	 	wmic computersystem where name="%computername%" set AutomaticManagedPagefile=False | Out-Null
    		wmic pagefileset where name="C:\\pagefile.sys" set InitialSize=$pagefileSize,MaximumSize=$pagefileSize | Out-Null

    		Write-LogOk "RAM virtuelle configuree a $pagefileSize Mo"
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
                                Write-LogAvert "desactiver manuellement : $($entry.Text)"
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

    Write-Log "Optimisation terminee."
}

function Uninstall-TargetedApps {
    $appsToRemove = @("OneDrive", "Java", "Driver Booster", "MacAfee")  # Liste de mots-clesafiltrer

    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $foundApps = @()

    foreach ($path in $registryPaths) {
        $entries = Get-ItemProperty $path -ErrorAction SilentlyContinue
        foreach ($entry in $entries) {
            foreach ($pattern in $appsToRemove) {
                if ($entry.DisplayName -and $entry.DisplayName -like "*$pattern*") {
                    if (-not ($foundApps | Where-Object { $_.DisplayName -eq $entry.DisplayName })) {
                        $foundApps += $entry
                    }
                    break
                }
            }
        }
    }

    if ($foundApps.Count -eq 0) {
        Write-LogOk "Aucune application ciblee trouvee pour desinstallation."
        return
    }

    # Creation de la fenetre de selection
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Desinstallation ciblee"
    $form.Size = New-Object System.Drawing.Size(500, 400)
    $form.StartPosition = "CenterParent"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Size = New-Object System.Drawing.Size(460, 280)
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
    $btnOK.Location = New-Object System.Drawing.Point(280, 310)
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(140, 310)
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
                Write-Log "Desinstallation de $($app.DisplayName)..."

                if ($cmd -match '^\"(.+?)\"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(" ")[0]
                    $args = $cmd.Substring($exe.Length)
                }

                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($app.DisplayName) desinstalle."
            } catch {
                Write-LogError "Erreur suppression de $($app.DisplayName) : $_"
            }
        } else {
            Write-LogAvert "Pas de commande de desinstallation pour $name"
        }
    }
}


function Uninstall-Bitdefender {
    Write-Log "Suppression de Bitdefender..."
    Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Bitdefender*" } | ForEach-Object {
        $_.Uninstall() | Out-Null
        Write-Log "$($_.Name) desinstalle."
    }
}

function Update-Apps {
    # Fenetre de selection
    $formUpdate = New-Object System.Windows.Forms.Form
    $formUpdate.Text = "Mise a jour des applications"
    $formUpdate.Size = New-Object System.Drawing.Size(360,220)
    $formUpdate.StartPosition = "CenterParent"
    $formUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Size = New-Object System.Drawing.Size(320,100)
    $checkList.Location = New-Object System.Drawing.Point(10,10)
    $checkList.CheckOnClick = $true
    $checkList.Items.AddRange(@("Google Chrome", "Adobe Acrobat Reader"))
    $formUpdate.Controls.Add($checkList)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Lancer la mise a jour"
    $btnOK.Location = New-Object System.Drawing.Point(180,120)
    $btnOK.Add_Click({ $formUpdate.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formUpdate.Close() })
    $formUpdate.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(60,120)
    $btnCancel.Add_Click({ $formUpdate.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formUpdate.Close() })
    $formUpdate.Controls.Add($btnCancel)

    if ($formUpdate.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Mise a jour annulee par l'utilisateur."
        return
    }

    # Fonction interne pour version Chrome
    function Get-ChromeVersion {
        $paths = @(
            "${env:ProgramFiles(x86)}\\Google\\Chrome\\Application\\chrome.exe",
            "${env:ProgramFiles}\\Google\\Chrome\\Application\\chrome.exe"
        )
        foreach ($p in $paths) {
            if (Test-Path $p) {
                return (Get-Item $p).VersionInfo.ProductVersion
            }
        }
        return "Inconnue"
    }

    # MAJ Chrome
    if ($checkList.CheckedItems -contains "Google Chrome") {
        Write-Log "Mise a jour de Google Chrome..."
        $oldVersion = Get-ChromeVersion
        Write-Log "Version actuelle de Chrome : $oldVersion"

        try {
            Get-Process -Name chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Write-LogOk "Chrome arrete avant mise a jour."
        } catch {
            Write-LogAvert "Chrome non detecte ou echec a l'arret."
        }

        try {
            $chromeResult = winget upgrade --id Google.Chrome --silent --accept-source-agreements --accept-package-agreements 2>&1
            if ($chromeResult -match "No applicable update found" -or $chromeResult -match "Aucune mise") {
                Write-LogOk "Chrome est deja a jour."
            } elseif ($chromeResult -match "error|echec|failed") {
                Write-LogError "Erreur mise a jour Chrome : $chromeResult"
            } else {
                Write-LogOk "Mise a jour Chrome terminee."
            }
        } catch {
            Write-LogError "Exception mise a jour Chrome : $_"
        }

        $newVersion = Get-ChromeVersion
        if ($newVersion -ne $oldVersion) {
            Write-LogOk "Chrome mise a jour : $oldVersion ? $newVersion"
        } else {
            Write-Log "Version Chrome inchangee apres MAJ : $newVersion"
        }
    }

    # MAJ Adobe
    if ($checkList.CheckedItems -contains "Adobe Acrobat Reader") {
        Write-Log "Mise a jour d'Adobe Acrobat Reader..."
        try {
            $adobeResult = winget upgrade --id Adobe.Acrobat.Reader.64-bit --silent --accept-source-agreements --accept-package-agreements 2>&1
            if ($adobeResult -match "No applicable update found" -or $adobeResult -match "Aucune mise") {
                Write-LogOk "Adobe Reader est deja a jour."
            } elseif ($adobeResult -match "error|echec|failed") {
                Write-LogError "Erreur mise a jour Adobe : $adobeResult"
            } else {
                Write-LogOk "Mise a jour Adobe Reader terminee."
            }
        } catch {
            Write-LogError "Exception mise a jour Adobe : $_"
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

    # Creer une interface pour choisir les servicesaredemarrer
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
            Write-Log "L'outil SCANPST a été lance. L'utilisateur doit choisir le fichier PST/OST a reparer."
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
        $form.Text = "Profils Outlook - Sélection pour reparation"
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

        # Réparation de base : suppression des fichiers .dat/.xml liés
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


# Boutons
New-TabButton $tabMaj "Scanner Windows Update" 20 20 { Scan-WindowsUpdate } "Information"
New-TabButton $tabMaj "Reinitialiser Windows Update" 400 20 { Repair-WindowsUpdate }
New-TabButton $tabMaj "Forcer Windows Update" 20 80 { Force-WindowsUpdateDetection }
New-TabButton $tabMaj "Creer un point de restauration" 400 80 { Create-SystemRestorePoint }
New-TabButton $tabMaj "Redemarrer le PC (20s)" 20 140 { Restart-PCCountdown }

New-TabButton $tabDiag "Scanner Logiciels Installes" 20 20 { Scan-InstalledApps } "Question"
New-TabButton $tabDiag "Verifier pilotes obsoletes" 400 20 { Check-ObsoleteDrivers } 
New-TabButton $tabDiag "Voir connexions reseau" 20 80 { Show-NetworkConnections }
New-TabButton $tabDiag "Tableau de bord sante PC" 400 80 { Show-SystemHealthDashboard } "Question"
New-TabButton $tabDiag "Lister Antivirus installes" 20 140 { Check-Antivirus } "Shield"

New-TabButton $tabNettoyage "Nettoyage rapide du systeme" 20 20 { Quick-SystemClean }
New-TabButton $tabNettoyage "Desinstaller applis nefaste" 20 80 { Uninstall-TargetedApps }
New-TabButton $tabNettoyage "Desinstaller Bitdefender (test)" 400 80 { Uninstall-Bitdefender } "Warning"
New-TabButton $tabNettoyage "Mettre a jour applis" 20 140 { Update-Apps }

New-TabButton $tabBoost "Booster le PC" 20 20 { Boost-PCPerformance }

New-TabButton $tabRapport "Exporter le rapport (HTML)" 20 20 { Export-LogHtml }
New-TabButton $tabRapport "Verifier services critiques" 400 20 { Check-CriticalServices } "Warning"
New-TabButton $tabRapport "Installer Winget (si absent)" 20 80 { Install-WingetIfMissing }

New-TabButton $tabo365 "Verification la presence Office" 20 20 { Check-OfficeInstallation } "Application"
New-TabButton $tabo365 "Reparation Office" 20 80 { Repair-OfficeClickToRun }
New-TabButton $tabo365 "Vider le cache d'Outlook" 20 140 { Clear-OutlookCache }
New-TabButton $tabo365 "Vider les dossiers temporaires d'Outlook" 400 20 { Clean-OutlookTempFolder }
New-TabButton $tabo365 "Reparation fichier PST/OST" 400 80 { Repair-OutlookPST }
New-TabButton $tabo365 "Reparation profil" 400 140 { Show-OutlookProfilesWithRepair }

$form.ShowDialog()
