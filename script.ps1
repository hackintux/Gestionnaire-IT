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
        4 { [System.Drawing.Brushes]::LightCoral }
        default { [System.Drawing.Brushes]::LightGray }
    }

    $tabPage = $sender.TabPages[$e.Index]
    $e.Graphics.FillRectangle($brush, $e.Bounds)

    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'
    $sf.LineAlignment = 'Center'

    $e.Graphics.DrawString($tabPage.Text, $form.Font, [System.Drawing.Brushes]::Black, [System.Drawing.RectangleF]$e.Bounds, $sf)
})



$tabMaj = New-ColoredTabPage "Mise a jour" ([System.Drawing.Color]::LightGoldenrodYellow)
$tabDiag = New-ColoredTabPage "Diagnostic" ([System.Drawing.Color]::Lavender)
$tabNettoyage = New-ColoredTabPage "Nettoyage" ([System.Drawing.Color]::Honeydew)
$tabBoost = New-ColoredTabPage "Boost" ([System.Drawing.Color]::Honeydew)
$tabRapport = New-ColoredTabPage "Rapports" ([System.Drawing.Color]::MistyRose)

$tabControl.TabPages.AddRange(@($tabMaj, $tabDiag, $tabNettoyage, $tabBoost, $tabRapport))
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
        Write-LogError "ProgressBar non d�finie."
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
Write-Log "Bienvenue dans le Gestionnaire IT de ClicOnLine. Toutes les actions effectu�es s'afficheront ici."

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
            Write-LogOk "Aucune mise � jour disponible."
            return
        }

        $updateList = ""
        for ($i = 0; $i -lt $results.Count; $i++) {
            $title = $results.Item($i).Title
            $updateList += "$($i+1). $title`n"
        }

        Write-LogAvert "Mises � jour disponibles :"
        Write-Log $updateList.Trim()

        $dialogResult = [System.Windows.Forms.MessageBox]::Show(
            "Les mises � jour suivantes sont disponibles :`n`n$updateList`nVoulez-vous les installer ?",
            "Mises � jour d�tect�es",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $downloader = $session.CreateUpdateDownloader()
            $downloader.Updates = $results
            Write-Log "T�l�chargement des mises � jour..."
            $downloader.Download()

            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $results
            Write-Log "Installation des mises � jour..."
            $result = $installer.Install()

            Write-LogOk "R�sultat de l'installation : $($result.ResultCode)"
        } else {
            Write-LogAvert "Installation des mises � jour annul�e par l'utilisateur."
        }
    } catch {
        Write-LogError "Erreur durant le scan Windows Update : $_"
    }
}

function Repair-WindowsUpdate {
    Write-Log "R�initialisation compl�te des composants Windows Update..."

    try {
        # Arr�t des services n�cessaires
        Write-Log "Arr�t des services wuauserv et bits..."
        Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service bits -Force -ErrorAction SilentlyContinue

        # Suppression du cache de MAJ
        Write-Log "Suppression du contenu de SoftwareDistribution..."
        Remove-Item -Path "C:\Windows\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue

        # Suppression de catroot2 (base de validation cryptographique)
        Write-Log "Suppression du dossier catroot2..."
        Remove-Item -Path "C:\Windows\System32\catroot2" -Recurse -Force -ErrorAction SilentlyContinue

        # R�enregistrement des composants Update (optionnel mais utile)
        Write-Log "R�enregistrement des DLLs Windows Update..."
        $dlls = @("atl.dll","urlmon.dll","mshtml.dll","shdocvw.dll","browseui.dll","jscript.dll","vbscript.dll","scrrun.dll","msxml.dll","msxml3.dll","msxml6.dll","wuapi.dll","wuaueng.dll","wucltui.dll","wups.dll","wups2.dll","wuweb.dll","qmgr.dll","qmgrprxy.dll","wuaueng1.dll")
        foreach ($dll in $dlls) {
            regsvr32 /s $dll 2>$null
        }

        # Red�marrage des services
        Write-Log "Red�marrage des services..."
        Start-Service bits -ErrorAction SilentlyContinue
        Start-Service wuauserv -ErrorAction SilentlyContinue

        Write-LogOk "R�initialisation compl�te termin�e."
    } catch {
        Write-LogError "Erreur lors de la r�initialisation : $_"
    }
}


function Force-WindowsUpdateDetection {
    Write-Log "D�tection des mises � jour forc�e..."

    try {
        $process = Start-Process -FilePath "UsoClient.exe" -ArgumentList "StartScan" -NoNewWindow -PassThru -ErrorAction Stop
        $process.WaitForExit()

        if ($process.ExitCode -eq 0) {
            Write-LogOk "D�tection des mises � jour lanc�e avec succ�s."
        } else {
            Write-LogAvert "Commande UsoClient termin�e avec un code inattendu : $($process.ExitCode)"
        }
    } catch {
        Write-LogError "Erreur lors du lancement de la d�tection : $_"
    }
}


function Restart-PCCountdown {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous red�marrer l�ordinateur dans 20 secondes ?`nVous pouvez encore annuler avec shutdown /a.",
        "Confirmation de red�marrage",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        shutdown /r /t 20
        Write-Log "Red�marrage programm� dans 20 secondes."
        [System.Windows.Forms.MessageBox]::Show(
            "Le red�marrage est pr�vu dans 20 secondes.`nUtilisez shutdown /a pour l'annuler.",
            "Red�marrage en attente",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } else {
        Write-Log "Red�marrage annul� par l�utilisateur."
    }
}


function Create-SystemRestorePoint {
    try {
        Write-Log "Cr�ation point de restauration syst�me..."
        Checkpoint-Computer -Description "Point_Restauration_Outil_IT" -RestorePointType "MODIFY_SETTINGS"
        Write-LogOk "Point de restauration cr��."
    } catch {
        Write-LogError "Erreur cr�ation point de restauration : $_"
    }
}


function Scan-InstalledApps {
    Write-Log "Scan des logiciels install�s..."
    $softwares = @("Google Chrome", "Adobe Acrobat Reader")
    foreach ($software in $softwares) {
        $found = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
                 Where-Object { $_.DisplayName -like "*$software*" }
        if (!$found) {
            $found = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
                     Where-Object { $_.DisplayName -like "*$software*" }
        }
        if ($found) {
            Write-LogOk "$software trouv� : version $($found.DisplayVersion)"
        } else {
            Write-LogAvert "$software non trouv�."
        }
    }
}

function Check-ObsoleteDrivers {
    Write-Log "Scan des pilotes obsol�tes..."

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
                Write-LogAvert "$($driver.DeviceName) - $driverDate - $($driver.DriverVersion) - POTENTIELLEMENT OBSOL�TE"
            }
        }
    }

    if ($obsolete.Count -eq 0) {
        Write-LogOk "Aucun pilote obsol�te d�tect� ou d�sactivable en toute s�curit�."
        return
    }

    Write-LogAvert "Nombre total de pilotes potentiellement obsol�tes : $($obsolete.Count)"

    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Il y a $($obsolete.Count) pilotes potentiellement obsol�tes.`nSouhaitez-vous les d�sinstaller ?",
        "Suppression des pilotes obsol�tes",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        foreach ($d in $obsolete) {
            try {
                $infName = $d.InfName
                Write-Log "Suppression du pilote : $($d.DeviceName) ($infName)..."
                pnputil /delete-driver "$infName" /uninstall /force /quiet
                Write-Log "Pilote supprim� : $infName"
            } catch {
                Write-LogError "Erreur suppression pilote $($d.DeviceName) : $_"
            }
        }
    } else {
        Write-Log "Suppression des pilotes annul�e par l'utilisateur."
    }
}

function Show-NetworkConnections {
    Write-Log "Analyse des connexions r�seau actives..."
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
            Write-Log "Aucune connexion active d�tect�e."
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
        Write-LogError "Erreur durant l'analyse r�seau : $_"
    }
}


function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-SystemHealthDashboard {
    Write-Log "�tat de sant� du syst�me :"

    if (-not (Test-IsAdmin)) {
        Write-LogAvert "ATTENTION : Script non lanc� en tant qu'administrateur. Certaines informations peuvent �tre inaccessibles."
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
        Write-Log "RAM : $usedRAM / $totalRAM Go utilis�s"
    } catch {
        Write-LogError "Erreur RAM : $_"
    }

    # Uptime
    try {
        $uptime = (Get-Date) - $os.LastBootUpTime
        Write-Log "Dernier d�marrage : $([math]::Floor($uptime.TotalHours)) h $($uptime.Minutes) min"
    } catch {}

    # Batterie
    try {
        $batt = Get-CimInstance Win32_Battery
        if ($batt) {
            Write-Log "Batterie : $($batt.EstimatedChargeRemaining)%"
        }
    } catch {}

    # Sant� disque (S.M.A.R.T.)
    try {
        $disks = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus -ErrorAction Stop
        $i = 0
        foreach ($disk in $disks) {
            $status = if ($disk.PredictFailure -eq $false) { "OK" } else { "? �chec pr�visible" }
            Write-Log "Disque $i : S.M.A.R.T. => $status"
            $i++
        }
    } catch {
        Write-LogAvert "Impossible de lire l'�tat S.M.A.R.T. des disques (acc�s refus� ?)"
    }

    # Temp�rature CPU (si dispo)
    try {
        $temps = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction SilentlyContinue
        foreach ($t in $temps) {
            $celsius = ($t.CurrentTemperature / 10) - 273.15
            $tempStr = [math]::Round($celsius, 1)
            Write-Log "Temp�rature CPU : $tempStr �C"
        }
    } catch {
        Write-LogAvert "Temp�rature CPU non accessible (capteur ou droits manquants)"
    }
}


function Check-Antivirus {
    $antivirus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if ($antivirus) {
        foreach ($av in $antivirus) {
            Write-Log "Antivirus : $($av.displayName)"
        }
    } else {
        Write-LogAvert "Aucun antivirus d�tect�."
    }
}

function Quick-SystemClean {
    $paths = @("$env:TEMP", "$env:WINDIR\Temp", "$env:WINDIR\Prefetch", "$env:WINDIR\SoftwareDistribution\Download")
    foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Write-Log "Nettoy� : $path"
        }
    }
}

function Boost-PCPerformance {
    # Bo�te de s�lection des actions
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
        "Arr�ter OneDrive",
        "Augmenter la RAM virtuelle (pagefile)",
        "G�rer les apps au d�marrage"
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
        Write-Log "Boost PC annul� par l�utilisateur."
        return
    }

    foreach ($item in $checkList.CheckedItems) {
        switch ($item) {
            "Activer le plan Haute Performance" {
                powercfg -setactive SCHEME_MIN
                Write-LogOk "Plan Haute performance activ�."
            }
            "Nettoyer le dossier Temp" {
                Remove-Item \"$env:TEMP\\*\" -Recurse -Force -ErrorAction SilentlyContinue
                Write-LogOk "Dossier TEMP nettoy�."
            }
            "Vider la corbeille" {
                (New-Object -ComObject Shell.Application).NameSpace(0x0a).Items() | ForEach-Object {
                    try { Remove-Item $_.Path -Recurse -Force -ErrorAction SilentlyContinue } catch {}
                }
                Write-LogOk "Corbeille vid�e."
            }
            "Arr�ter OneDrive" {
                Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
                Write-LogOk "OneDrive arr�t�."
            }
            "Augmenter la RAM virtuelle" {
          	try {
    		$totalRAM_MB = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB
    		$pagefileSize = [math]::Round($totalRAM_MB * 1.5)
    		Write-Log "RAM install�e : $([math]::Round($totalRAM_MB)) Mo => Fichier d��change : $pagefileSize Mo"

   	 	wmic computersystem where name="%computername%" set AutomaticManagedPagefile=False | Out-Null
    		wmic pagefileset where name="C:\\pagefile.sys" set InitialSize=$pagefileSize,MaximumSize=$pagefileSize | Out-Null

    		Write-LogOk "RAM virtuelle configur�e � $pagefileSize Mo"
		} catch {
    		Write-LogError "Erreur configuration RAM virtuelle : $_"
		}
		}

            "G�rer les apps au d�marrage" {
                try {
                    $startupApps = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location, User
                    $formStartup = New-Object System.Windows.Forms.Form
                    $formStartup.Text = "Applications au d�marrage"
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
                    $btnDisable.Text = "D�sactiver les coch�s"
                    $btnDisable.Location = New-Object System.Drawing.Point(300,320)
                    $btnDisable.Add_Click({
                        foreach ($entry in $listView.Items) {
                            if ($entry.Checked -eq $false) {
                                Write-LogAvert "� d�sactiver manuellement : $($entry.Text)"
                                # Pas d'API directe fiable pour les d�sactiver (demande t�che planifi�e ou registry selon contexte)
                            }
                        }
                        $formStartup.Close()
                    })
                    $formStartup.Controls.Add($btnDisable)

                    $formStartup.ShowDialog()
                } catch {
                    Write-LogError "Erreur affichage ou lecture des apps d�marrage : $_"
                }
            }
        }
    }

    Write-Log "Optimisation termin�e."
}

function Uninstall-TargetedApps {
    $appsToRemove = @("OneDrive", "Java", "Driver Booster", "MacAfee")  # Liste de mots-cl�s � filtrer

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
        Write-LogOk "Aucune application cibl�e trouv�e pour d�sinstallation."
        return
    }

    # Cr�ation de la fen�tre de s�lection
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "D�sinstallation cibl�e"
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
    $btnOK.Text = "D�sinstaller"
    $btnOK.Location = New-Object System.Drawing.Point(280, 310)
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(140, 310)
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "D�sinstallation annul�e par l�utilisateur."
        return
    }

    $selected = $list.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucune application s�lectionn�e pour d�sinstallation."
        return
    }

    foreach ($name in $selected) {
        $app = $foundApps | Where-Object { $_.DisplayName -eq $name } | Select-Object -First 1
        if ($app -and $app.UninstallString) {
            try {
                $cmd = $app.UninstallString
                Write-Log "D�sinstallation de $($app.DisplayName)..."

                if ($cmd -match '^\"(.+?)\"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(" ")[0]
                    $args = $cmd.Substring($exe.Length)
                }

                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($app.DisplayName) d�sinstall�."
            } catch {
                Write-LogError "Erreur suppression de $($app.DisplayName) : $_"
            }
        } else {
            Write-LogAvert "Pas de commande de d�sinstallation pour $name"
        }
    }
}


function Uninstall-Bitdefender {
    Write-Log "Suppression de Bitdefender..."
    Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Bitdefender*" } | ForEach-Object {
        $_.Uninstall() | Out-Null
        Write-Log "$($_.Name) d�sinstall�."
    }
}

function Update-Apps {
    # Fen�tre de s�lection
    $formUpdate = New-Object System.Windows.Forms.Form
    $formUpdate.Text = "Mise � jour des applications"
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
    $btnOK.Text = "Lancer la mise � jour"
    $btnOK.Location = New-Object System.Drawing.Point(180,120)
    $btnOK.Add_Click({ $formUpdate.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formUpdate.Close() })
    $formUpdate.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(60,120)
    $btnCancel.Add_Click({ $formUpdate.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formUpdate.Close() })
    $formUpdate.Controls.Add($btnCancel)

    if ($formUpdate.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Mise � jour annul�e par l�utilisateur."
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
        Write-Log "Mise � jour de Google Chrome..."
        $oldVersion = Get-ChromeVersion
        Write-Log "Version actuelle de Chrome : $oldVersion"

        try {
            Get-Process -Name chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Write-LogOk "Chrome arr�t� avant mise � jour."
        } catch {
            Write-LogAvert "Chrome non d�tect� ou �chec � l'arr�t."
        }

        try {
            $chromeResult = winget upgrade --id Google.Chrome --silent --accept-source-agreements --accept-package-agreements 2>&1
            if ($chromeResult -match "No applicable update found" -or $chromeResult -match "Aucune mise") {
                Write-LogOk "Chrome est d�j� � jour."
            } elseif ($chromeResult -match "error|�chec|failed") {
                Write-LogError "Erreur mise � jour Chrome : $chromeResult"
            } else {
                Write-LogOk "Mise � jour Chrome termin�e."
            }
        } catch {
            Write-LogError "Exception mise � jour Chrome : $_"
        }

        $newVersion = Get-ChromeVersion
        if ($newVersion -ne $oldVersion) {
            Write-LogOk "Chrome mis � jour : $oldVersion ? $newVersion"
        } else {
            Write-Log "Version Chrome inchang�e apr�s MAJ : $newVersion"
        }
    }

    # MAJ Adobe
    if ($checkList.CheckedItems -contains "Adobe Acrobat Reader") {
        Write-Log "Mise � jour d'Adobe Acrobat Reader..."
        try {
            $adobeResult = winget upgrade --id Adobe.Acrobat.Reader.64-bit --silent --accept-source-agreements --accept-package-agreements 2>&1
            if ($adobeResult -match "No applicable update found" -or $adobeResult -match "Aucune mise") {
                Write-LogOk "Adobe Reader est d�j� � jour."
            } elseif ($adobeResult -match "error|�chec|failed") {
                Write-LogError "Erreur mise � jour Adobe : $adobeResult"
            } else {
                Write-LogOk "Mise � jour Adobe Reader termin�e."
            }
        } catch {
            Write-LogError "Exception mise � jour Adobe : $_"
        }
    }
}


function Export-LogHtml {
    $path = "$env:USERPROFILE\Desktop\rapport_maintenance.html"
    $html = $textBoxLogs.Lines -join "<br>"
    Set-Content -Path $path -Value "<html><body><pre>$html</pre></body></html>"
    Write-Log "Export HTML termin� : $path"
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
            Write-LogAvert "$svc NON d�marr�"
            $nonRunning += $s
        } else {
            Write-LogOk "$svc OK"
        }
    }

    if ($nonRunning.Count -eq 0) {
        Write-Log "Tous les services critiques sont actifs."
        return
    }

    # Cr�er une interface pour choisir les services � red�marrer
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "D�marrer les services arr�t�s"
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
    $btnOK.Text = "D�marrer s�lection"
    $btnOK.Location = New-Object System.Drawing.Point(220, 210)
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(80, 210)
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "D�marrage des services annul� par l'utilisateur."
        return
    }

    # D�marrage uniquement des services coch�s
    foreach ($selectedName in $checkList.CheckedItems) {
        try {
            Start-Service -Name $selectedName -ErrorAction Stop
            Write-LogOk "Service $selectedName d�marr� avec succ�s."
        } catch {
            Write-LogError "Erreur d�marrage service $selectedName : $_"
        }
    }
}


function Install-WingetIfMissing {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "Winget non install�. T�l�chargez-le depuis le Microsoft Store."
    } else {
        Write-Log "Winget est d�j� install�."
    }
}

# Boutons
New-TabButton $tabMaj "Scanner Windows Update" 20 20 { Scan-WindowsUpdate }
New-TabButton $tabMaj "R�initialiser Windows Update" 400 20 { Repair-WindowsUpdate }
New-TabButton $tabMaj "Forcer Windows Update" 20 80 { Force-WindowsUpdateDetection }
New-TabButton $tabMaj "Cr�er un point de restauration" 400 80 { Create-SystemRestorePoint }
New-TabButton $tabMaj "Red�marrer le PC (20s)" 20 140 { Restart-PCCountdown }

New-TabButton $tabDiag "Scanner Logiciels Install�s" 20 20 { Scan-InstalledApps }
New-TabButton $tabDiag "V�rifier pilotes obsol�tes" 400 20 { Check-ObsoleteDrivers }
New-TabButton $tabDiag "Voir connexions r�seau" 20 80 { Show-NetworkConnections }
New-TabButton $tabDiag "Tableau de bord sant� PC" 400 80 { Show-SystemHealthDashboard }
New-TabButton $tabDiag "Lister Antivirus install�s" 20 140 { Check-Antivirus }

New-TabButton $tabNettoyage "Nettoyage rapide du syst�me" 20 20 { Quick-SystemClean }
New-TabButton $tabNettoyage "D�sinstaller applis n�faste" 20 80 { Uninstall-TargetedApps }
New-TabButton $tabNettoyage "D�sinstaller Bitdefender (test)" 400 80 { Uninstall-Bitdefender }
New-TabButton $tabNettoyage "Mettre � jour applis" 20 140 { Update-Apps }

New-TabButton $tabBoost "Booster le PC" 400 20 { Boost-PCPerformance }

New-TabButton $tabRapport "Exporter le rapport (HTML)" 20 20 { Export-LogHtml }
New-TabButton $tabRapport "V�rifier services critiques" 400 20 { Check-CriticalServices }
New-TabButton $tabRapport "Installer Winget (si absent)" 20 80 { Install-WingetIfMissing }

$form.ShowDialog()
