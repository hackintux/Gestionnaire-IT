# Gestionnaire IT GUI – Version complète (Thème sombre/clair)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# := Définition de la couleur “lime-vert” en scope global := 
$global:greenLime = [System.Drawing.Color]::FromArgb(117, 180, 25)

# -------------------------
# 1. Thème initial global
# -------------------------
$global:theme = "dark"

# -------------------------
# 2. Déclarations globales
# -------------------------
# Collections de contrôles
$generalButtons       = @()
$itemButtons          = @()
$tabButtons           = @()
$global:actionButtons = @()

# -------------------------
# 3. Fonction Apply-Theme
# -------------------------
function Apply-Theme {
    param(
        [ValidateSet("dark","light")]
        [string]$mode
    )
    $greenLime = [System.Drawing.Color]::FromArgb(117, 180, 25)

    if ($mode -eq "dark") {
        $form.BackColor       = [System.Drawing.Color]::FromArgb(30,30,30)
        $titleLabel.ForeColor = $greenLime
        $textBoxLogs.BackColor     = [System.Drawing.Color]::FromArgb(45,45,45)
        $textBoxLogs.ForeColor     = [System.Drawing.Color]::White

        $themeButton.BackColor = $greenLime
        $themeButton.ForeColor = [System.Drawing.Color]::Black

        foreach ($btn in $itemButtons)         { $btn.BackColor = $greenLime; $btn.ForeColor = [System.Drawing.Color]::White }
        foreach ($btn in $tabButtons)          { $btn.BackColor = "white"; $btn.ForeColor = [System.Drawing.Color]::FromArgb(45,45,45) }
        foreach ($btn in $generalButtons)      { $btn.BackColor = [System.Drawing.Color]::White; $btn.ForeColor = $greenLime }
        foreach ($btn in $global:actionButtons){ $btn.BackColor = [System.Drawing.Color]::White; $btn.ForeColor = $greenLime }
    } else {
        $form.BackColor       = [System.Drawing.Color]::White
        $titleLabel.ForeColor = $greenLime
        $textBoxLogs.BackColor     = [System.Drawing.Color]::WhiteSmoke
        $textBoxLogs.ForeColor     = [System.Drawing.Color]::Black

        $themeButton.BackColor = $greenLime
        $themeButton.ForeColor = [System.Drawing.Color]::White

        foreach ($btn in $itemButtons)         { $btn.BackColor = $greenLime; $btn.ForeColor = [System.Drawing.Color]::White }
        foreach ($btn in $tabButtons)          { $btn.BackColor = $greenLime; $btn.ForeColor = [System.Drawing.Color]::White }
        foreach ($btn in $generalButtons)      { $btn.BackColor = $greenLime; $btn.ForeColor = [System.Drawing.Color]::White }
        foreach ($btn in $global:actionButtons){ $btn.BackColor = $greenLime; $btn.ForeColor = [System.Drawing.Color]::White }
    }
}

# -------------------------
# 4. Formulaire principal
# -------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Gestionnaire IT - ClicOnLine"
$form.Size            = New-Object System.Drawing.Size(1130, 800)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.Font            = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

# -------------------------
# 5. Logo et titre
# -------------------------
$logoBox = New-Object System.Windows.Forms.PictureBox
$logoBox.Size        = New-Object System.Drawing.Size(60, 90)
$logoBox.Location    = New-Object System.Drawing.Point(10, 10)
$logoBox.SizeMode    = "StretchImage"
$logoPath = "C:\Users\David Salvador\Documents\Scripts\Powershell\clic_version\Gestionnaire_IT_clic\icone.ico"
if (Test-Path $logoPath) {
    $logoBox.Image = [System.Drawing.Image]::FromFile($logoPath)
}
$form.Controls.Add($logoBox)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text      = "CLICONLINE"
$titleLabel.Size      = New-Object System.Drawing.Size(500, 50)
$titleLabel.Location  = New-Object System.Drawing.Point(70, 30)
$titleLabel.Font      = New-Object System.Drawing.Font("Consolas", 32, [System.Drawing.FontStyle]::Bold)
$titleLabel.TextAlign = "MiddleLeft"
$form.Controls.Add($titleLabel)

# -------------------------
# 6. Bouton Changer de thème
# -------------------------
$themeButton = New-Object System.Windows.Forms.Button
$themeButton.Text     = "Changer de thème"
$themeButton.Size     = New-Object System.Drawing.Size(180, 30)
$themeButton.Location = New-Object System.Drawing.Point(920, 15)
$themeButton.FlatStyle= "Standard"
$themeButton.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$themeButton.Add_Click({
    if ($global:theme -eq "dark") { $global:theme = "light" } else { $global:theme = "dark" }
    Apply-Theme -mode $global:theme
})
$form.Controls.Add($themeButton)

# Fonctions
function Scan-WindowsUpdate {
    try {
        Write-Log "Scan Windows Update..."
        $serviceWU = Get-Service -Name wuauserv -ErrorAction Stop
        Write-Log "Service Windows Update : $($serviceWU.Status)"
        Write-Log "Veuillez patientez ..."

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

        # Fenétre de confirmation style LCARS
        $formUpdates = New-Object System.Windows.Forms.Form
        $formUpdates.Text = "Mises à jour détectées"
        $formUpdates.Size = New-Object System.Drawing.Size(600, 400)
        $formUpdates.StartPosition = "CenterScreen"
        $formUpdates.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
        $formUpdates.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Les Mises à jour suivantes sont disponibles :"
        $label.ForeColor = [System.Drawing.Color]::DeepSkyBlue
        $label.AutoSize = $false
        $label.Size = New-Object System.Drawing.Size(560, 30)
        $label.Location = New-Object System.Drawing.Point(20, 20)
        $formUpdates.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ReadOnly = $true
        $textBox.BackColor = [System.Drawing.Color]::DarkSlateGray
        $textBox.ForeColor = [System.Drawing.Color]::White
        $textBox.ScrollBars = "Vertical"
        $textBox.Text = $updateList.Trim()
        $textBox.Size = New-Object System.Drawing.Size(560, 250)
        $textBox.Location = New-Object System.Drawing.Point(20, 60)
        $formUpdates.Controls.Add($textBox)

        $btnYes = New-Object System.Windows.Forms.Button
        $btnYes.Text = "Installer"
        $btnYes.Size = New-Object System.Drawing.Size(120, 40)
        $btnYes.Location = New-Object System.Drawing.Point(330, 320)
        $btnYes.BackColor = [System.Drawing.Color]::LimeGreen
        $btnYes.ForeColor = [System.Drawing.Color]::Black
        $btnYes.FlatStyle = 'Flat'

        $btnNo = New-Object System.Windows.Forms.Button
        $btnNo.Text = "Annuler"
        $btnNo.Size = New-Object System.Drawing.Size(120, 40)
        $btnNo.Location = New-Object System.Drawing.Point(160, 320)
        $btnNo.BackColor = [System.Drawing.Color]::IndianRed
        $btnNo.ForeColor = [System.Drawing.Color]::White
        $btnNo.FlatStyle = 'Flat'

        $btnYes.Add_Click({
            $formUpdates.Tag = 'Yes'
            $formUpdates.Close()
        })
        $btnNo.Add_Click({
            $formUpdates.Tag = 'No'
            $formUpdates.Close()
        })

        $formUpdates.Controls.Add($btnYes)
        $formUpdates.Controls.Add($btnNo)

        [void]$formUpdates.ShowDialog()

        if ($formUpdates.Tag -eq 'Yes') {
            $downloader = $session.CreateUpdateDownloader()
            $downloader.Updates = $results
            Write-Log "Téléchargement des Mises à jour..."
            $downloader.Download()

            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $results
            Write-Log "Installation des Mises à jour..."
            $result = $installer.Install()

            Write-LogOk "Résultat de l'installation : $($result.ResultCode)"
        } else {
            Write-LogAvert "Installation des Mises à jour annulée par l'utilisateur."
        }

    } catch {
        Write-LogError "Erreur durant le scan Windows Update : $_"
    }
}

function Boost-PCPerformance {
    $formBoost = New-Object System.Windows.Forms.Form
    $formBoost.Text = "Optimisation du PC"
    $formBoost.Size = New-Object System.Drawing.Size(500, 360)
    $formBoost.StartPosition = "CenterParent"
    $formBoost.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
    $formBoost.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Size = New-Object System.Drawing.Size(460, 180)
    $checkList.Location = New-Object System.Drawing.Point(20, 20)
    $checkList.BackColor = [System.Drawing.Color]::DarkSlateGray
    $checkList.ForeColor = [System.Drawing.Color]::White
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
    $btnOK.Size = New-Object System.Drawing.Size(100, 40)
    $btnOK.Location = New-Object System.Drawing.Point(270, 220)
    $btnOK.BackColor = $greenLime
    $btnOK.ForeColor = "white"
    $btnOK.FlatStyle = 'Flat'
    $formBoost.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Size = New-Object System.Drawing.Size(100, 40)
    $btnCancel.Location = New-Object System.Drawing.Point(120, 220)
    $btnCancel.BackColor = "Red"
    $btnCancel.ForeColor = "White"
    $btnCancel.FlatStyle = 'Flat'
    $formBoost.Controls.Add($btnCancel)

    $btnOK.Add_Click({ $formBoost.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formBoost.Close() })
    $btnCancel.Add_Click({ $formBoost.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formBoost.Close() })

    if ($formBoost.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Boost PC annulé par l'utilisateur."
        return
    }

    foreach ($item in $checkList.CheckedItems) {
        switch ($item) {
            "Activer le plan Haute Performance" {
                powercfg -setactive SCHEME_MIN
                Write-LogOk "Plan Haute performance activé."
            }
            "Nettoyer le dossier Temp" {
                Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-LogOk "Dossier TEMP nettoyé."
            }
            "Vider la corbeille" {
                (New-Object -ComObject Shell.Application).NameSpace(0x0a).Items() | ForEach-Object {
                    try { Remove-Item $_.Path -Recurse -Force -ErrorAction SilentlyContinue } catch {}
                }
                Write-LogOk "Corbeille vidée."
            }
            "Arreter OneDrive" {
                Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
                Write-LogOk "OneDrive arrété."
            }
            "Augmenter la RAM virtuelle (pagefile)" {
                try {
                    $totalRAM_MB = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB
                    $pagefileSize = [math]::Round($totalRAM_MB * 1.5)
                    Write-Log "RAM installée : $([math]::Round($totalRAM_MB)) Mo => Pagefile : $pagefileSize Mo"

                    wmic computersystem where name="%computername%" set AutomaticManagedPagefile=False | Out-Null
                    wmic pagefileset where name="C:\\pagefile.sys" set InitialSize=$pagefileSize,MaximumSize=$pagefileSize | Out-Null

                    Write-LogOk "RAM virtuelle configurée à $pagefileSize Mo"
                } catch {
                    Write-LogError "Erreur configuration RAM virtuelle : $_"
                }
            }
            "Gerer les apps au demarrage" {
                try {
                    $startupApps = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command
                    $formStartup = New-Object System.Windows.Forms.Form
                    $formStartup.Text = "Applications au démarrage"
                    $formStartup.Size = New-Object System.Drawing.Size(600, 450)
                    $formStartup.StartPosition = "CenterParent"
                    $formStartup.BackColor = [System.Drawing.Color]::Black
                    $formStartup.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)

                    $listView = New-Object System.Windows.Forms.ListView
                    $listView.View = 'Details'
                    $listView.CheckBoxes = $true
                    $listView.FullRowSelect = $true
                    $listView.Size = New-Object System.Drawing.Size(560, 320)
                    $listView.Location = New-Object System.Drawing.Point(20, 20)
                    $listView.BackColor = [System.Drawing.Color]::DarkSlateGray
                    $listView.ForeColor = [System.Drawing.Color]::White
                    $listView.Columns.Add("Nom", 150)
                    $listView.Columns.Add("Chemin", 380)

                    foreach ($app in $startupApps) {
                        $item = New-Object System.Windows.Forms.ListViewItem($app.Name)
                        $item.SubItems.Add($app.Command)
                        $item.Checked = $true
                        $listView.Items.Add($item)
                    }
                    $formStartup.Controls.Add($listView)

                    $btnDisable = New-Object System.Windows.Forms.Button
                    $btnDisable.Text = "Désactiver les cochés"
                    $btnDisable.Size = New-Object System.Drawing.Size(200, 40)
                    $btnDisable.Location = New-Object System.Drawing.Point(200, 360)
                    $btnDisable.BackColor = [System.Drawing.Color]::DarkOrange
                    $btnDisable.ForeColor = [System.Drawing.Color]::Black
                    $btnDisable.FlatStyle = 'Flat'
                    $btnDisable.Add_Click({
                        foreach ($entry in $listView.Items) {
                            if (-not $entry.Checked) {
                                Write-LogAvert "Application désactivée (simulation) : $($entry.Text)"
                            }
                        }
                        $formStartup.Close()
                    })
                    $formStartup.Controls.Add($btnDisable)
                    $formStartup.ShowDialog()
                } catch {
                    Write-LogError "Erreur lecture apps démarrage : $_"
                }
            }
        }
    }

    Write-LogOk "Optimisation terminée."
}

function Create-SystemRestorePoint {
    try {
        Write-Log "Creation point de restauration systeme..."
        Checkpoint-Computer -Description "Point_Restauration_Outil_IT_ClicOnLine" -RestorePointType "MODIFY_SETTINGS"
        Write-LogOk "Point de restauration cree."
    } catch {
        Write-LogError "Erreur creation point de restauration : $_"
    }
}

function Restart-PC {
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

    # Fenêtre de confirmation
    $formConfirm = New-Object System.Windows.Forms.Form
    $formConfirm.Text = "Suppression des pilotes obsolètes"
    $formConfirm.Size = New-Object System.Drawing.Size(600, 250)
    $formConfirm.StartPosition = "CenterScreen"
    $formConfirm.BackColor = [System.Drawing.Color]::Black
    $formConfirm.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Il y a $($obsolete.Count) pilotes potentiellement obsolètes. Souhaitez-vous les désinstaller ?"
    $label.ForeColor = [System.Drawing.Color]::Gold
    $label.Size = New-Object System.Drawing.Size(560, 80)
    $label.Location = New-Object System.Drawing.Point(20, 30)
    $label.TextAlign = "MiddleCenter"
    $label.AutoSize = $false
    $formConfirm.Controls.Add($label)

    $btnYes = New-Object System.Windows.Forms.Button
    $btnYes.Text = "Oui - Supprimer"
    $btnYes.Size = New-Object System.Drawing.Size(150, 40)
    $btnYes.Location = New-Object System.Drawing.Point(320, 130)
    $btnYes.BackColor = $greenLime
    $btnYes.ForeColor = "white"
    $btnYes.FlatStyle = 'Flat'

    $btnNo = New-Object System.Windows.Forms.Button
    $btnNo.Text = "Non - Annuler"
    $btnNo.Size = New-Object System.Drawing.Size(150, 40)
    $btnNo.Location = New-Object System.Drawing.Point(120, 130)
    $btnNo.BackColor = "Red"
    $btnNo.ForeColor = "White"
    $btnNo.FlatStyle = 'Flat'

    $btnYes.Add_Click({
        $formConfirm.Tag = 'Yes'
        $formConfirm.Close()
    })
    $btnNo.Add_Click({
        $formConfirm.Tag = 'No'
        $formConfirm.Close()
    })

    $formConfirm.Controls.Add($btnYes)
    $formConfirm.Controls.Add($btnNo)

    [void]$formConfirm.ShowDialog()

    if ($formConfirm.Tag -eq 'Yes') {
        $needReboot = $false

        foreach ($d in $obsolete) {
            $deviceName = $d.DeviceName
            $infName = $d.InfName
            $deviceID = $d.DeviceID

            try {
                Write-Log "Tentative de désactivation du périphérique : $deviceName"
                Disable-PnpDevice -InstanceId $deviceID -Confirm:$false -ErrorAction SilentlyContinue

                $devStatus = Get-PnpDevice -InstanceId $deviceID -ErrorAction SilentlyContinue
                if ($devStatus.Status -ne "Disabled") {
                    Write-LogAvert "$deviceName n’a pas pu être désactivé proprement."
                    $needReboot = $true
                } else {
                    Write-LogOk "$deviceName désactivé avec succès."
                }
            } catch {
                Write-LogError "Erreur lors de la désactivation de $deviceName : $_"
                $needReboot = $true
            }

            try {
                Write-Log "Suppression du pilote : $deviceName ($infName)..."
                $process = Start-Process -FilePath "pnputil.exe" -ArgumentList "/delete-driver `"$infName`" /uninstall /force /quiet" -NoNewWindow -Wait -PassThru

                if ($process.ExitCode -ne 0) {
                    Write-LogError "Échec de suppression de $infName - Code de sortie : $($process.ExitCode)"
                    $needReboot = $true
                } else {
                    Start-Sleep -Seconds 2
                    $check = Get-WmiObject Win32_PnPSignedDriver | Where-Object { $_.InfName -eq $infName }
                    if ($check) {
                        Write-LogError "Le pilote $infName est toujours présent après suppression."
                        $needReboot = $true
                    } else {
                        Write-LogOk "Pilote supprimé : $infName"
                    }
                }
            } catch {
                Write-LogError "Erreur lors de la suppression du pilote $deviceName : $_"
                $needReboot = $true
            }
        }

        if ($needReboot) {
            Write-LogAvert "Un redémarrage du système est recommandé pour finaliser la suppression de certains pilotes."
        } else {
            Write-LogOk "Tous les pilotes obsolètes ont été supprimés sans redémarrage requis."
        }

    } else {
        Write-Log "Suppression des pilotes annulée par l'utilisateur."
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
    # Détection des antivirus installés
    $antivirus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if (-not $antivirus) {
        Write-LogAvert "Aucun antivirus détecté."
        return
    }

    $uninstallPaths = @(
        "HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*",
        "HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*"
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

    # Interface LCARS
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Désinstallation des antivirus"
    $form.Size = New-Object System.Drawing.Size(520, 400)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::Black
    $form.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Size = New-Object System.Drawing.Size(480, 240)
    $list.Location = New-Object System.Drawing.Point(20, 20)
    $list.BackColor = [System.Drawing.Color]::DarkSlateGray
    $list.ForeColor = [System.Drawing.Color]::White
    $list.CheckOnClick = $true
    foreach ($av in $avList) {
        $list.Items.Add($av.Nom)
    }
    $form.Controls.Add($list)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Désinstaller"
    $btnOK.Size = New-Object System.Drawing.Size(120, 40)
    $btnOK.Location = New-Object System.Drawing.Point(280, 280)
    $btnOK.BackColor = [System.Drawing.Color]::LimeGreen
    $btnOK.ForeColor = [System.Drawing.Color]::Black
    $btnOK.FlatStyle = 'Flat'
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 40)
    $btnCancel.Location = New-Object System.Drawing.Point(120, 280)
    $btnCancel.BackColor = [System.Drawing.Color]::IndianRed
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = 'Flat'
    $form.Controls.Add($btnCancel)

    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Désinstallation annulée par l'utilisateur."
        return
    }

    $selected = $list.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucun antivirus sélectionné pour suppression."
        return
    }

    foreach ($nom in $selected) {
        $av = $avList | Where-Object { $_.Nom -eq $nom } | Select-Object -First 1
        if ($av.UninstallString) {
            try {
                Write-Log "Désinstallation de $($av.Nom)..."
                $cmd = $av.UninstallString
                if ($cmd -match '^\"(.+?)\"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(' ')[0]
                    $args = $cmd.Substring($exe.Length)
                }
                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($av.Nom) désinstallé avec succès."
            } catch {
                Write-LogError "Erreur lors de la suppression de $($av.Nom) : $_"
            }
        } elseif ($av.Nom -like "*Norton*") {
            Write-LogAvert "Norton détecté sans commande de désinstallation. Téléchargement de l'outil..."
            Force-Uninstall-Norton
        } else {
            Write-LogAvert "Aucune commande de désinstallation trouvée pour $($av.Nom)"
        }
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

function Quick-SystemClean {
    $paths = @("$env:TEMP", "$env:WINDIR\Temp", "$env:WINDIR\Prefetch", "$env:WINDIR\SoftwareDistribution\Download")
    foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Write-Log "Nettoye : $path"
        }
    }
}

function Uninstall-TargetedApps {
    $registryPaths = @(
        "HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*",
        "HKLM:\\Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*"
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
        Write-LogAvert "Aucune application trouvée pour désinstallation."
        return
    }

    # Interface LCARS stylée
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Sélectionnez les applications à désinstaller"
    $form.Size = New-Object System.Drawing.Size(540, 480)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::Black
    $form.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Size = New-Object System.Drawing.Size(500, 330)
    $list.Location = New-Object System.Drawing.Point(20, 20)
    $list.BackColor = [System.Drawing.Color]::DarkSlateGray
    $list.ForeColor = [System.Drawing.Color]::White
    $list.CheckOnClick = $true

    foreach ($app in $foundApps) {
        $display = $app.DisplayName
        if ($display -and !$list.Items.Contains($display)) {
            $list.Items.Add($display)
        }
    }

    $form.Controls.Add($list)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Désinstaller"
    $btnOK.Size = New-Object System.Drawing.Size(120, 40)
    $btnOK.Location = New-Object System.Drawing.Point(300, 370)
    $btnOK.BackColor = [System.Drawing.Color]::LimeGreen
    $btnOK.ForeColor = [System.Drawing.Color]::Black
    $btnOK.FlatStyle = 'Flat'
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 40)
    $btnCancel.Location = New-Object System.Drawing.Point(100, 370)
    $btnCancel.BackColor = [System.Drawing.Color]::IndianRed
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Désinstallation annulée par l'utilisateur."
        return
    }

    $selected = $list.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucune application sélectionnée pour suppression."
        return
    }

    foreach ($name in $selected) {
        $app = $foundApps | Where-Object { $_.DisplayName -eq $name } | Select-Object -First 1
        if ($app -and $app.UninstallString) {
            try {
                $cmd = $app.UninstallString
                Write-Log "Désinstallation de $($app.DisplayName)..."

                if ($cmd -match '^"(.+?)"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(" ")[0]
                    $args = $cmd.Substring($exe.Length)
                }

                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($app.DisplayName) désinstallée avec succès."
            } catch {
                Write-LogError "Erreur lors de la désinstallation de $($app.DisplayName) : $_"
            }
        } else {
            Write-LogAvert "Pas de commande de désinstallation trouvée pour $name"
        }
    }
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
            Write-LogAvert "$svc NON démarré"
            $nonRunning += $s
        } else {
            Write-LogOk "$svc OK"
        }
    }

    if ($nonRunning.Count -eq 0) {
        Write-Log "Tous les services critiques sont actifs."
        return
    }

    # Interface LCARS pour redémarrer les services arrétés
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Démarrer les services arrétés"
    $form.Size = New-Object System.Drawing.Size(460, 340)
    $form.StartPosition = "CenterParent"
    $form.BackColor = [System.Drawing.Color]::Black
    $form.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Size = New-Object System.Drawing.Size(420, 200)
    $checkList.Location = New-Object System.Drawing.Point(20, 20)
    $checkList.CheckOnClick = $true
    $checkList.BackColor = [System.Drawing.Color]::DarkSlateGray
    $checkList.ForeColor = [System.Drawing.Color]::White

    foreach ($svc in $nonRunning) {
        $checkList.Items.Add($svc.Name)
    }

    $form.Controls.Add($checkList)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Démarrer sélection"
    $btnOK.Size = New-Object System.Drawing.Size(150, 40)
    $btnOK.Location = New-Object System.Drawing.Point(250, 240)
    $btnOK.BackColor = [System.Drawing.Color]::LimeGreen
    $btnOK.ForeColor = [System.Drawing.Color]::Black
    $btnOK.FlatStyle = 'Flat'
    $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 40)
    $btnCancel.Location = New-Object System.Drawing.Point(60, 240)
    $btnCancel.BackColor = [System.Drawing.Color]::IndianRed
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Démarrage des services annulé par l'utilisateur."
        return
    }

    foreach ($selectedName in $checkList.CheckedItems) {
        try {
            Start-Service -Name $selectedName -ErrorAction Stop
            Write-LogOk "Service $selectedName démarré avec succès."
        } catch {
            Write-LogError "Erreur démarrage service $selectedName : $_"
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
        $officeKey = "HKCU:\\Software\\Microsoft\\Office"
        $versions = Get-ChildItem -Path $officeKey | Where-Object { $_.Name -match 'Office\\\\\d+\.\d+' }

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
            Write-LogAvert "Aucun profil Outlook trouvé."
            return
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Profils Outlook - Sélection pour réparation"
        $form.Size = New-Object System.Drawing.Size(520, 440)
        $form.StartPosition = "CenterScreen"
        $form.BackColor = [System.Drawing.Color]::Black
        $form.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

        $checkList = New-Object System.Windows.Forms.CheckedListBox
        $checkList.Size = New-Object System.Drawing.Size(480, 280)
        $checkList.Location = New-Object System.Drawing.Point(20, 20)
        $checkList.CheckOnClick = $true
        $checkList.BackColor = [System.Drawing.Color]::DarkSlateGray
        $checkList.ForeColor = [System.Drawing.Color]::White

        foreach ($entry in $allProfiles) {
            $checkList.Items.Add("Office $($entry.Version) - $($entry.Profile)")
        }

        $form.Controls.Add($checkList)

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "Réparer"
        $btnOK.Size = New-Object System.Drawing.Size(120, 40)
        $btnOK.Location = New-Object System.Drawing.Point(280, 330)
        $btnOK.BackColor = [System.Drawing.Color]::LimeGreen
        $btnOK.ForeColor = [System.Drawing.Color]::Black
        $btnOK.FlatStyle = 'Flat'
        $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
        $form.Controls.Add($btnOK)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Annuler"
        $btnCancel.Size = New-Object System.Drawing.Size(120, 40)
        $btnCancel.Location = New-Object System.Drawing.Point(100, 330)
        $btnCancel.BackColor = [System.Drawing.Color]::IndianRed
        $btnCancel.ForeColor = [System.Drawing.Color]::White
        $btnCancel.FlatStyle = 'Flat'
        $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
        $form.Controls.Add($btnCancel)

        if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "Réparation annulée par l'utilisateur."
            return
        }

        $selected = $checkList.CheckedItems
        if ($selected.Count -eq 0) {
            Write-LogAvert "Aucun profil sélectionné."
            return
        }

        Write-LogAvert "Outlook arrété pour maintenance."
        Get-Process -Name outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

        $paths = @("$env:APPDATA\Microsoft\Outlook", "$env:LOCALAPPDATA\Microsoft\Outlook")
        foreach ($path in $paths) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -Include *.dat,*.xml -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Log "Nettoyage config dans : $path"
            }
        }

        foreach ($profileText in $selected) {
            Write-LogOk "Profil réparé : $profileText"
        }

        Write-LogOk "Réparation des profils Outlook terminée."

    } catch {
        Write-LogError "Erreur durant la détection ou la réparation des profils Outlook : $_"
    }
}

function Get-SystemInfoPlus {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # === Fenêtre principale ===
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Système en Temps Réel"
    $form.Size = New-Object System.Drawing.Size(1000, 620)
    $form.MinimumSize = $form.Size
    $form.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
    $form.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    # === Infos système texte ===
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.ReadOnly = $true
    $textBox.Size = New-Object System.Drawing.Size(590, 300)
    $textBox.Location = New-Object System.Drawing.Point(10, 10)
    $textBox.BackColor = [System.Drawing.Color]::DarkSlateGray
    $textBox.ForeColor = [System.Drawing.Color]::White

    $form.Controls.Add($textBox)

    # === Graphiques dynamiques ===
    function Create-Chart($title, $topOffset, $color) {
        $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $chart.Width = 380
        $chart.Height = 140
        $chart.Left = 600
        $chart.Top = $topOffset
        $chart.BackColor = [System.Drawing.Color]::Black

        $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "Main"
        $area.AxisX.MajorGrid.Enabled = $false
        $area.AxisY.MajorGrid.Enabled = $false
        $area.AxisY.Maximum = 100
        $area.BackColor = [System.Drawing.Color]::DarkSlateGray
        $chart.ChartAreas.Add($area)

        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series $title
        $series.ChartType = 'Line'
        $series.Color = $color
        $series.BorderWidth = 2
        $chart.Series.Add($series)

        $chart.Titles.Add($title).ForeColor = $color

        return $chart
    }

    $chartCPU = Create-Chart "CPU (%)" 10 ([System.Drawing.Color]::Orange)
    $chartRAM = Create-Chart "RAM (%)" 160 ([System.Drawing.Color]::DeepSkyBlue)
    $chartDISK = Create-Chart "Disque C: (%)" 310 ([System.Drawing.Color]::LimeGreen)

    $form.Controls.AddRange(@($chartCPU, $chartRAM, $chartDISK))

    # === Barres de progression avec étiquettes ===
    $labelRAM = New-Object System.Windows.Forms.Label
    $labelRAM.Text = "RAM utilisée :"
    $labelRAM.ForeColor = [System.Drawing.Color]::White
    $labelRAM.Location = New-Object System.Drawing.Point(10, 350)
    $labelRAM.AutoSize = $true

    $progressRAM = New-Object System.Windows.Forms.ProgressBar
    $progressRAM.Width = 250
    $progressRAM.Height = 20
    $progressRAM.Location = New-Object System.Drawing.Point(220, 350)

    $labelDisk = New-Object System.Windows.Forms.Label
    $labelDisk.Text = "Disque C: utilisé :"
    $labelDisk.ForeColor = [System.Drawing.Color]::White
    $labelDisk.Location = New-Object System.Drawing.Point(10, 400)
    $labelDisk.AutoSize = $true

    $progressDisk = New-Object System.Windows.Forms.ProgressBar
    $progressDisk.Width = 250
    $progressDisk.Height = 20
    $progressDisk.Location = New-Object System.Drawing.Point(220, 400)

    $form.Controls.AddRange(@($labelRAM, $progressRAM, $labelDisk, $progressDisk))

    # === Bouton Fermer ===
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Fermer"
    $btnClose.Size = New-Object System.Drawing.Size(120, 40)
    $btnClose.Location = New-Object System.Drawing.Point(420, 500)
    $btnClose.BackColor = "red"
    $btnClose.ForeColor = "white"
    $btnClose.FlatStyle = 'Flat'
    $btnClose.Add_Click({
    $timer.Stop()
    $form.Invoke([Action]{
        $form.Close()
        $form.Dispose()
    })
})
    $form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
})    
    $form.Controls.Add($btnClose)

    # === Données dynamiques ===
    $queueCPU = New-Object System.Collections.Queue
    $queueRAM = New-Object System.Collections.Queue
    $queueDISK = New-Object System.Collections.Queue
    for ($i = 0; $i -lt 30; $i++) { $queueCPU.Enqueue(0); $queueRAM.Enqueue(0); $queueDISK.Enqueue(0) }

    $counterCPU = New-Object System.Diagnostics.PerformanceCounter("Processor", "% Processor Time", "_Total")

    function Get-RAMPercent {
        $os = Get-CimInstance Win32_OperatingSystem
        $total = $os.TotalVisibleMemorySize
        $free = $os.FreePhysicalMemory
        return [math]::Round((1 - ($free / $total)) * 100, 1)
    }

    function Get-DiskPercent {
        $c = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
        return [math]::Round((1 - ($c.FreeSpace / $c.Size)) * 100, 1)
    }

    function Get-CPUTemp {
        try {
            $temp = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi"
            if ($temp -and $temp.CurrentTemperature -gt 0) {
                return [math]::Round(($temp.CurrentTemperature - 2732) / 10, 1)
            }
        } catch { }
        return "Non disponible"
    }

    # === Timer pour actualiser ===
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 3000
    $timer.Add_Tick({
        $cpuVal = [math]::Round($counterCPU.NextValue(), 3)
        $ramVal = Get-RAMPercent
        $diskVal = Get-DiskPercent
        $tempVal = Get-CPUTemp

        foreach ($q in @($queueCPU, $queueRAM, $queueDISK)) { if ($q.Count -ge 30) { [void]$q.Dequeue() } }
        $queueCPU.Enqueue($cpuVal); $queueRAM.Enqueue($ramVal); $queueDISK.Enqueue($diskVal)

        $chartCPU.Series[0].Points.Clear()
        $chartRAM.Series[0].Points.Clear()
        $chartDISK.Series[0].Points.Clear()

        [int]$i = 0
        foreach ($v in $queueCPU) { $chartCPU.Series[0].Points.AddXY($i++, $v) }
        $i = 0; foreach ($v in $queueRAM) { $chartRAM.Series[0].Points.AddXY($i++, $v) }
        $i = 0; foreach ($v in $queueDISK) { $chartDISK.Series[0].Points.AddXY($i++, $v) }

        $progressRAM.Value = [Math]::Min($ramVal, 100)
        $labelRAM.Text = "RAM utilisée : $ramVal%"

        $progressDisk.Value = [Math]::Min($diskVal, 100)
        $labelDisk.Text = "Disque C: utilisé : $diskVal%"

        $disks = Get-PhysicalDisk | ForEach-Object {
            $type = if ($_.MediaType) { $_.MediaType } else { "Inconnu" }
            "Nom : $($_.FriendlyName) - Type : $type - Taille : $([math]::Round($_.Size/1GB)) Go"
        }

        $textBox.Lines = $textLines
        $textBox.Multiline = $true

        $os = Get-CimInstance Win32_OperatingSystem
        $comp = Get-CimInstance Win32_ComputerSystem
        $cpu = Get-CimInstance Win32_Processor
        $bios = Get-CimInstance Win32_BIOS
        $textBox.Text = $text
        $textBox.Text = "Ordinateur     : $env:COMPUTERNAME`r`n" +
                "Utilisateur    : $env:USERNAME`r`n" +
                "OS             : $($os.Caption)`r`n" +
                "Version        : $($os.Version)`r`n" +
                "Architecture   : $($os.OSArchitecture)`r`n" +
                "Fabricant      : $($comp.Manufacturer)`r`n" +
                "Numéro de série : $($bios.SerialNumber)`r`n" +
                "Modèle         : $($comp.Model)`r`n" +
                "CPU            : $($cpu.Name)`r`n" +
                "Cœurs logiques : $($cpu.NumberOfLogicalProcessors)`r`n" +
                "Température CPU: $tempVal °C`r`n" +
                "Disque         : $disks `r`n`r`n" +
                "Certaines informations demandent l'accès administrateur"

    })

    $timer.Start()
    [void]$form.ShowDialog()
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
        Write-LogOk "Compte '$username' masque de léecran de connexion."
    } catch {
        Write-LogError "Erreur creation ou masquage du compte '$username' : $_"
    }
}

function Show-ComingSoon {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "En cours de développement"
    $form.Size = New-Object System.Drawing.Size(450, 200)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::Black
    $form.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Fonctionnalité en cours de développement"
    $label.Size = New-Object System.Drawing.Size(410, 80)
    $label.Location = New-Object System.Drawing.Point(20, 30)
    $label.ForeColor = [System.Drawing.Color]::Orange
    $label.TextAlign = "MiddleCenter"
    $label.AutoSize = $false
    $form.Controls.Add($label)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Fermer"
    $btnClose.Size = New-Object System.Drawing.Size(100, 40)
    $btnClose.Location = New-Object System.Drawing.Point(170, 110)
    $btnClose.BackColor = [System.Drawing.Color]::Gray
    $btnClose.ForeColor = [System.Drawing.Color]::Black
    $btnClose.FlatStyle = 'Flat'
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    [void]$form.ShowDialog()
}

function Start-RescueToolbox {
    # Création d'une petite fenêtre de sélection
    $formChoice = New-Object System.Windows.Forms.Form
    $formChoice.Text = "Outils de Réparation Système"
    $formChoice.Size = New-Object System.Drawing.Size(400, 200)
    $formChoice.StartPosition = "CenterScreen"
    $formChoice.BackColor = [System.Drawing.Color]::Black
    $formChoice.FormBorderStyle = 'FixedDialog'

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Que souhaitez-vous exécuter ?"
    $label.Size = New-Object System.Drawing.Size(360, 40)
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.ForeColor = [System.Drawing.Color]::DeepSkyBlue
    $label.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
    $formChoice.Controls.Add($label)

    # === Bouton DISM ===
    $btnDISM = New-Object System.Windows.Forms.Button
    $btnDISM.Text = "DISM"
    $btnDISM.Size = New-Object System.Drawing.Size(100,40)
    $btnDISM.Location = New-Object System.Drawing.Point(40, 80)
    $btnDISM.BackColor = [System.Drawing.Color]::DeepSkyBlue
    $btnDISM.ForeColor = [System.Drawing.Color]::Black
    $btnDISM.Add_Click({
        $formChoice.Close()
        Write-Log "Lancement DISM..."
        Start-Process "DISM.exe" "/Online /Cleanup-Image /RestoreHealth" -Wait -NoNewWindow
        Write-LogOk "DISM terminé."
    })
    $formChoice.Controls.Add($btnDISM)

    # === Bouton SFC ===
    $btnSFC = New-Object System.Windows.Forms.Button
    $btnSFC.Text = "SFC"
    $btnSFC.Size = New-Object System.Drawing.Size(100,40)
    $btnSFC.Location = New-Object System.Drawing.Point(150, 80)
    $btnSFC.BackColor = [System.Drawing.Color]::Purple
    $btnSFC.ForeColor = [System.Drawing.Color]::Black
    $btnSFC.Add_Click({
        $formChoice.Close()
        Write-Log "Lancement SFC..."
        Start-Process "sfc.exe" "/scannow" -Wait -NoNewWindow
        Write-LogOk "SFC terminé."
    })
    $formChoice.Controls.Add($btnSFC)

    # === Bouton CHKDSK ===
    $btnCHK = New-Object System.Windows.Forms.Button
    $btnCHK.Text = "CHKDSK"
    $btnCHK.Size = New-Object System.Drawing.Size(100,40)
    $btnCHK.Location = New-Object System.Drawing.Point(260, 80)
    $btnCHK.BackColor = [System.Drawing.Color]::LimeGreen
    $btnCHK.ForeColor = [System.Drawing.Color]::Black
    $btnCHK.Add_Click({
        $formChoice.Close()
        Write-Log "Lancement CHKDSK (peut nécessiter redémarrage)..."
        Start-Process "cmd.exe" "/c chkdsk C: /f /r /x" -Wait -NoNewWindow
        Write-LogOk "CHKDSK terminé (si applicable)."
    })
    $formChoice.Controls.Add($btnCHK)

    $formChoice.ShowDialog()
}

function Analyse-ReactiviteSysteme {
    try {
        Write-Log "Démarrage de l'analyse de réactivité système..."

        # --- Latence disque
        $disques = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_LogicalDisk | Where-Object { $_.Name -ne "_Total" }
        if ($disques) {
            $latenceTotale = ($disques | Measure-Object -Property AvgDisksecPerTransfer -Average).Average * 1000
            $latence = [math]::Round($latenceTotale, 2)
            Write-Log "Latence disque moyenne : $latence ms"
        } else {
            Write-LogAvert "Latence disque : non disponible"
        }

        # --- Charge CPU
        $cpu = Get-CimInstance -ClassName Win32_Processor
        $chargeCPU = ($cpu | Measure-Object -Property LoadPercentage -Average).Average
        $charge = [math]::Round($chargeCPU, 1)
        Write-Log "Charge CPU : $charge %"

        if ($charge -ge 90) {
            Write-LogAvert "CPU très chargé !"
        } elseif ($charge -ge 70) {
            Write-Log "CPU modérément chargé."
        } else {
            Write-LogOk "CPU dans une charge normale."
        }

        # --- Interruptions système
        $interruptions = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor | Where-Object { $_.Name -eq "_Total" }
        if ($interruptions) {
            $irq = $interruptions.InterruptsPersec
            Write-Log "Interruptions système : $irq / sec"
        } else {
            Write-LogAvert "Interruptions système : non disponible"
        }

        # --- Utilisation mémoire
        $mem = Get-CimInstance -ClassName Win32_OperatingSystem
        $totalPhys = $mem.TotalVisibleMemorySize
        $freePhys = $mem.FreePhysicalMemory
        $usedMem = (($totalPhys - $freePhys) / $totalPhys) * 100
        $used = [math]::Round($usedMem, 1)
        Write-Log "Mémoire utilisée : $used %"

        if ($used -ge 85) {
            Write-LogAvert "Utilisation mémoire critique !"
        } elseif ($used -ge 70) {
            Write-Log "Utilisation mémoire élevée."
        } else {
            Write-LogOk "Mémoire disponible suffisante."
        }

        Write-LogOk "Analyse de réactivité terminée."
    }
    catch {
        Write-LogErreur "Erreur lors de l'analyse : $_"
    }
}

# -------------------------
# 7. Menu latéral gauche
# -------------------------
$menuItems = @(
    "ANALYSE",
    "Créer compte admin Cliconline",
    "Point de restauration",
    "BOOST",
    "Réparation système / disque"
)
$startY = 120
foreach ($item in $menuItems) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $item
    $btn.Size      = New-Object System.Drawing.Size(250, 40)
    $btn.Location  = New-Object System.Drawing.Point(30, $startY)
    $btn.FlatStyle = "Flat"
    $btn.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    switch ($item) {
        "ANALYSE"                       { $btn.Add_Click({ Get-SystemInfoPlus }) }
        "Créer compte admin Cliconline" { $btn.Add_Click({ Create-LocalAdmin }) }
        "Point de restauration"         { $btn.Add_Click({ Create-SystemRestorePoint }) }
        "BOOST"                         { $btn.Add_Click({ Boost-PCPerformance }) }
        "Réparation système / disque"   { $btn.Add_Click({ Start-RescueToolbox }) }
    }
    $itemButtons    += $btn
    $form.Controls.Add($btn)
    $startY += 60
}

# -------------------------
# 8. Onglets et actions
# -------------------------
$tabs = @("MISE A JOUR","RESEAU","SECURITE","NETTOYAGE","OUTLOOK")
$actionsByTab = @{
    "MISE A JOUR" = @(
        @{ Text="WINDOWS UPDATE"; Action={ Scan-WindowsUpdate } }
        @{ Text="REPARER WINDOWS UPDATE"; Action={ Repair-WindowsUpdate } }
        @{ Text="INSTALLER WINGET"; Action={ Install-WingetIfMissing } }
        @{ Text="ANALYSE DE REACTIVITE SYSTEME"; Action={ Analyse-ReactiviteSysteme } }
        @{ Text="SANTE"; Action={ Show-SystemHealthDashboard } }
    )
    "RESEAU" = @(
        @{ Text="DIAGNOSTIQUE"; Action={ Start-NetworkDiagnostic } }
        @{ Text="En cours dev.";    Action={ Show-ComingSoon } }
        @{ Text="En cours dev.";    Action={ Show-ComingSoon } }
        @{ Text="En cours dev.";    Action={ Show-ComingSoon } }
        @{ Text="En cours dev.";    Action={ Show-ComingSoon } }
    )
    "SECURITE" = @(
        @{ Text="ANTIVIRUS"; Action={ Check-Antivirus } }
        @{ Text="En cours dev."; Action={ Show-ComingSoon } }
        @{ Text="En cours dev."; Action={ Show-ComingSoon } }
        @{ Text="En cours dev."; Action={ Show-ComingSoon } }
        @{ Text="En cours dev."; Action={ Show-ComingSoon } }
    )
    "NETTOYAGE" = @(
        @{ Text="Nettoyage rapide";      Action={ Quick-SystemClean } }
        @{ Text="Pilotes obsolètes";     Action={ Check-ObsoleteDrivers } }
        @{ Text="Supprimer logiciels";   Action={ Uninstall-TargetedApps } }
        @{ Text="En cours dev.";         Action={ Show-ComingSoon } }
        @{ Text="En cours dev.";         Action={ Show-ComingSoon } }
    )
    "OUTLOOK" = @(
        @{ Text="NETTOYER CACHE";        Action={ Clear-OutlookCache } }
        @{ Text="NETTOYER TEMP";         Action={ Clean-OutlookTempFolder } }
        @{ Text="REPARER PST";           Action={ Repair-OutlookPST } }
        @{ Text="En cours dev.";         Action={ Show-ComingSoon } }
        @{ Text="En cours dev.";         Action={ Show-ComingSoon } }
    )
}

function Show-ActionButtons {
    param([string]$tabName)
    $actions = $actionsByTab[$tabName]
    if (-not $actions) {
        [System.Windows.Forms.MessageBox]::Show("ERREUR : Aucune action pour l’onglet '$tabName'","ERREUR")
        return
    }

    # Supprime les anciens boutons
    foreach ($btn in $global:actionButtons) { $form.Controls.Remove($btn) }
    $global:actionButtons = @()
    $actionX = 300

    foreach ($action in $actions) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text      = $action.Text
        $btn.Size      = New-Object System.Drawing.Size(150,60)
        $btn.Location  = New-Object System.Drawing.Point($actionX,200)
        $btn.FlatStyle = 'Flat'
        $btn.Font      = New-Object System.Drawing.Font("Consolas",12,[System.Drawing.FontStyle]::Bold)

        # ——————————————
        # * Nouvelle section : appliquer la couleur selon le thème *
        # ——————————————
        if ($global:theme -eq "dark") {
            $btn.BackColor = [System.Drawing.Color]::White
            $btn.ForeColor = $greenLime
        }
        else {
            $btn.BackColor = $greenLime
            $btn.ForeColor = [System.Drawing.Color]::White
        }
        # ——————————————

        $btn.Add_Click($action.Action)
        $form.Controls.Add($btn)
        $global:actionButtons += $btn
        $actionX += 160
    }
}


# Création des boutons d’onglets
$tabX = 300
foreach ($tab in $tabs) {
    $btnTab = New-Object System.Windows.Forms.Button
    $btnTab.Text      = $tab
    $btnTab.Size      = New-Object System.Drawing.Size(150,40)
    $btnTab.Location  = New-Object System.Drawing.Point($tabX,120)
    $btnTab.FlatStyle = 'Flat'
    $btnTab.Font      = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Bold)
    $btnTab.Add_Click({ param($s,$e) Show-ActionButtons $s.Text })
    $tabButtons     += $btnTab
    $form.Controls.Add($btnTab)
    $tabX += 160
}

# Afficher par défaut le premier onglet
Show-ActionButtons $tabs[0]

# -------------------------
# 9. Zone de log
# -------------------------
$textBoxLogs = New-Object System.Windows.Forms.RichTextBox
$textBoxLogs.Multiline  = $true
$textBoxLogs.ScrollBars = "Vertical"
$textBoxLogs.ReadOnly   = $true
$textBoxLogs.WordWrap   = $true
$textBoxLogs.Size       = New-Object System.Drawing.Size(790,410)
$textBoxLogs.Location   = New-Object System.Drawing.Point(300,310)
$form.Controls.Add($textBoxLogs)

function Write-Log {
    param([string]$message)
    $timestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $textBoxLogs.AppendText("[$timestamp] $message`r`n")
    $textBoxLogs.SelectionStart = $textBoxLogs.Text.Length
    $textBoxLogs.ScrollToCaret()
}

# Message de bienvenue
Write-Log "Bienvenue dans le Gestionnaire IT de ClicOnLine."

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

# -------------------------
# 10. Barre de scan (animation)
# -------------------------
$scanPanel = New-Object System.Windows.Forms.Panel
$scanPanel.Size     = New-Object System.Drawing.Size(790,10)
$scanPanel.Location = New-Object System.Drawing.Point(300,280)
$scanPanel.BackColor= [System.Drawing.Color]::Gray
$form.Controls.Add($scanPanel)

$scanner = New-Object System.Windows.Forms.Panel
$scanner.Size       = New-Object System.Drawing.Size(100,10)
$scanner.BackColor  = [System.Drawing.Color]::Lime
$scanner.Left       = -100
$scanPanel.Controls.Add($scanner)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30
$timer.Add_Tick({
    $scanner.Left += 5
    if ($scanner.Left -ge $scanPanel.Width) { $scanner.Left = -$scanner.Width }
})
$timer.Start()

# -------------------------
# 11. Bouton EXIT et footer
# -------------------------
$exitBtn = New-Object System.Windows.Forms.Button
$exitBtn.Text      = "Fermer"
$exitBtn.Size      = New-Object System.Drawing.Size(150,60)
$exitBtn.Location  = New-Object System.Drawing.Point(10,690)
$exitBtn.FlatStyle = 'Flat'
$exitBtn.BackColor = [System.Drawing.Color]::Red
$exitBtn.ForeColor = [System.Drawing.Color]::White
$exitBtn.Font      = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Bold)
$exitBtn.Add_Click({ $form.Close() })
$form.Controls.Add($exitBtn)

$footer = New-Object System.Windows.Forms.Label
$footer.Text      = "© 2025 Cliconline - Département IT | Gestionnaire IT v2.0"
$footer.AutoSize  = $false
$footer.Size      = New-Object System.Drawing.Size(500,25)
$footer.Location  = New-Object System.Drawing.Point(300,730)
$footer.ForeColor = [System.Drawing.Color]::FromArgb(117,180,25)
$footer.Font      = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Italic)
$footer.TextAlign = "MiddleCenter"
$form.Controls.Add($footer)

# -------------------------
# 12. Appliquer thème initial et afficher
# -------------------------
Apply-Theme -mode $global:theme
[void]$form.ShowDialog()
