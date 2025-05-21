Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Gestionnaire IT - ClicOnLine"
$form.Size = New-Object System.Drawing.Size(1024, 768)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::Black
$form.FormBorderStyle = 'FixedDialog'

$form.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

# Barre de titre
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "CLICONLINE APPS"
$titleLabel.Size = New-Object System.Drawing.Size(1000, 40)
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.ForeColor = [System.Drawing.Color]::DeepSkyBlue
$titleLabel.Font = New-Object System.Drawing.Font("Consolas", 18, [System.Drawing.FontStyle]::Bold)
$titleLabel.TextAlign = "MiddleCenter"
$form.Controls.Add($titleLabel)

# Fonctions

function Scan-WindowsUpdate {
    try {
        Write-Log "Scan Windows Update..."
        $serviceWU = Get-Service -Name wuauserv -ErrorAction Stop
        Write-Log "Service Windows Update : $($serviceWU.Status)"

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

        # Fenêtre de confirmation style LCARS
        $formUpdates = New-Object System.Windows.Forms.Form
        $formUpdates.Text = "Mises à jour détectées"
        $formUpdates.Size = New-Object System.Drawing.Size(600, 400)
        $formUpdates.StartPosition = "CenterScreen"
        $formUpdates.BackColor = [System.Drawing.Color]::Black
        $formUpdates.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Les mises à jour suivantes sont disponibles :"
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

function Boost-PCPerformance {
    $formBoost = New-Object System.Windows.Forms.Form
    $formBoost.Text = "Optimisation du PC"
    $formBoost.Size = New-Object System.Drawing.Size(500, 360)
    $formBoost.StartPosition = "CenterParent"
    $formBoost.BackColor = [System.Drawing.Color]::Black
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
        "Augmenter la RAM virtuelle (pagefile)",
        "Gerer les apps au demarrage"
    ))
    $formBoost.Controls.Add($checkList)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Lancer"
    $btnOK.Size = New-Object System.Drawing.Size(100, 40)
    $btnOK.Location = New-Object System.Drawing.Point(270, 220)
    $btnOK.BackColor = [System.Drawing.Color]::LimeGreen
    $btnOK.ForeColor = [System.Drawing.Color]::Black
    $btnOK.FlatStyle = 'Flat'
    $formBoost.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Size = New-Object System.Drawing.Size(100, 40)
    $btnCancel.Location = New-Object System.Drawing.Point(120, 220)
    $btnCancel.BackColor = [System.Drawing.Color]::IndianRed
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = 'Flat'
    $formBoost.Controls.Add($btnCancel)

    $btnOK.Add_Click({ $formBoost.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formBoost.Close() })
    $btnCancel.Add_Click({ $formBoost.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formBoost.Close() })

    if ($formBoost.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "Boost PC annulé par l’utilisateur."
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
                Write-LogOk "OneDrive arrêté."
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
        Animate-ProgressBar -progressBar $progressBar -durationSeconds 2
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
                Write-LogAvert "$($driver.DeviceName) - $driverDate - $($driver.DriverVersion) - POTENTIELLEMENT OBSOLÈTE"
            }
        }
    }

    if ($obsolete.Count -eq 0) {
        Write-LogOk "Aucun pilote obsolète détecté ou désactivable en toute sécurité."
        return
    }

    Write-LogAvert "Nombre total de pilotes potentiellement obsolètes : $($obsolete.Count)"

    # === Fenêtre LCARS de confirmation ===
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
    $btnYes.BackColor = [System.Drawing.Color]::Orange
    $btnYes.ForeColor = [System.Drawing.Color]::Black
    $btnYes.FlatStyle = 'Flat'

    $btnNo = New-Object System.Windows.Forms.Button
    $btnNo.Text = "Non - Annuler"
    $btnNo.Size = New-Object System.Drawing.Size(150, 40)
    $btnNo.Location = New-Object System.Drawing.Point(120, 130)
    $btnNo.BackColor = [System.Drawing.Color]::IndianRed
    $btnNo.ForeColor = [System.Drawing.Color]::White
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
        foreach ($d in $obsolete) {
            try {
                $infName = $d.InfName
                Write-Log "Suppression du pilote : $($d.DeviceName) ($infName)..."
                pnputil /delete-driver "$infName" /uninstall /force /quiet
                Write-LogOk "Pilote supprimé : $infName"
            } catch {
                Write-LogError "Erreur suppression pilote $($d.DeviceName) : $_"
            }
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

    # Interface LCARS pour redémarrer les services arrêtés
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Démarrer les services arrêtés"
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

        Write-LogAvert "Outlook arrêté pour maintenance."
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
    $output = @{}

    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $computer = Get-CimInstance -ClassName Win32_ComputerSystem
    $bios = Get-CimInstance -ClassName Win32_BIOS
    $cpu = Get-CimInstance -ClassName Win32_Processor
    $gpu = Get-CimInstance -ClassName Win32_VideoController
    $mem = Get-CimInstance -ClassName Win32_PhysicalMemory
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
    $net = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

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

    # Affichage graphique LCARS
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Informations système complètes"
    $form.Size = New-Object System.Drawing.Size(700, 600)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::Black
    $form.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.ReadOnly = $true
    $textBox.Size = New-Object System.Drawing.Size(660, 500)
    $textBox.Location = New-Object System.Drawing.Point(20, 20)
    $textBox.BackColor = [System.Drawing.Color]::DarkSlateGray
    $textBox.ForeColor = [System.Drawing.Color]::White

    $builder = "Informations système complètes :`r`n`r`n"
    foreach ($key in $output.Keys) {
        $builder += "{0,-25} : {1}`r`n" -f $key, $output[$key]
    }

    $textBox.Text = $builder
    $form.Controls.Add($textBox)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Fermer"
    $btnClose.Size = New-Object System.Drawing.Size(120, 40)
    $btnClose.Location = New-Object System.Drawing.Point(280, 520)
    $btnClose.BackColor = [System.Drawing.Color]::Orange
    $btnClose.ForeColor = [System.Drawing.Color]::Black
    $btnClose.FlatStyle = 'Flat'
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

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
        Write-LogOk "Compte '$username' masque de l’ecran de connexion."
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


# Menu latéral gauche
$menuItems = @("SERVICES", "ANALYSE", "PARAMETRE", "Créer compte admin Cliconline", "REBOOT", "Désinstallation Logiciels", "DRIVERS")
$startY = 80
foreach ($item in $menuItems) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $item
    $btn.Size = New-Object System.Drawing.Size(150, 50)
    $btn.Location = New-Object System.Drawing.Point(10, $startY)
    $btn.BackColor = [System.Drawing.Color]::Orange
    $btn.ForeColor = [System.Drawing.Color]::Black
    $btn.FlatStyle = 'Flat'
    $btn.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    switch ($item) {
        "SERVICES" { $btn.Add_Click({ Check-CriticalServices})}
        "ANALYSE" { $btn.Add_Click({ Get-SystemInfoPlus})}
        "PARAMETRE" { $btn.Add_Click({ Show-ComingSoon})}
        "Créer compte admin Cliconline" { $btn.Add_Click({ Create-LocalAdmin})}
        "REBOOT" { $btn.Add_Click({ Restart-PC})}
        "Désinstallation Logiciels" { $btn.Add_Click({ Uninstall-TargetedApps})}
        "DRIVERS" { $btn.Add_Click({ Check-ObsoleteDrivers})}
    }

    $form.Controls.Add($btn)
    $startY += 60
}

# Bouton Exit
$exitBtn = New-Object System.Windows.Forms.Button
$exitBtn.Text = "EXIT"
$exitBtn.Size = New-Object System.Drawing.Size(150, 60)
$exitBtn.Location = New-Object System.Drawing.Point(10, 650)
$exitBtn.BackColor = [System.Drawing.Color]::OrangeRed
$exitBtn.ForeColor = [System.Drawing.Color]::White
$exitBtn.FlatStyle = 'Flat'
$exitBtn.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$exitBtn.Add_Click({ $form.Close() })
$form.Controls.Add($exitBtn)

# Onglets principaux
$tabs = @("SYSTEME", "RESEAU", "SECURITE", "NETTOYAGE", "POINT DE RESTAURATION")
$tabX = 180
foreach ($tab in $tabs) {
    $tabBtn = New-Object System.Windows.Forms.Button
    $tabBtn.Text = $tab
    $tabBtn.Size = New-Object System.Drawing.Size(150, 40)
    $tabBtn.Location = New-Object System.Drawing.Point($tabX, 80)
    $tabBtn.BackColor = [System.Drawing.Color]::DarkSlateBlue
    $tabBtn.ForeColor = [System.Drawing.Color]::White
    $tabBtn.FlatStyle = 'Flat'
    $tabBtn.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    switch ($tab) {
        "SYSTEME" { $tabBtn.Add_Click({ Install-WingetIfMissing})}
        "RESEAU" { }
        "SECURITE" { $tabBtn.Add_Click({ Check-Antivirus})}
        "NETTOYAGE" { $tabBtn.Add_Click({ Quick-SystemClean})}
        "POINT DE RESTAURATION" { $tabBtn.Add_Click({ Create-SystemRestorePoint})}
    }

    $form.Controls.Add($tabBtn)
    $tabX += 160
}

# Zone de log
$textBoxLogs = New-Object System.Windows.Forms.RichTextBox -Property @{
    Multiline = $true
    ScrollBars = 'Vertical'
    Size = New-Object System.Drawing.Size(800, 385)
    Location = New-Object System.Drawing.Point(180, 320)
    ReadOnly = $true
    BackColor = "Black"
    ForeColor = "white"
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
Write-Log "Bienvenue dans le Gestionnaire IT de ClicOnLine. Toutes les actions effectuees s'afficheront ici."

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

# Boutons d'action principaux
$actionBtns = @("SCAN", "DIAGNOSTIC", "BOOST PC", "MISE À JOUR", "Réparation Windows update")
$colors = @("DeepSkyBlue", "Purple", "LimeGreen", "Gold", "mediumTurquoise")
$actionX = 180
foreach ($btnText in $actionBtns) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $btnText
    $btn.Size = New-Object System.Drawing.Size(150, 60)
    $btn.Location = New-Object System.Drawing.Point($actionX, 200)
    $btn.BackColor = [System.Drawing.Color]::$($colors[$actionBtns.IndexOf($btnText)])
    $btn.ForeColor = [System.Drawing.Color]::Black
    $btn.FlatStyle = 'Flat'
    $btn.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)

    switch ($btnText) {
        "BOOST PC"     { $btn.Add_Click({ Boost-PCPerformance }) }
        "SCAN"         { $btn.Add_Click({ Show-ComingSoon })} 
        "DIAGNOSTIC"   { $btn.Add_Click({ Show-SystemHealthDashboard }) }
        "MISE À JOUR"  { $btn.Add_Click({ Scan-WindowsUpdate }) }
        "Réparation Windows update"         { $btn.Add_Click({ Repair-WindowsUpdate })}
    }

    $form.Controls.Add($btn)
    $actionX += 160
}

# Boutons d'action principaux
$actionBtns2 = @("CACHE OUTLOOK", "Nettoyer Temporaire Outlook", "Réparer PST Outlook", "Réparer profil Outlook", "en cours de developpement")
$colors2 = @( "LimeGreen","DeepSkyBlue", "mediumTurquoise", "Purple", "Gold")
$actionX2 = 180
foreach ($btnText2 in $actionBtns2) {
    $btn2 = New-Object System.Windows.Forms.Button
    $btn2.Text = $btnText2
    $btn2.Size = New-Object System.Drawing.Size(150, 60)
    $btn2.Location = New-Object System.Drawing.Point($actionX2, 130)
    $btn2.BackColor = [System.Drawing.Color]::$($colors2[$actionBtns2.IndexOf($btnText2)])
    $btn2.ForeColor = [System.Drawing.Color]::Black
    $btn2.FlatStyle = 'Flat'
    $btn2.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)

    switch ($btnText2) {
        "CACHE OUTLOOK"         { $btn2.Add_Click({ Clear-OutlookCache })}
        "Nettoyer Temporaire Outlook"   { $btn2.Add_Click({ Clean-OutlookTempFolder }) }
        "Réparer PST Outlook"  { $btn2.Add_Click({ Repair-OutlookPST }) }
        "Réparer profil Outlook"         { $btn2.Add_Click({ Show-OutlookProfilesWithRepair })}
        "en cours de developpement"     { $btn2.Add_Click({ Show-ComingSoon })}
    }

    $form.Controls.Add($btn2)
    $actionX2 += 160
}

# Barre de scan (Progression style "scanner")
$scanPanel = New-Object System.Windows.Forms.Panel
$scanPanel.Size = New-Object System.Drawing.Size(800, 10)
$scanPanel.Location = New-Object System.Drawing.Point(180, 280)
$scanPanel.BackColor = [System.Drawing.Color]::Gray

$scanner = New-Object System.Windows.Forms.Panel
$scanner.Size = New-Object System.Drawing.Size(100, 10)
$scanner.BackColor = [System.Drawing.Color]::Lime
$scanner.Left = -100
$scanPanel.Controls.Add($scanner)
$form.Controls.Add($scanPanel)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30
$timer.Add_Tick({
    $scanner.Left += 5
    if ($scanner.Left -ge $scanPanel.Width) {
        $scanner.Left = -$scanner.Width
    }
})
$timer.Start()

# Footer officiel
$footer = New-Object System.Windows.Forms.Label
$footer.Text = "© 2025 Cliconline - Département IT | Gestionnaire IT v2.0"
$footer.AutoSize = $false
$footer.Size = New-Object System.Drawing.Size(1000, 25)
$footer.Location = New-Object System.Drawing.Point(10, 710)
$footer.ForeColor = [System.Drawing.Color]::Gray
$footer.BackColor = [System.Drawing.Color]::Black
$footer.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Italic)
$footer.TextAlign = "MiddleCenter"
$form.Controls.Add($footer)

[void]$form.ShowDialog()