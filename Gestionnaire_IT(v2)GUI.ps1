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
$titleLabel.Text = "CLICONLINE"
$titleLabel.Size = New-Object System.Drawing.Size(1000, 40)
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.ForeColor = [System.Drawing.Color]::LimeGreen
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

        # Fen�tre de confirmation style LCARS
        $formUpdates = New-Object System.Windows.Forms.Form
        $formUpdates.Text = "Mises � jour d�tect�es"
        $formUpdates.Size = New-Object System.Drawing.Size(600, 400)
        $formUpdates.StartPosition = "CenterScreen"
        $formUpdates.BackColor = [System.Drawing.Color]::Black
        $formUpdates.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Les mises � jour suivantes sont disponibles :"
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
                Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-LogOk "Dossier TEMP nettoy�."
            }
            "Vider la corbeille" {
                (New-Object -ComObject Shell.Application).NameSpace(0x0a).Items() | ForEach-Object {
                    try { Remove-Item $_.Path -Recurse -Force -ErrorAction SilentlyContinue } catch {}
                }
                Write-LogOk "Corbeille vid�e."
            }
            "Arreter OneDrive" {
                Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
                Write-LogOk "OneDrive arr�t�."
            }
            "Augmenter la RAM virtuelle (pagefile)" {
                try {
                    $totalRAM_MB = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB
                    $pagefileSize = [math]::Round($totalRAM_MB * 1.5)
                    Write-Log "RAM install�e : $([math]::Round($totalRAM_MB)) Mo => Pagefile : $pagefileSize Mo"

                    wmic computersystem where name="%computername%" set AutomaticManagedPagefile=False | Out-Null
                    wmic pagefileset where name="C:\\pagefile.sys" set InitialSize=$pagefileSize,MaximumSize=$pagefileSize | Out-Null

                    Write-LogOk "RAM virtuelle configur�e � $pagefileSize Mo"
                } catch {
                    Write-LogError "Erreur configuration RAM virtuelle : $_"
                }
            }
            "Gerer les apps au demarrage" {
                try {
                    $startupApps = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command
                    $formStartup = New-Object System.Windows.Forms.Form
                    $formStartup.Text = "Applications au d�marrage"
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
                    $btnDisable.Text = "D�sactiver les coch�s"
                    $btnDisable.Size = New-Object System.Drawing.Size(200, 40)
                    $btnDisable.Location = New-Object System.Drawing.Point(200, 360)
                    $btnDisable.BackColor = [System.Drawing.Color]::DarkOrange
                    $btnDisable.ForeColor = [System.Drawing.Color]::Black
                    $btnDisable.FlatStyle = 'Flat'
                    $btnDisable.Add_Click({
                        foreach ($entry in $listView.Items) {
                            if (-not $entry.Checked) {
                                Write-LogAvert "Application d�sactiv�e (simulation) : $($entry.Text)"
                            }
                        }
                        $formStartup.Close()
                    })
                    $formStartup.Controls.Add($btnDisable)
                    $formStartup.ShowDialog()
                } catch {
                    Write-LogError "Erreur lecture apps d�marrage : $_"
                }
            }
        }
    }

    Write-LogOk "Optimisation termin�e."
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
    Write-Log "Scan des pilotes obsol�tes..."
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
                Write-LogAvert "$($driver.DeviceName) - $driverDate - $($driver.DriverVersion) - POTENTIELLEMENT OBSOL�TE"
            }
        }
    }

    if ($obsolete.Count -eq 0) {
        Write-LogOk "Aucun pilote obsol�te d�tect� ou d�sactivable en toute s�curit�."
        return
    }

    Write-LogAvert "Nombre total de pilotes potentiellement obsol�tes : $($obsolete.Count)"

    # === Fen�tre LCARS de confirmation ===
    $formConfirm = New-Object System.Windows.Forms.Form
    $formConfirm.Text = "Suppression des pilotes obsol�tes"
    $formConfirm.Size = New-Object System.Drawing.Size(600, 250)
    $formConfirm.StartPosition = "CenterScreen"
    $formConfirm.BackColor = [System.Drawing.Color]::Black
    $formConfirm.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Il y a $($obsolete.Count) pilotes potentiellement obsol�tes. Souhaitez-vous les d�sinstaller ?"
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
                Write-LogOk "Pilote supprim� : $infName"
            } catch {
                Write-LogError "Erreur suppression pilote $($d.DeviceName) : $_"
            }
        }
    } else {
        Write-Log "Suppression des pilotes annul�e par l'utilisateur."
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
    # D�tection des antivirus install�s
    $antivirus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if (-not $antivirus) {
        Write-LogAvert "Aucun antivirus d�tect�."
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
    $form.Text = "D�sinstallation des antivirus"
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
    $btnOK.Text = "D�sinstaller"
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
        Write-Log "D�sinstallation annul�e par l'utilisateur."
        return
    }

    $selected = $list.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucun antivirus s�lectionn� pour suppression."
        return
    }

    foreach ($nom in $selected) {
        $av = $avList | Where-Object { $_.Nom -eq $nom } | Select-Object -First 1
        if ($av.UninstallString) {
            try {
                Write-Log "D�sinstallation de $($av.Nom)..."
                $cmd = $av.UninstallString
                if ($cmd -match '^\"(.+?)\"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(' ')[0]
                    $args = $cmd.Substring($exe.Length)
                }
                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($av.Nom) d�sinstall� avec succ�s."
            } catch {
                Write-LogError "Erreur lors de la suppression de $($av.Nom) : $_"
            }
        } elseif ($av.Nom -like "*Norton*") {
            Write-LogAvert "Norton d�tect� sans commande de d�sinstallation. T�l�chargement de l'outil..."
            Force-Uninstall-Norton
        } else {
            Write-LogAvert "Aucune commande de d�sinstallation trouv�e pour $($av.Nom)"
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
        Write-LogAvert "Aucune application trouv�e pour d�sinstallation."
        return
    }

    # Interface LCARS styl�e
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "S�lectionnez les applications � d�sinstaller"
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
    $btnOK.Text = "D�sinstaller"
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
        Write-Log "D�sinstallation annul�e par l'utilisateur."
        return
    }

    $selected = $list.CheckedItems
    if ($selected.Count -eq 0) {
        Write-LogAvert "Aucune application s�lectionn�e pour suppression."
        return
    }

    foreach ($name in $selected) {
        $app = $foundApps | Where-Object { $_.DisplayName -eq $name } | Select-Object -First 1
        if ($app -and $app.UninstallString) {
            try {
                $cmd = $app.UninstallString
                Write-Log "D�sinstallation de $($app.DisplayName)..."

                if ($cmd -match '^"(.+?)"') {
                    $exe = $matches[1]
                    $args = $cmd.Substring($exe.Length + 2)
                } else {
                    $exe = $cmd.Split(" ")[0]
                    $args = $cmd.Substring($exe.Length)
                }

                Start-Process -FilePath $exe -ArgumentList $args -NoNewWindow -Wait -ErrorAction Stop
                Write-LogOk "$($app.DisplayName) d�sinstall�e avec succ�s."
            } catch {
                Write-LogError "Erreur lors de la d�sinstallation de $($app.DisplayName) : $_"
            }
        } else {
            Write-LogAvert "Pas de commande de d�sinstallation trouv�e pour $name"
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

    # Interface LCARS pour red�marrer les services arr�t�s
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "D�marrer les services arr�t�s"
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
    $btnOK.Text = "D�marrer s�lection"
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
        Write-Log "D�marrage des services annul� par l'utilisateur."
        return
    }

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
            Write-LogAvert "Aucun profil Outlook trouv�."
            return
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Profils Outlook - S�lection pour r�paration"
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
        $btnOK.Text = "R�parer"
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
            Write-Log "R�paration annul�e par l'utilisateur."
            return
        }

        $selected = $checkList.CheckedItems
        if ($selected.Count -eq 0) {
            Write-LogAvert "Aucun profil s�lectionn�."
            return
        }

        Write-LogAvert "Outlook arr�t� pour maintenance."
        Get-Process -Name outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

        $paths = @("$env:APPDATA\Microsoft\Outlook", "$env:LOCALAPPDATA\Microsoft\Outlook")
        foreach ($path in $paths) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -Include *.dat,*.xml -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Log "Nettoyage config dans : $path"
            }
        }

        foreach ($profileText in $selected) {
            Write-LogOk "Profil r�par� : $profileText"
        }

        Write-LogOk "R�paration des profils Outlook termin�e."

    } catch {
        Write-LogError "Erreur durant la d�tection ou la r�paration des profils Outlook : $_"
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
    $form.BackColor = [System.Drawing.Color]::Black
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
    $btnClose.BackColor = [System.Drawing.Color]::Orange
    $btnClose.ForeColor = [System.Drawing.Color]::Black
    $btnClose.FlatStyle = 'Flat'
    $btnClose.Add_Click({
    $timer.Stop()
    $form.Invoke([Action]{
        $form.Close()
        $form.Dispose()
        [System.Windows.Forms.Application]::Exit()
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

        $os = Get-CimInstance Win32_OperatingSystem
        $comp = Get-CimInstance Win32_ComputerSystem
        $cpu = Get-CimInstance Win32_Processor
        $bios = Get-CimInstance Win32_BIOS
        $text = @"
Ordinateur     : $env:COMPUTERNAME
Utilisateur    : $env:USERNAME
OS             : $($os.Caption)
Version        : $($os.Version)
Architecture   : $($os.OSArchitecture)
Fabricant      : $($comp.Manufacturer)
Numéro de série : $($bios.SerialNumber)
Modèle         : $($comp.Model)
CPU            : $($cpu.Name)
Cœurs logiques : $($cpu.NumberOfLogicalProcessors)
Température CPU: $tempVal °C



Certaines informations demandent l'accès administrateur
"@
        $textBox.Text = $text
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
        Write-LogOk "Compte '$username' masque de l�ecran de connexion."
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

# Menu latéral gauche
$menuItems = @( "ANALYSE", "Créer compte admin Cliconline", "Point de restauration", "BOOST", "Réparation fichiers système / disque")
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
        "ANALYSE" { $btn.Add_Click({ Get-SystemInfoPlus})}
        "Créer compte admin Cliconline" { $btn.Add_Click({ Create-LocalAdmin})}
        "Point de restauration" { $btn.Add_Click({ Create-SystemRestorePoint })}
        "BOOST" { $btn.Add_Click({ Boost-PCPerformance })}
        "Réparation fichiers système / disque" { $btn.Add_Click({ Start-RescueToolbox })}
    }

    $form.Controls.Add($btn)
    $startY += 60
}

$tabs = @("MISE A JOUR", "RESEAU", "SECURITE", "NETTOYAGE", "OUTLOOK")
$actionsByTab = @{
    "MISE A JOUR" = @(
        @{ Text="WINDOWS UPDATE"; Color="DeepSkyBlue"; Action={ Scan-WindowsUpdate } }
        @{ Text="REPARER WINDOWS UPDATE"; Color="Purple"; Action={ Repair-WindowsUpdate } }
        @{ Text="INSTALLER WINGET"; Color="LimeGreen"; Action={ Install-WingetIfMissing } }
        @{ Text="En cours de développement"; Color="Gold"; Action={ Show-ComingSoon } }
        @{ Text="REDEMARRER"; Color="MediumTurquoise"; Action={ Restart-PC } }
    )
    "RESEAU" = @(
        @{ Text="DIAGNOSTIQUE"; Color="MediumTurquoise"; Action={ Start-NetworkDiagnostic } }
        @{ Text="En cours de développement"; Color="DeepSkyBlue"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="Purple"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="LimeGreen"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="Gold"; Action={ Show-ComingSoon } }
    )
    "SECURITE" = @(
        @{ Text="ANTIVIRUS"; Color="Gold"; Action={ Check-Antivirus} }
        @{ Text="En cours de développement"; Color="MediumTurquoise"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="DeepSkyBlue"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="Purple"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="LimeGreen"; Action={ Show-ComingSoon } }
    )
    "NETTOYAGE" = @(
        @{ Text="Nettoyage rapide"; Color="LimeGreen"; Action={ Quick-SystemClean} }
        @{ Text="Nettoyage des pilotes obsoletes"; Color="Gold"; Action={ Check-ObsoleteDrivers } }
        @{ Text="Supprimer logiciels"; Color="MediumTurquoise"; Action={ Uninstall-TargetedApps } }
        @{ Text="En cours de développement"; Color="DeepSkyBlue"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="Purple"; Action={ Show-ComingSoon } }
    )
    "OUTLOOK" = @(
        @{ Text="NETTOYER CACHE"; Color="Purple"; Action={ Clear-OutlookCache } }
        @{ Text="NETTOYER FICHIER TEMP"; Color="LimeGreen"; Action={ Clean-OutlookTempFolder } }
        @{ Text="REPARER FICHIER PST"; Color="Gold"; Action={ Repair-OutlookPST } }
        @{ Text="En cours de développement"; Color="MediumTurquoise"; Action={ Show-ComingSoon } }
        @{ Text="En cours de développement"; Color="DeepSkyBlue"; Action={ Show-ComingSoon } }
    )
}

$global:actionButtons = @()

function Show-ActionButtons {
    param($tabName)
    $actions = $actionsByTab[$tabName]
    if ($null -eq $actions) {
        [System.Windows.Forms.MessageBox]::Show("ERREUR : Aucune action trouvée pour l’onglet '$tabName' !","ERREUR")
        return
    }
    foreach ($btn in $global:actionButtons) { $form.Controls.Remove($btn) }
    $global:actionButtons = @()
    $actionX = 180
    foreach ($action in $actions) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $action.Text
        $btn.Size = New-Object System.Drawing.Size(150, 60)
        $btn.Location = New-Object System.Drawing.Point($actionX, 200)
        $btn.BackColor = [System.Drawing.Color]::$($action.Color)
        $btn.ForeColor = [System.Drawing.Color]::Black
        $btn.FlatStyle = 'Flat'
        $btn.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
        $btn.Add_Click($action.Action)
        $form.Controls.Add($btn)
        $global:actionButtons += $btn
        $actionX += 160
    }
}

$tabX = 180
foreach ($tab in $tabs) {
    $btnTab = New-Object System.Windows.Forms.Button
    $btnTab.Text = $tab
    $btnTab.Size = New-Object System.Drawing.Size(150, 40)
    $btnTab.Location = New-Object System.Drawing.Point($tabX, 80)
    $btnTab.BackColor = [System.Drawing.Color]::DarkSlateBlue
    $btnTab.ForeColor = [System.Drawing.Color]::White
    $btnTab.FlatStyle = 'Flat'
    $btnTab.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    # --- SOLUTION : récupérer le texte du bouton cliqué ---
    $btnTab.Add_Click( {
        param($sender, $event)
        $tabClicked = $sender.Text
        Show-ActionButtons $tabClicked
    })
    $form.Controls.Add($btnTab)
    $tabX += 160
}

Show-ActionButtons $tabs[0]

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


# Footer officiel
$footer = New-Object System.Windows.Forms.Label
$footer.Text = "� 2025 Cliconline - Département IT | Gestionnaire IT v2.0"
$footer.AutoSize = $false
$footer.Size = New-Object System.Drawing.Size(1000, 25)
$footer.Location = New-Object System.Drawing.Point(10, 710)
$footer.ForeColor = [System.Drawing.Color]::Gray
$footer.BackColor = [System.Drawing.Color]::Black
$footer.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Italic)
$footer.TextAlign = "MiddleCenter"
$form.Controls.Add($footer)

[void]$form.ShowDialog()
