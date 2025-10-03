Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Software Packager"
$form.Size = New-Object System.Drawing.Size(850,650)  # Adjusted height
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# Set fonts and colors
$defaultFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
$tabFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 12)
$blueFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 12)
$softwaress = New-Object System.Drawing.Font("Microsoft Sans Serif", 13)
$blueColor = [System.Drawing.Color]::FromArgb(0, 102, 204)

# Global variable to track process
$global:operationRunning = $false
$global:shouldStop = $false

# Tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10,10)
$tabControl.Size = New-Object System.Drawing.Size(820,500)  # Adjusted height
$tabControl.Font = $tabFont
$form.Controls.Add($tabControl)

## INSTALLATION TAB ##
$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = "Installation"
$tabInstall.Font = $tabFont
$tabControl.Controls.Add($tabInstall)

# Source folder selection
$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Location = New-Object System.Drawing.Point(20,20)
$labelSource.Size = New-Object System.Drawing.Size(300,25)
$labelSource.Text = "Software Source Folder:"
$labelSource.Font = $defaultFont
$tabInstall.Controls.Add($labelSource)

$textboxSource = New-Object System.Windows.Forms.TextBox
$textboxSource.Location = New-Object System.Drawing.Point(20,50)
$textboxSource.Size = New-Object System.Drawing.Size(650,30)
$textboxSource.Font = $defaultFont
$tabInstall.Controls.Add($textboxSource)

$buttonBrowseFolder = New-Object System.Windows.Forms.Button
$buttonBrowseFolder.Location = New-Object System.Drawing.Point(680,48)
$buttonBrowseFolder.Size = New-Object System.Drawing.Size(120,27)
$buttonBrowseFolder.Text = "Browse..."
$buttonBrowseFolder.Font = $defaultFont
$buttonBrowseFolder.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Software Source Folder"
    $folderBrowser.RootFolder = "MyComputer"
    $folderBrowser.ShowNewFolderButton = $false
    $folderBrowser.SelectedPath = [Environment]::GetFolderPath("Desktop")
    if($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textboxSource.Text = $folderBrowser.SelectedPath
    }
})
$tabInstall.Controls.Add($buttonBrowseFolder)

# Target EXE selection
$labelExe = New-Object System.Windows.Forms.Label
$labelExe.Location = New-Object System.Drawing.Point(20,100)
$labelExe.Size = New-Object System.Drawing.Size(300,25)
$labelExe.Text = "Main Application EXE:"
$labelExe.Font = $defaultFont
$tabInstall.Controls.Add($labelExe)

$textboxExe = New-Object System.Windows.Forms.TextBox
$textboxExe.Location = New-Object System.Drawing.Point(20,130)
$textboxExe.Size = New-Object System.Drawing.Size(650,30)
$textboxExe.Font = $defaultFont
$textboxExe.Add_TextChanged({
    if ($textboxExe.Text -ne "") {
        $exeName = [System.IO.Path]::GetFileNameWithoutExtension($textboxExe.Text)
        $textboxShortcutName.Text = $exeName
    }
})
$tabInstall.Controls.Add($textboxExe)

$buttonBrowseExe = New-Object System.Windows.Forms.Button
$buttonBrowseExe.Location = New-Object System.Drawing.Point(680,128)
$buttonBrowseExe.Size = New-Object System.Drawing.Size(120,27)
$buttonBrowseExe.Text = "Browse..."
$buttonBrowseExe.Font = $defaultFont
$buttonBrowseExe.Add_Click({
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $fileBrowser.Filter = "Executable files (*.exe)|*.exe|All files (*.*)|*.*"
    $fileBrowser.Title = "Select Main Application EXE"
    $fileBrowser.InitialDirectory = if ($textboxSource.Text -ne "") { $textboxSource.Text } else { [Environment]::GetFolderPath("Desktop") }
    if($fileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textboxExe.Text = $fileBrowser.FileName
    }
})
$tabInstall.Controls.Add($buttonBrowseExe)

# Shortcut checkboxes
$checkboxStartMenu = New-Object System.Windows.Forms.CheckBox
$checkboxStartMenu.Location = New-Object System.Drawing.Point(40,180)
$checkboxStartMenu.Size = New-Object System.Drawing.Size(300,25)
$checkboxStartMenu.Text = "Create Start Menu Shortcut"
$checkboxStartMenu.Font = $defaultFont
$checkboxStartMenu.Checked = $true
$checkboxStartMenu.Add_CheckedChanged({
    UpdateShortcutNameFieldVisibility
})
$tabInstall.Controls.Add($checkboxStartMenu)

$checkboxDesktop = New-Object System.Windows.Forms.CheckBox
$checkboxDesktop.Location = New-Object System.Drawing.Point(40,210)
$checkboxDesktop.Size = New-Object System.Drawing.Size(300,25)
$checkboxDesktop.Text = "Create Desktop Shortcut"
$checkboxDesktop.Font = $defaultFont
$checkboxDesktop.Checked = $true
$checkboxDesktop.Add_CheckedChanged({
    UpdateShortcutNameFieldVisibility
})
$tabInstall.Controls.Add($checkboxDesktop)

# Function to update shortcut name field visibility
function UpdateShortcutNameFieldVisibility {
    $anyChecked = $checkboxStartMenu.Checked -or $checkboxDesktop.Checked
    $labelShortcutName.Visible = $anyChecked
    $textboxShortcutName.Visible = $anyChecked
}

# Shortcut name
$labelShortcutName = New-Object System.Windows.Forms.Label
$labelShortcutName.Location = New-Object System.Drawing.Point(20,250)
$labelShortcutName.Size = New-Object System.Drawing.Size(300,25)
$labelShortcutName.Text = "Shortcut Name:"
$labelShortcutName.Font = $defaultFont
$tabInstall.Controls.Add($labelShortcutName)

$textboxShortcutName = New-Object System.Windows.Forms.TextBox
$textboxShortcutName.Location = New-Object System.Drawing.Point(20,280)
$textboxShortcutName.Size = New-Object System.Drawing.Size(300,30)
$textboxShortcutName.Text = ""
$textboxShortcutName.Font = $blueFont
$textboxShortcutName.ForeColor = $blueColor
$tabInstall.Controls.Add($textboxShortcutName)

# Installation status label
$labelInstallStatus = New-Object System.Windows.Forms.Label
$labelInstallStatus.Location = New-Object System.Drawing.Point(20,330)
$labelInstallStatus.Size = New-Object System.Drawing.Size(760,30)
$labelInstallStatus.Text = "Ready to install."
$labelInstallStatus.Font = $defaultFont
$tabInstall.Controls.Add($labelInstallStatus)

# Installation progress bar
$progressBarInstall = New-Object System.Windows.Forms.ProgressBar
$progressBarInstall.Location = New-Object System.Drawing.Point(20,370)
$progressBarInstall.Size = New-Object System.Drawing.Size(760,25)
$progressBarInstall.Style = "Continuous"
$tabInstall.Controls.Add($progressBarInstall)

# Install button (position matched with uninstall button)
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Location = New-Object System.Drawing.Point(20,410)
$buttonInstall.Size = New-Object System.Drawing.Size(150,40)
$buttonInstall.Text = "Install"
$buttonInstall.Font = $defaultFont
$buttonInstall.Add_Click({
    InstallSoftware
})
$tabInstall.Controls.Add($buttonInstall)

## UNINSTALLATION TAB ##
$tabUninstall = New-Object System.Windows.Forms.TabPage
$tabUninstall.Text = "Uninstallation"
$tabUninstall.Font = $tabFont
$tabControl.Controls.Add($tabUninstall)

# Installed software selection
$labelUninstall = New-Object System.Windows.Forms.Label
$labelUninstall.Location = New-Object System.Drawing.Point(20,20)
$labelUninstall.Size = New-Object System.Drawing.Size(300,25)
$labelUninstall.Text = "Installed Software to Remove:"
$labelUninstall.Font = $defaultFont
$tabUninstall.Controls.Add($labelUninstall)

$comboInstalledSoftware = New-Object System.Windows.Forms.ComboBox
$comboInstalledSoftware.Location = New-Object System.Drawing.Point(20,50)
$comboInstalledSoftware.Size = New-Object System.Drawing.Size(650,30)
$comboInstalledSoftware.Font = $softwaress
$comboInstalledSoftware.DropDownStyle = "DropDownList"
$tabUninstall.Controls.Add($comboInstalledSoftware)

$buttonRefreshList = New-Object System.Windows.Forms.Button
$buttonRefreshList.Location = New-Object System.Drawing.Point(680,48)
$buttonRefreshList.Size = New-Object System.Drawing.Size(125,27)
$buttonRefreshList.Text = "Refresh List"
$buttonRefreshList.Font = $defaultFont
$buttonRefreshList.Add_Click({
    RefreshInstalledSoftwareList
})
$tabUninstall.Controls.Add($buttonRefreshList)

function RefreshInstalledSoftwareList {
    $comboInstalledSoftware.Items.Clear()
    if (Test-Path "C:\Program Files") {
        Get-ChildItem "C:\Program Files" -Directory | ForEach-Object {
            $comboInstalledSoftware.Items.Add($_.Name)
        }
    }
    if ($comboInstalledSoftware.Items.Count -gt 0) {
        $comboInstalledSoftware.SelectedIndex = 0
    }
}

# Uninstallation status label
$labelUninstallStatus = New-Object System.Windows.Forms.Label
$labelUninstallStatus.Location = New-Object System.Drawing.Point(20,100)
$labelUninstallStatus.Size = New-Object System.Drawing.Size(760,30)
$labelUninstallStatus.Text = "Select software to uninstall."
$labelUninstallStatus.Font = $defaultFont
$tabUninstall.Controls.Add($labelUninstallStatus)

# Uninstallation progress bar
$progressBarUninstall = New-Object System.Windows.Forms.ProgressBar
$progressBarUninstall.Location = New-Object System.Drawing.Point(20,140)
$progressBarUninstall.Size = New-Object System.Drawing.Size(760,25)
$progressBarUninstall.Style = "Continuous"
$tabUninstall.Controls.Add($progressBarUninstall)

# Uninstall button (position matched with install button)
$buttonUninstall = New-Object System.Windows.Forms.Button
$buttonUninstall.Location = New-Object System.Drawing.Point(20,180)
$buttonUninstall.Size = New-Object System.Drawing.Size(150,40)
$buttonUninstall.Text = "Uninstall"
$buttonUninstall.Font = $defaultFont
$buttonUninstall.Add_Click({
    UninstallSoftware
})
$tabUninstall.Controls.Add($buttonUninstall)

# Button panel (for shared buttons)
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(10,520)
$buttonPanel.Size = New-Object System.Drawing.Size(820,100)
$form.Controls.Add($buttonPanel)

# Stop button (shared)
$buttonStop = New-Object System.Windows.Forms.Button
$buttonStop.Location = New-Object System.Drawing.Point(20,20)  # Position adjusted
$buttonStop.Size = New-Object System.Drawing.Size(150,40)
$buttonStop.Text = "Stop"
$buttonStop.Font = $defaultFont
$buttonStop.Enabled = $false
$buttonStop.Add_Click({
    $global:shouldStop = $true
    $buttonStop.Enabled = $false
})
$buttonPanel.Controls.Add($buttonStop)

# About button (shared)
$buttonAbout = New-Object System.Windows.Forms.Button
$buttonAbout.Location = New-Object System.Drawing.Point(190,20)  # Position adjusted
$buttonAbout.Size = New-Object System.Drawing.Size(150,40)
$buttonAbout.Text = "About"
$buttonAbout.Font = $defaultFont
$buttonAbout.Add_Click({
    ShowAboutDialog
})
$buttonPanel.Controls.Add($buttonAbout)

# Exit button (shared)
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Location = New-Object System.Drawing.Point(360,20)  # Position adjusted
$buttonExit.Size = New-Object System.Drawing.Size(150,40)
$buttonExit.Text = "Exit"
$buttonExit.Font = $defaultFont
$buttonExit.Add_Click({
    if ($global:operationRunning) {
        $result = [System.Windows.Forms.MessageBox]::Show("Operation is in progress. Are you sure you want to exit?", "Confirmation", "YesNo", "Question")
        if ($result -eq "Yes") {
            $form.Close()
        }
    }
    else {
        $form.Close()
    }
})
$buttonPanel.Controls.Add($buttonExit)




# About dialog function - Final spacing adjustments
function ShowAboutDialog {
    # Add assembly for Windows Forms
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "About Software Packager"
    $aboutForm.Size = New-Object System.Drawing.Size(435,680)
    $aboutForm.StartPosition = "CenterScreen"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false
    $aboutForm.MinimizeBox = $false
    $aboutForm.BackColor = [System.Drawing.Color]::White
    
    # Title label
    $labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Location = New-Object System.Drawing.Point(10,20)
    $labelTitle.Size = New-Object System.Drawing.Size(370,30)
    $labelTitle.Text = "SOFTWARE PACKAGER"
    $labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $labelTitle.ForeColor = [System.Drawing.Color]::DarkBlue
    $labelTitle.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center # This one is correct for title
    $aboutForm.Controls.Add($labelTitle)
    
    # Version label - adjusted spacing from title
    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Location = New-Object System.Drawing.Point(10,65)
    $labelVersion.Size = New-Object System.Drawing.Size(370,20)
    $labelVersion.Text = "Version 1.0"
    $labelVersion.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $labelVersion.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center # This one is correct for version
    $aboutForm.Controls.Add($labelVersion)
    
    # Separator line - moved down to follow version with good spacing
    $separator = New-Object System.Windows.Forms.Label
    $separator.Location = New-Object System.Drawing.Point(20,95)
    $separator.Size = New-Object System.Drawing.Size(360,1)
    $separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $aboutForm.Controls.Add($separator)
    
    # Info panel - large enough for all lines
    $panelInfo = New-Object System.Windows.Forms.Panel
    $panelInfo.Location = New-Object System.Drawing.Point(20,110)
    $panelInfo.Size = New-Object System.Drawing.Size(360,340)
    $panelInfo.BackColor = [System.Drawing.Color]::White
    # Uncomment for debugging to see the panel's boundaries
    # $panelInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $aboutForm.Controls.Add($panelInfo)
    
    # Info text with better formatting and increased line spacing
    $infoLines = @(
        "A tool for software deployment that:",
        "",
        "✓ Installs applications to Program Files",
        "",
        "✓ Creates Start Menu shortcuts",
        "",
        "✓ Adds Desktop icons when needed",
        "",
        "✓ Completely removes applications",
        "",
        "✓ Tracks progress of all operations"
    )
    
    $yPos = 5 # Start a bit higher within the panel
    $lineHeight = 25 # Consistent line spacing
    
    foreach ($line in $infoLines) {
        $infoLabel = New-Object System.Windows.Forms.Label
        $infoLabel.Location = New-Object System.Drawing.Point(5,$yPos)
        $infoLabel.Size = New-Object System.Drawing.Size(350,$lineHeight)
        $infoLabel.Text = $line
        $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        
        # Note: The TextAlign for these labels within panelInfo is implicitly TopLeft by default, 
        # or it uses the label's default setting if not explicitly set.
        # If you want to explicitly left align, you'd also need ContentAlignment here.
        
        if ($line.StartsWith("✓")) {
            $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $infoLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        }
        elseif ($line -eq "A tool for software deployment that:") {
            $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $infoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
        }
        
        $panelInfo.Controls.Add($infoLabel)
        $yPos += $lineHeight
    }

    # --- Creator and Copyright Information (using multiple labels for robustness) ---
    $copyrightFont = New-Object System.Drawing.Font("Segoe UI", 8)
    $copyrightColor = [System.Drawing.Color]::Gray
    $copyrightStartX = 20 # Align with the separator and info panel left edge
    $copyrightLineHeight = 20 # Adjust line height for smaller font

    # Calculate starting Y for this section
    # Bottom of panelInfo is 110 (top) + 340 (height) = 450.
    $currentY = 460 # Start with a 10px gap from the panel bottom

    # Created by: P.Avarwand
    $labelCreatedBy = New-Object System.Windows.Forms.Label
    $labelCreatedBy.Location = New-Object System.Drawing.Point($copyrightStartX, $currentY)
    $labelCreatedBy.Size = New-Object System.Drawing.Size(410, $copyrightLineHeight)
    $labelCreatedBy.Text = "Created by: P.Avarwand"
    $labelCreatedBy.Font = $copyrightFont
    $labelCreatedBy.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $labelCreatedBy.ForeColor = $copyrightColor
    $aboutForm.Controls.Add($labelCreatedBy)
    $currentY += $copyrightLineHeight

    # Contact: payam.avarwand@zit-bb.brandenburg.de
    $labelContact = New-Object System.Windows.Forms.Label
    $labelContact.Location = New-Object System.Drawing.Point($copyrightStartX, $currentY)
    $labelContact.Size = New-Object System.Drawing.Size(410, $copyrightLineHeight)
    $labelContact.Text = "Contact: payam.avarwand@zit-bb.brandenburg.de"
    $labelContact.Font = $copyrightFont
    $labelContact.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $labelContact.ForeColor = $copyrightColor
    $aboutForm.Controls.Add($labelContact)
    $currentY += $copyrightLineHeight

    # © June 2025 by Avarwand
    $labelCopyright = New-Object System.Windows.Forms.Label
    $labelCopyright.Location = New-Object System.Drawing.Point($copyrightStartX, $currentY)
    $labelCopyright.Size = New-Object System.Drawing.Size(410, $copyrightLineHeight)
    $labelCopyright.Text = "© June 2025"
    $labelCopyright.Font = $copyrightFont
    $labelCopyright.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $labelCopyright.ForeColor = [System.Drawing.Color]::Gray
    $aboutForm.Controls.Add($labelCopyright)
    $currentY += $copyrightLineHeight + 20
    # --- END NEW ---
    
    # OK button - adjusted based on precise calculation
    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Location = New-Object System.Drawing.Point(160, 570)
    $buttonOK.Size = New-Object System.Drawing.Size(100,30)
    $buttonOK.Text = "OK"
    $buttonOK.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $buttonOK.BackColor = [System.Drawing.Color]::LightSteelBlue
    $buttonOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonOK.Add_Click({ $aboutForm.Close() })
    $aboutForm.Controls.Add($buttonOK)
    
    $aboutForm.ShowDialog()
}





# Install function
function InstallSoftware {
    if ($global:operationRunning) {
        [System.Windows.Forms.MessageBox]::Show("Operation is already running!", "Information", "OK", "Information")
        return
    }
    
    # Validate inputs
    $sourcePath = $textboxSource.Text
    $exePath = $textboxExe.Text
    
    if (-not (Test-Path $sourcePath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid source folder!", "Error", "OK", "Error")
        return
    }
    
    if (-not (Test-Path $exePath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid EXE file!", "Error", "OK", "Error")
        return
    }
    
    $shortcutName = $textboxShortcutName.Text
    $createShortcuts = $checkboxStartMenu.Checked -or $checkboxDesktop.Checked
    
    if ($createShortcuts -and [string]::IsNullOrWhiteSpace($shortcutName)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a shortcut name!", "Error", "OK", "Error")
        return
    }
    
    # Ensure shortcut name doesn't have .lnk extension
    $shortcutName = $shortcutName -replace '\.lnk$',''
    
    $global:operationRunning = $true
    $global:shouldStop = $false
    $buttonStop.Enabled = $true
    $buttonInstall.Enabled = $false
    
    try {
        $labelInstallStatus.Text = "Preparing installation..."
        $progressBarInstall.Value = 0
        $form.Refresh()
        
        # Get folder name from path
        $folderName = Split-Path $sourcePath -Leaf
        
        # Copy files to Program Files
        $destinationPath = Join-Path "C:\Program Files" $folderName
        $labelInstallStatus.Text = "Copying files to Program Files..."
        $progressBarInstall.Value = 20
        $form.Refresh()
        
        if ($global:shouldStop) { throw "Operation stopped by user" }
        
        Copy-Item -Path $sourcePath -Destination $destinationPath -Recurse -Force
        
        # Create shortcuts if selected
        $Shell = New-Object -ComObject WScript.Shell
        $TargetExe = Join-Path $destinationPath (Split-Path $exePath -Leaf)
        
        if ($checkboxStartMenu.Checked -and -not $global:shouldStop) {
            $labelInstallStatus.Text = "Creating Start Menu shortcut..."
            $progressBarInstall.Value = 60
            $form.Refresh()
            
            $StartMenuShortcut = Join-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs" "$shortcutName.lnk"
            $Shortcut = $Shell.CreateShortcut($StartMenuShortcut)
            $Shortcut.TargetPath = $TargetExe
            $Shortcut.IconLocation = "$TargetExe, 0"
            $Shortcut.Description = $shortcutName
            $Shortcut.WorkingDirectory = $destinationPath
            $Shortcut.Save()
        }
        
        if ($checkboxDesktop.Checked -and -not $global:shouldStop) {
            $labelInstallStatus.Text = "Creating Desktop shortcut..."
            $progressBarInstall.Value = 80
            $form.Refresh()
            
            $DesktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$shortcutName.lnk"
            $Shortcut = $Shell.CreateShortcut($DesktopShortcut)
            $Shortcut.TargetPath = $TargetExe
            $Shortcut.IconLocation = "$TargetExe, 0"
            $Shortcut.Description = $shortcutName
            $Shortcut.WorkingDirectory = $destinationPath
            $Shortcut.Save()
        }
        
        if (-not $global:shouldStop) {
            $labelInstallStatus.Text = "Installation completed successfully!"
            $progressBarInstall.Value = 100
            [System.Windows.Forms.MessageBox]::Show("Software installed successfully!", "Success", "OK", "Information")
            RefreshInstalledSoftwareList
        }
    }
    catch {
        $labelInstallStatus.Text = if ($global:shouldStop) { "Installation stopped by user" } else { "Error during installation: $($_.Exception.Message)" }
        $progressBarInstall.Value = 0
        if (-not $global:shouldStop) {
            [System.Windows.Forms.MessageBox]::Show("Error during installation: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    }
    finally {
        $global:operationRunning = $false
        $buttonStop.Enabled = $false
        $buttonInstall.Enabled = $true
    }
}

# Uninstall function
function UninstallSoftware {
    if ($global:operationRunning) {
        [System.Windows.Forms.MessageBox]::Show("Operation is already running!", "Information", "OK", "Information")
        return
    }
    
    if ($comboInstalledSoftware.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select software to uninstall!", "Error", "OK", "Error")
        return
    }
    
    $softwareName = $comboInstalledSoftware.SelectedItem.ToString()
    
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to completely uninstall $softwareName?`nThis will remove the program folder and all shortcuts.", "Confirmation", "YesNo", "Warning")
    if ($result -ne "Yes") {
        return
    }
    
    $global:operationRunning = $true
    $global:shouldStop = $false
    $buttonStop.Enabled = $true
    $buttonUninstall.Enabled = $false
    
    try {
        $labelUninstallStatus.Text = "Starting uninstallation..."
        $progressBarUninstall.Value = 10
        $form.Refresh()
        
        # Remove from Program Files
        $programFolder = Join-Path "C:\Program Files" $softwareName
        $labelUninstallStatus.Text = "Removing program files..."
        $progressBarUninstall.Value = 30
        $form.Refresh()
        
        if ($global:shouldStop) { throw "Operation stopped by user" }
        
        if (Test-Path $programFolder) {
            Remove-Item $programFolder -Recurse -Force -ErrorAction Stop
        }
        
        # Remove Start Menu shortcuts
        $labelUninstallStatus.Text = "Removing Start Menu shortcuts..."
        $progressBarUninstall.Value = 60
        $form.Refresh()
        
        if ($global:shouldStop) { throw "Operation stopped by user" }
        
        $startMenuShortcuts = Get-ChildItem "C:\ProgramData\Microsoft\Windows\Start Menu\Programs" -Filter "*.lnk" -File
        foreach ($shortcut in $startMenuShortcuts) {
            $shell = New-Object -ComObject WScript.Shell
            $lnk = $shell.CreateShortcut($shortcut.FullName)
            if ($lnk.TargetPath -like "*\$softwareName\*") {
                Remove-Item $shortcut.FullName -Force
            }
        }
        
        # Remove Desktop shortcuts
        $labelUninstallStatus.Text = "Removing Desktop shortcuts..."
        $progressBarUninstall.Value = 80
        $form.Refresh()
        
        if ($global:shouldStop) { throw "Operation stopped by user" }
        
        $desktopShortcuts = Get-ChildItem ([Environment]::GetFolderPath("Desktop")) -Filter "*.lnk" -File
        foreach ($shortcut in $desktopShortcuts) {
            $shell = New-Object -ComObject WScript.Shell
            $lnk = $shell.CreateShortcut($shortcut.FullName)
            if ($lnk.TargetPath -like "*\$softwareName\*") {
                Remove-Item $shortcut.FullName -Force
            }
        }
        
        if (-not $global:shouldStop) {
            $labelUninstallStatus.Text = "Uninstallation completed successfully!"
            $progressBarUninstall.Value = 100
            [System.Windows.Forms.MessageBox]::Show("Software uninstalled successfully!", "Success", "OK", "Information")
            RefreshInstalledSoftwareList
        }
    }
    catch {
        $labelUninstallStatus.Text = if ($global:shouldStop) { "Uninstallation stopped by user" } else { "Error during uninstallation: $($_.Exception.Message)" }
        $progressBarUninstall.Value = 0
        if (-not $global:shouldStop) {
            [System.Windows.Forms.MessageBox]::Show("Error during uninstallation: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    }
    finally {
        $global:operationRunning = $false
        $buttonStop.Enabled = $false
        $buttonUninstall.Enabled = $true
    }
}

# Tab change handler to update UI
$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabInstall) {
        $buttonInstall.Visible = $true
        $buttonUninstall.Visible = $false
    }
    elseif ($tabControl.SelectedTab -eq $tabUninstall) {
        $buttonInstall.Visible = $false
        $buttonUninstall.Visible = $true
    }
})

# Initialize visibility
UpdateShortcutNameFieldVisibility
RefreshInstalledSoftwareList
$buttonInstall.Visible = $true
$buttonUninstall.Visible = $false

# Show the form
[void]$form.ShowDialog()