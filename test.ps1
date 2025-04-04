# Browse button click event
$browseButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder"
    $folderDialog.ShowNewFolderButton = $true

    # Set initial directory from config
    if ($config -and $config.DefaultBrowsePath -and (Test-Path $config.DefaultBrowsePath)) {
        $folderDialog.SelectedPath = $config.DefaultBrowsePath
    }
    if ($folderDialog.ShowDialog() -eq "OK") {
        $zipFilePathBox.Text = $folderDialog.SelectedPath
    }
})


# Update the Browse button click event
$browseButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.OpenFileDialog
    $folderDialog.Filter = "All Files (*.*)|*.*|Executable Files (*.exe)|*.exe|Installation Files (*.msi)|*.msi"
    $folderDialog.FilterIndex = 1
    $folderDialog.Multiselect = $false
    $folderDialog.Title = "Select Installation File"
    
    # Set initial directory from config
    if ($config -and $config.DefaultBrowsePath -and (Test-Path $config.DefaultBrowsePath)) {
        $folderDialog.InitialDirectory = $config.DefaultBrowsePath
    }
    else {
        $folderDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    }

    # Set dialog properties
    $folderDialog.AutoUpgradeEnabled = $true
    $folderDialog.CheckFileExists = $true
    $folderDialog.CheckPathExists = $true
    $folderDialog.DereferenceLinks = $true
    $folderDialog.ShowHelp = $false
    $folderDialog.SupportMultiDottedExtensions = $true
    $folderDialog.ValidateNames = $true
    
    if ($folderDialog.ShowDialog() -eq "OK") {
        $zipFilePathBox.Text = Split-Path $folderDialog.FileName -Parent
    }
})