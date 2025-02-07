# Requires -Modules Wpf, PSDesiredStateConfiguration

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Server File Copy and Command Execution"
$form.Width = 800
$form.Height = 600

# Drop-down list (ComboBox)
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(10, 10)
$dropdown.Width = 600
$dropdown.Height = 40
$dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Add items to the dropdown (server list files)
$dropdown.Items.Add("Select Server List") # Default item
$dropdown.SelectedIndex = 0 # Select the default item initially
$textFiles = Get-ChildItem -Path . -Filter "*.txt" # Get all .txt files in the current directory
foreach ($file in $textFiles) {
    $dropdown.Items.Add($file.BaseName) # Add the base name (without extension) to the dropdown
}
$form.Controls.Add($dropdown)

# File browse button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse File"
$browseButton.Location = New-Object System.Drawing.Point(10, 60)
$browseButton.Width = 160
$browseButton.Height = 40
$form.Controls.Add($browseButton)

# TextBox to display the selected file path
$filePathBox = New-Object System.Windows.Forms.TextBox
$filePathBox.Location = New-Object System.Drawing.Point(180, 60)
$filePathBox.Width = 600
$form.Controls.Add($filePathBox)

# Browse button click event
$browseButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    if ($fileDialog.ShowDialog() -eq "OK") {
        $filePathBox.Text = $fileDialog.FileName
    }
})

# Debug mode checkbox
$debugCheckbox = New-Object System.Windows.Forms.CheckBox
$debugCheckbox.Text = "Debug Mode"
$debugCheckbox.Location = New-Object System.Drawing.Point(10, 110)
$form.Controls.Add($debugCheckbox)

# Output TextBox
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 150)
$outputBox.Width = 760
$outputBox.Height = 350
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$form.Controls.Add($outputBox)

# Copy button
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "Copy"
$copyButton.Location = New-Object System.Drawing.Point(10, 520)
$copyButton.Width = 160
$copyButton.Height = 40
$form.Controls.Add($copyButton)

# Copy button click event
$copyButton.Add_Click({
    $filePath = $filePathBox.Text
    if (-not $filePath) {
        $outputBox.Text += "Please select a file to copy." + [Environment]::NewLine
        return
    }

    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        $outputBox.Text += "Please select a server list." + [Environment]::NewLine
        return
    }

    $cred = Get-Credential -Message "Enter credentials for remote servers"
    $isDebug = $debugCheckbox.Checked

    try {
        $serverListFile = "$selectedItem.txt"
        $servers = Get-Content $serverListFile

        foreach ($server in $servers) {
            if (-not [string]::IsNullOrWhiteSpace($server)) {
                if ($isDebug) {
                    $outputBox.Text += "Copying to ${server}..." + [Environment]::NewLine
                }
                try {
                    $destinationPath = "\\${server}\C$\temp"

                    # Copy with Progress Reporting (using Copy-Item)
                    $job = Copy-Item -Path $filePath -Destination $destinationPath -Credential $cred -Force -PassThru -AsJob

                    $progressSplat = @{
                        EventName = "Progress"
                        SourceIdentifier = "FileCopyProgress"
                        Action = {
                            param([object] $progressEvent)
                            $progressRecord = $progressEvent.ProgressRecord
                            $percentComplete = $progressRecord.PercentComplete
                            # Update the outputBox with progress
                            $outputBox.Text += "Copying to ${server}: $percentComplete% complete" + [Environment]::NewLine
                        }
                    }

                    Register-ObjectEvent -InputObject $job -EventName "ProgressChanged" @progressSplat

                    # Wait for the job to complete
                    $job | Wait-Job

                    if ($job.State -eq 'Completed') {
                        $outputBox.Text += "Copy to ${server} completed successfully." + [Environment]::NewLine
                    } else {
                        $outputBox.Text += "Copy to ${server} failed." + [Environment]::NewLine
                    }

                    # Clean up the job
                    Remove-Job -Job $job
                } catch {
                    $outputBox.Text += "Error copying to ${server}: $($_.Exception.Message)" + [Environment]::NewLine
                }
            }
        }
    } catch {
        $outputBox.Text += "Error reading server list file '$serverListFile': $($_.Exception.Message)" + [Environment]::NewLine
    }
})

# Show the form
$form.ShowDialog()