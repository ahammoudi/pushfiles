# Requires -Modules Wpf, PSDesiredStateConfiguration

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Server File Copy and Command Execution"
$form.Width = 500
$form.Height = 450

# Drop-down list (ComboBox) - Double the size
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(10, 10)
$dropdown.Width = 200
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

# File browse button - Double the size, on a new line
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse File"
$browseButton.Location = New-Object System.Drawing.Point(10, 60)
$browseButton.Width = 160
$browseButton.Height = 40
$form.Controls.Add($browseButton)

$filePathLabel = New-Object System.Windows.Forms.Label
$filePathLabel.Location = New-Object System.Drawing.Point(10, 100)
$filePathLabel.Width = 400
$form.Controls.Add($filePathLabel)

# Output box (TextBox)
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 140)
$outputBox.Width = 460
$outputBox.Height = 200
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$form.Controls.Add($outputBox)

# Debug Checkbox
$debugCheckbox = New-Object System.Windows.Forms.CheckBox
$debugCheckbox.Text = "Debug Mode"
$debugCheckbox.Location = New-Object System.Drawing.Point(10, 350)
$form.Controls.Add($debugCheckbox)

# Submit button - Moved to the bottom
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(200, 380)
$form.Controls.Add($submitButton)

# Event handlers
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "ZIP files (*.zip)|*.zip|All files (*.*)|*.*"
    $openFileDialog.Title = "Select a ZIP File"

    if ($openFileDialog.ShowDialog() -eq "OK") {
        $filePath = $openFileDialog.FileName

        if ($filePath.EndsWith(".zip", [System.StringComparison]::OrdinalIgnoreCase)) {
            $filePathLabel.Text = $filePath
            $fileName = [System.IO.Path]::GetFileName($filePath)
            $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
            $outputBox.Text += "Selected file: $fileName" + [Environment]::NewLine
            $outputBox.Text += "Selected file (without extension): $fileNameWithoutExtension" + [Environment]::NewLine
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a ZIP file.", "Invalid File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $filePathLabel.Text = ""
        }

    }
})

$submitButton.Add_Click({
    $filePath = $filePathLabel.Text

    # *** KEY FIX: Check dropdown *before* anything else ***
    if ($dropdown.SelectedIndex -eq 0) {  #Check index, not selected items
        $outputBox.Text += "Please select a server list." + [Environment]::NewLine
        return
    }

    if (-not $filePath) {
        $outputBox.Text += "Please select a file to copy." + [Environment]::NewLine
        return
    }

    $selectedItems = $dropdown.SelectedItems # Get selected items AFTER the check

    $cred = Get-Credential -Message "Enter credentials for remote servers"
    $isDebug = $debugCheckbox.Checked

    foreach ($item in $selectedItems) {
        try {
            $serverListFile = "$item.txt"
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
                                $status = $progressRecord.StatusDescription
                                if ($isDebug) {
                                    $outputBox.Text += "  Copy Progress on ${server}: $($percentComplete)% - $status" + [Environment]::NewLine
                                    $outputBox.CaretIndex = $outputBox.Text.Length # Keep the output box scrolled to the bottom
                                }
                            }
                        }

                        Register-ObjectEvent -InputObject $job -Splat $progressSplat

                        # *** KEY FIX: Capture output and errors from the job ***
                        $job | Wait-Job | Receive-Job -Keep | Out-String | ForEach-Object { $outputBox.Text += $_ + [Environment]::NewLine } # Capture output

                        Unregister-Event -SourceIdentifier "FileCopyProgress"

                        if ($job.State -eq "Failed") {
                            $outputBox.Text += "  Copy failed on ${server}: $($job.Error.Exception.Message)" + [Environment]::NewLine
                        } elseif ($isDebug) {
                            $outputBox.Text += "  Copy completed successfully on ${server}." + [Environment]::NewLine
                        }

                        if ($isDebug) {
                            $outputBox.Text += "Executing command(s) on ${server}..." + [Environment]::NewLine
                        }

                        $commands = @(
                            "Get-Process",
                            "Get-Service -Name *spool*",
                            "Write-Host 'Hello from remote server!'" # Example
                        )

                        foreach ($command in $commands) {
                            try {
                                # *** KEY FIX: Capture output from Invoke-Command ***
                                $result = Invoke-Command -ComputerName ${server} -ScriptBlock { & $args[0] } -ArgumentList $command -Credential $cred

                                # *** KEY FIX: Output to textbox ***
                                $result | Out-String | ForEach-Object { $outputBox.Text += "  Command: $command`n  Result: $_" + [Environment]::NewLine }

                            } catch {
                                $outputBox.Text += "  Error executing command '$command' on ${server}: $($_.Exception.Message)" + [Environment]::NewLine
                            }
                        }

                    } catch {
                        $outputBox.Text += "Error with ${server}: $($_.Exception.Message)" + [Environment]::NewLine
                    }
                }
            }

        } catch {
             $outputBox.Text += "Error reading server list file '$item.txt': $($_.Exception.Message)" + [Environment]::NewLine
        }
    }

})

# Show the form
$form.ShowDialog()