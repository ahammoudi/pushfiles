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
    $fileDialog.Filter = "ZIP files (*.zip)|*.zip"
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

# Execute Commands button
$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Text = "Execute Commands"
$executeButton.Location = New-Object System.Drawing.Point(180, 520)
$executeButton.Width = 160
$executeButton.Height = 40
$form.Controls.Add($executeButton)

# Variable to store credentials
$global:cred = $null

# Function to get credentials
function Get-GlobalCredential {
    if (-not $global:cred) {
        $global:cred = Get-Credential -Message "Enter credentials for remote servers"
    }
    return $global:cred
}

# Copy button click event
$copyButton.Add_Click({
    $filePath = $filePathBox.Text
    if (-not $filePath) {
        $outputBox.Text += "Please select a file to copy." + [Environment]::NewLine
        return
    }

    $fileName = Split-Path -Path $filePath -Leaf

    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        $outputBox.Text += "Please select a server list." + [Environment]::NewLine
        return
    }

    $cred = Get-GlobalCredential
    $isDebug = $debugCheckbox.Checked

    try {
        $serverListFile = "$selectedItem.txt"
        $servers = Get-Content $serverListFile

        foreach ($server in $servers) {
            if (-not [string]::IsNullOrWhiteSpace($server)) {
                if ($isDebug) {
                    $outputBox.Text += "Copying ${fileName} to ${server}..." + [Environment]::NewLine
                }
                try {
                    $destinationPath = "\\${server}\C$\temp"

                    # Start the copy operation as a job
                    $job = Start-Job -ScriptBlock {
                        param ($filePath, $destinationPath, $cred)
                        Copy-Item -Path $filePath -Destination $destinationPath -Credential $cred -Force
                    } -ArgumentList $filePath, $destinationPath, $cred

                    # Wait for the job to complete
                    $job | Wait-Job

                    # Get the job output
                    $jobOutput = Receive-Job -Job $job

                    # Check the job state
                    if ($job.State -eq 'Completed') {
                        $outputBox.Text += "Copy of ${fileName} to ${server} completed successfully." + [Environment]::NewLine
                    } else {
                        $outputBox.Text += "Copy of ${fileName} to ${server} failed." + [Environment]::NewLine
                    }

                    # Display the job output
                    $outputBox.Text += $jobOutput + [Environment]::NewLine

                    # Clean up the job
                    Remove-Job -Job $job
                } catch {
                    $outputBox.Text += "Error copying ${fileName} to ${server}: $($_.Exception.Message)" + [Environment]::NewLine
                }
            }
        }
    } catch {
        $outputBox.Text += "Error reading server list file '$serverListFile': $($_.Exception.Message)" + [Environment]::NewLine
    }
})

# Execute Commands button click event
$executeButton.Add_Click({
    $filePath = $filePathBox.Text
    $fileName = Split-Path -Path $filePath -Leaf

    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        $outputBox.Text += "Please select a server list." + [Environment]::NewLine
        return
    }

    $cred = Get-GlobalCredential
    $isDebug = $debugCheckbox.Checked

    try {
        $serverListFile = "$selectedItem.txt"
        $servers = Get-Content $serverListFile

        foreach ($server in $servers) {
            if (-not [string]::IsNullOrWhiteSpace($server)) {
                if ($isDebug) {
                    $outputBox.Text += "Executing commands on ${server}..." + [Environment]::NewLine
                }
                try {
                    # Start the command execution as a job
                    $job = Start-Job -ScriptBlock {
                        param ($server, $cred, $fileName)
                        $session = New-PSSession -ComputerName $server -Credential $cred
                        Invoke-Command -Session $session -ScriptBlock {
                            param ($fileName)
                            # Add your additional commands here
                            Write-Host "Executing additional commands on $env:COMPUTERNAME"
                            # Example: Unzip the file
                            $zipPath = "C:\temp\$fileName"
                            $extractPath = "C:\temp\extracted"
                            Add-Type -AssemblyName System.IO.Compression.FileSystem
                            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractPath)
                            # Example: Write a log file
                            "Commands executed on $env:COMPUTERNAME" | Out-File "C:\temp\log.txt"
                        } -ArgumentList $fileName
                        Remove-PSSession -Session $session
                    } -ArgumentList $server, $cred, $fileName

                    # Wait for the job to complete
                    $job | Wait-Job

                    # Get the job output
                    $jobOutput = Receive-Job -Job $job

                    # Check the job state
                    if ($job.State -eq 'Completed') {
                        $outputBox.Text += "Commands executed on ${server} successfully." + [Environment]::NewLine
                    } else {
                        $outputBox.Text += "Commands execution on ${server} failed." + [Environment]::NewLine
                    }

                    # Display the job output
                    $outputBox.Text += $jobOutput + [Environment]::NewLine

                    # Clean up the job
                    Remove-Job -Job $job
                } catch {
                    $outputBox.Text += "Error executing commands on ${server}: $($_.Exception.Message)" + [Environment]::NewLine
                }
            }
        }
    } catch {
        $outputBox.Text += "Error reading server list file '$serverListFile': $($_.Exception.Message)" + [Environment]::NewLine
    }
})

# Show the form
$form.ShowDialog()