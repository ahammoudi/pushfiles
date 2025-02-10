Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Load configuration
try {
    $config = Get-Content -Path ".\config\config.json" | ConvertFrom-Json
    
    $LogFile = $config.LogFile
    $TargetMSIFileName = $config.TargetMSIFileName
    $TargetMSIPathInZip = $config.TargetMSIPathInZip
    $YourAppPoolName = $config.AppPoolName
    $YourServiceName = $config.ServiceName

    if (-not (Test-Path $LogFile)) {
        New-Item -ItemType File -Path $LogFile -Force
    }

    # Add log rotation function and check here
    function Move-Logs {
        param(
            [string]$LogFile,
            [int]$MaxSize = 10MB,
            [int]$MaxFiles = 5
        )

        if ((Get-Item $LogFile).Length -gt $MaxSize) {
            for ($i = $MaxFiles; $i -gt 0; $i--) {
                $oldFile = "${LogFile}.${i}"
                $newFile = "${LogFile}.$($i + 1)"
                if (Test-Path $oldFile) {
                    Move-Item -Path $oldFile -Destination $newFile -Force
                }
            }
            Move-Item -Path $LogFile -Destination "${LogFile}.1" -Force
            New-Item -ItemType File -Path $LogFile -Force | Out-Null
        }
    }

    # Check for log rotation
    Move-Logs -LogFile $LogFile

} catch {
    Write-Error "Failed to load configuration: $_"
    exit 1
}

# Cleanup function for remote servers
function Remove-TempFiles {
    param (
        [string]$server
    )
    try {
        $session = New-PSSession -ComputerName $server -Credential $global:cred -ErrorAction Stop
        if ($session) {
            try {
                Invoke-Command -Session $session -ScriptBlock {
                    Remove-Item -Path "C:\temp\extracted" -Recurse -Force -ErrorAction SilentlyContinue
                    Remove-Item -Path "C:\temp\*.zip" -Force -ErrorAction SilentlyContinue
                    Remove-Item -Path "C:\temp\log.txt" -Force -ErrorAction SilentlyContinue
                }
            }
            finally {
                Remove-PSSession -Session $session
            }
        } else {
            Write-LogMessage "Failed to create session for ${server}" -IsError
        }
    } catch {
        Write-LogMessage "Error accessing ${server}: $($_.Exception.Message)" -IsError
    }
}

#Update progress 
function Update-Progress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Operation
    )
    $percentComplete = [math]::Floor(($Current / $Total) * 100)
    Write-LogMessage "${Operation} (${Current}/${Total})" -ProgressValue $percentComplete
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Deployment Tool"
$form.Width = 800
$form.Height = 600
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)

# Add form cleanup on close
$form.Add_FormClosing({
    param($formSender, $e)
    if ($global:cred) { $global:cred = $null }
    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
})

# Add status strip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Add progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 85)
$progressBar.Width = 760
$progressBar.Height = 20
$progressBar.Style = 'Continuous'
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Drop-down list (ComboBox) for server list files
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(10, 10)
$dropdown.Width = 200
$dropdown.Height = 40
$dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Add items to the dropdown (server list files)
$dropdown.Items.Add("Select Server List") # Default item
$dropdown.SelectedIndex = 0 # Select the default item initially
$textFiles = Get-ChildItem -Path . -Filter ".\config\*.txt" # Get all .txt files in the current directory
foreach ($file in $textFiles) {
    $dropdown.Items.Add($file.BaseName) # Add the base name (without extension) to the dropdown
}
$form.Controls.Add($dropdown)

$syncButton = New-Object System.Windows.Forms.Button
$syncButton.Text = "Sync Releases"
$syncButton.Width = 160
$syncButton.Height = 40
$form.Controls.Add($syncButton)

# File browse button for MSI ZIP file
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Select folder"
$browseButton.Location = New-Object System.Drawing.Point(10, 60)
$browseButton.Width = 160
$browseButton.Height = 20
$form.Controls.Add($browseButton)

# TextBox to display the selected ZIP file path
$zipFilePathBox = New-Object System.Windows.Forms.TextBox
$zipFilePathBox.Location = New-Object System.Drawing.Point(175, 60)
$zipFilePathBox.Width = 430
$form.Controls.Add($zipFilePathBox)

# Browse button click event
$browseButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder"
    $folderDialog.ShowNewFolderButton = $true
    
    if ($folderDialog.ShowDialog() -eq "OK") {
        $zipFilePathBox.Text = $folderDialog.SelectedPath
    }
})
# Output TextBox
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 110)
$outputBox.Width = 760
$outputBox.Height = 380
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($outputBox)

# Push ZIP button
$pushZipButton = New-Object System.Windows.Forms.Button
$pushZipButton.Text = "Push ZIP File"
$pushZipButton.Width = 160
$pushZipButton.Height = 40
$form.Controls.Add($pushZipButton)

# Install MSI button
$installMsiButton = New-Object System.Windows.Forms.Button
$installMsiButton.Text = "Install MSI"
$installMsiButton.Width = 160
$installMsiButton.Height = 40
$form.Controls.Add($installMsiButton)

# Function to center buttons
function Set-ButtonsAlignment {
    $formWidth = [int]$form.ClientSize.Width
    $buttonWidth = [int]$pushZipButton.Width
    $spacing = [int]20
    $totalWidth = [int](($buttonWidth * 3) + ($spacing * 2))
    $startX = [int](($formWidth - $totalWidth) / 2)
    $y = [int]($outputBox.Location.Y + $outputBox.Height + 10)

    $pushZipButton.Location = New-Object System.Drawing.Point($startX, $y)
    $installMsiButton.Location = New-Object System.Drawing.Point(($startX + $buttonWidth + $spacing), $y)
    $syncButton.Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $spacing) * 2), $y)
}

# Add resize event handler
$form.Add_Resize({ Set-ButtonsAlignment })

# Variable to store credentials
$global:cred = $null

# Function to get credentials
function Get-GlobalCredential {
    if (-not $global:cred) {
        $global:cred = Get-Credential -Message "Enter credentials for remote servers"
    }
    return $global:cred
}

# Function to write to the log file and output box
function Write-LogMessage {
    param(
        [string]$Message,
        [switch]$IsError,
        [int]$ProgressValue = -1
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if ($IsError) {
        $LogEntry = "[$Timestamp] ERROR: $Message"
        Write-Error $LogEntry
        $outputBox.SelectionColor = [System.Drawing.Color]::Red
        $statusLabel.Text = "Error: $Message"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    } else {
        $LogEntry = "[$Timestamp] INFO: $Message"
        Write-Host $LogEntry
        $outputBox.SelectionColor = [System.Drawing.Color]::Black
        $statusLabel.Text = $Message
        $statusLabel.ForeColor = [System.Drawing.Color]::Black
    }

    if ($ProgressValue -ge 0) {
        $progressBar.Visible = $true
        $progressBar.Value = $ProgressValue
    } elseif ($ProgressValue -eq 100) {
        $progressBar.Visible = $false
        $progressBar.Value = 0
    }

    Add-Content -Path $LogFile -Value $LogEntry
    $outputBox.AppendText($LogEntry + [Environment]::NewLine)
    $outputBox.ScrollToCaret()
}
$syncButton.Add_Click({
    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        Write-LogMessage "Please select a server list." -IsError
        return
    }

    $cred = Get-GlobalCredential
    
    # Get remote server path from config
    $remotePath = $config.RemoteServer
    if (-not $remotePath) {
        Write-LogMessage "Remote server path not configured in config.json" -IsError
        return
    }

    # Create local releases directory if it doesn't exist
    $localPath = Join-Path $PSScriptRoot "releases"
    if (-not (Test-Path $localPath)) {
        New-Item -ItemType Directory -Path $localPath | Out-Null
    }

    try {
        Write-LogMessage "Starting sync from $remotePath..." -ProgressValue 0

        $job = Start-Job -ScriptBlock {
            param ($remotePath, $localPath, [PSCredential]$cred)
            
            # Get all folders in remote path
            $remoteFolders = Get-ChildItem -Path $remotePath -Directory -Credential $cred
            $totalItems = $remoteFolders.Count
            $current = 0

            foreach ($folder in $remoteFolders) {
                $current++
                $targetPath = Join-Path $localPath $folder.Name
                
                # Create folder if it doesn't exist
                if (-not (Test-Path $targetPath)) {
                    New-Item -ItemType Directory -Path $targetPath | Out-Null
                }

                # Copy all contents recursively
                Copy-Item -Path (Join-Path $folder.FullName "*") -Destination $targetPath -Recurse -Force -Credential $cred
                
                Write-Output "Synced folder: $($folder.Name)"
            }
        } -ArgumentList $remotePath, $localPath, $cred

        # Wait for sync to complete
        $job | Wait-Job
        $jobOutput = Receive-Job -Job $job

        if ($job.State -eq 'Completed') {
            Write-LogMessage "Sync completed successfully." -ProgressValue 100
            $jobOutput | ForEach-Object { Write-LogMessage $_ }
        } else {
            Write-LogMessage "Sync operation failed." -IsError
        }

        Remove-Job -Job $job

    } catch {
        Write-LogMessage "Error during sync: $($_.Exception.Message)" -IsError
        $progressBar.Visible = $false
    }
})

# Push ZIP button click event
$pushZipButton.Add_Click({
    $zipFilePath = $zipFilePathBox.Text
    if (-not $zipFilePath) {
        Write-LogMessage "Please select a ZIP file to push." -IsError
        return
    }

    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        Write-LogMessage "Please select a server list." -IsError
        return
    }

    $cred = Get-GlobalCredential

    try {
        $serverListFile = ".\config\$selectedItem.txt"
        $servers = Get-Content $serverListFile
        $totalServers = ($servers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
        $currentServer = 0

        foreach ($server in $servers) {
            if (-not [string]::IsNullOrWhiteSpace($server)) {
                $currentServer++
                Update-Progress -Current $currentServer -Total $totalServers -Operation "Pushing ZIP to servers"

                try {
                    $destinationPath = "\\${server}\C$\temp"
                    $job = Start-Job -ScriptBlock {
                        param ($zipFilePath, $destinationPath, [PSCredential]$cred)
                        Copy-Item -Path $zipFilePath -Destination $destinationPath -Credential $cred -Force
                    } -ArgumentList $zipFilePath, $destinationPath, $cred

                    $job | Wait-Job
                    $jobOutput = Receive-Job -Job $job

                    if ($job.State -eq 'Completed') {
                        Write-LogMessage "Copy of ${zipFilePath} to ${server} completed successfully."
                    } else {
                        Write-LogMessage "Copy of ${zipFilePath} to ${server} failed." -IsError
                    }

                    $outputBox.AppendText($jobOutput + [Environment]::NewLine)
                    Remove-Job -Job $job
                } catch {
                    Write-LogMessage "Error copying ${zipFilePath} to ${server}: $($_.Exception.Message)" -IsError
                }
            }
        }

        Write-LogMessage "Push operation completed." -ProgressValue 100
    } catch {
        Write-LogMessage "Error reading server list file '$serverListFile': $($_.Exception.Message)" -IsError
        $progressBar.Visible = $false
    }
})

# Install MSI button click event
$installMsiButton.Add_Click({
    $zipFilePath = $zipFilePathBox.Text
    $fileName = Split-Path -Path $zipFilePath -Leaf

    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        Write-LogMessage "Please select a server list." -IsError
        return
    }

    $cred = Get-GlobalCredential

    try {
        $serverListFile = ".\config\$selectedItem.txt"
        $servers = Get-Content $serverListFile
        $totalServers = ($servers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
        $currentServer = 0

        foreach ($server in $servers) {
            if (-not [string]::IsNullOrWhiteSpace($server)) {
                $currentServer++
                Update-Progress -Current $currentServer -Total $totalServers -Operation "Installing MSI on servers"

                try {
                    $job = Start-Job -ScriptBlock {
                        param ($server, [PSCredential]$cred, $fileName)
                        $session = New-PSSession -ComputerName $server -Credential $cred
                        Invoke-Command -Session $session -ScriptBlock {
                            param ($fileName)
                            Write-Host "Executing commands on $env:COMPUTERNAME"
                            $zipPath = "C:\temp\$fileName"
                            $extractPath = "C:\temp\extracted"
                            Add-Type -AssemblyName System.IO.Compression.FileSystem
                            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractPath)
                            $msiPath = Join-Path -Path $extractPath -ChildPath $TargetMSIFileName
                            Start-Process msiexec.exe -ArgumentList "/i `"$msiPath`" /quiet /norestart" -Wait -NoNewWindow
                            Restart-Service -Name 'W3SVC'
                            Import-Module WebAdministration
                            Restart-WebAppPool -Name $YourAppPoolName
                            Restart-Service -Name $YourServiceName
                            "Commands executed on $env:COMPUTERNAME" | Out-File "C:\temp\log.txt"
                        } -ArgumentList $fileName
                        Remove-PSSession -Session $session
                    } -ArgumentList $server, $cred, $fileName

                    $job | Wait-Job
                    $jobOutput = Receive-Job -Job $job

                    if ($job.State -eq 'Completed') {
                        Write-LogMessage "Commands executed on ${server} successfully."
                        Remove-TempFiles -server $server
                    } else {
                        Write-LogMessage "Commands execution on ${server} failed." -IsError
                    }

                    $outputBox.AppendText($jobOutput + [Environment]::NewLine)
                    Remove-Job -Job $job
                } catch {
                    Write-LogMessage "Error executing commands on ${server}: $($_.Exception.Message)" -IsError
                }
            }
        }

        Write-LogMessage "Installation completed." -ProgressValue 100
    } catch {
        Write-LogMessage "Error reading server list file '$serverListFile': $($_.Exception.Message)" -IsError
        $progressBar.Visible = $false
    }
})

# Center buttons initially
Set-ButtonsAlignment

# Show the form
$form.ShowDialog()