<#
.SYNOPSIS
    Deployment Tool for pushing and installing MSI packages on remote servers.

.DESCRIPTION
    This script provides a graphical user interface (GUI) for deploying MSI packages to remote servers.
    It includes functionalities for selecting server lists, pushing ZIP files, installing MSI packages,
    and backing out installations. The script also handles logging, progress updates, and error handling.

.PARAMETER LogFile
    The path to the log file where log messages will be written.

.PARAMETER InstallerName
    The name of the installer to be used for installation.

.PARAMETER AppPoolName
    The name of the application pool to be restarted after installation.

.PARAMETER ServiceName
    The name of the service to be restarted after installation.

.FUNCTION Move-Logs
    Rotates the log file if it exceeds a specified size and maintains a specified number of log files.

.FUNCTION Remove-TempFiles
    Cleans up temporary files on remote servers.

.FUNCTION Update-Progress
    Updates the progress bar and log messages with the current progress of an operation.

.FUNCTION Get-GlobalCredential
    Prompts the user for credentials and validates them against a test server.

.FUNCTION Test-Credential
    Tests the provided credentials against a specified server.

.FUNCTION Write-LogMessage
    Writes log messages to the log file and the output box in the GUI.

.FUNCTION Set-ButtonsAlignment
    Centers the buttons on the form and adjusts their positions when the form is resized.

.FUNCTION Get-GlobalCredential
    Prompts the user for credentials and validates them against a test server.

.FUNCTION Test-Credential
    Tests the provided credentials against a specified server.

.FUNCTION Write-LogMessage
    Writes log messages to the log file and the output box in the GUI.

.FUNCTION Set-ButtonsAlignment
    Centers the buttons on the form and adjusts their positions when the form is resized.

.NOTES
    Created by Ayad.

.EXAMPLE
    To run the script, simply execute it in PowerShell:
    .\push_install_v2.ps1

#>
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Load configuration
try {
    $config = Get-Content -Path "$PSScriptRoot\config\config.json" | ConvertFrom-Json
    
    $LogFile = $config.LogFile
    $InstallerName = $config.InstallerName
    $AppPoolName = $config.AppPoolName
    $ServiceName = $config.ServiceName

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

}
catch {
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
        }
        else {
            Write-LogMessage "Failed to create session for ${server}" -IsError
        }
    }
    catch {
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

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Update progress bar properties
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 85)
$progressBar.Width = [int]$form.ClientSize.Width - 20
$progressBar.Height = 25  # Increased height to accommodate text
$progressBar.Style = 'Continuous'
$progressBar.Visible = $false
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
[System.Windows.Forms.AnchorStyles]::Left -bor 
[System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progressBar)

# Progress label with proper integer casting
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Width = 150
$progressLabel.Height = 20
$progressLabel.Text = "0%"
$progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$progressLabel.BackColor = [System.Drawing.Color]::Transparent
$progressLabel.ForeColor = [System.Drawing.Color]::Black
$progressLabel.Visible = $false
$progressLabel.BringToFront()  # Make sure label is on top
$progressLabel.Location = New-Object System.Drawing.Point(
    ([int]$progressBar.Location.X + ([int]$progressBar.Width - [int]$progressLabel.Width) / 2),
    ([int]$progressBar.Location.Y + ([int]$progressBar.Height - [int]$progressLabel.Height) / 2)
)
# Make the label transparent
$progressLabel.Parent = $progressBar
$form.Controls.Add($progressLabel)

# Make sure to call BringToFront() after adding both controls
$progressBar.SendToBack()  # Send progress bar to back
$progressLabel.BringToFront()  # Bring label to front

# Drop-down list (ComboBox) for server list files
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(10, 10)
$dropdown.Width = 200
$dropdown.Height = 40
$dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Add items to the dropdown (server list files)
$dropdown.Items.Add("Select Server List") # Default item
$dropdown.SelectedIndex = 0 # Select the default item initially
$textFiles = Get-ChildItem -Path "$PSScriptRoot\config" -Filter "*.txt" # Get all .txt files in the config directory
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
$zipFilePathBox.Width = $form.ClientSize.Width - 195  # Dynamic width
$zipFilePathBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
[System.Windows.Forms.AnchorStyles]::Left -bor 
[System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($zipFilePathBox)

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
# Output TextBox
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 110)
$outputBox.Width = 760
$outputBox.Height = 370
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($outputBox)

# Add signature label
$signatureLabel = New-Object System.Windows.Forms.Label
$signatureLabel.Text = "Created by Ayad"
$signatureLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
$signatureLabel.ForeColor = [System.Drawing.Color]::Gray
$signatureLabel.AutoSize = $true
$signatureLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$signatureLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$signatureLabel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 10, 5)

# Calculate initial position
$signatureLabel.Location = New-Object System.Drawing.Point(
    ($form.ClientSize.Width - $signatureLabel.Width - 10),
    ($form.ClientSize.Height - $signatureLabel.Height - 20)
)

# Add the label to form controls
$form.Controls.Add($signatureLabel)

# Version label
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "Version 1.0.1"  # Update version number as needed
$versionLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Regular)
$versionLabel.ForeColor = [System.Drawing.Color]::Gray
$versionLabel.AutoSize = $true
$versionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$versionLabel.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 5)

# Calculate initial position (bottom-left)
$versionLabel.Location = New-Object System.Drawing.Point(
    10, # Left margin
    ($form.ClientSize.Height - $versionLabel.Height - 20)  # Bottom margin
)

# Add the label to form controls
$form.Controls.Add($versionLabel)


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

# Back Out button
$backOutButton = New-Object System.Windows.Forms.Button
$backOutButton.Text = "Back Out"
$backOutButton.Width = 160
$backOutButton.Height = 40
$form.Controls.Add($backOutButton)

# Documentation link button
$docButton = New-Object System.Windows.Forms.LinkLabel
$docButton.Text = "Documentation"
$docButton.AutoSize = $true
$docButton.Font = New-Object System.Drawing.Font("Arial", 9)
$docButton.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$docButton.LinkColor = [System.Drawing.Color]::Blue
$docButton.ActiveLinkColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($docButton)

# Function to center buttons
function Set-ButtonsAlignment {
    $formWidth = [int]$form.ClientSize.Width
    $buttonWidth = [int]$pushZipButton.Width
    $spacing = [int]20
    $totalWidth = [int](($buttonWidth * 4) + ($spacing * 3))
    $startX = [int](($formWidth - $totalWidth) / 2)
    $y = [int]($outputBox.Location.Y + $outputBox.Height + 10)

    $pushZipButton.Location = New-Object System.Drawing.Point($startX, $y)
    $installMsiButton.Location = New-Object System.Drawing.Point(($startX + $buttonWidth + $spacing), $y)
    $backOutButton.Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $spacing) * 2), $y)
    $syncButton.Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $spacing) * 3), $y)

    # Position doc button below other buttons
    $docButton.Location = New-Object System.Drawing.Point(
        [int](($formWidth - $docButton.Width) / 2),
        [int]($y + $pushZipButton.Height + 10)
    )
}

# Add resize event handler
$form.Add_Resize({ 
        Set-ButtonsAlignment
        # Recenter progress label on progress bar
        $progressLabel.Location = New-Object System.Drawing.Point(
        ([int]$progressBar.Location.X + ([int]$progressBar.Width - [int]$progressLabel.Width) / 2),
        ([int]$progressBar.Location.Y + ([int]$progressBar.Height - [int]$progressLabel.Height) / 2)
        )
        $signatureLabel.Location = New-Object System.Drawing.Point(
        ($form.ClientSize.Width - $signatureLabel.Width - 10),
        ($form.ClientSize.Height - $signatureLabel.Height - 20)
        )
        $versionLabel.Location = New-Object System.Drawing.Point(10, ($form.ClientSize.Height - $versionLabel.Height - 20))
        $zipFilePathBox.Width = $form.ClientSize.Width - 195
        $form.Refresh()
    })

# Variable to store credentials
$global:cred = $null

# Function to get credentials
function Get-GlobalCredential {
    param(
        [string]$TestServer
    )
    
    if (-not $global:cred) {
        $global:cred = Get-Credential -Message "Enter credentials for remote servers"
        
        # Test credentials against first available server
        if ($global:cred -and $TestServer) {
            if (-not (Test-Credential -Credential $global:cred -TestServer $TestServer)) {
                return $null
            }
        }
    }
    return $global:cred
}

# Test credentials is valid.
function Test-Credential {
    param(
        [PSCredential]$Credential,
        [string]$TestServer
    )
    
    try {
        Write-LogMessage "Testing credentials..."
        $session = New-PSSession -ComputerName $TestServer -Credential $Credential -ErrorAction Stop
        if ($session) {
            Remove-PSSession -Session $session
            Write-LogMessage "Credential validation successful"
            return $true
        }
    }
    catch {
        Write-LogMessage "Credential validation failed: $($_.Exception.Message)" -IsError
        $global:cred = $null  # Clear invalid credentials
        return $false
    }
    return $false
}
# Function to write to the log file and output box
function Write-LogMessage {
    param(
        [string]$Message,
        [switch]$IsError,
        [int]$ProgressValue = -1
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Add spacing for new tasks
    if ($Message -match "^(Starting|Completed|Push completed|Installation completed)") {
        $LogEntry = "`n[$Timestamp] ----------------------------------------`n[$Timestamp] INFO: $Message"
    }
    else {
        $LogEntry = "[$Timestamp] INFO: $Message"
    }
    
    if ($IsError) {
        $LogEntry = "[$Timestamp] ERROR: $Message"
        Write-Error $LogEntry
        $outputBox.SelectionColor = [System.Drawing.Color]::Red
    }
    else {
        Write-Host $LogEntry
        $outputBox.SelectionColor = [System.Drawing.Color]::Black
    }

    if ($ProgressValue -ge 0) {
        $progressBar.Visible = $true
        $progressBar.Value = $ProgressValue
        $progressLabel.Text = "$ProgressValue%"
        $progressLabel.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()
    }
    elseif ($ProgressValue -eq 100) {
        $progressBar.Visible = $false
        $progressBar.Value = 0
        $progressLabel.Visible = $false
    }

    Add-Content -Path $LogFile -Value $LogEntry
    $outputBox.AppendText($LogEntry + [Environment]::NewLine)
    $outputBox.ScrollToCaret()
}

# click event for documentation
$docButton.Add_Click({
        $docPath = Join-Path $PSScriptRoot "docs\index.html"
        if (Test-Path $docPath) {
            Start-Process $docPath
        }
        else {
            Write-LogMessage "Documentation file not found at: $docPath" -IsError
        }
    })

$syncButton.Add_Click({
        $selectedItem = $dropdown.SelectedItem
        if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
            Write-LogMessage "Please select a server list." -IsError
            return
        }

        # Define remote path here
        $remotePath = "C:\Releases"  # Path on remote server
        if (-not $remotePath) {
            Write-LogMessage "Remote server path is not defined" -IsError
            return
        }

        $cred = Get-GlobalCredential

        # Create local releases directory if it doesn't exist
        $localPath = Join-Path $PSScriptRoot "releases"
        if (-not (Test-Path $localPath)) {
            New-Item -ItemType Directory -Path $localPath | Out-Null
        }

        try {
            Write-LogMessage "Starting sync from $remotePath..." -ProgressValue 0
        
            $job = Start-Job -ScriptBlock {
                param ($remotePath, $localPath, [PSCredential]$cred)
            
                try {
                    # Create session with the remote server
                    $session = New-PSSession -ComputerName $TestServer -Credential $cred
                    if (-not $session) {
                        throw "Failed to create session"
                    }

                    # Get remote folders using session
                    $remoteFolders = Invoke-Command -Session $session -ScriptBlock {
                        param($path)
                        Get-ChildItem -Path $path -Directory
                    } -ArgumentList $remotePath

                    $totalFolders = $remoteFolders.Count
                    $currentFolder = 0
                    $totalProgress = 0

                    foreach ($folder in $remoteFolders) {
                        $currentFolder++
                    
                        # Calculate base progress for folder (0-50%)
                        $folderProgress = [math]::Floor(($currentFolder / $totalFolders) * 50)
                        Write-Output "###PROGRESS###:$folderProgress"
                        Write-Output "###STATUS###:Processing folder ($currentFolder/$totalFolders): $($folder.Name)"
                    
                        $targetPath = Join-Path $localPath $folder.Name
                    
                        if (-not (Test-Path $targetPath)) {
                            New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                            Write-Output "###STATUS###:Created new folder: $($folder.Name)"
                        }

                        # Get remote files using session with error handling
                        $remoteFiles = Invoke-Command -Session $session -ScriptBlock {
                            param($folderPath)
                            if (Test-Path $folderPath) {
                                Get-ChildItem -Path $folderPath -Recurse -File
                            }
                            else {
                                Write-Output "###ERROR###:Remote folder not found: $folderPath"
                                return $null
                            }
                        } -ArgumentList (Join-Path $remotePath $folder.Name)

                        # Get total files for this folder
                        $totalFiles = ($remoteFiles | Measure-Object).Count
                        $currentFile = 0
                    
                        foreach ($remoteFile in $remoteFiles) {
                            $currentFile++
                            # Calculate file progress (50-100%)
                            $fileProgress = 50 + [math]::Floor(($currentFile / $totalFiles) * 50)
                            Write-Output "###PROGRESS###:$fileProgress"
                        
                            $relativePath = $remoteFile.FullName.Substring($remotePath.Length + 1)
                            $localFilePath = Join-Path $localPath $relativePath
                        
                            $shouldCopy = $false
                        
                            if (-not (Test-Path $localFilePath)) {
                                $shouldCopy = $true
                            }
                            else {
                                $localFile = Get-Item $localFilePath
                                if ($remoteFile.LastWriteTime -gt $localFile.LastWriteTime) {
                                    $shouldCopy = $true
                                    Write-Output "###STATUS###:Updated file found: $relativePath"
                                }
                            }

                            if ($shouldCopy) {
                                $localFolder = Split-Path $localFilePath -Parent
                                if (-not (Test-Path $localFolder)) {
                                    New-Item -ItemType Directory -Path $localFolder -Force | Out-Null
                                }
                                Copy-Item -Path $remoteFile.FullName -Destination $localFilePath -FromSession $session -Force
                            }
                        }
                        Write-Output "###COMPLETE###:$($folder.Name)"
                    }

                    # Cleanup session
                    Remove-PSSession -Session $session

                }
                catch {
                    Write-Output "###ERROR###:$($_.Exception.Message)"
                    if ($session) {
                        Remove-PSSession -Session $session
                    }
                }
            } -ArgumentList $remotePath, $localPath, $cred

            # Monitor job progress
            while ($job.State -eq 'Running') {
                $jobOutput = Receive-Job -Job $job
                foreach ($line in $jobOutput) {
                    if ($line.StartsWith('###PROGRESS###:')) {
                        $progress = [int]($line.Split(':')[1])
                        # Silently update progress bar
                        $progressBar.Value = $progress
                        $progressLabel.Text = "$progress%"
                        $progressLabel.Visible = $true
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    elseif ($line.StartsWith('###STATUS###:')) {
                        $status = $line.Split(':')[1]
                        Write-LogMessage $status
                    }
                    elseif ($line.StartsWith('###COMPLETE###:')) {
                        $folder = $line.Split(':')[1]
                        Write-LogMessage "Completed syncing folder: $folder"
                    }
                    elseif ($line.StartsWith('###ERROR###:')) {
                        Write-LogMessage $line.Split(':')[1] -IsError
                    }
                }
                Start-Sleep -Milliseconds 100
                [System.Windows.Forms.Application]::DoEvents()
            }

            # Process final job output - simplified
            $finalOutput = Receive-Job -Job $job
            if ($finalOutput) {
                foreach ($line in $finalOutput) {
                    if ($line.StartsWith('###STATUS###:')) {
                        $status = $line.Split(':')[1]
                        Write-LogMessage $status
                    }
                    elseif ($line.StartsWith('###COMPLETE###:')) {
                        $folder = $line.Split(':')[1]
                        Write-LogMessage "Completed syncing folder: $folder"
                    }
                    elseif ($line.StartsWith('###ERROR###:')) {
                        Write-LogMessage $line.Split(':')[1] -IsError
                    }
                }
            }
            Remove-Job -Job $job

            Write-LogMessage "Sync operation completed." -ProgressValue 100
        }
        catch {
            Write-LogMessage "Error during sync: $($_.Exception.Message)" -IsError
            $progressBar.Visible = $false
            if ($job) {
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
            Get-PSSession | Where-Object { $_.State -eq 'Broken' } | Remove-PSSession
        }
    })

# Create temp directory if it doesn't exist
$tempPath = Join-Path $PSScriptRoot "temp"
if (-not (Test-Path $tempPath)) {
    New-Item -ItemType Directory -Path $tempPath | Out-Null
}

# Update Push ZIP button click event
$pushZipButton.Add_Click({
        # Cleanup any existing jobs and sessions first
        Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
        Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue

        $selectedFolder = $zipFilePathBox.Text
        if (-not $selectedFolder) {
            Write-LogMessage "Please select a folder to zip and push." -IsError
            return
        }

        $selectedItem = $dropdown.SelectedItem
        if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
            Write-LogMessage "Please select a server list." -IsError
            return
        }

        try {
            # Clear previous progress
            $progressBar.Value = 0
            $progressLabel.Text = "0%"
            # Create ZIP file
            $folderName = Split-Path $selectedFolder -Leaf
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $global:currentZipFileName = "${folderName}_${timestamp}.zip"
            $zipFilePath = Join-Path $tempPath $global:currentZipFileName

            Write-LogMessage "Creating ZIP file..." -ProgressValue 0
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::CreateFromDirectory($selectedFolder, $zipFilePath)
            Write-LogMessage "ZIP file created: $global:currentZipFileName" -ProgressValue 20

            # Get server list and credentials
            $serverListFile = ".\config\$selectedItem.txt"
            $servers = Get-Content $serverListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $cred = Get-GlobalCredential -TestServer $servers[0]
            if (-not $cred) {
                Remove-Item -Path $zipFilePath -Force
                return
            }

            Write-LogMessage "Starting parallel file push..." -ProgressValue 25

            # Create all copy jobs simultaneously
            $jobs = foreach ($server in $servers) {
                Start-Job -ScriptBlock {
                    param ($zipFilePath, $server, [PSCredential]$cred)
                    try {
                        Write-Output "###STATUS###:Starting copy to $server"
                        $destinationPath = "\\${server}\C$\temp"
                        
                        # Check if this is the local machine
                        if ($server -eq $env:COMPUTERNAME -or $server -eq "localhost" -or $server -eq ".") {
                            # Local copy
                            if (-not (Test-Path "C:\temp")) {
                                New-Item -Path "C:\temp" -ItemType Directory -Force | Out-Null
                            }
                            Copy-Item -Path $zipFilePath -Destination "C:\temp" -Force
                            
                            # Verify local copy
                            $destFile = Join-Path "C:\temp" (Split-Path $zipFilePath -Leaf)
                            if (Test-Path -Path $destFile) {
                                Write-Output "###SUCCESS###:$server"
                            }
                            else {
                                throw "Local file copy failed"
                            }
                        }
                        else {
                            # Remote copy - use existing code
                            if (-not (Test-Path -Path $destinationPath -Credential $cred)) {
                                Write-Output "###STATUS###:Creating temp directory on $server"
                                New-Item -Path $destinationPath -ItemType Directory -Credential $cred -Force
                            }
                            
                            Copy-Item -Path $zipFilePath -Destination $destinationPath -Credential $cred -Force
                            
                            # Verify remote copy
                            $destFile = Join-Path $destinationPath (Split-Path $zipFilePath -Leaf)
                            if (Test-Path -Path $destFile -Credential $cred) {
                                Write-Output "###SUCCESS###:$server"
                            }
                            else {
                                throw "Remote file copy failed"
                            }
                        }
                    }
                    catch {
                        Write-Output "###ERROR###:${server}:$($_.Exception.Message)"
                    }
                } -ArgumentList $zipFilePath, $server, $cred
            }
        
            $totalJobs = $jobs.Count
            $lastProgress = 25
            $processedJobs = @()

            # Monitor jobs until all are complete
            while ($jobs | Where-Object { $_.State -eq 'Running' -or ($_.State -eq 'Completed' -and $_ -notin $processedJobs) }) {
                $completed = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
                $progress = 25 + [math]::Floor(($completed / $totalJobs) * 75)
    
                if ($progress -ne $lastProgress) {
                    Write-LogMessage "Copying to servers... ($completed/$totalJobs complete)" -ProgressValue $progress
                    $lastProgress = $progress
                }

                # Process newly completed jobs
                foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                    $output = Receive-Job -Job $job
                    foreach ($line in $output) {
                        if ($line.StartsWith('###SUCCESS###:')) {
                            $server = $line.Split(':')[1]
                            Write-LogMessage "Push completed: $server"
                        }
                        elseif ($line.StartsWith('###ERROR###:')) {
                            $parts = $line.Split(':')
                            Write-LogMessage "Push failed: $($parts[1]) - $($parts[2])" -IsError
                        }
                        elseif ($line.StartsWith('###STATUS###:')) {
                            $status = $line.Split(':')[1]
                            Write-LogMessage $status
                        }
                    }
                    $processedJobs += $job
                    Remove-Job -Job $job
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 100
            }

            # Final check for any remaining jobs
            foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                Receive-Job -Job $job | Out-Null
                Remove-Job -Job $job
            }

            Write-LogMessage "Push operation completed." -ProgressValue 100
            Remove-Item -Path $zipFilePath -Force
        }
        catch {
            Write-LogMessage "Error: $($_.Exception.Message)" -IsError
            if (Test-Path $zipFilePath) {
                Remove-Item -Path $zipFilePath -Force
            }
        }
    })

# Update Install MSI button click event
$installMsiButton.Add_Click({
        # Disable the button during installation
        $installMsiButton.Enabled = $false
    
        if (-not $global:currentZipFileName) {
            Write-LogMessage "Please push a ZIP file first." -IsError
            return
        }

        $selectedItem = $dropdown.SelectedItem
        if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
            Write-LogMessage "Please select a server list." -IsError
            return
        }

        # Add confirmation dialog
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to install on the selected servers?`n`nServer List: $selectedItem`nPackage: $global:currentZipFileName",
            "Confirm Installation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            Write-LogMessage "Installation cancelled by user."
            return
        }
    
        $cred = Get-GlobalCredential

        try {
            # Clear previous progress
            $progressBar.Value = 0
            $progressLabel.Text = "0%"
        
            Write-LogMessage "Starting installation process..." -ProgressValue 0

            $serverListFile = ".\config\$selectedItem.txt"
            $servers = Get-Content $serverListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

            Write-LogMessage "Preparing for parallel installation..." -ProgressValue 10

            # Create all installation jobs simultaneously
            $jobs = foreach ($server in $servers) {
                Start-Job -ScriptBlock {
                    param ($server, [PSCredential]$cred, $zipFileName, $AppPoolName, $ServiceName)
                    try {
                        Write-Output "###STATUS###:Starting installation on $server"
                        $session = New-PSSession -ComputerName $server -Credential $cred
                    
                        Invoke-Command -Session $session -ScriptBlock {
                            param ($zipFileName, $AppPoolName, $ServiceName)
                            
                            $zipPath = "C:\temp\$zipFileName"
                            $extractPath = "C:\temp\extracted"
                        
                            Add-Type -AssemblyName System.IO.Compression.FileSystem
                            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractPath)
                            
                            $exeFile = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | Select-Object -First 1
                            if ($exeFile) {
                                # Change to the executable's directory before running it
                                $originalLocation = Get-Location
                                try {
                                    Set-Location -Path $exeFile.DirectoryName
                                    $process = Start-Process -FilePath $exeFile.FullName -ArgumentList "install" -Wait -NoNewWindow -PassThru
                        
                                    # Service check and restart section in Invoke-Command block
                                    if ($process.ExitCode -eq 0) {
                                        Write-Output "###STATUS###:Checking services on $env:COMPUTERNAME"
    
                                        # Check initial states and store service objects
                                        $initialStates = @{
                                            IIS           = @{
                                                IsRunning = $false
                                                Service   = $null
                                            }
                                            AppPool       = @{
                                                IsRunning = $false
                                                Object    = $null
                                            }
                                            CustomService = @{
                                                IsRunning = $false
                                                Service   = $null
                                            }
                                        }
    
                                        # Get service states once
                                        Import-Module WebAdministration
                                        $initialStates.IIS.Service = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
                                        $initialStates.AppPool.Object = Get-Item "IIS:\AppPools\$using:AppPoolName" -ErrorAction SilentlyContinue
                                        $initialStates.CustomService.Service = Get-Service -Name $using:ServiceName -ErrorAction SilentlyContinue

                                        # Store initial running states
                                        if ($initialStates.IIS.Service -and $initialStates.IIS.Service.Status -eq 'Running') {
                                            $initialStates.IIS.IsRunning = $true
                                            Write-Output "###STATUS###:Restarting IIS..."
                                            Restart-Service -Name 'W3SVC' -Force
                                        }
                                        else {
                                            Write-Output "###STATUS###:IIS is already stopped, leaving as is"
                                        }

                                        if ($initialStates.AppPool.Object -and $initialStates.AppPool.Object.State -eq 'Started') {
                                            $initialStates.AppPool.IsRunning = $true
                                            Write-Output "###STATUS###:Restarting Application Pool..."
                                            Restart-WebAppPool -Name $using:AppPoolName
                                        }
                                        else {
                                            Write-Output "###STATUS###:AppPool is already stopped, leaving as is"
                                        }

                                        if ($initialStates.CustomService.Service -and $initialStates.CustomService.Service.Status -eq 'Running') {
                                            $initialStates.CustomService.IsRunning = $true
                                            Write-Output "###STATUS###:Restarting Custom Service..."
                                            Restart-Service -Name $using:ServiceName -Force
                                        }
                                        else {
                                            Write-Output "###STATUS###:Custom Service is already stopped, leaving as is"
                                        }

                                        # Wait for services to stabilize
                                        Write-Output "###STATUS###:Waiting for services to stabilize..."
                                        Start-Sleep -Seconds 5

                                        # Verify only services that were running
                                        Write-Output "###STATUS###:Verifying services status..."
                                        $success = $true

                                        # Verify each service that was running
                                        if ($initialStates.IIS.IsRunning) {
                                            $iisService = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
                                            Write-Output "###STATUS###:IIS Status: $($iisService.Status)"
                                            if (-not $iisService -or $iisService.Status -ne 'Running') {
                                                $success = $false
                                                Write-Output "###ERROR###:$env:COMPUTERNAME:IIS failed to start"
                                            }
                                        }

                                        if ($initialStates.AppPool.IsRunning) {
                                            $appPool = Get-Item "IIS:\AppPools\$using:AppPoolName" -ErrorAction SilentlyContinue
                                            Write-Output "###STATUS###:AppPool Status: $($appPool.State)"
                                            if (-not $appPool -or $appPool.State -ne 'Started') {
                                                $success = $false
                                                Write-Output "###ERROR###:$env:COMPUTERNAME:AppPool failed to start"
                                            }
                                        }

                                        if ($initialStates.CustomService.IsRunning) {
                                            $customService = Get-Service -Name $using:ServiceName -ErrorAction SilentlyContinue
                                            Write-Output "###STATUS###:Custom Service Status: $($customService.Status)"
                                            if (-not $customService -or $customService.Status -ne 'Running') {
                                                $success = $false
                                                Write-Output "###ERROR###:$env:COMPUTERNAME:Custom Service failed to start"
                                            }
                                        }

                                        if ($success) {
                                            Write-Output "###SUCCESS###:$env:COMPUTERNAME"
                                        }
                                        else {
                                            throw "One or more previously running services failed to start properly"
                                        }
                                    }
                                    else {
                                        throw "Installation failed with exit code: $($process.ExitCode)"
                                    }
                                }
                                finally {
                                    Set-Location -Path $originalLocation
                                }
                            }
                            else {
                                throw "No EXE file found in extracted contents"
                            }
                        } -ArgumentList $global:currentZipFileName, $AppPoolName, $ServiceName
                        Remove-PSSession -Session $session
                    }
                    catch {
                        Write-Output "###ERROR###:${server}:$($_.Exception.Message)"
                    }
                } -ArgumentList $server, $cred, $global:currentZipFileName, $AppPoolName, $ServiceName
            }

            Write-LogMessage "Starting parallel installation on servers..." -ProgressValue 20
        
            $totalJobs = $jobs.Count
            $lastProgress = 20
            $processedJobs = @()

            # Monitor jobs until all are complete
            while ($jobs | Where-Object { $_.State -eq 'Running' -or ($_.State -eq 'Completed' -and $_ -notin $processedJobs) }) {
                $completed = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
                $progress = 20 + [math]::Floor(($completed / $totalJobs) * 80)

                if ($progress -ne $lastProgress) {
                    Write-LogMessage "Installing on servers... ($completed/$totalJobs complete)" -ProgressValue $progress
                    $lastProgress = $progress
                }

                # Process completed jobs
                foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                    $output = Receive-Job -Job $job
                    foreach ($line in $output) {
                        if ($line.StartsWith('###SUCCESS###:')) {
                            $server = $line.Split(':')[1]
                            Write-LogMessage "Installation completed: $server"
                        }
                        elseif ($line.StartsWith('###ERROR###:')) {
                            $parts = $line.Split(':')
                            Write-LogMessage "Installation failed: $($parts[1]) - $($parts[2])" -IsError
                        }
                        elseif ($line.StartsWith('###STATUS###:')) {
                            $status = $line.Split(':')[1]
                            Write-LogMessage $status
                        }
                    }
                    $processedJobs += $job
                    Remove-Job -Job $job
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 100
            }

            # Final check for any remaining jobs
            foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                Receive-Job -Job $job | Out-Null
                Remove-Job -Job $job
            }
            #Clean up temp files
            Write-LogMessage "Cleaning up temporary files..." -ProgressValue 95
            foreach ($server in $servers) {
                try {
                    Remove-TempFiles -server $server
                }
                catch {
                    Write-LogMessage "Warning: Cleanup failed on $server" -IsError
                }
            }
            Write-LogMessage "Installation completed on all servers." -ProgressValue 100
        }
        catch {
            Write-LogMessage "Error during installation: $($_.Exception.Message)" -IsError
            $progressBar.Visible = $false 
        }
        finally {
            # Cleanup any remaining jobs
            $jobs | Where-Object { $_ } | Remove-Job -Force -ErrorAction SilentlyContinue
            Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
            $installMsiButton.Enabled = $true
            $progressBar.Value = 0
            $progressLabel.Text = "0%"
        }
    })

# Add Back Out button click event
$backOutButton.Add_Click({
        if (-not $global:currentZipFileName) {
            Write-LogMessage "Please push a ZIP file first." -IsError
            return
        }

        $selectedItem = $dropdown.SelectedItem
        if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
            Write-LogMessage "Please select a server list." -IsError
            return
        }

        # Add confirmation dialog
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to Backout on the selected servers?`n`nServer List: $selectedItem`nPackage: $global:currentZipFileName",
            "Confirm Installation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            Write-LogMessage "Installation cancelled by user."
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
                    Update-Progress -Current $currentServer -Total $totalServers -Operation "Executing back out on servers"

                    try {
                        $job = Start-Job -ScriptBlock {
                            param ($server, [PSCredential]$cred, $zipFileName)
                            $session = New-PSSession -ComputerName $server -Credential $cred
                            Invoke-Command -Session $session -ScriptBlock {
                                param ($zipFileName)
                                Write-Host "###PROGRESS###:Preparing back out on $env:COMPUTERNAME"
                                $zipPath = "C:\temp\$zipFileName"
                                $extractPath = "C:\temp\extracted"

                                Write-Host "###PROGRESS###:Extracting ZIP file"
                                Add-Type -AssemblyName System.IO.Compression.FileSystem
                                [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractPath)
                                $exeFile = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | Select-Object -First 1
                                if ($exeFile) {
                                    Write-Host "###PROGRESS###:Found executable: $($exeFile.Name)"
                                    try {
                                        Write-Host "###PROGRESS###:Starting back out"
                                        $process = Start-Process -FilePath $exeFile.FullName -ArgumentList "uninstall" -Wait -NoNewWindow -PassThru

                                        if ($process.ExitCode -eq 0) {
                                            Write-Host "###PROGRESS###:Back out completed successfully (Exit Code: 0)"
                                        }
                                        else {
                                            throw "Back out failed with exit code: $($process.ExitCode)"
                                        }
                                        Write-Host "###PROGRESS###:Restarting services"
                                        Restart-Service -Name 'W3SVC'
                                        Import-Module WebAdministration
                                        Restart-WebAppPool -Name $using:AppPoolName
                                        Restart-Service -Name $using:ServiceName
                                        Write-Host "###PROGRESS###:Services restarted successfully"
                                    }
                                    catch {
                                        throw "Back out failed: $_"
                                    }
                                }
                                else {
                                    throw "No EXE file found in extracted contents"
                                }
                            } -ArgumentList $zipFileName
                            Remove-PSSession -Session $session
                        } -ArgumentList $server, $cred, $global:currentZipFileName

                        $job | Wait-Job
                        $jobOutput = Receive-Job -Job $job
                    
                        foreach ($line in $jobOutput) {
                            if ($line.StartsWith('###PROGRESS###:')) {
                                $status = $line.Split(':')[1]
                                Write-LogMessage $status -ProgressValue ([math]::Floor(($currentServer / $totalServers) * 100))
                                [System.Windows.Forms.Application]::DoEvents()
                            }
                            else {
                                $outputBox.AppendText($line + [Environment]::NewLine)
                            }
                        }
                    }
                    catch {
                        Write-LogMessage "Error executing commands on ${server}: $($_.Exception.Message)" -IsError
                    }
                }
            }
            # Clean up temp files
            Write-LogMessage "Cleaning up temporary files..." -ProgressValue 95
            foreach ($server in $servers) {
                try {
                    Remove-TempFiles -server $server
                }
                catch {
                    Write-LogMessage "Warning: Cleanup failed on $server" -IsError
                }
            }
            Write-LogMessage "Back out completed." -ProgressValue 100
        }
        catch {
            Write-LogMessage "Error reading server list file '$serverListFile': $($_.Exception.Message)" -IsError
            $progressBar.Visible = $false
        }
    })

# Center buttons initially
Set-ButtonsAlignment

# Show the form
$form.ShowDialog()
