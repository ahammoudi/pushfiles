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
    $AppPoolName = $config.AppPoolName
    $AppPoolName2 = $config.AppPoolName2       # Add second app pool
    $AppPoolName3 = $config.AppPoolName3       # Add third app pool
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
$form.Height = 700  # Increase form height to show all buttons (was 600)
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)


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
$browseButton.Height = 30
$form.Controls.Add($browseButton)

# TextBox to display the selected ZIP file path
$zipFilePathBox = New-Object System.Windows.Forms.TextBox
$zipFilePathBox.Location = New-Object System.Drawing.Point(175, 60)
$zipFilePathBox.Width = $form.ClientSize.Width - 195  # Dynamic width
$zipFilePathBox.Height = 30  # Increased height
$zipFilePathBox.Multiline = $true  # Enable multiline for increased height
$zipFilePathBox.ScrollBars = "Vertical"  # Add scrollbars for longer paths
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
    else {
        $folderDialog.SelectedPath = [Environment]::GetFolderPath('Desktop')
    }

    # Set the root folder to Desktop for better navigation
    $folderDialog.RootFolder = [System.Environment+SpecialFolder]::Desktop
    
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $zipFilePathBox.Text = $folderDialog.SelectedPath
    }
})

# Output TextBox
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 110)
$outputBox.Width = 760
$outputBox.Height = 300  # Reduce height to make room for two rows of buttons
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

# Rollback button
$rollbackButton = New-Object System.Windows.Forms.Button
$rollbackButton.Text = "Rollback"
$rollbackButton.Width = 160
$rollbackButton.Height = 40
$form.Controls.Add($rollbackButton)

# Documentation link button
$docButton = New-Object System.Windows.Forms.LinkLabel
$docButton.Text = "Documentation"
$docButton.AutoSize = $true
$docButton.Font = New-Object System.Drawing.Font("Arial", 9)
$docButton.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$docButton.LinkColor = [System.Drawing.Color]::Blue
$docButton.ActiveLinkColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($docButton)

# Replace the Restart Services button with separate Stop and Start buttons
$stopServicesButton = New-Object System.Windows.Forms.Button
$stopServicesButton.Text = "Stop Services"
$stopServicesButton.Width = 160
$stopServicesButton.Height = 40
$form.Controls.Add($stopServicesButton)

$startServicesButton = New-Object System.Windows.Forms.Button
$startServicesButton.Text = "Start Services"
$startServicesButton.Width = 160
$startServicesButton.Height = 40
$form.Controls.Add($startServicesButton)

# Remove the old restart button
$form.Controls.Remove($restartServicesButton)

# Rename the Back Out button to Check Version button
$checkVersionButton = New-Object System.Windows.Forms.Button
$checkVersionButton.Text = "Check Version"
$checkVersionButton.Width = 160
$checkVersionButton.Height = 40
$form.Controls.Add($checkVersionButton)

# Remove the old backOutButton if it exists
$form.Controls.Remove($backOutButton)

# Update the Set-ButtonsAlignment function for two rows of buttons
function Set-ButtonsAlignment {
    $formWidth = [int]$form.ClientSize.Width
    $buttonWidth = [int]$pushZipButton.Width
    $spacing = [int]20
    
    # First row - 3 buttons
    $totalWidthRow1 = [int](($buttonWidth * 3) + ($spacing * 2))
    $startXRow1 = [int](($formWidth - $totalWidthRow1) / 2)
    $yRow1 = [int]($outputBox.Location.Y + $outputBox.Height + 10)

    # Second row - 4 buttons
    $totalWidthRow2 = [int](($buttonWidth * 4) + ($spacing * 3))
    $startXRow2 = [int](($formWidth - $totalWidthRow2) / 2)
    $yRow2 = [int]($yRow1 + $pushZipButton.Height + 10)

    # Position buttons in first row: Sync Release, Push Zip File, Stop Services
    $syncButton.Location = New-Object System.Drawing.Point($startXRow1, $yRow1)
    $pushZipButton.Location = New-Object System.Drawing.Point(($startXRow1 + $buttonWidth + $spacing), $yRow1)
    $stopServicesButton.Location = New-Object System.Drawing.Point(($startXRow1 + ($buttonWidth + $spacing) * 2), $yRow1)

    # Position buttons in second row: Install MSI, Start Services, Check Version, Rollback
    $installMsiButton.Location = New-Object System.Drawing.Point($startXRow2, $yRow2)
    $startServicesButton.Location = New-Object System.Drawing.Point(($startXRow2 + $buttonWidth + $spacing), $yRow2)
    $checkVersionButton.Location = New-Object System.Drawing.Point(($startXRow2 + ($buttonWidth + $spacing) * 2), $yRow2)
    $rollbackButton.Location = New-Object System.Drawing.Point(($startXRow2 + ($buttonWidth + $spacing) * 3), $yRow2)

    # Position doc button below all buttons
    $docButton.Location = New-Object System.Drawing.Point(
        [int](($formWidth - $docButton.Width) / 2),
        [int]($yRow2 + $installMsiButton.Height + 10)
    )

    # Ensure all buttons are on the form
    $form.Controls.Add($syncButton)
    $form.Controls.Add($pushZipButton)
    $form.Controls.Add($stopServicesButton)
    $form.Controls.Add($installMsiButton)
    $form.Controls.Add($startServicesButton)
    $form.Controls.Add($checkVersionButton)
    $form.Controls.Add($rollbackButton)
    $form.Controls.Add($docButton)
}

# Function to get the status file path for a server
function Get-ServiceStatusFilePath {
    param(
        [string]$ServerList,
        [string]$Server
    )
    
    $statusDir = Join-Path $PSScriptRoot "status"
    if (-not (Test-Path $statusDir)) {
        New-Item -ItemType Directory -Path $statusDir | Out-Null
    }
    
    return Join-Path $statusDir "${ServerList}_${Server}_status.xml"
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
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [switch]$IsError,
        [ValidateRange(-1, 100)]
        [int]$ProgressValue = -1
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Add spacing for new tasks with improved pattern matching
    $LogEntry = if ($Message -match "^(Starting|Completed|Push|Installation|Sync)\s") {
        "`n[$Timestamp] ----------------------------------------`n[$Timestamp] $(if ($IsError) { 'ERROR:' } else { 'INFO:' }) $Message"
    }
    else {
        "[$Timestamp] $(if ($IsError) { 'ERROR:' } else { 'INFO:' }) $Message"
    }
    
    try {
        # Handle error messages
        if ($IsError) {
            Write-Error $Message
            $outputBox.SelectionStart = $outputBox.TextLength
            $outputBox.SelectionLength = 0
            $outputBox.SelectionColor = [System.Drawing.Color]::Red
        }
        else {
            Write-Host $LogEntry
            $outputBox.SelectionStart = $outputBox.TextLength
            $outputBox.SelectionLength = 0
            $outputBox.SelectionColor = [System.Drawing.Color]::Black
        }

        # Handle progress updates
        if ($ProgressValue -ge 0) {
            $progressBar.Visible = $true
            $progressBar.Value = [Math]::Min([Math]::Max($ProgressValue, 0), 100)
            $progressLabel.Text = "$($progressBar.Value)%"
            $progressLabel.Visible = $true
        }
        elseif ($ProgressValue -eq 100) {
            $progressBar.Visible = $false
            $progressBar.Value = 0
            $progressLabel.Visible = $false
        }

        # Write to log file and update output box
        Add-Content -Path $LogFile -Value $LogEntry
        $outputBox.AppendText($LogEntry + [Environment]::NewLine)
        $outputBox.ScrollToCaret()

        # Force UI update
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {
        Write-Error "Failed to write log message: $($_.Exception.Message)"
    }
}

# Update form closing event handler
$form.Add_FormClosing({
    param($formSender, $e)
    
    # Check if this is a user-initiated close (not a programmatic one)
    if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to exit?",
            "Confirm Exit",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            # Cancel the close operation
            $e.Cancel = $true
            return
        }
    }

    try {
        # Rotate log file
        Write-LogMessage "Performing log rotation..."
        Move-Logs -LogFile $LogFile -MaxSize 10MB -MaxFiles 10

        # Clean up remote servers
        Write-LogMessage "Cleaning up remote servers..."
        $serverFiles = Get-ChildItem -Path "$PSScriptRoot\config" -Filter "*.txt"
        foreach ($file in $serverFiles) {
            $servers = Get-Content $file.FullPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            foreach ($server in $servers) {
                try {
                    Remove-TempFiles -server $server
                    Write-LogMessage "Cleaned up server: $server"
                }
                catch {
                    Write-LogMessage "Failed to clean up server $server : $($_.Exception.Message)" -IsError
                }
            }
        }
    }
    catch {
        Write-LogMessage "Error during cleanup: $($_.Exception.Message)" -IsError
    }
    finally {
        # Cleanup operations when closing is confirmed
        if ($global:cred) { $global:cred = $null }
        Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
        Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
    }
})

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
                    param ($server, [PSCredential]$cred, $zipFileName, $AppPoolName1, $ServiceName)
                    try {
                        Write-Output "###STATUS###:Starting installation on $server"
                        $session = New-PSSession -ComputerName $server -Credential $cred
                    
                        Invoke-Command -Session $session -ScriptBlock {
                            param ($zipFileName, $AppPoolName1, $ServiceName)
                            
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
                        
                                    if ($process.ExitCode -eq 0) {
                                        Write-Output "###SUCCESS###:$env:COMPUTERNAME"
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
                        } -ArgumentList $ZipFileName, $AppPoolName, $AppPoolName2, $AppPoolName3, $ServiceName
                        Remove-PSSession -Session $session
                    }
                    catch {
                        Write-Output "###ERROR###:${server}:$($_.Exception.Message)"
                    }
                } -ArgumentList $server, $cred, $global:currentZipFileName, $AppPoolName, $AppPoolName2, $AppPoolName3, $ServiceName
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

# Add the Stop Services button click event
$stopServicesButton.Add_Click({
    # Disable the button during stop operation
    $stopServicesButton.Enabled = $false
    
    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        Write-LogMessage "Please select a server list." -IsError
        $stopServicesButton.Enabled = $true
        return
    }

    # Add confirmation dialog
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to stop services on the selected servers?`n`nServer List: $selectedItem",
        "Confirm Service Stop",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::No) {
        Write-LogMessage "Service stop cancelled by user."
        $stopServicesButton.Enabled = $true
        return
    }
    
    $cred = Get-GlobalCredential

    try {
        # Clear previous progress
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
        $progressBar.Visible = $true
        $progressLabel.Visible = $true
    
        Write-LogMessage "Starting service stop process..." -ProgressValue 0

        $serverListFile = ".\config\$selectedItem.txt"
        $servers = Get-Content $serverListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        Write-LogMessage "Preparing for parallel service stop..." -ProgressValue 10

        # Create all stop jobs simultaneously
        $jobs = foreach ($server in $servers) {
            Start-Job -ScriptBlock {
                param ($server, [PSCredential]$cred, $AppPoolName, $AppPoolName2, $AppPoolName3, $ServiceName, $selectedItem, $statusFilePath)
                try {
                    Write-Output "###STATUS###:Checking service status on $server"
                    $session = New-PSSession -ComputerName $server -Credential $cred
                
                    Invoke-Command -Session $session -ScriptBlock {
                        param ($AppPoolName, $AppPoolName2, $AppPoolName3, $ServiceName, $statusFilePath)
                        
                        # Check states and store service objects
                        $serviceStates = @{
                            IIS = @{
                                Name = "W3SVC"
                                IsRunning = $false
                            }
                            AppPools = @(
                                @{
                                    Name = $AppPoolName
                                    IsRunning = $false
                                },
                                @{
                                    Name = $AppPoolName2
                                    IsRunning = $false
                                },
                                @{
                                    Name = $AppPoolName3
                                    IsRunning = $false
                                }
                            )
                            CustomService = @{
                                Name = $ServiceName
                                IsRunning = $false
                            }
                        }
                        
                        # Get service states
                        Import-Module WebAdministration
                        
                        # Check IIS status
                        $iisService = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
                        if ($iisService -and $iisService.Status -eq 'Running') {
                            $serviceStates.IIS.IsRunning = $true
                            Write-Output "###STATUS###:IIS (W3SVC) is running, will be stopped"
                        }

                        # Check app pools
                        foreach ($i in 0..2) {
                            $appPoolName = $serviceStates.AppPools[$i].Name
                            if ($appPoolName) {
                                $appPool = Get-Item "IIS:\AppPools\$appPoolName" -ErrorAction SilentlyContinue
                                if ($appPool -and $appPool.State -eq 'Started') {
                                    $serviceStates.AppPools[$i].IsRunning = $true
                                    Write-Output "###STATUS###:Application Pool ($appPoolName) is running, will be stopped"
                                }
                            }
                        }

                        # Check custom service
                        $customService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                        if ($customService -and $customService.Status -eq 'Running') {
                            $serviceStates.CustomService.IsRunning = $true
                            Write-Output "###STATUS###:Custom Service ($ServiceName) is running, will be stopped"
                        }

                        # Stop services in reverse order
                        # 1. Stop custom service first
                        if ($serviceStates.CustomService.IsRunning) {
                            Write-Output "###STATUS###:Stopping Custom Service ($ServiceName)..."
                            Stop-Service -Name $ServiceName -Force
                        }

                        # 2. Stop app pools
                        foreach ($appPool in $serviceStates.AppPools) {
                            if ($appPool.IsRunning) {
                                Write-Output "###STATUS###:Stopping Application Pool ($($appPool.Name))..."
                                Stop-WebAppPool -Name $appPool.Name
                            }
                        }

                        # 3. Stop IIS last
                        if ($serviceStates.IIS.IsRunning) {
                            Write-Output "###STATUS###:Stopping IIS (W3SVC)..."
                            Stop-Service -Name 'W3SVC' -Force
                        }

                        # Verify all stopped correctly
                        $allStopped = $true
                        
                        # Verify custom service
                        if ($serviceStates.CustomService.IsRunning) {
                            $customService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                            if ($customService -and $customService.Status -ne 'Stopped') {
                                $allStopped = $false
                                Write-Output "###ERROR###:$env:COMPUTERNAME:Custom Service ($ServiceName) failed to stop"
                            }
                        }
                        
                        # Verify app pools
                        foreach ($appPool in $serviceStates.AppPools) {
                            if ($appPool.IsRunning) {
                                $poolObject = Get-Item "IIS:\AppPools\$($appPool.Name)" -ErrorAction SilentlyContinue
                                if ($poolObject -and $poolObject.State -ne 'Stopped') {
                                    $allStopped = $false
                                    Write-Output "###ERROR###:$env:COMPUTERNAME:Application Pool ($($appPool.Name)) failed to stop"
                                }
                            }
                        }
                        
                        # Verify IIS
                        if ($serviceStates.IIS.IsRunning) {
                            $iisService = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
                            if ($iisService -and $iisService.Status -ne 'Stopped') {
                                $allStopped = $false
                                Write-Output "###ERROR###:$env:COMPUTERNAME:IIS (W3SVC) failed to stop"
                            }
                        }

                        # Return status and save to file
                        if ($allStopped) {
                            # Convert to XML and save
                            $serviceStates | Export-Clixml -Path "C:\temp\service_state.xml" -Force
                            Write-Output "###STATEFILE###:C:\temp\service_state.xml"
                            Write-Output "###SUCCESS###:$env:COMPUTERNAME"
                        }
                        else {
                            throw "One or more services failed to stop properly"
                        }
                    } -ArgumentList $AppPoolName, $AppPoolName2, $AppPoolName3, $ServiceName, $statusFilePath
                    
                    # Copy the state file from the remote server to local
                    $remoteStateFile = Invoke-Command -Session $session -ScriptBlock {
                        Get-Item -Path "C:\temp\service_state.xml" -ErrorAction SilentlyContinue
                    }
                    
                    if ($remoteStateFile) {
                        Copy-Item -Path "C:\temp\service_state.xml" -Destination $statusFilePath -FromSession $session -Force
                        Write-Output "###LOCALSTATEFILE###:$statusFilePath"
                    }
                    
                    Remove-PSSession -Session $session
                }
                catch {
                    Write-Output "###ERROR###:${server}:$($_.Exception.Message)"
                }
            } -ArgumentList $server, $cred, $AppPoolName, $AppPoolName2, $AppPoolName3, $ServiceName, $selectedItem, (Get-ServiceStatusFilePath -ServerList $selectedItem -Server $server)
        }

        Write-LogMessage "Starting parallel service stop on servers..." -ProgressValue 20
    
        $totalJobs = $jobs.Count
        $lastProgress = 20
        $processedJobs = @()

        # Monitor jobs until all are complete
        while ($jobs | Where-Object { $_.State -eq 'Running' -or ($_.State -eq 'Completed' -and $_ -notin $processedJobs) }) {
            $completed = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
            $progress = 20 + [math]::Floor(($completed / $totalJobs) * 80)

            if ($progress -ne $lastProgress) {
                Write-LogMessage "Stopping services on servers... ($completed/$totalJobs complete)" -ProgressValue $progress
                $lastProgress = $progress
            }

            # Process completed jobs
            foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                $output = Receive-Job -Job $job
                foreach ($line in $output) {
                    if ($line.StartsWith('###SUCCESS###:')) {
                        $server = $line.Split(':')[1]
                        Write-LogMessage "Service stop completed: $server"
                    }
                    elseif ($line.StartsWith('###ERROR###:')) {
                        $parts = $line.Split(':')
                        Write-LogMessage "Service stop failed: $($parts[1]) - $($parts[2])" -IsError
                    }
                    elseif ($line.StartsWith('###STATUS###:')) {
                        $status = $line.Split(':')[1]
                        Write-LogMessage $status
                    }
                    elseif ($line.StartsWith('###LOCALSTATEFILE###:')) {
                        $stateFile = $line.Split(':')[1]
                        Write-LogMessage "Service state saved to: $stateFile"
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

        Write-LogMessage "Service stop completed on all servers." -ProgressValue 100
    }
    catch {
        Write-LogMessage "Error during service stop: $($_.Exception.Message)" -IsError
        $progressBar.Visible = $false 
    }
    finally {
        # Cleanup any remaining jobs
        $jobs | Where-Object { $_ } | Remove-Job -Force -ErrorAction SilentlyContinue
        Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
        $stopServicesButton.Enabled = $true
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
    }
})

# Add the Start Services button click event
$startServicesButton.Add_Click({
    # Disable the button during start operation
    $startServicesButton.Enabled = $false
    
    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        Write-LogMessage "Please select a server list." -IsError
        $startServicesButton.Enabled = $true
        return
    }

    # Add confirmation dialog
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to start previously stopped services on the selected servers?`n`nServer List: $selectedItem",
        "Confirm Service Start",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::No) {
        Write-LogMessage "Service start cancelled by user."
        $startServicesButton.Enabled = $true
        return
    }
    
    $cred = Get-GlobalCredential

    try {
        # Clear previous progress
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
        $progressBar.Visible = $true
        $progressLabel.Visible = $true
    
        Write-LogMessage "Starting service start process..." -ProgressValue 0

        $serverListFile = ".\config\$selectedItem.txt"
        $servers = Get-Content $serverListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        Write-LogMessage "Preparing for parallel service start..." -ProgressValue 10

        # Create all start jobs simultaneously
        $jobs = foreach ($server in $servers) {
            $statusFilePath = Get-ServiceStatusFilePath -ServerList $selectedItem -Server $server
            
            if (-not (Test-Path $statusFilePath)) {
                Write-LogMessage "No state file found for $server, skipping." -IsError
                continue
            }
            
            Start-Job -ScriptBlock {
                param ($server, [PSCredential]$cred, $statusFilePath)
                try {
                    Write-Output "###STATUS###:Starting services on $server based on saved state"
                    
                    # First, copy the local state file to the remote server
                    $session = New-PSSession -ComputerName $server -Credential $cred
                    
                    # Copy state file to remote server
                    Copy-Item -Path $statusFilePath -Destination "C:\temp\service_state.xml" -ToSession $session -Force
                    
                    Invoke-Command -Session $session -ScriptBlock {
                        param ($statusFilePath)
                        
                        # If the state file doesn't exist, report error
                        if (-not (Test-Path "C:\temp\service_state.xml")) {
                            throw "State file not found on remote server"
                        }
                        
                        # Import the saved state
                        $serviceStates = Import-Clixml -Path "C:\temp\service_state.xml"
                        
                        Import-Module WebAdministration
                        
                        # Start services in correct order:
                        # 1. Start IIS first
                        if ($serviceStates.IIS.IsRunning) {
                            Write-Output "###STATUS###:Starting IIS (W3SVC)..."
                            Start-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
                        }
                        
                        # 2. Start app pools
                        foreach ($appPool in $serviceStates.AppPools) {
                            if ($appPool.IsRunning) {
                                Write-Output "###STATUS###:Starting Application Pool ($($appPool.Name))..."
                                Start-WebAppPool -Name $appPool.Name -ErrorAction SilentlyContinue
                            }
                        }
                        
                        # 3. Start custom service last
                        if ($serviceStates.CustomService.IsRunning) {
                            Write-Output "###STATUS###:Starting Custom Service ($($serviceStates.CustomService.Name))..."
                            Start-Service -Name $serviceStates.CustomService.Name -ErrorAction SilentlyContinue
                        }
                        
                        # Wait for services to stabilize
                        Write-Output "###STATUS###:Waiting for services to stabilize..."
                        Start-Sleep -Seconds 5
                        
                        # Verify services are running as expected
                        $allStarted = $true
                        
                        # Check IIS
                        if ($serviceStates.IIS.IsRunning) {
                            $iisService = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
                            if ($iisService) {
                                $status = $iisService.Status.ToString()
                                Write-Output "###STATUS###:IIS (W3SVC) Status: $status"
                                if ($status -ne 'Running') {
                                    $allStarted = $false
                                    Write-Output "###ERROR###:$env:COMPUTERNAME:IIS (W3SVC) failed to start"
                                }
                            }
                        }
                        
                        # Check app pools
                        foreach ($appPool in $serviceStates.AppPools) {
                            if ($appPool.IsRunning) {
                                $poolObject = Get-Item "IIS:\AppPools\$($appPool.Name)" -ErrorAction SilentlyContinue
                                if ($poolObject) {
                                    $status = $poolObject.State.ToString()
                                    Write-Output "###STATUS###:Application Pool ($($appPool.Name)) Status: $status"
                                    if ($status -ne 'Started') {
                                        $allStarted = $false
                                        Write-Output "###ERROR###:$env:COMPUTERNAME:Application Pool ($($appPool.Name)) failed to start"
                                    }
                                }
                            }
                        }
                        
                        # Check custom service
                        if ($serviceStates.CustomService.IsRunning) {
                            $customService = Get-Service -Name $serviceStates.CustomService.Name -ErrorAction SilentlyContinue
                            if ($customService) {
                                $status = $customService.Status.ToString()
                                Write-Output "###STATUS###:Custom Service ($($serviceStates.CustomService.Name)) Status: $status"
                                if ($status -ne 'Running') {
                                    $allStarted = $false
                                    Write-Output "###ERROR###:$env:COMPUTERNAME:Custom Service ($($serviceStates.CustomService.Name)) failed to start"
                                }
                            }
                        }
                        
                        if ($allStarted) {
                            Write-Output "###SUCCESS###:$env:COMPUTERNAME"
                        }
                        else {
                            throw "One or more services failed to start properly"
                        }
                    } -ArgumentList $statusFilePath
                    
                    Remove-PSSession -Session $session
                }
                catch {
                    Write-Output "###ERROR###:${server}:$($_.Exception.Message)"
                }
            } -ArgumentList $server, $cred, $statusFilePath
        }

        # If no jobs were created (no state files found), exit early
        if (-not $jobs -or $jobs.Count -eq 0) {
            Write-LogMessage "No service state files found for servers in this list. Please stop services first." -IsError
            $startServicesButton.Enabled = $true
            $progressBar.Visible = $false
            return
        }

        Write-LogMessage "Starting services on servers..." -ProgressValue 20
    
        $totalJobs = $jobs.Count
        $lastProgress = 20
        $processedJobs = @()

        # Monitor jobs until all are complete
        while ($jobs | Where-Object { $_.State -eq 'Running' -or ($_.State -eq 'Completed' -and $_ -notin $processedJobs) }) {
            $completed = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
            $progress = 20 + [math]::Floor(($completed / $totalJobs) * 80)

            if ($progress -ne $lastProgress) {
                Write-LogMessage "Starting services on servers... ($completed/$totalJobs complete)" -ProgressValue $progress
                $lastProgress = $progress
            }

            # Process completed jobs
            foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                $output = Receive-Job -Job $job
                foreach ($line in $output) {
                    if ($line.StartsWith('###SUCCESS###:')) {
                        $server = $line.Split(':')[1]
                        Write-LogMessage "Service start completed: $server"
                    }
                    elseif ($line.StartsWith('###ERROR###:')) {
                        $parts = $line.Split(':')
                        Write-LogMessage "Service start failed: $($parts[1]) - $($parts[2])" -IsError
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

        Write-LogMessage "Service start completed on all servers." -ProgressValue 100
    }
    catch {
        Write-LogMessage "Error during service start: $($_.Exception.Message)" -IsError
        $progressBar.Visible = $false 
    }
    finally {
        # Cleanup any remaining jobs
        $jobs | Where-Object { $_ } | Remove-Job -Force -ErrorAction SilentlyContinue
        Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
        $startServicesButton.Enabled = $true
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
    }
})

# Add the Rollback button click event
$rollbackButton.Add_Click({
    # Disable the button during rollback
    $rollbackButton.Enabled = $false
    
    if (-not $global:currentZipFileName) {
        Write-LogMessage "Please push a ZIP file first." -IsError
        $rollbackButton.Enabled = $true
        return
    }

    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        Write-LogMessage "Please select a server list." -IsError
        $rollbackButton.Enabled = $true
        return
    }

    # Add confirmation dialog
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to rollback on the selected servers?`n`nServer List: $selectedItem`nPackage: $global:currentZipFileName",
        "Confirm Rollback",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::No) {
        Write-LogMessage "Rollback cancelled by user."
        $rollbackButton.Enabled = $true
        return
    }
    
    $cred = Get-GlobalCredential

    try {
        # Clear previous progress
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
        $progressBar.Visible = $true
        $progressLabel.Visible = $true
        
        Write-LogMessage "Starting rollback process..." -ProgressValue 0

        $serverListFile = ".\config\$selectedItem.txt"
        $servers = Get-Content $serverListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        Write-LogMessage "Preparing for parallel rollback..." -ProgressValue 10

        # Create all rollback jobs simultaneously
        $jobs = foreach ($server in $servers) {
            Start-Job -ScriptBlock {
                param ($server, [PSCredential]$cred, $zipFileName, $AppPoolName1, $AppPoolName2, $AppPoolName3, $ServiceName)
                try {
                    Write-Output "###STATUS###:Starting rollback on $server"
                    $session = New-PSSession -ComputerName $server -Credential $cred
                    
                    Invoke-Command -Session $session -ScriptBlock {
                        param ($zipFileName, $AppPoolName1, $AppPoolName2, $AppPoolName3, $ServiceName)
                        
                        $zipPath = "C:\temp\$zipFileName"
                        $extractPath = "C:\temp\extracted"
                        
                        Add-Type -AssemblyName System.IO.Compression.FileSystem
                        
                        # Check if the extracted files exist
                        if (-not (Test-Path $extractPath)) {
                            throw "Extracted files not found. Cannot perform rollback."
                        }
                        
                        $exeFile = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | Select-Object -First 1
                        if ($exeFile) {
                            # Change to the executable's directory before running it
                            $originalLocation = Get-Location
                            try {
                                Set-Location -Path $exeFile.DirectoryName
                                # Use 'uninstall' argument instead of 'install'
                                $process = Start-Process -FilePath $exeFile.FullName -ArgumentList "uninstall" -Wait -NoNewWindow -PassThru
                        
                                if ($process.ExitCode -eq 0) {
                                    Write-Output "###SUCCESS###:$env:COMPUTERNAME"
                                }
                                else {
                                    throw "Rollback failed with exit code: $($process.ExitCode)"
                                }
                            }
                            finally {
                                Set-Location -Path $originalLocation
                            }
                        }
                        else {
                            throw "No EXE file found in extracted contents for rollback"
                        }
                    } -ArgumentList $zipFileName, $AppPoolName1, $AppPoolName2, $AppPoolName3, $ServiceName
                    
                    Remove-PSSession -Session $session
                }
                catch {
                    Write-Output "###ERROR###:${server}:$($_.Exception.Message)"
                }
            } -ArgumentList $server, $cred, $global:currentZipFileName, $AppPoolName, $AppPoolName2, $AppPoolName3, $ServiceName
        }

        Write-LogMessage "Starting parallel rollback on servers..." -ProgressValue 20
        
        $totalJobs = $jobs.Count
        $lastProgress = 20
        $processedJobs = @()

        # Monitor jobs until all are complete
        while ($jobs | Where-Object { $_.State -eq 'Running' -or ($_.State -eq 'Completed' -and $_ -notin $processedJobs) }) {
            $completed = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
            $progress = 20 + [math]::Floor(($completed / $totalJobs) * 80)

            if ($progress -ne $lastProgress) {
                Write-LogMessage "Rolling back on servers... ($completed/$totalJobs complete)" -ProgressValue $progress
                $lastProgress = $progress
            }

            # Process completed jobs
            foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                $output = Receive-Job -Job $job
                foreach ($line in $output) {
                    if ($line.StartsWith('###SUCCESS###:')) {
                        $server = $line.Split(':')[1]
                        Write-LogMessage "Rollback completed: $server"
                    }
                    elseif ($line.StartsWith('###ERROR###:')) {
                        $parts = $line.Split(':')
                        Write-LogMessage "Rollback failed: $($parts[1]) - $($parts[2])" -IsError
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
        
        Write-LogMessage "Rollback completed on all servers." -ProgressValue 100
    }
    catch {
        Write-LogMessage "Error during rollback: $($_.Exception.Message)" -IsError
        $progressBar.Visible = $false 
    }
    finally {
        # Cleanup any remaining jobs
        $jobs | Where-Object { $_ } | Remove-Job -Force -ErrorAction SilentlyContinue
        Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
        $rollbackButton.Enabled = $true
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
    }
})

# Add the Check Version button click event
$checkVersionButton.Add_Click({
    # Disable the button during operation
    $checkVersionButton.Enabled = $false
    
    $selectedItem = $dropdown.SelectedItem
    if ($selectedItem -eq "Select Server List" -or -not $selectedItem) {
        Write-LogMessage "Please select a server list." -IsError
        $checkVersionButton.Enabled = $true
        return
    }

    $cred = Get-GlobalCredential

    try {
        # Clear previous progress
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
        $progressBar.Visible = $true
        $progressLabel.Visible = $true
    
        Write-LogMessage "Starting version check process..." -ProgressValue 0

        $serverListFile = ".\config\$selectedItem.txt"
        $servers = Get-Content $serverListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        Write-LogMessage "Preparing to check versions on servers..." -ProgressValue 10

        # Create all version check jobs simultaneously
        $jobs = foreach ($server in $servers) {
            Start-Job -ScriptBlock {
                param ($server, [PSCredential]$cred, $AppPoolName, $ServiceName)
                try {
                    Write-Output "###STATUS###:Checking version information on $server"
                    $session = New-PSSession -ComputerName $server -Credential $cred
                
                    Invoke-Command -Session $session -ScriptBlock {
                        param ($AppPoolName, $ServiceName)
                        
                        $versionInfo = @{
                            ComputerName = $env:COMPUTERNAME
                            OSVersion = [System.Environment]::OSVersion.Version.ToString()
                            ServiceInfo = $null
                            AppPoolInfo = $null
                            InstalledApps = @()
                        }
                        
                        # Check service version if possible
                        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                        if ($service) {
                            $versionInfo.ServiceInfo = @{
                                Name = $service.Name
                                DisplayName = $service.DisplayName
                                Status = $service.Status.ToString()
                            }
                            
                            # Try to get executable path and version
                            try {
                                $wmiService = Get-WmiObject -Class Win32_Service -Filter "Name='$ServiceName'"
                                if ($wmiService) {
                                    $pathName = $wmiService.PathName
                                    $exePath = $pathName -replace '^"([^"]+)".*$', '$1'
                                    if (Test-Path $exePath) {
                                        $fileVersion = (Get-Item $exePath).VersionInfo.FileVersion
                                        $versionInfo.ServiceInfo.Path = $exePath
                                        $versionInfo.ServiceInfo.Version = $fileVersion
                                    }
                                }
                            }
                            catch {
                                # Unable to get service executable info
                            }
                        }
                        
                        # Check app pool info
                        if (Get-Command Get-WebAppPoolState -ErrorAction SilentlyContinue) {
                            try {
                                Import-Module WebAdministration
                                $appPool = Get-Item "IIS:\AppPools\$AppPoolName" -ErrorAction SilentlyContinue
                                if ($appPool) {
                                    $versionInfo.AppPoolInfo = @{
                                        Name = $appPool.Name
                                        State = $appPool.State.ToString()
                                        RuntimeVersion = $appPool.ManagedRuntimeVersion
                                    }
                                    
                                    # Try to get application information
                                    $sites = Get-ChildItem "IIS:\Sites" -ErrorAction SilentlyContinue
                                    $appInfo = @()
                                    foreach ($site in $sites) {
                                        $apps = Get-WebApplication -Site $site.Name -ErrorAction SilentlyContinue
                                        foreach ($app in $apps) {
                                            if ($app.ApplicationPool -eq $AppPoolName) {
                                                $appInfo += @{
                                                    SiteName = $site.Name
                                                    AppName = $app.Path
                                                    PhysicalPath = $app.PhysicalPath
                                                }
                                                
                                                # Try to get version from dll or config file
                                                try {
                                                    $assemblyFiles = Get-ChildItem -Path $app.PhysicalPath -Filter "*.dll" -Recurse -ErrorAction SilentlyContinue | 
                                                                     Where-Object { $_.Name -like "*$ServiceName*" -or $_.Name -like "*Application*" }
                                                    
                                                    if ($assemblyFiles) {
                                                        foreach ($file in $assemblyFiles) {
                                                            $appInfo[-1].AssemblyVersion = (Get-Item $file.FullName).VersionInfo.FileVersion
                                                            break
                                                        }
                                                    }
                                                }
                                                catch {
                                                    # Unable to get assembly version
                                                }
                                            }
                                        }
                                    }
                                    
                                    if ($appInfo.Count -gt 0) {
                                        $versionInfo.AppPoolInfo.Applications = $appInfo
                                    }
                                }
                            }
                            catch {
                                # Unable to get app pool info
                            }
                        }
                        
                        # Get installed applications
                        try {
                            $installedApps = Get-WmiObject -Class Win32_Product |
                                             Where-Object { $_.Name -like "*$ServiceName*" -or $_.Name -like "*Application*" } |
                                             Select-Object Name, Version, Vendor, InstallDate
                            
                            if ($installedApps) {
                                foreach ($app in $installedApps) {
                                    $versionInfo.InstalledApps += @{
                                        Name = $app.Name
                                        Version = $app.Version
                                        Vendor = $app.Vendor
                                        InstallDate = $app.InstallDate
                                    }
                                }
                            }
                        }
                        catch {
                            # Unable to get installed apps
                        }
                        
                        # Return version information
                        return $versionInfo
                    } -ArgumentList $AppPoolName, $ServiceName
                    
                    # Process the returned version info
                    $output = "### VERSION INFO FOR $server ###`n"
                    $output += "OS Version: $($versionInfo.OSVersion)`n"
                    
                    if ($versionInfo.ServiceInfo) {
                        $output += "`nService Information:`n"
                        $output += "  Name: $($versionInfo.ServiceInfo.Name)`n"
                        $output += "  Display Name: $($versionInfo.ServiceInfo.DisplayName)`n"
                        $output += "  Status: $($versionInfo.ServiceInfo.Status)`n"
                        if ($versionInfo.ServiceInfo.Version) {
                            $output += "  Version: $($versionInfo.ServiceInfo.Version)`n"
                        }
                    }
                    
                    if ($versionInfo.AppPoolInfo) {
                        $output += "`nApp Pool Information:`n"
                        $output += "  Name: $($versionInfo.AppPoolInfo.Name)`n"
                        $output += "  State: $($versionInfo.AppPoolInfo.State)`n"
                        $output += "  Runtime Version: $($versionInfo.AppPoolInfo.RuntimeVersion)`n"
                        
                        if ($versionInfo.AppPoolInfo.Applications) {
                            $output += "  Applications:`n"
                            foreach ($app in $versionInfo.AppPoolInfo.Applications) {
                                $output += "    - Site: $($app.SiteName), App: $($app.AppName)`n"
                                $output += "      Physical Path: $($app.PhysicalPath)`n"
                                if ($app.AssemblyVersion) {
                                    $output += "      Assembly Version: $($app.AssemblyVersion)`n"
                                }
                            }
                        }
                    }
                    
                    if ($versionInfo.InstalledApps.Count -gt 0) {
                        $output += "`nInstalled Applications:`n"
                        foreach ($app in $versionInfo.InstalledApps) {
                            $output += "  - $($app.Name) (Version: $($app.Version))`n"
                            $output += "    Vendor: $($app.Vendor), Installed: $($app.InstallDate)`n"
                        }
                    }
                    
                    Write-Output "###VERSIONINFO###:$output"
                    Write-Output "###SUCCESS###:$server"
                    
                    Remove-PSSession -Session $session
                }
                catch {
                    Write-Output "###ERROR###:${server}:$($_.Exception.Message)"
                }
            } -ArgumentList $server, $cred, $AppPoolName, $ServiceName
        }

        Write-LogMessage "Checking versions on servers..." -ProgressValue 20
    
        $totalJobs = $jobs.Count
        $lastProgress = 20
        $processedJobs = @()

        # Monitor jobs until all are complete
        while ($jobs | Where-Object { $_.State -eq 'Running' -or ($_.State -eq 'Completed' -and $_ -notin $processedJobs) }) {
            $completed = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
            $progress = 20 + [math]::Floor(($completed / $totalJobs) * 80)

            if ($progress -ne $lastProgress) {
                Write-LogMessage "Checking versions on servers... ($completed/$totalJobs complete)" -ProgressValue $progress
                $lastProgress = $progress
            }

            # Process completed jobs
            foreach ($job in ($jobs | Where-Object { $_.State -eq 'Completed' -and $_ -notin $processedJobs })) {
                $output = Receive-Job -Job $job
                foreach ($line in $output) {
                    if ($line.StartsWith('###SUCCESS###:')) {
                        $server = $line.Split(':')[1]
                        Write-LogMessage "Version check completed: $server"
                    }
                    elseif ($line.StartsWith('###ERROR###:')) {
                        $parts = $line.Split(':')
                        Write-LogMessage "Version check failed: $($parts[1]) - $($parts[2])" -IsError
                    }
                    elseif ($line.StartsWith('###STATUS###:')) {
                        $status = $line.Split(':')[1]
                        Write-LogMessage $status
                    }
                    elseif ($line.StartsWith('###VERSIONINFO###:')) {
                        $versionInfoText = $line.Substring(17)  # Remove prefix
                        Write-LogMessage "`n$versionInfoText"
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

        Write-LogMessage "Version check completed on all servers." -ProgressValue 100
    }
    catch {
        Write-LogMessage "Error during version check: $($_.Exception.Message)" -IsError
        $progressBar.Visible = $false 
    }
    finally {
        # Cleanup any remaining jobs
        $jobs | Where-Object { $_ } | Remove-Job -Force -ErrorAction SilentlyContinue
        Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
        $checkVersionButton.Enabled = $true
        $progressBar.Value = 0
        $progressLabel.Text = "0%"
    }
})

# Center buttons initially
Set-ButtonsAlignment

# Show the form
$form.ShowDialog()
