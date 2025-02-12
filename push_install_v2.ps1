Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Load configuration
try {
    $config = Get-Content -Path ".\config\config.json" | ConvertFrom-Json
    
    $LogFile = $config.LogFile
    $InstallerName = $config.InstallerName
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

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(620, 65)  # Position above progress bar
$progressLabel.Width = 150
$progressLabel.Height = 20
$progressLabel.Text = "0%"
$progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$progressLabel.Visible = $false
$form.Controls.Add($progressLabel)

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
    
        # Set initial directory from config
        if ($config.DefaultBrowsePath -and (Test-Path $config.DefaultBrowsePath)) {
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
        $statusLabel.Text = "Error: $Message"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
    else {
        Write-Host $LogEntry
        $outputBox.SelectionColor = [System.Drawing.Color]::Black
        $statusLabel.Text = $Message
        $statusLabel.ForeColor = [System.Drawing.Color]::Black
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
                    $percentComplete = [math]::Floor(($current / $totalItems) * 100)
                
                    Write-Output "###PROGRESS###:$percentComplete"
                    Write-Output "###STATUS###:Processing folder ($current/$totalItems): $($folder.Name)"
                
                    $targetPath = Join-Path $localPath $folder.Name
                
                    if (-not (Test-Path $targetPath)) {
                        New-Item -ItemType Directory -Path $targetPath | Out-Null
                        Write-Output "###STATUS###:Created new folder: $($folder.Name)"
                    }
        
                    # Get remote files
                    $remoteFiles = Get-ChildItem -Path (Join-Path $folder.FullName "*") -Recurse -Credential $cred
        
                    foreach ($remoteFile in $remoteFiles) {
                        $relativePath = $remoteFile.FullName.Substring($folder.FullName.Length)
                        $localFilePath = Join-Path $targetPath $relativePath
                    
                        $shouldCopy = $false
                    
                        if (-not (Test-Path $localFilePath)) {
                            $shouldCopy = $true
                            Write-Output "###STATUS###:New file found: $relativePath"
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
                            Copy-Item -Path $remoteFile.FullName -Destination $localFilePath -Force -Credential $cred
                        }
                    }
                    Write-Output "###COMPLETE###:$($folder.Name)"
                }
            } -ArgumentList $remotePath, $localPath, $cred

            # Wait for sync to complete
            while ($job.State -eq 'Running') {
                $jobOutput = Receive-Job -Job $job
                foreach ($line in $jobOutput) {
                    if ($line.StartsWith('###PROGRESS###:')) {
                        $progress = [int]($line.Split(':')[1])
                        $progressBar.Value = $progress
                    }
                    elseif ($line.StartsWith('###STATUS###:')) {
                        $status = $line.Split(':')[1]
                        Write-LogMessage $status
                    }
                    elseif ($line.StartsWith('###COMPLETE###:')) {
                        $folder = $line.Split(':')[1]
                        Write-LogMessage "Completed copying folder: $folder"
                    }
                }
                Start-Sleep -Milliseconds 100
            }

            Receive-Job -Job $job | Out-Null
    
            if ($job.State -eq 'Completed') {
                Write-LogMessage "Sync completed successfully." -ProgressValue 100
            }
            else {
                Write-LogMessage "Sync operation failed." -IsError
            }

            Remove-Job -Job $job

        }
        catch {
            Write-LogMessage "Error during sync: $($_.Exception.Message)" -IsError
            $progressBar.Visible = $false
        }
    })

# Create temp directory if it doesn't exist
$tempPath = Join-Path $PSScriptRoot "temp"
if (-not (Test-Path $tempPath)) {
    New-Item -ItemType Directory -Path $tempPath | Out-Null
}

# Update Push ZIP button click event
$pushZipButton.Add_Click({
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
                            } else {
                                throw "Local file copy failed"
                            }
                        } else {
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
                            } else {
                                throw "Remote file copy failed"
                            }
                        }
                    } catch {
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

# Install MSI button click event
$installMsiButton.Add_Click({
        if (-not $global:currentZipFileName) {
            Write-LogMessage "Please push a ZIP file first." -IsError
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
                    Update-Progress -Current $currentServer -Total $totalServers -Operation "Installing MSI on servers"

                    try {
                        $job = Start-Job -ScriptBlock {
                            param ($server, [PSCredential]$cred, $zipFileName)
                            $session = New-PSSession -ComputerName $server -Credential $cred
                            Invoke-Command -Session $session -ScriptBlock {
                                param ($zipFileName)
                                Write-Host "###PROGRESS###:Preparing installation on $env:COMPUTERNAME"
                                $zipPath = "C:\temp\$zipFileName"
                                $extractPath = "C:\temp\extracted"
    
                                Write-Host "###PROGRESS###:Extracting ZIP file"
                                Add-Type -AssemblyName System.IO.Compression.FileSystem
                                [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractPath)
                                $exeFile = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | Select-Object -First 1
                                if ($exeFile) {
                                    Write-Host "###PROGRESS###:Found executable: $($exeFile.Name)"
                                    try {
                                        Write-Host "###PROGRESS###:Starting installation"
                                        $process = Start-Process -FilePath $exeFile.FullName -ArgumentList "install" -Wait -NoNewWindow -PassThru

                                        if ($process.ExitCode -eq 0) {
                                            Write-Host "###PROGRESS###:Installation completed successfully (Exit Code: 0)"
                                        }
                                        else {
                                            throw "Installation failed with exit code: $($process.ExitCode)"
                                        }
                                        Write-Host "###PROGRESS###:Restarting services"
                                        Restart-Service -Name 'W3SVC'
                                        Import-Module WebAdministration
                                        Restart-WebAppPool -Name $using:YourAppPoolName
                                        Restart-Service -Name $using:YourServiceName
                                        Write-Host "###PROGRESS###:Services restarted successfully"
                                    }
                                    catch {
                                        throw "Installation failed: $_"
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

            Write-LogMessage "Installation completed." -ProgressValue 100
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