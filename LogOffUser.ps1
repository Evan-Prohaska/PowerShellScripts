Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Threading
Add-Type -AssemblyName System.ComponentModel

function Test-ActiveSubnet {
    param([String]$subnet, $bgWorker = $null, $e = $null)

    # Strip off any numbers after the last dot
    $baseSubnet = $subnet -replace '\.\d+$'

    # Initialize an empty list to hold results
    $results = @()

    # Create a runspace pool
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 256)
    $runspacePool.Open()

    $runspaces = @()

    # Loop through 1-254 for the last octet
    1..254 | ForEach-Object {
        if ($bgWorker.CancellationPending) {
            $e.Cancel = $true
            return
        }

        $ip = "$baseSubnet.$_"

        $runspace = [powershell]::Create().AddScript({
            param($ip)

            # Ping machine
            $isActive = Test-Connection -ComputerName $ip -Count 1 -Quiet
            
            if ($isActive) {
                # Check if port 135 is open, which would indicate a potential Windows machine
                $tcpTest = Test-NetConnection -ComputerName $ip -Port 135 -WarningAction SilentlyContinue
                if ($tcpTest.TcpTestSucceeded) {
                    try {
                        # Get the computer name
                        $computerName = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ip -ErrorAction SilentlyContinue).Name
                        
                        # If we couldn't get the name, assign a placeholder
                        if (-not $computerName) {
                            $computerName = "Unknown"
                        }
                        $currentUsername = "$env:USERDOMAIN\$env:USERNAME"

                        $allLoggedUsers = Get-WmiObject -Class Win32_LoggedOnUser -ComputerName $ip |
                                          Where-Object { 
                                              $_.Antecedent -match 'Domain="(.+?)",Name="(.+?)"' -and 
                                              $matches[1] -eq 'SIMONMED' -and
                                              "$($matches[1])\$($matches[2])" -ne $currentUsername
                                          } |
                                          ForEach-Object {
                                              "$($matches[1])\$($matches[2])"
                                          } |
                                          Sort-Object -Unique
            
                        # If we couldn't get users, assign a placeholder
                        if (-not $allLoggedUsers) {
                            $allLoggedUsers = "Unknown"
                        }
            
                        # Create an object to hold details
                        $obj = New-Object PSObject -Property @{
                            IPAddress     = $ip
                            ComputerName  = $computerName
                            LoggedUsers   = $allLoggedUsers
                        }
            
                        return $obj
                    } catch {
                        Write-Verbose "Error occurred for IP: $ip. Error Message: $_"
                        $obj = New-Object PSObject -Property @{
                            IPAddress     = $ip
                            ComputerName  = "Error"
                            LoggedUsers   = "Error"
                        }
            
                        return $obj
                    }
                }
            }

            return $null

        }).AddArgument($ip)

        $runspace.RunspacePool = $runspacePool

        $runspaces += [PSCustomObject]@{
            Runspace = $runspace
            Status   = $runspace.BeginInvoke()
        }
    }

    # Gather results from the runspaces
    $runspaces | ForEach-Object {
        $result = $_.Runspace.EndInvoke($_.Status)
        if ($result) {
            $results += $result
        }
    }

    $RunspacePool.Close()
    $RunspacePool.Dispose()

    return $results
}

function HandleKeyPress {
    param(
        [System.Windows.Forms.KeyPressEventArgs]$e,
        [System.Windows.Forms.TextBox]$textbox
    )

    $currentText = $textbox.Text
    $newText = $currentText.Insert($textbox.SelectionStart, $e.KeyChar)

    # Simulate the removal of a character for backspace key press.
    if ([System.Char]::IsControl($e.KeyChar) -and $e.KeyChar -ne [char]13) {
        $newText = $currentText.Remove($textbox.SelectionStart - 1, 1)
    }

    # Check if more than 3 digits between periods or value over 255.
    $isValid = $true
    $newText -split "\." | ForEach-Object {
        if ($_.Length -gt 3 -or ($_ -as [int]) -gt 255) {
            $isValid = $false
            return
        }
    }

    # Check for consecutive periods.
    if ($newText -match "\.\.") {
        $isValid = $false
    }

    # Allow numbers, the period character, and control characters (like backspace).
    # Also check our additional validity criteria.
    if (-not ($e.KeyChar -match "[0-9\.]" -or [System.Char]::IsControl($e.KeyChar)) -or -not $isValid) {
        # If the character/action is not allowed, suppress it.
        $e.Handled = $true
    }
}

# DoWork event of the BackgroundWorker
$backgroundWorker_DoWork = {
    param([object]$sender, [System.ComponentModel.DoWorkEventArgs]$e)

    $subnet = $e.Argument
    $worker = $sender

    $results = Test-ActiveSubnet -subnet $subnet -bgWorker $worker -e $e

    $e.Result = $results
}

$backgroundWorker_RunWorkerCompleted = {
    param([object]$sender, [System.ComponentModel.RunWorkerCompletedEventArgs]$e)

    if ($e.Cancelled) {
        $lblScanStatus.Text = 'Scan Status: Cancelled'
    } else {
        $lblScanStatus.Text = 'Scan Status: Completed'
        $lstMachines.Items.Clear()
        foreach ($result in $e.Result) {
            # Populate list box based on the username entered
            if ($txtUsername.Text -in $result.LoggedUsers) {
                $lstMachines.Items.Add($result.ComputerName)
            }
        }
    }

    $btnScan.Enabled = $true
    $btnStop.Enabled = $false
}

$backgroundWorker_ProgressChanged = {
    param([object]$sender, [System.ComponentModel.ProgressChangedEventArgs]$e)
    # Assuming the percentage of completion is passed as progress data.
    $progressBar.Value = $e.ProgressPercentage
}

$backgroundWorker = New-Object System.ComponentModel.BackgroundWorker
$backgroundWorker.WorkerReportsProgress = $true
$backgroundWorker.WorkerSupportsCancellation = $true
$backgroundWorker.Add_DoWork($backgroundWorker_DoWork)
$backgroundWorker.Add_ProgressChanged($backgroundWorker_ProgressChanged)
$backgroundWorker.Add_RunWorkerCompleted($backgroundWorker_RunWorkerCompleted)

# Form Initialization
$formMain = [System.Windows.Forms.Form]@{
    Text       = 'User Logoff Tool'
    Size       = [System.Drawing.Size]::new(310, 510)
    StartPosition = 'CenterScreen'
    MaximizeBox   = $false
    FormBorderStyle = 'FixedSingle'
}

# Label Initializations
$labels = @{
    'lblUsername' = @{
        Location = [System.Drawing.Point]::new(10, 10)
        Size     = [System.Drawing.Size]::new(100, 20)
        Text     = 'Username:'
    }
    'lblSubnet' = @{
        Location = [System.Drawing.Point]::new(10, 40)
        Size     = [System.Drawing.Size]::new(100, 20)
        Text     = 'Subnet:'
    }
    'lblScanStatus' = @{
        Location = [System.Drawing.Point]::new(10, 130)
        Size     = [System.Drawing.Size]::new(200, 20)
        Text     = 'Scan Status: Idle'
    }
    'lblMachines' = @{
        Location = [System.Drawing.Point]::new(10, 160)
        Size     = [System.Drawing.Size]::new(280, 20)
        Text     = 'Machines with User Logged In:'
    }
    'lblUnreachableMachines' = @{
        Location = [System.Drawing.Point]::new(10, 320)
        Size     = [System.Drawing.Size]::new(280, 20)
        Text     = 'Machines Not Reachable via PSRemoting:'
    }
}

$labels.GetEnumerator() | ForEach-Object {
    $label = [System.Windows.Forms.Label]$_.Value
    $label.Name = $_.Key
    Set-Variable -Name $_.Key -Value $label -Scope Script
    $formMain.Controls.Add($label)
}

# TextBox Initializations
$textBoxes = @{
    'txtUsername' = @{
        Location = [System.Drawing.Point]::new(120, 10)
        Size     = [System.Drawing.Size]::new(170, 20)
    }
    'txtSubnet' = @{
        Location = [System.Drawing.Point]::new(120, 40)
        Size     = [System.Drawing.Size]::new(170, 20)
        KeyPressEvent = {
            param($sender, $e)
            HandleKeyPress -e $e -textbox $sender
        }
    }
}

# Enumerate through the textboxes and add them to the form
$textBoxes.GetEnumerator() | ForEach-Object {
    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Name = $_.Key
    $textbox.Location = $_.Value.Location
    $textbox.Size = $_.Value.Size
    
    # If the KeyPressEvent property exists, add it
    if ($_.Value.ContainsKey('KeyPressEvent')) {
        $textbox.Add_KeyPress($_.Value.KeyPressEvent)
    }

    # If it's the txtSubnet, set the default text value
    if ($_.Key -eq 'txtSubnet') {
        $textbox.Text = '0.0.0.0'
    }

    Set-Variable -Name $_.Key -Value $textbox -Scope Script
    $formMain.Controls.Add($textbox)
}

# ListBox Initializations
$listBoxes = @{
    'lstMachines' = @{
        Location = [System.Drawing.Point]::new(10, 180)
        Size     = [System.Drawing.Size]::new(280, 130)
    }
    'lstUnreachableMachines' = @{
        Location = [System.Drawing.Point]::new(10, 340)
        Size     = [System.Drawing.Size]::new(280, 130)
    }
}

$listBoxes.GetEnumerator() | ForEach-Object {
    $listBox = [System.Windows.Forms.ListBox]$_.Value
    $listBox.Name = $_.Key
    Set-Variable -Name $_.Key -Value $listBox -Scope Script
    $formMain.Controls.Add($listBox)
}

# Buttons Initializations
$buttons = @{
    'btnScan' = @{
        Location = [System.Drawing.Point]::new(10, 100)
        Size     = [System.Drawing.Size]::new(100, 30)
        Text     = 'Scan'
    }
    'btnStop' = @{
        Location = [System.Drawing.Point]::new(120, 100)
        Size     = [System.Drawing.Size]::new(100, 30)
        Text     = 'Stop'
        Enabled  = $false
    }
    'btnLogoffUser' = @{
        Location = [System.Drawing.Point]::new(230, 100)
        Size     = [System.Drawing.Size]::new(60, 30)
        Text     = 'Logoff'
    }
}

$buttons.GetEnumerator() | ForEach-Object {
    # Create the button
    $button = [System.Windows.Forms.Button]::new()
    $button.Name = $_.Key
    $button.Location = $_.Value.Location
    $button.Size = $_.Value.Size
    $button.Text = $_.Value.Text
    if ($_.Value.ContainsKey('Enabled')) {
        $button.Enabled = $_.Value.Enabled
    }

    # Attach the respective event handler based on the button key
    switch ($_.Key) {
        'btnScan' { $button.Add_Click($btnScan_Click) }
        'btnStop' { $button.Add_Click($btnStop_Click) }
        'btnLogoffUser' { $button.Add_Click($btnLogoffUser_Click) }
    }

    Set-Variable -Name $_.Key -Value $button -Scope Script
    $formMain.Controls.Add($button)
}

# ProgressBar Initialization
$progressBar = [System.Windows.Forms.ProgressBar]@{
    Location = [System.Drawing.Point]::new(10, 70)
    Size     = [System.Drawing.Size]::new(280, 20)
    Maximum = 100  # To represent percentage
    Minimum = 0
}
$formMain.Controls.Add($progressBar)

# Timer Initialization
$timer = [System.Windows.Forms.Timer]::new()
$timer.Interval = 1000

$btnScan_Click = {
    $lblScanStatus.Text = 'Scan Status: Scanning...'
    $lstMachines.Items.Clear()
    $btnScan.Enabled = $false
    $btnStop.Enabled = $true

    $subnet = $txtSubnet.Text

    # Start the scan in the background
    $backgroundWorker.RunWorkerAsync($subnet)
}
$btnStop_Click = {
    Write-Host "Stop button clicked. Attempting to cancel background worker..."
    
    # Cancel the background worker
    $backgroundWorker.CancelAsync()

    # Update the Scan Status label to 'Idle'
    $lblScanStatus.Text = 'Scan Status: Idle'

    # Show a message box to inform the user
    [System.Windows.Forms.MessageBox]::Show('Scan canceled.', 'Information', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

$btnLogoffUser_Click = {
    # Ensure a machine is selected
    if ($lstMachines.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show('Please select a machine from the list.', 'Warning', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedMachine = $lstMachines.SelectedItem.ToString()
    $selectedUsername = $txtUsername.Text

    # Confirm with the user
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to log off user '$selectedUsername' from '$selectedMachine'?", 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($result -eq 'Yes') {
        try {
            # Get user sessions from the remote computer
            $sessions = query user /server:$selectedMachine 2>&1
            foreach ($session in $sessions) {
                if ($session -match '\s+(\d+)\s+(\S+)$' -and $matches[2] -eq $selectedUsername) {
                    $sessionId = $matches[1]
                    # Logoff the user
                    logoff $sessionId /server:$selectedMachine
                }
            }
            [System.Windows.Forms.MessageBox]::Show("User '$selectedUsername' was logged off from '$selectedMachine'.", 'Success', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to log off user '$selectedUsername' from '$selectedMachine'. Error: $_", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Display the Form
$formMain.ShowDialog()
