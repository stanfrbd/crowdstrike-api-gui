<#
.SYNOPSIS
	Simple PowerShell GUI for Crowdstrike API machine actions.

.DESCRIPTION
	Sample response tool that benefits from APIs and are using PowerShell as the tool of choice to perform actions in bulk. It doesn't require installation and can easily be adapted by anyone with some scripting experience. The tool currently accepts CSVs as device input methods. Once devices are selected, these types of actions can be performed:

	- Get devices IDs from Crowdstrike API - OK
    - Tagging devices - OK
	- Performing Isolation/Release from Isolation - To be tested

	A Crowdstrike API client should be used with AppID and Secret is required to connect to API and the tool needs the following App Permissions:

	- Scopes still to be defined (seems to be Hosts)
    - Hosts - Read / Write

#>


#===========================================================[Classes]===========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class ProcessDPI {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetProcessDPIAware();      
}
'@
$null = [ProcessDPI]::SetProcessDPIAware()


#===========================================================[Variables]===========================================================


$script:selectedmachines = @{}
$credspath = 'c:\temp\crowdstrikeuicreds.txt'

$UnclickableColour = "#8d8989"
$ClickableColour = "#ff7b00"
$TextBoxFont = 'Microsoft Sans Serif,10'

#===========================================================[WinForm]===========================================================


[System.Windows.Forms.Application]::EnableVisualStyles()


$MainForm = New-Object system.Windows.Forms.Form
$MainForm.SuspendLayout()
$MainForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$MainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$MainForm.ClientSize = '950,800'
$MainForm.text = "Crowdstrike API GUI"
$MainForm.BackColor = "#ffffff"
$MainForm.TopMost = $false

$Title = New-Object system.Windows.Forms.Label
$Title.text = "1 - Connect with Crowdstrike API Credentials"
$Title.AutoSize = $true
$Title.width = 25
$Title.height = 10
$Title.location = New-Object System.Drawing.Point(20, 20)
$Title.Font = 'Microsoft Sans Serif,12,style=Bold'

$AppIdBoxLabel = New-Object system.Windows.Forms.Label
$AppIdBoxLabel.text = "App Id:"
$AppIdBoxLabel.AutoSize = $true
$AppIdBoxLabel.width = 25
$AppIdBoxLabel.height = 10
$AppIdBoxLabel.location = New-Object System.Drawing.Point(20, 50)
$AppIdBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppIdBox = New-Object system.Windows.Forms.TextBox
$AppIdBox.multiline = $false
$AppIdBox.width = 314
$AppIdBox.height = 20
$AppIdBox.location = New-Object System.Drawing.Point(100, 50)
$AppIdBox.Font = $TextBoxFont
$AppIdBox.Visible = $true

$AppSecretBoxLabel = New-Object system.Windows.Forms.Label
$AppSecretBoxLabel.text = "App Secret:"
$AppSecretBoxLabel.AutoSize = $true
$AppSecretBoxLabel.width = 25
$AppSecretBoxLabel.height = 10
$AppSecretBoxLabel.location = New-Object System.Drawing.Point(20, 75)
$AppSecretBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppSecretBox = New-Object system.Windows.Forms.TextBox
$AppSecretBox.multiline = $false
$AppSecretBox.width = 314
$AppSecretBox.height = 20
$AppSecretBox.location = New-Object System.Drawing.Point(100, 75)
$AppSecretBox.Font = $TextBoxFont
$AppSecretBox.Visible = $true
$AppSecretBox.PasswordChar = '*'

$CidBoxLabel = New-Object system.Windows.Forms.Label
$CidBoxLabel.text = "CID:"
$CidBoxLabel.AutoSize = $true
$CidBoxLabel.width = 25
$CidBoxLabel.height = 10
$CidBoxLabel.location = New-Object System.Drawing.Point(20, 100)
$CidBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$CidBox = New-Object system.Windows.Forms.TextBox
$CidBox.multiline = $false
$CidBox.width = 314
$CidBox.height = 20
$CidBox.location = New-Object System.Drawing.Point(100, 100)
$CidBox.Font = $TextBoxFont
$CidBox.Visible = $true

$ConnectionStatusLabel = New-Object system.Windows.Forms.Label
$ConnectionStatusLabel.text = "Status:"
$ConnectionStatusLabel.AutoSize = $true
$ConnectionStatusLabel.width = 25
$ConnectionStatusLabel.height = 10
$ConnectionStatusLabel.location = New-Object System.Drawing.Point(20, 135)
$ConnectionStatusLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$ConnectionStatus = New-Object system.Windows.Forms.Label
$ConnectionStatus.text = "Not Connected"
$ConnectionStatus.AutoSize = $true
$ConnectionStatus.width = 25
$ConnectionStatus.height = 10
$ConnectionStatus.location = New-Object System.Drawing.Point(100, 135)
$ConnectionStatus.Font = 'Microsoft Sans Serif,10'

$SaveCredCheckbox = new-object System.Windows.Forms.checkbox
$SaveCredCheckbox.Location = New-Object System.Drawing.Point(200, 135)
$SaveCredCheckbox.AutoSize = $true
$SaveCredCheckbox.width = 60
$SaveCredCheckbox.height = 10
$SaveCredCheckbox.Text = "Save Credentials"
$SaveCredCheckbox.Font = 'Microsoft Sans Serif,10'
$SaveCredCheckbox.Checked = $false

$ConnectBtn = New-Object system.Windows.Forms.Button
$ConnectBtn.BackColor = "#ff7b00"
$ConnectBtn.text = "Connect"
$ConnectBtn.width = 90
$ConnectBtn.height = 30
$ConnectBtn.location = New-Object System.Drawing.Point(325, 130)
$ConnectBtn.Font = 'Microsoft Sans Serif,10'
$ConnectBtn.ForeColor = "#ffffff"
$ConnectBtn.Visible = $True

$TenantNoteLabel = New-Object system.Windows.Forms.Label
$TenantNoteLabel.text = "To perform actions on devices in a child tenant, include only machines from that CID. If your hosts list span multiple tenants, run the script once per tenant and set the matching CID each time."
$TenantNoteLabel.AutoSize = $false
$TenantNoteLabel.width = 460
$TenantNoteLabel.height = 40
$TenantNoteLabel.location = New-Object System.Drawing.Point(20, 170)
$TenantNoteLabel.Font = 'Microsoft Sans Serif,9'
$TenantNoteLabel.ForeColor = "#D0021B"
$TenantNoteLabel.Visible = $true

$MainForm.Controls.Add($TenantNoteLabel)

$TitleActions = New-Object system.Windows.Forms.Label
$TitleActions.text = "3 - Perform Action on selected devices"
$TitleActions.AutoSize = $true
$TitleActions.width = 25
$TitleActions.height = 10
$TitleActions.location = New-Object System.Drawing.Point(500, 20)
$TitleActions.Font = 'Microsoft Sans Serif,12,style=Bold'

$TagDeviceGroupBox = New-Object System.Windows.Forms.GroupBox
$TagDeviceGroupBox.Location = New-Object System.Drawing.Point(500, 40)
$TagDeviceGroupBox.width = 400
$TagDeviceGroupBox.height = 50
$TagDeviceGroupBox.Text = "Falcon Grouping Tag"
$TagDeviceGroupBox.Font = 'Microsoft Sans Serif,10,style=Bold'

$DeviceTag = New-Object system.Windows.Forms.TextBox
$Devicetag.multiline = $false
$DeviceTag.width = 200
$DeviceTag.height = 25
$DeviceTag.location = New-Object System.Drawing.Point(20, 20)
$Devicetag.Font = 'Microsoft Sans Serif,10'
$DeviceTag.Visible = $true
$Devicetag.Enabled = $false
$DeviceTag.PlaceholderText = "FalconGroupingTags/..."

$TagDeviceBtn = New-Object system.Windows.Forms.Button
$TagDeviceBtn.BackColor = $UnclickableColour
$TagDeviceBtn.text = "Apply Tag"
$TagDeviceBtn.width = 110
$TagDeviceBtn.height = 30
$TagDeviceBtn.location = New-Object System.Drawing.Point(280, 15)
$TagDeviceBtn.Font = 'Microsoft Sans Serif,10'
$TagDeviceBtn.ForeColor = "#ffffff"
$TagDeviceBtn.Visible = $true
$TagDeviceBtn.Enabled = $false

# enlarge groupbox to fit new button
$TagDeviceGroupBox.Height = 90

$RemoveTagBtn = New-Object System.Windows.Forms.Button
$RemoveTagBtn.BackColor = $UnclickableColour
$RemoveTagBtn.Text = "Remove Tag"
$RemoveTagBtn.Width = 110
$RemoveTagBtn.Height = 30
$RemoveTagBtn.Location = New-Object System.Drawing.Point(280, 50)
$RemoveTagBtn.Font = 'Microsoft Sans Serif,10'
$RemoveTagBtn.ForeColor = "#ffffff"
$RemoveTagBtn.Visible = $true
$RemoveTagBtn.Enabled = $false

$TagDeviceGroupBox.Controls.AddRange(@($DeviceTag, $TagDeviceBtn, $RemoveTagBtn))

$IsolateGroupBox = New-Object System.Windows.Forms.GroupBox
$IsolateGroupBox.Location = New-Object System.Drawing.Point(500, 165)
$IsolateGroupBox.Width = 400
$IsolateGroupBox.Height = 60
$IsolateGroupBox.Text = "Network Containment Actions"
$IsolateGroupBox.Font = 'Microsoft Sans Serif,10,style=Bold'

$NetworkContainDeviceBtn = New-Object System.Windows.Forms.Button
$NetworkContainDeviceBtn.BackColor = $UnclickableColour
$NetworkContainDeviceBtn.Text = "Network Contain Devices"
$NetworkContainDeviceBtn.Width = 180
$NetworkContainDeviceBtn.Height = 30
$NetworkContainDeviceBtn.Location = New-Object System.Drawing.Point(20, 20)
$NetworkContainDeviceBtn.Font = 'Microsoft Sans Serif,10'
$NetworkContainDeviceBtn.ForeColor = "#ffffff"
$NetworkContainDeviceBtn.Visible = $true
$NetworkContainDeviceBtn.Enabled = $false

$ReleaseFromIsolationBtn = New-Object System.Windows.Forms.Button
$ReleaseFromIsolationBtn.BackColor = $UnclickableColour
$ReleaseFromIsolationBtn.Text = "Release Devices"
$ReleaseFromIsolationBtn.Width = 180
$ReleaseFromIsolationBtn.Height = 30
$ReleaseFromIsolationBtn.Location = New-Object System.Drawing.Point(210, 20)
$ReleaseFromIsolationBtn.Font = 'Microsoft Sans Serif,10'
$ReleaseFromIsolationBtn.ForeColor = "#ffffff"
$ReleaseFromIsolationBtn.Visible = $true
$ReleaseFromIsolationBtn.Enabled = $false

$IsolateGroupBox.Controls.AddRange(@($NetworkContainDeviceBtn, $ReleaseFromIsolationBtn))

$InputCsvFileBox = New-Object System.Windows.Forms.GroupBox
$InputCsvFileBox.width = 880
$InputCsvFileBox.height = 240
$InputCsvFileBox.location = New-Object System.Drawing.Point(20, 290)
$InputCsvFileBox.text = "2 - Select devices to perform action on (CSV)"
$InputCsvFileBox.Font = 'Microsoft Sans Serif,12,style=Bold'

$GetDevicesFromQueryBtn = New-Object System.Windows.Forms.Button
$GetDevicesFromQueryBtn.BackColor = $UnclickableColour
$GetDevicesFromQueryBtn.text = "Get Devices"
$GetDevicesFromQueryBtn.width = 180
$GetDevicesFromQueryBtn.height = 30
$GetDevicesFromQueryBtn.location = New-Object System.Drawing.Point(690, 190)
$GetDevicesFromQueryBtn.Font = 'Microsoft Sans Serif,10'
$GetDevicesFromQueryBtn.ForeColor = "#ffffff"
$GetDevicesFromQueryBtn.Visible = $true

$SelectedDevicesBtn = New-Object system.Windows.Forms.Button
$SelectedDevicesBtn.BackColor = $UnclickableColour
$SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
$SelectedDevicesBtn.width = 150
$SelectedDevicesBtn.height = 30
$SelectedDevicesBtn.location = New-Object System.Drawing.Point(530, 190)
$SelectedDevicesBtn.Font = 'Microsoft Sans Serif,10'
$SelectedDevicesBtn.ForeColor = "#ffffff"
$SelectedDevicesBtn.Visible = $false

$ClearSelectedDevicesBtn = New-Object system.Windows.Forms.Button
$ClearSelectedDevicesBtn.BackColor = $UnclickableColour
$ClearSelectedDevicesBtn.text = "Clear Selection"
$ClearSelectedDevicesBtn.width = 150
$ClearSelectedDevicesBtn.height = 30
$ClearSelectedDevicesBtn.location = New-Object System.Drawing.Point(370, 190)
$ClearSelectedDevicesBtn.Font = 'Microsoft Sans Serif,10'
$ClearSelectedDevicesBtn.ForeColor = "#ffffff"
$ClearSelectedDevicesBtn.Visible = $false

# CSV file picker controls (shown when InputRadioButton3 is selected)
$CsvPathBox = New-Object system.Windows.Forms.TextBox
$CsvPathBox.multiline = $false
$CsvPathBox.width = 700
$CsvPathBox.height = 25
$CsvPathBox.location = New-Object System.Drawing.Point(20, 60)
$CsvPathBox.Font = $TextBoxFont
$CsvPathBox.ReadOnly = $true
$CsvPathBox.Enabled = $false

$BrowseCsvBtn = New-Object system.Windows.Forms.Button
$BrowseCsvBtn.BackColor = $UnclickableColour
$BrowseCsvBtn.text = "Browse..."
$BrowseCsvBtn.width = 90
$BrowseCsvBtn.height = 25
$BrowseCsvBtn.location = New-Object System.Drawing.Point(730, 60)
$BrowseCsvBtn.Font = 'Microsoft Sans Serif,9'
$BrowseCsvBtn.ForeColor = "#ffffff"
$BrowseCsvBtn.Visible = $false
$BrowseCsvBtn.Enabled = $false

# add a label with short description of what to do with the CSV (should have "Name" header and only hostnames) under the browse button
$CsvDescLabel = New-Object system.Windows.Forms.Label
$CsvDescLabel.text = "Select a CSV file with a 'Name' header (one single column) containing hostnames (one per line)."
$CsvDescLabel.width = 700
$CsvDescLabel.height = 40
$CsvDescLabel.location = New-Object System.Drawing.Point(20, 90)
$CsvDescLabel.Font = 'Microsoft Sans Serif,9'
# use a visible colour (black) on the white form background
$CsvDescLabel.ForeColor = "#000000"
$CsvDescLabel.Visible = $true

# OpenFileDialog for CSV selection
$OpenCsvDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenCsvDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenCsvDialog.Multiselect = $false

$CsvPathBox.Visible = $true
$BrowseCsvBtn.Visible = $true

$InputCsvFileBox.Controls.AddRange(@(
        $CsvPathBox,
        $BrowseCsvBtn,
        $CsvDescLabel,
        $GetDevicesFromQueryBtn,
        $SelectedDevicesBtn,
        $ClearSelectedDevicesBtn
    ))

$LogBoxLabel = New-Object system.Windows.Forms.Label
$LogBoxLabel.text = "4 - Logs:"
$LogBoxLabel.width = 394
$LogBoxLabel.height = 20
$LogBoxLabel.location = New-Object System.Drawing.Point(20, 600)
$LogBoxLabel.Font = 'Microsoft Sans Serif,12,style=Bold'
$LogBoxLabel.Visible = $true

$LogBox = New-Object system.Windows.Forms.TextBox
$LogBox.multiline = $true
$LogBox.width = 880
$LogBox.height = 100
$LogBox.location = New-Object System.Drawing.Point(20, 630)
$LogBox.ScrollBars = 'Vertical'
$LogBox.Font = $TextBoxFont
$LogBox.Visible = $true

$ExportLogBtn = New-Object system.Windows.Forms.Button
$ExportLogBtn.BackColor = '#FFF0F8FF'
$ExportLogBtn.text = "Export Logs"
$ExportLogBtn.width = 90
$ExportLogBtn.height = 30
$ExportLogBtn.location = New-Object System.Drawing.Point(20, 750)
$ExportLogBtn.Font = 'Microsoft Sans Serif,10'
$ExportLogBtn.ForeColor = "#ff000000"
$ExportLogBtn.Visible = $true

$cancelBtn = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor = '#FFF0F8FF'
$cancelBtn.text = "Cancel"
$cancelBtn.width = 90
$cancelBtn.height = 30
$cancelBtn.location = New-Object System.Drawing.Point(810, 750)
$cancelBtn.Font = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor = "#ff000000"
$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$MainForm.CancelButton = $cancelBtn
$MainForm.Controls.Add($cancelBtn)

#$MainForm.AutoScaleMode = 'dpi'

$MainForm.controls.AddRange(@($Title,
        $ConnectionStatusLabel, 
        $ConnectionStatus,
        $cancelBtn, 
        $AppIdBox, 
        $AppSecretBox,
        $CidBox, 
        $AppIdBoxLabel, 
        $AppSecretBoxLabel, 
        $CidBoxLabel, 
        $ConnectBtn,
        $TenantNoteLabel,
        $TitleActions, 
        $LogBoxLabel, 
        $LogBox, 
        $IsolateGroupBox,
        $SaveCredCheckbox,
        $ScanGroupBox,
        $InputCsvFileBox,
        $TagDeviceGroupBox,
        $ExportLogBtn
    ))


#===========================================================[Functions]===========================================================


#Authentication - Get Crowdstrike token

function GetToken {
    $ConnectionStatus.ForeColor = "#000000"
    $ConnectionStatus.Text = 'Connecting...'
    $appId = $AppIdBox.Text
    $appSecret = $AppSecretBox.Text
    $cid = $CidBox.Text.Trim()

    $oAuthUri = "https://api.eu-1.crowdstrike.com/oauth2/token"
    $authBody = @{
        client_id     = $appId
        client_secret = $appSecret
        grant_type    = 'client_credentials'
    }

    if ($cid) {
        $authBody['member_cid'] = $cid
    }

    try {
        $authResponse = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $authBody -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    }
    catch {
        $ConnectionStatus.text = "Connection Failed"
        $ConnectionStatus.ForeColor = "#D0021B"
        $cancelBtn.text = "Close"
        if ($_.Exception.Response -and ($_.Exception.Response.Content)) {
            $msg = $_.Exception.Response.Content
        }
        else {
            $msg = $_.Exception.Message
        }
        [System.Windows.Forms.MessageBox]::Show("Error obtaining CrowdStrike token: $msg", "Error")
        return
    }

    $token = $authResponse.access_token
    if (-not $token) {
        $ConnectionStatus.text = "Connection Failed"
        $ConnectionStatus.ForeColor = "#D0021B"
        [System.Windows.Forms.MessageBox]::Show("No access token returned from CrowdStrike.", "Error")
        return
    }

    $script:headers = @{
        'Content-Type' = 'application/json'
        Accept         = 'application/json'
        Authorization  = "Bearer $token"
    }

    $ConnectionStatus.text = "Connected"
    $ConnectionStatus.ForeColor = "#7ed321"
    $LogBox.AppendText((get-date).ToString() + " Successfully connected to CrowdStrike (client_id: " + $appId + $(if ($cid) { ", cid: $cid" } else { "" }) + ")" + [Environment]::NewLine)
    # show token in logs for debugging purposes (remove in production)
    # $LogBox.AppendText((get-date).ToString() + " Token: " + $token + [Environment]::NewLine)
    ChangeButtonColours -Buttons $GetDevicesFromQueryBtn, $SelectedDevicesBtn, $ClearSelectedDevicesBtn, $BrowseCsvBtn
    SaveCreds
    $Devicetag.Enabled = $true
    $CsvPathBox.Enabled = $true
    $BrowseCsvBtn.Enabled = $true

    return $script:headers
}

function SaveCreds {
    if ($SaveCredCheckbox.Checked) {
        $securespassword = $AppSecretBox.Text | ConvertTo-SecureString -AsPlainText -Force
        $securestring = $securespassword | ConvertFrom-SecureString
        $creds = @($CidBox.Text, $AppIdBox.Text, $securestring)
        $creds | Out-File $credspath
    }
}

function ChangeButtonColours {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        $Buttons
    )
    $ButtonsToChangeColour = $Buttons

    foreach ( $Button in $ButtonsToChangeColour) {
        $Button.BackColor = $ClickableColour
    }
}

function AddTagDevice {
    # Validate selection and tag
    if (-not $script:selectedmachines -or $script:selectedmachines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No devices selected.", "Info")
        return
    }

    $MachineTag = $DeviceTag.Text.Trim()
    if (-not $MachineTag) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a tag to apply.", "Info")
        return
    }

    # if tag does not start with "FalconGroupingTags/", tell user to rewrite it because this is required by Crowdstrike API
    if ($MachineTag -notmatch '^FalconGroupingTags\/') {
        [System.Windows.Forms.MessageBox]::Show("Tag must begin with 'FalconGroupingTags/'. Please rewrite the tag.", "Info")
        return
    }

    # Get authentication headers (must contain Authorization). Ensure Accept header present.
    $authHeaders = if ($script:headers) { @($script:headers) } elseif ($headers) { @($headers) } else {
        [System.Windows.Forms.MessageBox]::Show("Not connected. Please connect first.", "Error")
        return
    }

    # Normalize headers hashtable (in case it's an array-wrapped hashtable)
    if ($authHeaders -is [System.Collections.Hashtable] -eq $false) {
        $authHeaders = @{}
        foreach ($k in $script:headers.Keys) { $authHeaders[$k] = $script:headers[$k] }
    }

    if (-not $authHeaders.ContainsKey('Accept')) {
        $authHeaders['Accept'] = 'application/json'
    }

    # Ensure Authorization begins with "Bearer "
    if ($authHeaders.ContainsKey('Authorization')) {
        if ($authHeaders['Authorization'] -notmatch '(?i)^Bearer\s') {
            $authHeaders['Authorization'] = "Bearer $($authHeaders['Authorization'])"
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Missing Authorization header. Please connect first.", "Error")
        return
    }

    $allIds = $script:selectedmachines.Values | ForEach-Object { $_.ToString() } | Select-Object -Unique
    $total = $allIds.Count

    if ($total -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No device ids found to tag.", "Info")
        return
    }

    # Enforce limitation: do not proceed if 5000 or more devices
    if ($total -ge 5000) {
        [System.Windows.Forms.MessageBox]::Show("Too many devices selected ($total). Due to API limits please use smaller CSV files and run the script multiple times with subsets (e.g. split CSV into multiple files).", "Limit Exceeded")
        $LogBox.AppendText((Get-Date).ToString() + " Tagging aborted: $total devices selected (limit is < 5000). User instructed to split CSV and retry." + [Environment]::NewLine)
        return
    }

    $url = "https://api.eu-1.crowdstrike.com/devices/entities/devices/tags/v1"
    $body = @{
        device_ids = @($allIds)
        action     = "add"
        tags       = @($MachineTag)
    }

    # show the query in logs for debugging purposes (remove in production)
    # $LogBox.AppendText((Get-Date).ToString() + " Tagging $total device(s) with tag '$MachineTag'. API URL: $url. Body: " + ($body | ConvertTo-Json -Depth 5) + [Environment]::NewLine)

    try {
        # Use Invoke-RestMethod to perform the PATCH (mirrors the curl behavior)
        $response = Invoke-RestMethod -Method Patch -Uri $url -Headers $authHeaders -Body ($body | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop

        # Log success. If API returns a body include a compacted representation.
        if ($null -ne $response) {
            $respText = ($response | Out-String).Trim() -replace "\s+", " "
            $LogBox.AppendText((Get-Date).ToString() + " Applied tag '$MachineTag' to $total device(s). Response: " + $respText + [Environment]::NewLine)
        }
        else {
            $LogBox.AppendText((Get-Date).ToString() + " Applied tag '$MachineTag' to $total device(s). No body returned." + [Environment]::NewLine)
        }

        [System.Windows.Forms.MessageBox]::Show("Tag operation completed for $total device(s). See logs for details.", "Info")
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.Exception.Response -and ($_.Exception.Response.Content)) {
            try {
                $raw = $_.Exception.Response.Content
                $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
                $errMsg = ($parsed | Out-String).Trim()
            }
            catch {
                # fallback to raw content string
                $errMsg = $_.Exception.Response.Content
            }
        }

        $LogBox.AppendText((Get-Date).ToString() + " Failed to apply tag '$MachineTag' to $total device(s). Error: " + ($errMsg -replace "\s+", " ") + [Environment]::NewLine)
        [System.Windows.Forms.MessageBox]::Show("Error applying tag to devices: " + $errMsg, "Error")
    }
}

function RemoveTagDevice {
    # Validate selection and tag
    if (-not $script:selectedmachines -or $script:selectedmachines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No devices selected.", "Info")
        return
    }

    $MachineTag = $DeviceTag.Text.Trim()
    if (-not $MachineTag) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a tag to remove.", "Info")
        return
    }

    # if tag does not start with "FalconGroupingTags/", tell user to rewrite it because this is required by Crowdstrike API
    if ($MachineTag -notmatch '^FalconGroupingTags\/') {
        [System.Windows.Forms.MessageBox]::Show("Tag must begin with 'FalconGroupingTags/'. Please rewrite the tag.", "Info")
        return
    }

    # Get authentication headers (must contain Authorization). Ensure Accept header present.
    $authHeaders = if ($script:headers) { @($script:headers) } elseif ($headers) { @($headers) } else {
        [System.Windows.Forms.MessageBox]::Show("Not connected. Please connect first.", "Error")
        return
    }

    # Normalize headers hashtable (in case it's an array-wrapped hashtable)
    if ($authHeaders -is [System.Collections.Hashtable] -eq $false) {
        $authHeaders = @{}
        foreach ($k in $script:headers.Keys) { $authHeaders[$k] = $script:headers[$k] }
    }

    if (-not $authHeaders.ContainsKey('Accept')) {
        $authHeaders['Accept'] = 'application/json'
    }

    # Ensure Authorization begins with "Bearer "
    if ($authHeaders.ContainsKey('Authorization')) {
        if ($authHeaders['Authorization'] -notmatch '(?i)^Bearer\s') {
            $authHeaders['Authorization'] = "Bearer $($authHeaders['Authorization'])"
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Missing Authorization header. Please connect first.", "Error")
        return
    }

    $allIds = $script:selectedmachines.Values | ForEach-Object { $_.ToString() } | Select-Object -Unique
    $total = $allIds.Count

    if ($total -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No device ids found to remove tag from.", "Info")
        return
    }

    # Enforce limitation: do not proceed if 5000 or more devices
    if ($total -ge 5000) {
        [System.Windows.Forms.MessageBox]::Show("Too many devices selected ($total). Due to API limits please use smaller CSV files and run the script multiple times with subsets (e.g. split CSV into multiple files).", "Limit Exceeded")
        $LogBox.AppendText((Get-Date).ToString() + " Remove tag aborted: $total devices selected (limit is < 5000). User instructed to split CSV and retry." + [Environment]::NewLine)
        return
    }

    $url = "https://api.eu-1.crowdstrike.com/devices/entities/devices/tags/v1"
    $body = @{
        device_ids = @($allIds)
        action     = "remove"
        tags       = @($MachineTag)
    }

    try {
        $response = Invoke-RestMethod -Method Patch -Uri $url -Headers $authHeaders -Body ($body | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop

        if ($null -ne $response) {
            $respText = ($response | Out-String).Trim() -replace "\s+", " "
            $LogBox.AppendText((Get-Date).ToString() + " Removed tag '$MachineTag' from $total device(s). Response: " + $respText + [Environment]::NewLine)
        }
        else {
            $LogBox.AppendText((Get-Date).ToString() + " Removed tag '$MachineTag' from $total device(s). No body returned." + [Environment]::NewLine)
        }

        [System.Windows.Forms.MessageBox]::Show("Remove tag operation completed for $total device(s). See logs for details.", "Info")
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.Exception.Response -and ($_.Exception.Response.Content)) {
            try {
                $raw = $_.Exception.Response.Content
                $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
                $errMsg = ($parsed | Out-String).Trim()
            }
            catch {
                $errMsg = $_.Exception.Response.Content
            }
        }

        $LogBox.AppendText((Get-Date).ToString() + " Failed to remove tag '$MachineTag' from $total device(s). Error: " + ($errMsg -replace "\s+", " ") + [Environment]::NewLine)
        [System.Windows.Forms.MessageBox]::Show("Error removing tag from devices: " + $errMsg, "Error")
    }
}

# This function uses Crowdstrike API to isolate devices but has not been tested yet
function NetworkContainDevice {
    # Use CrowdStrike "contain" action on selected device AIDs (ids). Batches of 99.
    if (-not $script:selectedmachines -or $script:selectedmachines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No devices selected.", "Info")
        return
    }

    # Normalize/auth headers
    $authHeaders = if ($script:headers) { @($script:headers) } elseif ($headers) { @($headers) } else {
        [System.Windows.Forms.MessageBox]::Show("Not connected. Please connect first.", "Error")
        return
    }
    if ($authHeaders -isnot [hashtable]) {
        $tmp = @{}
        foreach ($k in $script:headers.Keys) { $tmp[$k] = $script:headers[$k] }
        $authHeaders = $tmp
    }
    if (-not $authHeaders.ContainsKey('Accept')) { $authHeaders['Accept'] = 'application/json' }
    if ($authHeaders.ContainsKey('Authorization')) {
        if ($authHeaders['Authorization'] -notmatch '(?i)^Bearer\s') {
            $authHeaders['Authorization'] = "Bearer $($authHeaders['Authorization'])"
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Missing Authorization header. Please connect first.", "Error")
        return
    }

    $allIds = $script:selectedmachines.Values | ForEach-Object { $_.ToString() } | Select-Object -Unique
    $total = $allIds.Count
    if ($total -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No device ids found to contain.", "Info")
        return
    }

    $LogBox.AppendText((Get-Date).ToString() + " Starting contain action for $total device(s)..." + [Environment]::NewLine)

    $urlBase = "https://api.eu-1.crowdstrike.com/devices/entities/devices-actions/v2"
    $batchSize = 99

    for ($i = 0; $i -lt $allIds.Count; $i += $batchSize) {
        $end = [math]::Min($allIds.Count - 1, $i + $batchSize - 1)
        # force $allIds to be an array before slicing to avoid getting characters or unexpected elements
        $batch = @($allIds)[$i..$end]
        # pass the batch array directly so ConvertTo-Json serializes full ID strings
        $body = @{ ids = $batch }

        $LogBox.AppendText((Get-Date).ToString() + " Sending contain request for batch " + ([int]($i / $batchSize) + 1) + " (count: " + $batch.Count + ")." + [Environment]::NewLine)

        try {
            $resp = Invoke-RestMethod -Method Post -Uri ($urlBase + "?action_name=contain") -Headers $authHeaders -Body ($body | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop

            # Log a compact response if present
            if ($null -ne $resp) {
                $respText = ($resp | Out-String).Trim() -replace "\s+", " "
                $LogBox.AppendText((Get-Date).ToString() + " Batch contained. Response: " + $respText + [Environment]::NewLine)
            }
            else {
                $LogBox.AppendText((Get-Date).ToString() + " Batch contained. No body returned." + [Environment]::NewLine)
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($_.Exception.Response -and ($_.Exception.Response.Content)) {
                try {
                    $raw = $_.Exception.Response.Content
                    $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
                    $errMsg = ($parsed | Out-String).Trim() -replace "\s+", " "
                }
                catch {
                    $errMsg = $_.Exception.Response.Content
                }
            }
            $LogBox.AppendText((Get-Date).ToString() + " Failed to contain batch starting at index $i. Error: " + $errMsg + [Environment]::NewLine)
            [System.Windows.Forms.MessageBox]::Show("Error performing contain action: " + $errMsg, "Error")
            # continue to next batch (don't abort entirely)
        }

        Start-Sleep -Milliseconds 500
    }

    $LogBox.AppendText((Get-Date).ToString() + " Contain operation completed for $total device(s)." + [Environment]::NewLine)
    [System.Windows.Forms.MessageBox]::Show("Contain operation completed for $total device(s). See logs for details.", "Info")
}

# This function uses Crowdstrike API to release devices from isolation but has not been tested yet
function ReleaseFromIsolation {
    # Use CrowdStrike "lift_containment" action on selected device AIDs (ids). Batches of 99.
    if (-not $script:selectedmachines -or $script:selectedmachines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No devices selected.", "Info")
        return
    }

    # Normalize/auth headers
    $authHeaders = if ($script:headers) { @($script:headers) } elseif ($headers) { @($headers) } else {
        [System.Windows.Forms.MessageBox]::Show("Not connected. Please connect first.", "Error")
        return
    }
    if ($authHeaders -isnot [hashtable]) {
        $tmp = @{}
        foreach ($k in $script:headers.Keys) { $tmp[$k] = $script:headers[$k] }
        $authHeaders = $tmp
    }
    if (-not $authHeaders.ContainsKey('Accept')) { $authHeaders['Accept'] = 'application/json' }
    if ($authHeaders.ContainsKey('Authorization')) {
        if ($authHeaders['Authorization'] -notmatch '(?i)^Bearer\s') {
            $authHeaders['Authorization'] = "Bearer $($authHeaders['Authorization'])"
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Missing Authorization header. Please connect first.", "Error")
        return
    }

    $allIds = $script:selectedmachines.Values | ForEach-Object { $_.ToString() } | Select-Object -Unique
    $total = $allIds.Count
    if ($total -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No device ids found to lift containment.", "Info")
        return
    }

    $LogBox.AppendText((Get-Date).ToString() + " Starting lift_containment action for $total device(s)..." + [Environment]::NewLine)

    $urlBase = "https://api.eu-1.crowdstrike.com/devices/entities/devices-actions/v2"
    $batchSize = 99

    for ($i = 0; $i -lt $allIds.Count; $i += $batchSize) {
        $end = [math]::Min($allIds.Count - 1, $i + $batchSize - 1)
        # force $allIds to be an array before slicing to avoid getting characters or unexpected elements
        $batch = @($allIds)[$i..$end]
        # pass the batch array directly so ConvertTo-Json serializes full ID strings
        $body = @{ ids = $batch }

        $LogBox.AppendText((Get-Date).ToString() + " Sending lift_containment request for batch " + ([int]($i / $batchSize) + 1) + " (count: " + $batch.Count + ")." + [Environment]::NewLine)

        try {
            $resp = Invoke-RestMethod -Method Post -Uri ($urlBase + "?action_name=lift_containment") -Headers $authHeaders -Body ($body | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop

            if ($null -ne $resp) {
                $respText = ($resp | Out-String).Trim() -replace "\s+", " "
                $LogBox.AppendText((Get-Date).ToString() + " Batch lift_containment completed. Response: " + $respText + [Environment]::NewLine)
            }
            else {
                $LogBox.AppendText((Get-Date).ToString() + " Batch lift_containment completed. No body returned." + [Environment]::NewLine)
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($_.Exception.Response -and ($_.Exception.Response.Content)) {
                try {
                    $raw = $_.Exception.Response.Content
                    $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
                    $errMsg = ($parsed | Out-String).Trim() -replace "\s+", " "
                }
                catch {
                    $errMsg = $_.Exception.Response.Content
                }
            }
            $LogBox.AppendText((Get-Date).ToString() + " Failed to lift containment for batch starting at index $i. Error: " + $errMsg + [Environment]::NewLine)
            [System.Windows.Forms.MessageBox]::Show("Error performing lift_containment action: " + $errMsg, "Error")
            # continue to next batch
        }

        Start-Sleep -Milliseconds 500
    }

    $LogBox.AppendText((Get-Date).ToString() + " Lift containment operation completed for $total device(s)." + [Environment]::NewLine)
    [System.Windows.Forms.MessageBox]::Show("Lift containment operation completed for $total device(s). See logs for details.", "Info")
}



# Implemented with Crowdstrike API results
function ViewSelectedDevices {
    $filtermachines = $script:selectedmachines | Out-GridView -Title "Select devices to perform action on:" -PassThru 
    $script:selectedmachines.clear()
    foreach ($machine in $filtermachines) {
        $script:selectedmachines.Add($machine.Name, $machine.Value)
    }
    $SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
    if ($null -eq $script:selectedmachines.Keys.Count) {
        $SelectedDevicesBtn.Visible = $false
        $SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
        $ClearSelectedDevicesBtn.Visible = $false
    }
    $LogBox.AppendText((get-date).ToString() + " Devices selected count: " + ($script:selectedmachines.Keys.count -join [Environment]::NewLine) + [Environment]::NewLine + ($script:selectedmachines.Keys -join [Environment]::NewLine) + [Environment]::NewLine)
}

function ClearSelectedDevices {
    $script:selectedmachines = @{}
    $ClearSelectedDevicesBtn.Visible = $false
    $SelectedDevicesBtn.Visible = $false
    $LogBox.AppendText((get-date).ToString() + " Devices selected count: " + $script:selectedmachines.Keys.count + [Environment]::NewLine)
}

# Implemented with Crowdstrike API results
function GetDevicesFromCsv {
    if ((Test-Path $CsvPathBox.Text) -and ($CsvPathBox.Text).EndsWith(".csv")) {
        $machines = Import-Csv -Path $CsvPathBox.Text
        $script:selectedmachines = @{}
        $LogBox.AppendText((get-date).ToString() + " Querying " + $machines.count + " machines from CSV file... Please wait" + [Environment]::NewLine)

        foreach ($machine in $machines) {
            $Hostname = $machine.Name.Trim()
            if (-not $Hostname) { continue }

            # CrowdStrike devices query endpoint: filter on hostname

            $encodedFilter = [uri]::EscapeDataString("hostname:'$Hostname'")
            $fullUri = "https://api.eu-1.crowdstrike.com/devices/queries/devices/v1?filter=$encodedFilter&sort=last_seen.desc&limit=1"

            # write in the logs the query being made (for debugging purposes)
            # $LogBox.AppendText((get-date).ToString() + " Querying URI: " + $fullUri + [Environment]::NewLine)

            try {
                $resp = Invoke-RestMethod -Method Get -Uri $fullUri -Headers $headers -ErrorAction Stop
            }
            catch {
                $LogBox.AppendText((get-date).ToString() + " Error querying hostname '$Hostname': " + $_.Exception.Message + [Environment]::NewLine)
                continue
            }

            $aids = @()
            if ($null -ne $resp.resources) {
                # resources usually contains AIDs (agent IDs)
                $aids = $resp.resources
            }

            if ($aids.Count -eq 0) {
                $LogBox.AppendText((get-date).ToString() + " No AID found for hostname: $Hostname" + [Environment]::NewLine)
                continue
            }

            # If multiple AIDs returned for the same hostname, add each as a separate selectable entry
            foreach ($aid in $aids) {
                $key = if ($aids.Count -gt 1) { "$Hostname - $aid" } else { $Hostname }
                if (-not $script:selectedmachines.ContainsKey($key)) {
                    $script:selectedmachines.Add($key, $aid)
                    $LogBox.AppendText((get-date).ToString() + " Found AID for $Hostname : $aid" + [Environment]::NewLine)
                }
            }
        }

        if ($script:selectedmachines.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No devices found for hostnames in CSV.", "Info")
            return
        }

        # Let user pick which discovered devices/AIDs to act on
        $filtermachines = $script:selectedmachines.GetEnumerator() |
        ForEach-Object { [PSCustomObject]@{ Name = $_.Key; Value = $_.Value } } |
        Out-GridView -Title "Select devices to perform action on:" -PassThru

        $script:selectedmachines.Clear()
        foreach ($machine in $filtermachines) {
            $script:selectedmachines.Add($machine.Name, $machine.Value)
        }

        if ($script:selectedmachines.Keys.Count -gt 0) {
            ChangeButtonColours -Buttons $TagDeviceBtn, $RemoveTagBtn, $NetworkContainDeviceBtn, $ReleaseFromIsolationBtn
            $SelectedDevicesBtn.Visible = $true
            $SelectedDevicesBtn.Text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
            $ClearSelectedDevicesBtn.Visible = $true
            $TagDeviceBtn.Enabled = $true
            $RemoveTagBtn.Enabled = $true
            $NetworkContainDeviceBtn.Enabled = $true
            $ReleaseFromIsolationBtn.Enabled = $true
        }

        $LogBox.AppendText((get-date).ToString() + " Devices selected count: " + $script:selectedmachines.Keys.count + [Environment]::NewLine)
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($CsvPathBox.Text + " is not a valid CSV path." , "Error")
    }
}

function ExportLog {
    $LogBox.Text | Out-file .\crowdstrike_ui_log.txt
    $LogBox.AppendText((get-date).ToString() + " Log file created: " + (Get-Item .\crowdstrike_ui_log.txt).FullName + [Environment]::NewLine)
}

#===========================================================[Script]===========================================================


if (test-path $credspath) {
    $creds = Get-Content $credspath
    $pass = $creds[2] | ConvertTo-SecureString
    $unsecurePassword = [PSCredential]::new(0, $pass).GetNetworkCredential().Password
    $CidBox.Text = $creds[0]
    $AppIdBox.Text = $creds[1]
    $AppSecretBox.Text = $unsecurePassword
}

$ConnectBtn.Add_Click({ GetToken })

$TagDeviceBtn.Add_Click({ AddTagDevice })

$RemoveTagBtn.Add_Click({ RemoveTagDevice })

# set text to "FalconGroupingTags/" when clicking on the DeviceTag box for user convenience
$DeviceTag.Add_Click({
        if ($DeviceTag.Text -eq "") {
            $DeviceTag.Text = "FalconGroupingTags/"
        }
    })

$NetworkContainDeviceBtn.Add_Click({ NetworkContainDevice })

$ReleaseFromIsolationBtn.Add_Click({ ReleaseFromIsolation })

$GetDevicesFromQueryBtn.Add_Click({ GetDevicesFromCsv })

$BrowseCsvBtn.Add_Click({
        if ($OpenCsvDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $CsvPathBox.Text = $OpenCsvDialog.FileName
            # show selected path in log
            $LogBox.AppendText((get-date).ToString() + " Selected CSV: " + $CsvPathBox.Text + [Environment]::NewLine)
            # enable the GetDevicesFromQueryBtn so user can proceed
            $GetDevicesFromQueryBtn.BackColor = $ClickableColour
        }
    })

$SelectedDevicesBtn.Add_Click({ ViewSelectedDevices })

$ClearSelectedDevicesBtn.Add_Click({ ClearSelectedDevices })

$ExportLogBtn.Add_Click({ ExportLog })

$MainForm.ResumeLayout()
[void]$MainForm.ShowDialog()