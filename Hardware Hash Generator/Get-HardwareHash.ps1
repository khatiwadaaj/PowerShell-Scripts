function Get-WindowsAutoPilotInfo {

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
	[Parameter(Mandatory=$False,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=0)][alias("DNSHostName","ComputerName","Computer")] [String[]] $Name = @("localhost"),
	[Parameter(Mandatory=$False)] [String] $OutputFile = "", 
	[Parameter(Mandatory=$False)] [String] $GroupTag = "",
	[Parameter(Mandatory=$False)] [String] $AssignedUser = "",
	[Parameter(Mandatory=$False)] [Switch] $Append = $false,
	[Parameter(Mandatory=$False)] [System.Management.Automation.PSCredential] $Credential = $null,
	[Parameter(Mandatory=$False)] [Switch] $Partner = $false,
	[Parameter(Mandatory=$False)] [Switch] $Force = $false,
	[Parameter(Mandatory=$True,ParameterSetName = 'Online')] [Switch] $Online = $false,
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $TenantId = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AppId = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AppSecret = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AddToGroup = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [String] $AssignedComputerName = "",
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [Switch] $Assign = $false, 
	[Parameter(Mandatory=$False,ParameterSetName = 'Online')] [Switch] $Reboot = $false
)

Begin
{
	# Initialize empty list
	$computers = @()

	# If online, make sure we are able to authenticate
	if ($Online) {

		# Get NuGet
		$provider = Get-PackageProvider NuGet -ErrorAction Ignore
		if (-not $provider) {
			Write-Host "Installing provider NuGet"
			Find-PackageProvider -Name NuGet -ForceBootstrap -IncludeDependencies
		}
        
		# Get WindowsAutopilotIntune module (and dependencies)
		$module = Import-Module WindowsAutopilotIntune -MinimumVersion 5.4.0 -PassThru -ErrorAction Ignore
		if (-not $module) {
			Write-Host "Installing module WindowsAutopilotIntune"
			Install-Module WindowsAutopilotIntune -Force
		}
		Import-Module WindowsAutopilotIntune -Scope Global
		
        	# Get Graph Authentication module (and dependencies)
        	$module = Import-Module microsoft.graph.authentication -PassThru -ErrorAction Ignore
        	if (-not $module) {
            		Write-Host "Installing module microsoft.graph.authentication"
            		Install-Module microsoft.graph.authentication -Force
        	}
        	Import-Module microsoft.graph.authentication -Scope Global

		# Get required modules for AddToGroup switch
		if ($AddToGroup)
		{
			$module = Import-Module Microsoft.Graph.Groups -PassThru -ErrorAction Ignore
			if (-not $module)
			{
				Write-Host "Installing module Microsoft.Graph.Groups"
				Install-Module Microsoft.Graph.Groups -Force
			}

            		$module = Import-Module Microsoft.Graph.Identity.DirectoryManagement -PassThru -ErrorAction Ignore
			if (-not $module)
			{
				Write-Host "Installing module Microsoft.Graph.Identity.DirectoryManagement"
				Install-Module Microsoft.Graph.Identity.DirectoryManagement -Force
			}
		}

        	# Connect
		if ($AppId -ne "")
		{
			$graph = Connect-MSGraphApp -Tenant $TenantId -AppId $AppId -AppSecret $AppSecret
			Write-Host "Connected to Intune tenant $TenantId using app-based authentication (Azure AD authentication not supported)"
		}
		else {
			Connect-MgGraph -Scopes "DeviceManagementServiceConfig.ReadWrite.All", "DeviceManagementManagedDevices.ReadWrite.All", "Device.ReadWrite.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All"
            		$graph = Get-MgContext
			Write-Host "Connected to Intune tenant" $graph.TenantId
		}

		# Force the output to a file
		if ($OutputFile -eq "")
		{
			$OutputFile = "$($env:TEMP)\autopilot.csv"
		} 
	}
}

Process
{
	foreach ($comp in $Name)
	{
		$bad = $false

		# Get a CIM session
		if ($comp -eq "localhost") {
			$session = New-CimSession
		}
		else
		{
			$session = New-CimSession -ComputerName $comp -Credential $Credential
		}

		# Get the common properties.
		Write-Verbose "Checking $comp"
		$serial = (Get-CimInstance -CimSession $session -Class Win32_BIOS).SerialNumber

		# Get the hash (if available)
		$devDetail = (Get-CimInstance -CimSession $session -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'")
		if ($devDetail -and (-not $Force))
		{
			$hash = $devDetail.DeviceHardwareData
		}
		else
		{
			$bad = $true
			$hash = ""
		}

		# If the hash isn't available, get the make and model
		if ($bad -or $Force)
		{
			$cs = Get-CimInstance -CimSession $session -Class Win32_ComputerSystem
			$make = $cs.Manufacturer.Trim()
			$model = $cs.Model.Trim()
			if ($Partner)
			{
				$bad = $false
			}
		}
		else
		{
			$make = ""
			$model = ""
		}

		# Getting the PKID is generally problematic for anyone other than OEMs, so let's skip it here
		$product = ""

		# Depending on the format requested, create the necessary object
		if ($Partner)
		{
			# Create a pipeline object
			$c = New-Object psobject -Property @{
				"Device Serial Number" = $serial
				"Windows Product ID" = $product
				"Hardware Hash" = $hash
				"Manufacturer name" = $make
				"Device model" = $model
			}
			# From spec:
			#	"Manufacturer Name" = $make
			#	"Device Name" = $model

		}
		else
		{
			# Create a pipeline object
			$c = New-Object psobject -Property @{
				"Device Serial Number" = $serial
				"Windows Product ID" = $product
				"Hardware Hash" = $hash
			}
			
			if ($GroupTag -ne "")
			{
				Add-Member -InputObject $c -NotePropertyName "Group Tag" -NotePropertyValue $GroupTag
			}
			if ($AssignedUser -ne "")
			{
				Add-Member -InputObject $c -NotePropertyName "Assigned User" -NotePropertyValue $AssignedUser
			}
		}

		# Write the object to the pipeline or array
		if ($bad)
		{
			# Report an error when the hash isn't available
			Write-Error -Message "Unable to retrieve device hardware data (hash) from computer $comp" -Category DeviceError
		}
		elseif ($OutputFile -eq "")
		{
			$c
		}
		else
		{
			$computers += $c
			Write-Host "Gathered details for device with serial number: $serial"
		}

		Remove-CimSession $session
	}
}

End
{
	if ($OutputFile -ne "")
	{
		if ($Append)
		{
			if (Test-Path $OutputFile)
			{
				$computers += Import-CSV -Path $OutputFile
			}
		}
		if ($Partner)
		{
			$computers | Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash", "Manufacturer name", "Device model" | ConvertTo-CSV -NoTypeInformation | ForEach-Object {$_ -replace '"',''} | Out-File $OutputFile
		}
		elseif ($AssignedUser -ne "")
		{
			$computers | Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag", "Assigned User" | ConvertTo-CSV -NoTypeInformation | ForEach-Object {$_ -replace '"',''} | Out-File $OutputFile
		}
		elseif ($GroupTag -ne "")
		{
			$computers | Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" | ConvertTo-CSV -NoTypeInformation | ForEach-Object {$_ -replace '"',''} | Out-File $OutputFile
		}
		else
		{
			$computers | Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash" | ConvertTo-CSV -NoTypeInformation | ForEach-Object {$_ -replace '"',''} | Out-File $OutputFile
		}
	}
    if ($Online)
    {
        # Add the devices
		$importStart = Get-Date
		$imported = @()
		$computers | ForEach-Object {
			$imported += Add-AutopilotImportedDevice -serialNumber $_.'Device Serial Number' -hardwareIdentifier $_.'Hardware Hash' -groupTag $_.'Group Tag' -assignedUser $_.'Assigned User'
		}

		# Wait until the devices have been imported
		$processingCount = 1
		while ($processingCount -gt 0)
		{
			$current = @()
			$processingCount = 0
			$imported | ForEach-Object {
				$device = Get-AutopilotImportedDevice -id $_.id
				if ($device.state.deviceImportStatus -eq "unknown") {
					$processingCount = $processingCount + 1
				}
				$current += $device
			}
			$deviceCount = $imported.Length
			Write-Host "Waiting for $processingCount of $deviceCount to be imported"
			if ($processingCount -gt 0){
				Start-Sleep 30
			}
		}
		$importDuration = (Get-Date) - $importStart
		$importSeconds = [Math]::Ceiling($importDuration.TotalSeconds)
		$successCount = 0
		$current | ForEach-Object {
			Write-Host "$($device.serialNumber): $($device.state.deviceImportStatus) $($device.state.deviceErrorCode) $($device.state.deviceErrorName)"
			if ($device.state.deviceImportStatus -eq "complete") {
				$successCount = $successCount + 1
			}
		}
		Write-Host "$successCount devices imported successfully.  Elapsed time to complete import: $importSeconds seconds"
		
		# Wait until the devices can be found in Intune (should sync automatically)
		$syncStart = Get-Date
		$processingCount = 1
		while ($processingCount -gt 0)
		{
			$autopilotDevices = @()
			$processingCount = 0
			$current | ForEach-Object {
				if ($device.state.deviceImportStatus -eq "complete") {
					$device = Get-AutopilotDevice -id $_.state.deviceRegistrationId
					if (-not $device) {
						$processingCount = $processingCount + 1
					}
					$autopilotDevices += $device
				}	
			}
			$deviceCount = $autopilotDevices.Length
			Write-Host "Waiting for $processingCount of $deviceCount to be synced"
			if ($processingCount -gt 0){
				Start-Sleep 30
			}
		}
		$syncDuration = (Get-Date) - $syncStart
		$syncSeconds = [Math]::Ceiling($syncDuration.TotalSeconds)
		Write-Host "All devices synced.  Elapsed time to complete sync: $syncSeconds seconds"
        
        # Add the device to the specified AAD group
		if ($AddToGroup)
		{
			$aadGroup = Get-MgGroup -Filter "DisplayName eq '$AddToGroup'"
			if ($aadGroup)
			{
				$autopilotDevices | ForEach-Object {
					$aadDevice = Get-MgDevice -Search "deviceId:$($_.azureActiveDirectoryDeviceId)" -ConsistencyLevel eventual
					if ($aadDevice) {
						Write-Host "Adding device $($_.serialNumber) to group $AddToGroup"
						New-MgGroupMember -GroupId $($aadGroup.Id) -DirectoryObjectId $($aadDevice.Id)
                        			Write-Host "Added devices to group '$AddToGroup' $($aadGroup.Id)"
					}
					else {
						Write-Error "Unable to find Azure AD device with ID $($_.azureActiveDirectoryDeviceId)"
					}
				}				
			}
			else {
				Write-Error "Unable to find group $AddToGroup"
			}
		}

		# Assign the computer name 
		if ($AssignedComputerName -ne "")
		{
			$autopilotDevices | ForEach-Object {
				Set-AutopilotDevice -id $_.id -displayName $AssignedComputerName
			}
		}

		# Wait for assignment (if specified)
		if ($Assign)
		{
			$assignStart = Get-Date
			$processingCount = 1
			while ($processingCount -gt 0)
			{
				$processingCount = 0
				$autopilotDevices | ForEach-Object {
					$device = Get-AutopilotDevice -id $_.id -Expand
					if (-not ($device.deploymentProfileAssignmentStatus.StartsWith("assigned"))) {
						$processingCount = $processingCount + 1
					}
				}
				$deviceCount = $autopilotDevices.Length
				Write-Host "Waiting for $processingCount of $deviceCount to be assigned"
				if ($processingCount -gt 0){
					Start-Sleep 30
				}	
			}
			$assignDuration = (Get-Date) - $assignStart
			$assignSeconds = [Math]::Ceiling($assignDuration.TotalSeconds)
			Write-Host "Profiles assigned to all devices.  Elapsed time to complete assignment: $assignSeconds seconds"	
			if ($Reboot)
			{
				Restart-Computer -Force
			}
		}
	}
}

}


$env:Path +=”;c:\Program Files\WindowsPowerShell\Scripts”
New-Item -Type Directory -Path ".\HardwareHash" -Force

$computerBIOS = $(Get-WmiObject Win32_BIOS)
$SerialNumber = $computerBIOS.SerialNumber

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

Get-WindowsAutoPilotInfo -OutputFile ".\HardwareHash\$SerialNumber.csv"

Add-Type -AssemblyName System.Windows.Forms

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Notification"
$form.Size = New-Object System.Drawing.Size(400, 150)
$form.StartPosition = "CenterScreen"

# Create a label to display the message
$label = New-Object System.Windows.Forms.Label
$label.Text = "Hardware hash for $SerialNumber has been generated"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(75, 30)
$form.Controls.Add($label)

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100, 70)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

# Show the form
$form.ShowDialog()