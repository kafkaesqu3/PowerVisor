[CmdletBinding(DefaultParameterSetName="ServerList")]
Param(
	[Parameter(Position = 0,         ValueFromPipeline=$true)]
	[String]

	$serverList,
	
	[Parameter(Position = 1)]
	[ValidateSet('Basic', 'Roasting')]
	[String]
	$Mode = 'Basic',
	
	[Parameter(ParameterSetName = 'Credential')]
	[Management.Automation.PSCredential]
        [Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,
	
	[Parameter(Position = 3)]
	[Switch]
	$NoExcel
)


$Path = 'C:\users\david\desktop\PowerVisor-output\'

### Main script begin ###
# Test path
If (!(Test-Path $path)) {
	Write-Host "Path not found - $path" -ForegroundColor Red
	Exit
}

$ServerList = '10.210.1.90'
#,'10.210.1.96','10.210.1.30','10.210.1.31'

# Obtain Credentials

#$creds = $Credential
$creds = Get-Credential -Message "Enter ESX/ESXi host credentials used to connect"

# Cycle through servers gathering and exporting info
Foreach ($ESX in $ServerList) {
	# Create full path to Excel workbook
	# Connect to ESX/ESXi host
	try {
        $connection = Connect-VIServer $ESX -Credential $Credential
    }
    catch
    {

    }

	# Ensure connection
	If ($connection) {	
		Write-Host "Connected to $ESX"
		# Get ESX(i) Host Info
		$VMHost = Get-VMHost $ESX
		
		$VMHostProp = [ordered]@{}
		$VMHostProp.Name = $VMHost.NetworkInfo.HostName
		$VMHostProp."Domain Name" = $VMHost.NetworkInfo.DomainName
        $VMHostProp."IP address" = $ESX
		$VMHostProp.Hypervisor = $VMHost.Extensiondata.Config.Product.FullName
		$VMHostProp."CPU Current Usage (Mhz)" = $VMHost.CpuUsageMhz
		$VMHostProp."Memory (GB)" = [Math]::Round($VMHost.MemoryTotalGB,0)
		$VMHostProp."Memory Current Usage (GB)" = [Math]::Round($VMHost.MemoryUsageGB,0)
		$VMHostProp."Physical NICs" = $VMHost.ExtensionData.Config.Network.Pnic.Count
		If ($VMHost.NetworkInfo.VMKernelGateway) {
			$VMHostProp."VMKernel Gateway" = $VMHost.NetworkInfo.VMKernelGateway
		} Else {
			$VMHostProp."Console Gateway" = $VMHost.NetworkInfo.ConsoleGateway
		}
		Try {
			$domainStatus = ""
			$domainStatus = $VMHost | Get-VMHostAuthentication -ErrorAction Stop
			$VMHostProp."Authentication Domain" = If ($domainStatus.Domain){$domainStatus.Domain} Else {"Not Configured"}
			$VMHostProp."Authentication Domain Status" = If ($domainStatus.DomainMembershipStatus){$domainStatus.DomainMembershipStatus} Else {"Unknown"}
			$VMHostProp."Trusted Domains" = If ($domainStatus.Trusteddomains) {$domainStatus | Select -ExpandProperty TrustedDomains | Out-String} Else {"None"}
		} Catch {}
		$VMHostProp."DNS Server(s)" = $VMHost.ExtensionData.Config.Network.DnsConfig.Address | Out-String
		
		Try {
			$VMHostProp."Syslog Server(s)" = $VMHost | Get-VMHostSysLogServer -ErrorAction Stop | %{$_.Host + ":" + $_.Port} | Out-String
		} Catch {
			$VMHostProp."Syslog Server(s)" = "Unavailable"
		}
		Try {
			$VMHostProp."Firewall Default Allow Incoming" = $VMHost | Get-VMHostFirewallDefaultPolicy -ErrorAction Stop | Select -ExpandProperty IncomingEnabled
			$VMHostProp."Firewall Default Allow Outgoing" = $VMHost | Get-VMHostFirewallDefaultPolicy -ErrorAction Stop | Select -ExpandProperty OutgoingEnabled
		} Catch {}
		$VMHostProp."VM Count" = ($VMHost | Get-VM).Count
        
		$VMHostInfo = New-Object –TypeName PSObject –Prop $VMHostProp
		

	    $FullPath = $Path + "esxi.csv"
        $VMHostInfo | Export-csv -Path $FullPath -append -NoTypeInformation









        # Get Datastore Info
		$DSInfo = @()
		$DSs = $VMHost | Get-Datastore
		Foreach ($DS in $DSs) {
			$DSProp = [ordered]@{}
			$DSProp.Name = $DS.Name
			$DSProp.Type = $DS.Type
			$DSProp."Size (GB)" = [Math]::Round($DS.CapacityGB,2)
			$DSProp."Used (GB)" = [Math]::Round($DS.CapacityGB - $DS.FreeSpaceGB,2)
			$DSProp."Free (GB)" = [Math]::Round($DS.FreeSpaceGB,2)
			$usedPerc = ($DS.CapacityGB - $DS.FreeSpaceGB)/$DS.CapacityGB
			$DSProp."Used %" = "{0:P1}" -f $usedPerc
			$DSProp."Mount Point" = $DS.ExtensionData.Info.Url
			$DSProp."VM Count" = ($DS | Get-VM).Count
			$DSProp.State = $DS.State
			
			$DSTemp = New-Object –TypeName PSObject –Prop $DSProp
			$DSInfo += $DSTemp
		}
		If ($DSInfo) {
			$DSInfo | Sort Name | Export-Xlsx -Path $FullPath -WorksheetName "Datastores" -Title "Datastores" -AppendWorksheet -SheetPosition "end"

	    $FullPath = $Path +  "datastores.csv"
        $DSInfo | Export-csv -Path $FullPath -append -NoTypeInformation

		}






# Get Firewall Exceptions
		$fwInfo = @()
		Try {
			$fwExceptions = $VMHost | Get-VMHostFirewallException -ErrorAction Stop
			Foreach ($fwException in $fwExceptions) {
				$fwProp = [ordered]@{}
				$fwProp.Enabled = $fwException.Enabled
				$fwProp.Name = $fwException.Name	
				$fwProp."Incoming Ports" = $fwException.IncomingPorts
				$fwProp."Outgoing Ports" = $fwException.OutgoingPorts
				$fwProp."Protocol(s)" = $fwException.Protocols
				$fwProp."Allow all IPs"  = $fwException.ExtensionData.AllowedHosts.AllIp
				$fwProp.Service = $fwException.ExtensionData.Service
				$fwProp."Service Running" = $fwException.ServiceRunning
				
				$fwTemp = New-Object –TypeName PSObject –Prop $fwProp
				$fwInfo += $fwTemp
			}
		} Catch {}
		If ($fwInfo) {

	    $FullPath = $Path +  "firewall_info.csv"
        $fwInfo | Export-csv -Path $FullPath -append -NoTypeInformation
			
		}






		# Get Services Information
		$svcInfo = @()
		$services = $VMHost | Get-VMHostService
		Foreach ($service in $services) {
			$svcProp = [ordered]@{}
			$svcProp.Service = $service.Label
			$svcProp.Key = $service.Key	
			$svcProp.Running = $service.Running		
			switch ($service.Policy)
				{
					"on" {$svcProp."Startup Policy" = "Start and stop with host"}
					"off" {$svcProp."Startup Policy" = "Start and stop manually"}
					"automatic" {$svcProp."Startup Policy" = "Start and stop with port usage"}
					default {$svcProp."Startup Policy" = "Unknown"}
				}
			$svcProp.Required = $service.Required
			$svcProp.Uninstallable = $service.Uninstallable
			$svcProp."Source Package" = $service.ExtensionData.SourcePackage.SourcePackageName
			$svcProp."Source Package Desc" = $service.ExtensionData.SourcePackage.Description
			
			$svcTemp = New-Object –TypeName PSObject –Prop $svcProp
			$svcInfo += $svcTemp
		}
		If ($svcInfo) {


	    $FullPath = $Path +  "services.csv"
        $svcInfo | Export-csv -Path $FullPath -append -NoTypeInformation
			
		}





		# Get SNMP Info
		$snmpInfo = @()
		$snmpDetails = Get-VMHostSnmp -Server $ESX

		$snmpProp = [ordered]@{}
		$snmpProp.Enabled = $snmpDetails.Enabled
		$snmpProp.Port = $snmpDetails.Port	
		$snmpProp."Read Only Communities" = $snmpDetails.ReadOnlyCommunities | Out-String
		$snmpProp."Trap Targets" = $snmpDetails.TrapTargets | Out-String

		$snmpTemp = New-Object –TypeName PSObject –Prop $snmpProp
		$snmpInfo += $snmpTemp

		If ($snmpInfo) {

	        $FullPath = $Path +  "snmp.csv"
            $snmpInfo | Export-csv -Path $FullPath -append -NoTypeInformation
			
		}






		# Get Installed ESX(i) Patches
		$patchInfo = @()
		Try {
			$patches = (Get-EsxCli).software.vib.list()
			ForEach ($patch in $patches) {
				$patchProp = [ordered]@{}
				$patchProp.Name = $patch.Name
				$patchProp.ID = $patch.ID
				$patchProp.Vendor = $patch.Vendor
				$patchProp.Version = $patch.Version
				$patchProp."Acceptance Level" = $patch.AcceptanceLevel
				$patchProp."Created Date" = $patch.CreationDate
				$patchProp."Install Date" = $patch.InstallDate
				$patchProp.Status = $patch.Status
				
				$patchTemp = New-Object –TypeName PSObject –Prop $patchProp
				$patchInfo += $patchTemp
			}
		} Catch {
			$patches = $VMHost | Get-VMHostPatch
			ForEach ($patch in $patches) {
				$patchProp = [ordered]@{}
				$patchProp.Name = $patch.Description
				$patchProp.ID = $patch.ID
				$patchProp."Install Date" = $patch.InstallDate
				
				$patchTemp = New-Object –TypeName PSObject –Prop $patchProp
				$patchInfo += $patchTemp
			}	
		}
		If ($patchInfo) {
	        $FullPath = $Path +  "patches.csv"
            $patchInfo | Export-csv -Path $FullPath -append -NoTypeInformation
		}






		# Get Local Users
		$accountInfo = @()
		$accounts = Get-VMHostAccount -Server $ESX
		ForEach ($account in $accounts) {
			$accountProp = [ordered]@{}
			$accountProp.Name = $account.Name
			$accountProp.Description = $account.Description
			$accountProp.ShellAccess = $account.ShellAccessEnabled
			$accountProp.Groups = $account.Groups | Out-String
			
			$accountTemp = New-Object –TypeName PSObject –Prop $accountProp
			$accountInfo += $accountTemp
		}
		If ($accountInfo) {
	        $FullPath = $Path +  "accounts.csv"
            $accountInfo | Export-csv -Path $FullPath -append -NoTypeInformation
		}




		# Get Local Groups
		$grpInfo = @()
		Try {
			$grps = Get-VMHostAccount -Server $ESX -Group -ErrorAction Stop
			ForEach ($grp in $grps) {
				$grpProp = [ordered]@{}
				$grpProp.Name = $grp.Name
				$grpProp.Description = $grp.Description
				$grpProp.Users = $grp.Users | Out-String
				
				$grpTemp = New-Object –TypeName PSObject –Prop $grpProp
				$grpInfo += $grpTemp
			}
		} Catch {}
		If ($grpInfo) {
	        $FullPath = $Path +  "groups.csv"
            $grpInfo | Export-csv -Path $FullPath -append -NoTypeInformation
		}




# Get VI Roles
		$roleInfo = @()
		$roles = Get-VIRole -Server $ESX
		ForEach ($role in $roles) {
			$roleProp = [ordered]@{}
			$roleProp.Name = $role.Name
			$roleProp.Description = $role.Description
			$roleProp.IsSystem = $role.IsSystem
			$roleProp.Privileges = $role.PrivilegeList | Out-String
			
			$roleTemp = New-Object –TypeName PSObject –Prop $roleProp
			$roleInfo += $roleTemp
		}
		If ($roleInfo) {
	        $FullPath = $Path +  "roles.csv"
            $roleInfo | Export-csv -Path $FullPath -append -NoTypeInformation
			
		}




		# Get Permissions
		$permInfo = @()
		$perms = Get-VIPermission -Server $ESX
		ForEach ($perm in $perms) {
			$permProp = [ordered]@{}
			$permProp.Entity = $perm.Entity
			$permProp.Role = $perm.Role
			$permProp.Principal = $perm.Principal
			$permProp.IsGroup = $perm.IsGroup
			$permProp.Propagate = $perm.Propagate
			
			$permTemp = New-Object –TypeName PSObject –Prop $permProp
			$permInfo += $permTemp
		}
		If ($permInfo) {
	        $FullPath = $Path +  "permissions.csv"
            $permInfo | Export-csv -Path $FullPath -append -NoTypeInformation
		}





# Get VM Info
		$vmInfo = @()
		$vms = $VMHost | Get-VM
		ForEach ($vm in $vms) {
			$vmProp = [ordered]@{}
			$vmProp.Name = $vm.Name
			$vmProp.State = $vm.PowerState
			$vmProp.FullName = If (!$VM.Guest.hostname) {"Tools Not Running\Unknown"} Else {$VM.Guest.hostname}
			$vmProp.GuestOS = If (!$VM.Guest.OSFullName) {"Tools Not Running\Unknown"} Else {$VM.Guest.OSFullName}
			$vmProp.IP = If (!$VM.Guest.IPAddress[0]) {"Tools Not Running\Unknown"} Else {$VM.Guest.IPAddress[0]}
			$vmProp.NumCPU = $vm.NumCPU
			[int]$vmProp."Memory (GB)" = $vm.MemoryGB
			$vmProp."Disk (GB)" = [Math]::Round((($vm.HardDisks | Measure-Object -Property CapacityKB -Sum).Sum * 1KB / 1GB),2)
		  	$vmProp."DiskFree (GB)" = If (![Math]::Round((($vm.Guest.Disks | Measure-Object -Property FreeSpace -Sum).Sum / 1GB),2)) `
				{"Tools Not Running\Unknown"} Else {[Math]::Round((($vm.Guest.Disks | Measure-Object -Property FreeSpace -Sum).Sum / 1GB),2)}
		  	$vmProp."DiskUsed (GB)" = If ($vmProp."DiskFree (GB)" -eq "Tools Not Running\Unknown") `
				{"Tools Not Running\Unknown"} Else {$vmProp."Disk (GB)" - $vmProp."DiskFree (GB)"}
			$vmProp.Notes = $VM.Notes	
			
			$vmTemp = New-Object –TypeName PSObject –Prop $vmProp
			$vmInfo += $vmTemp
		}
		If ($vmInfo) {
            $FullPath = $Path +  "guests.csv"
            $vmInfo | Export-csv -Path $FullPath -append -NoTypeInformation
		}




		# Get VM Snapshots
		$snapInfo = @()
		$snaps = $VMHost | Get-VM | Get-Snapshot
		ForEach ($snap in $snaps) {
			$snapProp = [ordered]@{}
			$snapProp.VM = $snap.VM
			$snapProp.Name = $snap.Name
			$snapProp.Description = $snap.Description
			$snapProp.Created = $snap.Created
			$snapProp."Size (GB)" = [Math]::Round($snap.SizeGB,2)
			$snapProp.IsCurrent = $snap.IsCurrent
			
			$snapTemp = New-Object –TypeName PSObject –Prop $snapProp
			$snapInfo += $snapTemp
		}
		If ($snapInfo) {
            $FullPath = $Path +  "snapshots.csv"
            $snapInfo | Export-csv -Path $FullPath -append -NoTypeInformation

		}



		Disconnect-VIServer $ESX -Confirm:$false
		Write-Host "Report Complete - $FullPath"
	} Else {
		Write-Host "Unable to connect to $ESX" -ForegroundColor Red
	}	
}