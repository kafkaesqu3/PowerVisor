
$Path = 'C:\users\david\desktop\'

### Main script begin ###
# Test path
If (!(Test-Path $path)) {
	Write-Host "Path not found - $path" -ForegroundColor Red
	Exit
}


# Obtain Credentials
$creds = Get-Credential -Message "Enter ESX/ESXi host credentials used to connect"

# Cycle through servers gathering and exporting info
Foreach ($ESX in $Server) {
	# Create full path to Excel workbook
	# Connect to ESX/ESXi host
	$connection = Connect-VIServer $ESX -Credential $creds
	# Ensure connection
	If ($connection) {	
		Write-Host "Connected to $ESX"
		# Get ESX(i) Host Info
		$VMHost = Get-VMHost $ESX
		
		$VMHostProp = [ordered]@{}
		$VMHostProp.Name = $VMHost.NetworkInfo.HostName
		$VMHostProp."Domain Name" = $VMHost.NetworkInfo.DomainName
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
		

	    $FullPath = $Path.Trimend('\') + "esxi.csv"
        $VMHostInfo | Export-csv -Path $FullPath -append -NoTypeInformation

		Disconnect-VIServer $ESX -Confirm:$false
		Write-Host "Report Complete - $FullPath"
	} Else {
		Write-Host "Unable to connect to $ESX" -ForegroundColor Red
	}	
}