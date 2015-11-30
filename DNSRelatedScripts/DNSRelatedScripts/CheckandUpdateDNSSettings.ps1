#
# CheckandUpdateDNSSettings.ps1
#
param
(
	[Parameter(Mandatory=$true)]
	[string]$OUDistinguishedName,

	[Parameter(Mandatory=$false)]
	[switch]$ExportToCSV,

	[Parameter(Mandatory=$false)]
	[string]$CSVPath,

	[Parameter(Mandatory=$true)]
	[string]$LogPath
)
function Get-IPDNSSettings ($DiscoveredServers)
{
	Try
	{
		#this function is based on the powershell script found on Technet Gallery:
		#https://gallery.technet.microsoft.com/Gather-DNS-settings-from-fec23eaa#content

		#initialize arrays
		$AllServers = @()
		$ServerObj  = @()
		$Member = @{
			MemberType = "NoteProperty"
			Force = $true
		}

		#start foreach to go through all discovered server objects
		foreach ($server in $DiscoveredServers)
		{
			Try
			{
				#check whether discovered server object is still online by checken DNS and net connection to WinRM port
				if (Resolve-DnsName -Name $server.Name -ErrorAction SilentlyContinue)
				{
					#server name is discovered in DNS, test whether server is online by testing connection to WinRM port
					Write-Host "Server $($server.Name) is found in DNS, checking net connections ..." -ForegroundColor Yellow

					if (Test-NetConnection -ComputerName $server.Name -CommonTCPPort WINRM -InformationLevel Quiet -ErrorAction SilentlyContinue)
					{
						Write-Host "$($server.Name) is online, getting DNS client settings ..." -ForegroundColor Yellow
						#start discovery of DNS client settings on this server

						$StrComputer = $server.Name
						Write-Host "Checking $StrComputer" -ForegroundColor Yellow
        
						$NetItems = $null
						Write-Progress -Status "Working on $StrComputer" -Activity "Gathering Data"
						$ServerObj = New-Object psObject
						$ServerObj | Add-Member @Member -Name "Hostname" -Value $StrComputer
						$NetItems = @(Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'" -ComputerName $StrComputer)
						$intRowNet = 0
						$ServerObj | Add-Member -MemberType NoteProperty -Name "NIC's" -Value $NetItems.Length -Force
						[STRING]$MACAddresses = @()
						[STRING]$IpAddresses = @()
						[STRING]$DNS = @()
						[STRING]$DNSSuffix = @()
						foreach ($objItem in $NetItems)
						{
							if ($objItem.IPAddress.Count -gt 1)
							{
								$TempIpAdderesses = [STRING]$objItem.IPAddress
								$TempIpAdderesses  = $TempIpAdderesses.Trim().Replace(" ", " - ")
								$IpAddresses += $TempIpAdderesses
							}
							else
							{
								$IpAddresses += $objItem.IPAddress +"- "
							}
							if ($objItem.{MacAddress}.Count -gt 1)
							{
								$TempMACAddresses = [STRING]$objItem.MACAddress
								$TempMACAddresses = $TempMACAddresses.Replace(" ", " - ")
								$MACAddresses += $TempMACAddresses +"- "
							}
							else
							{
								$MACAddresses += $objItem.MACAddress +"- "
							}
							if ($objItem.{DNSServerSearchOrder}.Count -gt 1)
							{
								$TempDNSAddresses = [STRING]$objItem.DNSServerSearchOrder
								$TempDNSAddresses = $TempDNSAddresses.Replace(" ", " - ")
								$DNS += $TempDNSAddresses +"- "
							}
							else
							{
								$DNS += $objItem.{DNSServerSearchOrder} +"- "
							}
							if ($objItem.DNSDomainSuffixSearchOrder.Count -gt 1)
							{
								$TempDNSSuffixes = [STRING]$objItem.DNSDomainSuffixSearchOrder
								$TempDNSSuffixes = $TempDNSSuffixes.Replace(" ", " - ")
								$DNSSuffix += $TempDNSSuffixes +"- "
							}
							else
							{
								$DNSSuffix += $objItem.DNSDomainSuffixSearchOrder +"- "
							}

							$SubNet = [STRING]$objItem.IPSubnet[0]
							$intRowNet = $intRowNet + 1
						}

						$ServerObj | Add-Member @Member -Name "IP Address" -Value $IpAddresses.substring(0,$IpAddresses.LastIndexOf("-"))
						$ServerObj | Add-Member @Member -Name "IP Subnet" -Value $SubNet
						$ServerObj | Add-Member @Member -Name "MAC Address" -Value $MACAddresses.substring(0,$MACAddresses.LastIndexOf("-"))
						$ServerObj | Add-Member @Member -Name "DNS" -Value $DNS
						$ServerObj | Add-Member @Member -Name "DNS Suffix Search Order" -Value $DNSSuffix
						$ServerObj | Add-Member @Member -Name "DNS Enabled For Wins" -Value $objItem.DNSEnabledForWINSResolution
						$ServerObj | Add-Member @Member -Name "Domain DNS Registration Enabled" -Value $objItem.DomainDNSRegistrationEnabled
						$ServerObj | Add-Member @Member -Name "Full DNS Registration Enabled" -Value $objItem.FullDNSRegistrationEnabled
						$ServerObj | Add-Member @Member -Name "DHCP Enabled" -Value $objItem.DHCPEnabled
						$ServerObj | Add-Member @Member -Name "DHCP Lease Obtained" -Value $objItem.DHCPLeaseObtained
						$ServerObj | Add-Member @Member -Name "DHCP Lease Expires" -Value $objItem.DHCPLeaseExpires
						$ServerObj | Add-Member @Member -Name "Operating System" -Value $server.OperatingSystem
						$AllServers += $ServerObj
					}
					else
					{
						Write-Host "Net connection test for server $($server.Name) failed." -ForegroundColor Cyan
					}
				}
				else
				{
					Write-Host "Server $($Server.Name) is not found in DNS" -ForegroundColor Green
				}
			}

			Catch [Excpetion]
			{
				$ErrorActionPreference = "SilentlyContinue"
				$errorLog = $_.Exception.InnerException.Message
				$errorLineInScript = $_.InvocationInfo.ScriptLineNumber
				$WriteToLog = "Error occured during discovery of server $($server.Name):`n$errorLog`n$errorLineInScript`n`n"
				Add-Content -Path $LogPath -Value $WriteToLog
			}
		}

		return $AllServers
	}

	Catch [Excpetion]
	{
		$errorLog = $_.Exception.InnerException.Message
		$errorLineInScript = $_.InvocationInfo.ScriptLineNumber
		$WriteToLog = "Error occured:`n$errorLog`n$errorLineInScript`n`n"
		Add-Content -Path $LogPath -Value $WriteToLog
		$ErrorActionPreference = "Stop"
		return $null
	}
}

#create inventory of windows server 2008 R2 servers
$Servers = Get-ADComputer -SearchBase "$OUDistinguishedName" -Filter {operatingSystem -like "Windows Server*"} -Properties operatingSystem

if ($Servers.Count -gt 0)
{
	$IPSettings = Get-IPDNSSettings -DiscoveredServers $Servers
}
else
{
	Write-Host "No server objects found for the given OU distinguished name" -ForegroundColor Red
}

if ($IPSettings -ne $null)
{
	if ($ExportToCSV)
	{
		if ($CSVPath)
		{
			Write-Host "Creating a CSV export of the gathered IP information ..." -ForegroundColor Yellow
			$IPSettings | Export-Csv -Path $CSVPath -NoTypeInformation
		}
		else
		{
			Write-Host "No CSV export can be generated. The CSVPath parameter is empty." -ForegroundColor Red
		}
	}

	$IPSettings | Out-GridView
}