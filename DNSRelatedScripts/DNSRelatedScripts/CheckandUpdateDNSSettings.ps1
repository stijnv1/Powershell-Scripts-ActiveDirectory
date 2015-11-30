#
# CheckandUpdateDNSSettings.ps1
#
function Get-IPDNSSettings ($DiscoveredServers)
{
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
		#check whether discovered server object is still online by checken DNS and net connection to WinRM port


        $StrComputer = $server.Name
        Write-Host "Checking $StrComputer" -ForegroundColor DarkGreen
        
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
        foreach ($objItem in $NetItems){
            if ($objItem.IPAddress.Count -gt 1){
                $TempIpAdderesses = [STRING]$objItem.IPAddress
                $TempIpAdderesses  = $TempIpAdderesses.Trim().Replace(" ", " ; ")
                $IpAddresses += $TempIpAdderesses
            }
            else{
                $IpAddresses += $objItem.IPAddress +"; "
            }
            if ($objItem.{MacAddress}.Count -gt 1){
                $TempMACAddresses = [STRING]$objItem.MACAddress
                $TempMACAddresses = $TempMACAddresses.Replace(" ", " ; ")
                $MACAddresses += $TempMACAddresses +"; "
            }
            else{
                $MACAddresses += $objItem.MACAddress +"; "
            }
            if ($objItem.{DNSServerSearchOrder}.Count -gt 1){
                $TempDNSAddresses = [STRING]$objItem.DNSServerSearchOrder
                $TempDNSAddresses = $TempDNSAddresses.Replace(" ", " ; ")
                $DNS += $TempDNSAddresses +"; "
            }
            else{
                $DNS += $objItem.{DNSServerSearchOrder} +"; "
            }
            if ($objItem.DNSDomainSuffixSearchOrder.Count -gt 1){
                $TempDNSSuffixes = [STRING]$objItem.DNSDomainSuffixSearchOrder
                $TempDNSSuffixes = $TempDNSSuffixes.Replace(" ", " ; ")
                $DNSSuffix += $TempDNSSuffixes +"; "
                }
            else{
                $DNSSuffix += $objItem.DNSDomainSuffixSearchOrder +"; "
                }
                $SubNet = [STRING]$objItem.IPSubnet[0]
            $intRowNet = $intRowNet + 1
        }
        $ServerObj | Add-Member @Member -Name "IP Address" -Value $IpAddresses.substring(0,$IpAddresses.LastIndexOf(";"))
        $ServerObj | Add-Member @Member -Name "IP Subnet" -Value $SubNet
        $ServerObj | Add-Member @Member -Name "MAC Address" -Value $MACAddresses.substring(0,$MACAddresses.LastIndexOf(";"))
        $ServerObj | Add-Member @Member -Name "DNS" -Value $DNS
        $ServerObj | Add-Member @Member -Name "DNS Suffix Search Order" -Value $DNSSuffix
        $ServerObj | Add-Member @Member -Name "DNS Enabled For Wins" -Value $objItem.DNSEnabledForWINSResolution
        $ServerObj | Add-Member @Member -Name "Domain DNS Registration Enabled" -Value $objItem.DomainDNSRegistrationEnabled
        $ServerObj | Add-Member @Member -Name "Full DNS Registration Enabled" -Value $objItem.FullDNSRegistrationEnabled
        $ServerObj | Add-Member @Member -Name "DHCP Enabled" -Value $objItem.DHCPEnabled
        $ServerObj | Add-Member @Member -Name "DHCP Lease Obtained" -Value $objItem.DHCPLeaseObtained
        $ServerObj | Add-Member @Member -Name "DHCP Lease Expires" -Value $objItem.DHCPLeaseExpires
        $AllServers += $ServerObj
    }

    return $AllServers
}

#log path
$logPath = "D:\Sources"

#create inventory of windows server 2008 R2 servers
$Servers = Get-ADComputer -SearchBase "OU=File Server Cluster Nodes,OU=Servers,OU=Brussel,OU=My Business,DC=water-netwerk,DC=local" -Filter {operatingSystem -like "Windows Server*"} -Properties operatingSystem

$IPSettings = Get-IPDNSSettings -DiscoveredServers $Servers

$IPSettings | Out-GridView