#
# ExportUserInfo.ps1
#
param
(
	[Parameter(Mandatory=$true)]
	[string]$CSVExportFilePath,

	[Parameter(Mandatory=$true)]
	[string]$OUDistinguishedName,

	[Parameter(Mandatory=$true)]
	[string]$LogDirPath,

    [Parameter(Mandatory=$true)]
	[string]$ExchangeServerName,

    [Parameter(Mandatory=$false)]
    [string]$PostFixFilterEmailAliasses,

	[Parameter(Mandatory=$true)]
	[int]$MaxNumberOfAliasColumns
)

Function WriteToLog
{
	param
	(
		[string]$LogPath,
		[string]$TextValue,
		[bool]$WriteError
	)

	Try
	{
		#create log file name
		$thisDate = (Get-Date -DisplayHint Date).ToLongDateString()
		$LogFileName = "ExportUserInfo_$thisDate.log"

		#write content to log file
		if ($WriteError)
		{
			Add-Content -Value "[ERROR $(Get-Date -DisplayHint Time)] $TextValue" -Path "$LogPath\$LogFileName"
		}
		else
		{
			Add-Content -Value "[INFO $(Get-Date -DisplayHint Time)] $TextValue" -Path "$LogPath\$LogFileName"
		}
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message
		Write-Host "Error occured in WriteToLog function: $ErrorMessage" -ForegroundColor Red
	}

}

Function GetUserMailboxEmailAddresses
{
	param
	(
		[string]$ADSamAccountName,
        [string]$PostFixFilterEmailAliasses,
		[string]$LogDirPatch
	)

	Try
	{
		$emailaddresses = @()
        $emailAddressObject = New-Object PSObject

		if ($usermailbox = (Get-Mailbox -ResultSize unlimited $ADSamAccountName -ErrorAction silentlycontinue))
		{
            #create emailaddress object
			$emailAddressObject = GetEmailAliasses -userMailBox $usermailbox -emailAddressObject $emailAddressObject
            return $emailAddressObject
		}
		elseif ($usermailbox = (Get-RemoteMailbox -ResultSize unlimited $ADSamAccountName -ErrorAction stop))
		{
            #create emailaddress object
			$emailAddressObject = GetEmailAliasses -userMailBox $usermailbox -IsO365Mailbox -emailAddressObject $emailAddressObject
            return $emailAddressObject
		}

	}
	Catch
	{
		return $emailaddresses
		$ErrorMessage = $_.Exception.Message
		WriteToLog -LogPath $LogDirPath -TextValue "Error occured in GetUserMailboxEmailAddresses function: $ErrorMessage" -WriteError $true
		Write-Host "Error occured in GetUserMailboxEmailAddresses function: $ErrorMessage" -ForegroundColor Red
	}
}

Function GetEmailAliasses
{
    param
    (
        [object]$userMailBox,
        [switch]$IsO365Mailbox,
        [PSObject]$emailAddressObject
    )

    Try
    {
        #add primary smtp address
        $emailAddressObject | Add-Member -MemberType NoteProperty -Name PrimaryEmailAddress -Value $usermailbox.PrimarySmtpAddress

        #add secondary aliasses to mailbox object. Only specific postfix must be added
        #primary smtp address is not filtered out
        $emailAliasses = $usermailbox.EmailAddresses | ? {($_ -like "smtp*") -and ($_ -like "*$PostFixFilterEmailAliasses")}
            
        $aliasCounter = 1

        foreach ($emailAlias in $emailAliasses)
        {
            $emailAddressObject | Add-Member -MemberType NoteProperty -Name Alias$aliasCounter -Value $emailAlias.toLower().Replace("smtp:","")
            $aliasCounter++
        }

        #create additional alias columns to make the CSV export work
        if ($aliasCounter -lt $MaxNumberOfAliasColumns)
        {
            do
            {
                $emailAddressObject | Add-Member -MemberType NoteProperty -Name Alias$aliasCounter -Value ""
                $aliasCounter++
            }
            while ($aliasCounter -lt $MaxNumberOfAliasColumns)
        }

        if ($IsO365Mailbox)
        {
            $forwardingAddress = GetForwardingAddress -LogDirPath $LogDirPath -userMailbox $userMailBox
            $emailAddressObject | Add-Member -MemberType NoteProperty -Name ForwardingEmailAddress -Value $forwardingAddress
        }
        else
        {
            $emailAddressObject | Add-Member -MemberType NoteProperty -Name ForwardingEmailAddress -Value ""
        }

        return $emailAddressObject
    }

    Catch
    {
        return $emailaddresses
		$ErrorMessage = $_.Exception.Message
		WriteToLog -LogPath $LogDirPath -TextValue "Error occured in GetEmailAliasses function while processing user $($userMailBox.UserPrincipalName): $ErrorMessage" -WriteError $true
		Write-Host "Error occured in GetEmailAliasses function: $ErrorMessage" -ForegroundColor Red   
    }
}

Function GetForwardingAddress
{
    param
    (
        [string]$LogDirPath,
        [object]$userMailbox
    )

    Try
    {
        if ((Get-O365Mailbox $userMailbox.UserPrincipalName).ForwardingSmtpAddress)
        {
            return (Get-O365Mailbox $userMailbox.UserPrincipalName).ForwardingSmtpAddress.toLower().Replace("smtp:","")
        }
        else
        {
            return "No forwarding e-mail address configured"
        }
    }

    Catch
    {
        $ErrorMessage = $_.Exception.Message
	    WriteToLog -LogPath $LogDirPath -TextValue "Error occured in function GetForwardingAddress while processing user $($userMailbox.UserPrincipalName): $ErrorMessage" -WriteError $true
	    Write-Host "Error occured in function GetForwardingAddress: $ErrorMessage"
    }
}




Try
{
    WriteToLog -LogPath $LogDirPath -TextValue "Start of export script ..." -WriteError $false

    #user object arrays
    $UserObjectArray = @()

    #connect to O365
    $UserCredential = Get-Credential
    $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $O365Session -Prefix O365

    #connect to Exchange remote powershell
    $OnPremExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServerName/PowerShell/ -Authentication Kerberos
    Import-PSSession $OnPremExchangeSession -AllowClobber

	#get all user objects in specified OU
	$adusers = Get-ADUser -SearchBase $OUDistinguishedName -Filter * -Properties DisplayName

    $aduserCounter = 0
	foreach ($aduser in $adusers)
	{
        #create user object
        $UserObject = New-Object PSObject
        
        #add properties to user object, including emailaddress object
        $UserObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $aduser.DisplayName
        $UserObject | Add-Member -MemberType NoteProperty -Name GivenName -Value $aduser.GivenName
        $UserObject | Add-Member -MemberType NoteProperty -Name SurName -Value $aduser.SurName

        #get email address array
		$emailAddressObject = GetUserMailboxEmailAddresses -ADSamAccountName $aduser.SamAccountName -LogDirPatch $LogDirPath -PostFixFilterEmailAliasses $PostFixFilterEmailAliasses
        
        #Write-Host "emailaddress object = $emailAddressObject"

        #add email addresses to userobject
        if ($emailAddressObject)
        {
            $Aliasses = $emailAddressObject | Get-Member -MemberType NoteProperty

            foreach ($alias in $Aliasses)
            {
                $aliasValue = $emailAddressObject."$($alias.Name)"
                $UserObject | Add-Member -MemberType NoteProperty -Name $alias.Name -Value $emailAddressObject."$($alias.Name)"
            }
        }
        else
        {
            WriteToLog -LogPath $LogDirPath -TextValue "No email addresses were found for user $($aduser.DisplayName)" -WriteError $false
        }

        #add user object to array of users and email addresses to address array
        $UserObjectArray += $UserObject

        $aduserCounter++
        Write-Progress -Activity "Processed user $($aduser.DisplayName). Collecting user info ..." -Status "Progress:" -PercentComplete ($aduserCounter/$adusers.Count*100)
	}
    
    #remove pssession
    Remove-PSSession $OnPremExchangeSession
    Remove-PSSession $O365Session
	
    #export to CSV file
    $UserObjectArray | Export-Csv -Path $CSVExportFilePath -NoTypeInformation

    WriteToLog -LogPath $LogDirPath -TextValue "End of export script." -WriteError $false

    #grid out
    $UserObjectArray | Out-GridView -Title "User Info of users in OU $OUDistinguishedName"
}

Catch
{
    Remove-PSSession $OnPremExchangeSession
    Remove-PSSession $O365Session

	$ErrorMessage = $_.Exception.Message
	WriteToLog -LogPath $LogDirPath -TextValue "Error occured: $ErrorMessage" -WriteError $true
	Write-Host "Error occured: $ErrorMessage"
}