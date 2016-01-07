#
# CreateADUser.ps1
#
param
(
	[Parameter(Mandatory=$true)]
	[string[]]$CSVPath,

	[Parameter(Mandatory=$true)]
	[string[]]$OUDistinguishedName,

	[Parameter(Mandatory=$true)]
	[string[]]$LogPath
)


#this function creates a log file in the given directory, date is part of the log file name
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
		$LogFileName = "CreateADUser_$thisDate.log"

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

#this function creates an active directory user
Function CreateADUser
{
	param
	(
		[string]$UserGivenName,
		[string]$UserSurName,
		[string]$UserUPN,
		[string]$UserPassword,
		[string]$OU,
		[string]$oldMailDomain,
		[string]$oldEntity,
		[string]$UserSamAccountName,
		[string]$UserDisplayName
	)

	Try
	{
		#convert password to securestring
		$passwordSec = ConvertTo-SecureString "$UserPassword" -AsPlainText -Force

		#create user in correct OU
		#check if username is more than 20 characters (=limit of SamAccountName). Split string if greater than 20
		if ($UserSamAccountName.Length -gt 20)
		{
			$UserSamAccountName = $UserSamAccountName.Substring(0,20)
		}

		$createdUserObject = New-ADUser -Name "$UserGivenName $UserSurName" -GivenName $UserGivenName -SurName $UserSurName -UserPrincipalName $UserUPN.ToLower() -StreetAddress $oldMailDomain -Company $oldEntity -Enabled $true -AccountPassword $passwordSec -Path $OU -SamAccountName $UserSamAccountName.ToLower() -DisplayName $UserDisplayName -PassThru
		WriteToLog -LogPath $LogPath -TextValue "User $UserUPN with SamAccountName $UserSamAccountName is successfully created" -WriteError $false
		Write-Host "User with UPN $($createdUserObject.UserPrincipalName) and SamAccountName $($createdUserObject.samaccountname) is created" -ForegroundColor Yellow
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message
		WriteToLog -LogPath $LogPath -TextValue "Error occured in function CreateADUser while creating account for $UserGivenName $UserSurName with following error : $ErrorMessage" -WriteError $true
	}
}

Try
{
	#import CSV data
	$UserData = Import-Csv $CSVPath -Delimiter ";"

	#start looping through user data and create objects
	foreach ($userObject in $UserData)
	{
		CreateADUser -UserGivenName $userObject.Voornaam -UserSurName $userObject.Achternaam -UserUPN $userObject.UPN -UserPassword $userObject.Wachtwoord -oldMailDomain $userObject.Huidigmaildomein -oldEntity $userObject.Entiteit -OU $OUDistinguishedName -UserSamAccountName $userObject.Username -UserDisplayName $userObject.Naam
	}
}
Catch
{
	$ErrorMessage = $_.Exception.Message
	WriteToLog -LogPath $LogPath -TextValue "Error occured: $ErrorMessage" -WriteError $true
	Write-Host "Error occured: $error"
}