#
# CreateADUser.ps1
#
param
(
	[Parameter(Mandatory=$true)]
	[string[]]$CSVPath,

	[Parameter(Mandatory=$true)]
	[string[]]$OUDistinguishedName
)

Function CreateADUser
{
	param
	(
		[string]$UserGivenName,
		[string]$UserSurName,
		[string]$UserSamAccountName,
		[string]$UserUPN,
		[string]$UserPassword,
		[string]$OU
	)

	#convert password to securestring
	$passwordSec = ConvertTo-SecureString $UserPassword -AsPlainText -Force

	#create user in correct OU
	$createdUserObject = New-ADUser 
}

Try
{
	#import CSV data
	$UserData = Import-Csv $CSVPath

	#start looping through user data and create objects
	foreach ($userObject in $UserData)
	{

	}
}
Catch
{
	Write-Host "Error occured: $error"
}