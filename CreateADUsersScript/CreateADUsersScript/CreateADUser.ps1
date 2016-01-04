#
# CreateADUser.ps1
#
param
(
	[Parameter(Mandatory=$true)]
	[string[]]$CSVPath
)

Try
{
	#import CSV data

}
Catch
{
	Write-Host "Error occured: $error"
}