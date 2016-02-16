#
# ReadDNSLogFiles.ps1
#
#used functions from following website:
# http://dollarunderscore.azurewebsites.net/?p=291
#

param
(
	[string]$DNSLogPath = "\\lwm91man02\DNS Logs",
	[string]$ScriptLogDir = "C:\Temp",
	[string]$ExportCSVPath = "C:\Temp",

	[Parameter(Mandatory=$true)]
	[string]$SMTPRelayServer,

	[Parameter(Mandatory=$true)]
	[string]$MailAddress
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
		$LogFileName = "ScanDNSLogFiles_$thisDate.log"

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

Function Get-DNSDebugLog
{
    <#
    .SYNOPSIS
    This cmdlet parses a Windows DNS Debug log.

    .DESCRIPTION
    When a DNS log is converted with this cmdlet it will be turned into objects for further parsing.

    .EXAMPLE
    Get-DNSDebugLog -DNSLog ".\Something.log" | Format-Table

    Outputs the contents of the dns debug file "Something.log" as a table.

    .EXAMPLE
    Get-DNSDebugLog -DNSLog ".\Something.log" | Export-Csv .\ProperlyFormatedLog.csv

    Turns the debug file into a csv-file.

    .PARAMETER DNSLog
    Path to the DNS log or DNS log data. Allows pipelining from for example Get-ChildItem for files, and supports pipelining DNS log data.

    #>

    [CmdletBinding()]
    param(
      [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
      [Alias('Fullname')]
      [string] $DNSLog = 'StringMode')


    BEGIN { }

    PROCESS {

        $TheReverseRegExString='\(\d\)in-addr\(\d\)arpa\(\d\)'

        ReturnDNSLogLines -DNSLog $DNSLog | % {
                if ( $_ -match '^\d\d|^\d/\d' -AND $_ -notlike '*EVENT*' -AND $_ -notlike '* Note: *') 
				{
                    $Date=$null
                    $Time=$null
                    $DateTime=$null
                    $Protocol=$null
                    $Client=$null
                    $SendReceive=$null
                    $QueryType=$null
                    $RecordType=$null
                    $Query=$null
                    $Result=$null

                    $Date=($_ -split ' ')[0]

                    # Check log time format and set properties
                    if ($_ -match ':\d\d AM|:\d\d  PM') 
					{
                        $Time=($_ -split ' ')[1,2] -join ' '
                        $Protocol=($_ -split ' ')[7]
                        $Client=($_ -split ' ')[9]
                        $SendReceive=($_ -split ' ')[8]
                        $RecordType=(($_ -split ']')[1] -split ' ')[1]
                        $Query=($_.ToString().Substring(110)) -replace '\s' -replace '\(\d?\d\)','.' -replace '^\.' -replace "\.$"
                        $Result=(((($_ -split '\[')[1]).ToString().Substring(9)) -split ']')[0] -replace ' '
                    }
                    elseif ($_ -match '^\d\d\d\d\d\d\d\d \d\d:') 
					{
                        $Date=$Date.Substring(0,4) + '-' + $Date.Substring(4,2) + '-' + $Date.Substring(6,2)
                        $Time=($_ -split ' ')[1] -join ' '
                        $Protocol=($_ -split ' ')[6]
                        $Client=($_ -split ' ')[8]
                        $SendReceive=($_ -split ' ')[7]
                        $RecordType=(($_ -split ']')[1] -split ' ')[1]
                        $Query=($_.ToString().Substring(110)) -replace '\s' -replace '\(\d?\d\)','.' -replace '^\.' -replace "\.$"
                        $Result=(((($_ -split '\[')[1]).ToString().Substring(9)) -split ']')[0] -replace ' '
                    }
                    else 
					{
                        $Time=($_ -split ' ')[1]
                        $Protocol=($_ -split ' ')[6]
                        $Client=($_ -split ' ')[8]
                        $SendReceive=($_ -split ' ')[7]
                        $RecordType=(($_ -split ']')[1] -split ' ')[1]
                        $Query=($_.ToString().Substring(110)) -replace '\s' -replace '\(\d?\d\)','.' -replace '^\.' -replace "\.$"
                        $Result=(((($_ -split '\[')[1]).ToString().Substring(9)) -split ']')[0] -replace ' '
                    }

                    #$DateTime=Get-Date("$Date $Time") -Format 'M/dd/yyyy HH:mm:ss'


                    if ($_ -match $TheReverseRegExString) 
					{
                        $QueryType='Reverse'
                    }
                    else 
					{
                        $QueryType='Forward'
                    }

                    $returnObj = New-Object System.Object
                    $returnObj | Add-Member -Type NoteProperty -Name Date -Value $Date
					$returnObj | Add-Member -Type NoteProperty -Name Hour -Value $Time
                    $returnObj | Add-Member -Type NoteProperty -Name QueryType -Value $QueryType
                    $returnObj | Add-Member -Type NoteProperty -Name Client -Value $Client
                    $returnObj | Add-Member -Type NoteProperty -Name SendReceive -Value $SendReceive
                    $returnObj | Add-Member -Type NoteProperty -Name Protocol -Value $Protocol
                    $returnObj | Add-Member -Type NoteProperty -Name RecordType -Value $RecordType
                    $returnObj | Add-Member -Type NoteProperty -Name Query -Value $Query
                    $returnObj | Add-Member -Type NoteProperty -Name Results -Value $Result

                    if ($returnObj.Query -ne $null) 
					{
                        Write-Output $returnObj
                    }
                }
            }

    }

    END { }
}

Function ReturnDNSLogLines
{
param
	(
		$DNSLog
	)

	$PathCorrect=try { Test-Path $DNSLog -ErrorAction Stop } catch { $false }

    if ($DNSLog -match '^\d\d|^\d/\d' -AND $DNSLog -notlike '*EVENT*' -AND $PathCorrect -ne $true) 
	{
        $DNSLog
    }
    elseif ($PathCorrect -eq $true) 
	{
        Get-Content $DNSLog | % { $_ }
    }
}

Function SendCSVToEMail
{
	param
	(
		$CSVFilePath,
		$mailAddress,
		$SMTPRelayServer
	)

	$mailMessage = "This CSV file contains DNS entry logs of the following DNS servers:`n`t-LWM91DC02`n`t-LWM91DC03`n`t-LWM91SERVER02`n`t-LWM91SERVER04"
	$mailSubject = "DNS Log Entries Overview"
	Send-MailMessage -Attachments $CSVFilePath -Body $mailMessage -Subject $mailSubject -From "dnslogentries@lambweston.eu" -To $mailAddress -SmtpServer $SMTPRelayServer
}

Try
{
	#create CSV file name
	$thisDate = (Get-Date -DisplayHint Date).ToLongDateString()
	$CSVFileName = "DNSLogEntriesExport_$thisDate.csv"

	#get content of DNS log files
	Write-Verbose "DNS Log path = $DNSLogPath"
	WriteToLog -LogPath $ScriptLogDir -TextValue "`n`nStart reading DNS log files ..." -WriteError $false
	Get-ChildItem "$DNSLogPath\*.log" | Get-DNSDebugLog | Export-Csv -Path "$ExportCSVPath\$CSVFileName" -NoTypeInformation

	#send CSV file via e-mail
	WriteToLog -LogPath $ScriptLogDir -TextValue "Send mail to $MailAddress with following CSV file: $CSVFileName ..." -WriteError $false
	SendCSVToEMail -CSVFilePath "$ExportCSVPath\$CSVFileName" -mailAddress $MailAddress -SMTPRelayServer $SMTPRelayServer

	#delete CSV file
	Remove-Item -Path "$ExportCSVPath\$CSVFileName" -Force

}
Catch
{
	$ErrorMessage = $_.Exception.Message
    WriteToLog -LogPath $ScriptLogDir -TextValue "Error occured during execution of the script.`nError message = $ErrorMessage" -WriteError $true
}