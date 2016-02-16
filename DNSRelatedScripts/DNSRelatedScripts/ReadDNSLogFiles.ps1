#
# ReadDNSLogFiles.ps1
#
#used functions from following website:
# http://dollarunderscore.azurewebsites.net/?p=291
#

param
(
	[string]$DNSLogPath = "\\lwm91man02\DNS Logs",
	[string]$ScriptLogDir = "C:\temp\CheckDNSEntriesLogs",
	[string]$ExportCSVPath = "C:\Temp",

	[Parameter(Mandatory=$true)]
	[string]$SMTPRelayServer,

	[Parameter(Mandatory=$true)]
	[string[]]$MailAddress,

	[Parameter(Mandatory=$false)]
	[switch]$CleanUpDNSLogs
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
		[string]$CSVFilePath,
		[string[]]$mailAddress,
		[string]$SMTPRelayServer,
		[string]$ScriptLogDir,
		$DomainControllers
	)

	Try
	{
		$mailMessage = "This CSV file contains DNS entry logs of the following DNS servers:"
		$DomainControllers | % {$mailMessage += "`n`t-$($_.DNSHostName)"}
		$mailSubject = "DNS Log Entries Overview"
		WriteToLog -LogPath $ScriptLogDir -TextValue "Send mail to $mailAddress with following CSV file: $CSVFileName ..." -WriteError $false
		Send-MailMessage -Attachments $CSVFilePath -Body $mailMessage -Subject $mailSubject -From "dnslogentries@lambweston.eu" -To $mailAddress -SmtpServer $SMTPRelayServer
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message
		WriteToLog -LogPath $ScriptLogDir -TextValue "Error occured during execution of the function SendCSVToEMail.`nError message = $ErrorMessage" -WriteError $true
	}
}

Function CleanUpDNSLogFiles
{
	#this function cleans up the DNS debug log files.
	#to be able to cleanup, the DNS service needs to be stopped, file needs to be deleted, and DNS service needs to be started again
	param
	(
		$DNSServerName,
		$ScriptLogDir
	)

	Try
	{
		WriteToLog -LogPath $ScriptLogDir -TextValue "Stopping DNS service on server $DNSServerName ..." -WriteError $false
		Invoke-Command -ComputerName $DNSServerName -ScriptBlock {Stop-Service -Name "DNS"} -ErrorAction Stop

		WriteToLog -LogPath $ScriptLogDir -TextValue "Deleting DNS debug log file of server $DNSServerName ..." -WriteError $false
		Remove-Item -Path "$DNSLogPath\$DNSServerName*.log"

		WriteToLog -LogPath $ScriptLogDir -TextValue "Starting DNS service on server $DNSServerName ..." -WriteError $false
		Invoke-Command -ComputerName $DNSServerName -ScriptBlock {Start-Service -Name "DNS"}
	}

	Catch
	{
		$ErrorMessage = $_.Exception.Message
		WriteToLog -LogPath $ScriptLogDir -TextValue "Error occured during execution of function CleanUpDNSLogFiles.`r`nError message = $ErrorMessage" -WriteError $true
	}
}

Try
{
	#get all old domain controllers
	$DomainControllers = get-adcomputer -SearchBase "OU=Domain Controllers,DC=lwmeijer,DC=cag" -Filter {operatingSystem -like "Windows Server 2008*"}

	#create CSV file name
	$thisDate = (Get-Date -DisplayHint Date).ToLongDateString()
	$CSVFileName = "DNSLogEntriesExport_$thisDate.csv"

	#get content of DNS log files
	Write-Verbose "DNS Log path = $DNSLogPath"
	WriteToLog -LogPath $ScriptLogDir -TextValue "Start reading DNS log files ..." -WriteError $false
	Get-ChildItem "$DNSLogPath\*.log" | Get-DNSDebugLog | Export-Csv -Path "$ExportCSVPath\$CSVFileName" -NoTypeInformation

	#send CSV file via e-mail
	SendCSVToEMail -CSVFilePath "$ExportCSVPath\$CSVFileName" -mailAddress $MailAddress -SMTPRelayServer $SMTPRelayServer -DomainControllers $DomainControllers -ScriptLogDir $ScriptLogDir

	#delete CSV file
	Remove-Item -Path "$ExportCSVPath\$CSVFileName" -Force

	#if the CleanUpDNSLogs switch parameter is specified, the log files are deleted by stopping DNS services, deleting the log files, and starting the DNS services again
	if ($CleanUpDNSLogs)
	{
		$DomainControllers | % {CleanUpDNSLogFiles -DNSServerName $_.DNSHostName -ScriptLogDir $ScriptLogDir}
	}

}
Catch
{
	$ErrorMessage = $_.Exception.Message
    WriteToLog -LogPath $ScriptLogDir -TextValue "Error occured during execution of the script.`r`nError message = $ErrorMessage" -WriteError $true
}