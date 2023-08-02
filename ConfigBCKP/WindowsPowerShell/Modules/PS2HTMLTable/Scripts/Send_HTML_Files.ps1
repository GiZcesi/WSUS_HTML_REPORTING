#Requires -Modules @{ModuleName="PS2HTMLTable";ModuleVersion="1.0.0.0"}
#Requires -PSEdition Desktop

[CmdletBinding()]
param (
    [switch]$SendEmail = $true,
    [string]$FromAddress = "MAILSENDER",
    $RecipientAddress = @("Informatique <MAILRECEIVER>", "CDS <support.cds@novenci.fr>"),
    [string]$SMTPServer = "SERVERSMTP",
    [int32]$SMTPPort = 123456,
    $SMTPEncoding =[System.Text.Encoding]::UTF8

)


process {
        $currentDate = Get-Date -Format "yyyy-MM-dd"
        $File1 = "C:\Users\DefaultUsername\Documents\WSUS_Script\Export\Reporting_serveurs_$currentDate.html"
        Send-MailMessage -From $FromAddress -To $RecipientAddress -Subject "WSUS Reporting Novenci - $currentDate"  -Attachments "$File1" -SmtpServer $SMTPServer -Port $SMTPPort -Encoding $SMTPEncoding
    }
    
