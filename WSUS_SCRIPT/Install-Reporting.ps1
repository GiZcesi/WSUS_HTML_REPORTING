#Disable showing errors
$ErrorActionPreference= 'silentlycontinue'

#Launch the script .ps1 as admin 
param([switch]$Elevated)
function Test-Admin 
{
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
if ((Test-Admin) -eq $false) 
{
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}


$asciiArt = "
        __      __  _____________ ___  _________
       /  \    /  \/   _____/    |   \/   _____/
       \   \/\/   /\_____  \|    |   /\_____  \ 
        \        / /        \    |  / /        \
         \__/\  / /_______  /______/ /_______  /
              \/          \/                 \/  
        _______________________________________                          
"
Write-Host $asciiArt -ForegroundColor White


function CheckUser #Check if user is automatically found if not it prompt the user to enter his username
{
    [CmdletBinding()]
    param()

    $validUser = $false

    do {
        $Validusername = $env:UserName
        Write-Host
        $userPath = "C:\Users\$Validusername"

        if (-not (Test-Path $userPath)) {
            $Validusername = Read-Host "Enter your username (domain not required)"
            $userPath = "C:\Users\$Validusername"
            if (-not (Test-Path $userPath)) {
                Write-Host
                Write-Host "The user '$Validusername' does not exist on the server. Try again."
            } else {
                Write-Host
                Write-Host "The user '$Validusername' is valid."
                $validUser = $true
            }
        } else {
            Write-Host
            Write-Host "The user '$Validusername' was automatically found "
            $validUser = $true
        }
    } while (-not $validUser)

    # Return the validated username as output of the function
    $Validusername
}

# Check if User is Valid

$username = CheckUser
Write-Host
Write-Host "The current username is: $username"
Write-Host
$SID = (New-Object System.Security.Principal.NTAccount($username)).Translate([System.Security.Principal.SecurityIdentifier]).value
Write-Host "The current SID is : $SID"

function ClearConf 
{
    [CmdletBinding()]
    param ()

    $BCKPDirectory = "C:\Users\$username\Documents\WSUS_Script\ConfigBCKP\"
    $ReportDirectory = "C:\Users\$username\Documents\WSUS_Script\"

    Copy-Item -Path "$BCKPDirectory\*" -Destination $ReportDirectory -Recurse -Force
} #Make a clear config by overwriting files when launching the install script
ClearConf


function CheckGroupsRequirement #Check if the user meet the requirement (WSUS Groups)
{
[CmdletBinding()]
# Define the group and user variables
$groupwsusadm = "WSUS Administrators"
# Check if the user is already a member of the group
$isMember = Get-LocalGroupMember -Group $groupwsusadm | Where-Object { $_.Name -eq $username } 
# If the user is not a member of the group, add them
if (-not $isMember) {
  Add-LocalGroupMember -Group $groupwsusadm -Member $username  
}

# Define the group and user variables
$groupreportadm = "WSUS Reporters"
# Check if the user is already a member of the group
$isMemberofreportwsus = Get-LocalGroupMember -Group $groupreportadm | Where-Object { $_.Name -eq $username } 
# If the user is not a member of the group, add them
if (-not $isMemberofreportwsus) { 
  Add-LocalGroupMember -Group $groupreportadm -Member $username 
}

}
CheckGroupsRequirement


function Update-WSUSScript 
{

#Path of files
$xmlFilePath1 = "C:\Users\$username\Documents\WSUS_Script\XML\Reporting_SRV.xml"
$xmlFilePath2 = "C:\Users\$username\Documents\WSUS_Script\XML\Send_Reporting_HTML.xml"
$ps1FilePath1 = "C:\Users\$username\Documents\WSUS_Script\WindowsPowerShell\Modules\PS2HTMLTable\Scripts\report_wsus_SRV_Tache_planif.ps1"
$ps1FilePath2 = "C:\Users\$username\Documents\WSUS_Script\WindowsPowerShell\Modules\PS2HTMLTable\Scripts\Get-WSUSComputerStatus_Serveurs.ps1"
$ps1FilePath3 = "C:\Users\$username\Documents\WSUS_Script\WindowsPowerShell\Modules\PS2HTMLTable\Scripts\Send_HTML_Files.ps1"

function Get-WSUSPort {
    param($server)

    try {
        # Try the first configuration (HTTP)
        $wsusConfig = Get-WsusServer -Name $server -PortNumber 8530 

        # If the above command succeeds, the WSUS port is 8530 (HTTP)
        $wsusPort = 8530
    }
    catch {
        # If the first configuration fails, try the second configuration (HTTPS)
        try {
            $wsusConfig = Get-WsusServer -Name $server -PortNumber 8531 -UseSsl

            # If the above command succeeds, the WSUS port is 8531 (HTTPS)
            $wsusPort = 8531
        }
        catch {
            Write-Host
            Write-Host "Failed to connect to WSUS server. Check server name and permissions."
            return
        }
    }

    return $wsusPort
}

function Get-WSUSFQDN {
    [CmdletBinding()]
    param()

    $domainName = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
    $global:server = $env:COMPUTERNAME

    if (-not [string]::IsNullOrWhiteSpace($domainName)) {
        $global:server += "."
    }

    $global:server += $domainName

    Write-Host
    Write-Host "Detected FQDN of the server: $server"

    # Ask the user to confirm the FQDN or enter a different one
    Write-Host
    $confirm = Read-Host "Is this the correct FQDN of the WSUS server? (Y/N)"
    if ($confirm -ne 'Y') {
        Write-Host
        $global:server = Read-Host 'Enter FQDN of the WSUS server (e.g., server.domain.local)'
        Write-Host
    }

    # Return the final FQDN as output of the function
     $global:server
}

# Call the function to get the WSUS FQDN
$wsusFQDN = Get-WSUSFQDN


# Call the function to get the WSUS port, passing the $wsusFQDN variable as an argument
$port = Get-WSUSPort -server $wsusFQDN

if (-not $port) {
    # Handle the case where WSUS port couldn't be determined
     Write-Host
     Write-Host "WSUS port not detected automatically."
     Write-Host
     $port = Read-Host 'Please enter the port of the WSUS Server'
}

Write-Host
Write-Host "WSUS port is set on : $port"
Write-Host

#Replace all desired values

(Get-Content $xmlFilePath1) -replace "DefaultUsername", "$username" | Set-Content $xmlFilePath1
(Get-Content $xmlFilePath2) -replace "DefaultUsername", "$username" | Set-Content $xmlFilePath2
(Get-Content $xmlFilePath1) -replace "USER-SID", "$SID" | Set-Content $xmlFilePath1
(Get-Content $xmlFilePath2) -replace "USER-SID", "$SID" | Set-Content $xmlFilePath2

(Get-Content $ps1FilePath1) -replace "DefaultUsername", "$username" | Set-Content $ps1FilePath1

(Get-Content $ps1FilePath2) -replace "DefaultServer", "$server" | Set-Content $ps1FilePath2
(Get-Content $ps1FilePath2) -replace "123456", "$port" | Set-Content $ps1FilePath2

if($port -eq "8530")
{
$SSLContent = Get-Content $ps1FilePath2 -Raw
$updatedSSLContent = $SSLContent -replace '\[switch\]\$UseSSL = "ValueSSL",', '[switch]$UseSSL = $false,'
$updatedSSLContent | Set-Content -Path $ps1FilePath2
}elseif($port -eq "8531")
{
$SSLContent = Get-Content $ps1FilePath2 -Raw
$updatedSSLContent = $SSLContent -replace '\[switch\]\$UseSSL = "ValueSSL",', '[switch]$UseSSL = $true,'
$updatedSSLContent | Set-Content -Path $ps1FilePath2
}else {
    Write-Host "Please enter 1 for SSL = False or 2 for SSL = True:"
    $sslChoice = Read-Host

    if ($sslChoice -eq "1") {
        $updatedSSLContent = '[switch]$UseSSL = $false,'
    } elseif ($sslChoice -eq "2") {
        $updatedSSLContent = '[switch]$UseSSL = $true,'
    } else {
        Write-Host "Invalid input. Exiting..."
        exit 1
    }

    $SSLContent = Get-Content $ps1FilePath2 -Raw
    $updatedSSLContent = $SSLContent -replace '\[switch\]\$UseSSL = "ValueSSL",', $updatedSSLContent
    $updatedSSLContent | Set-Content -Path $ps1FilePath2
}


(Get-Content $ps1FilePath3) -replace "DefaultUsername", "$username" | Set-Content $ps1FilePath3

$mailsender = Read-Host 'Please specify the e-mail sender (exemple : notification@domain.fr/com) '
Write-Host

(Get-Content $ps1FilePath3) -replace "MAILSENDER", "$mailsender" | Set-Content $ps1FilePath3

$mailreceiver = Read-Host 'Please specify the e-mail recipient (exemple : it-support@domain.fr/com) '
Write-Host

(Get-Content $ps1FilePath3) -replace "MAILRECEIVER", "$mailreceiver" | Set-Content $ps1FilePath3

$serversmtp = Read-Host 'Please specify an smtp server  '
Write-Host

(Get-Content $ps1FilePath3) -replace "SERVERSMTP", "$serversmtp" | Set-Content $ps1FilePath3

$portsmtp = Read-Host 'Please specify the port of the smtp server (default: 25)  '
Write-Host

(Get-Content $ps1FilePath3) -replace "123456", "$portsmtp" | Set-Content $ps1FilePath3
} #Ask the user for parameters
Update-WSUSScript

function CreateXMLTasks #Create the scheduled tasks
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [PSCredential]$UserCredential
    )

    $XMLTasks = "C:\Users\$username\Documents\WSUS_Script\XML"

    # Namespace manager to handle the namespace in the XML
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager((New-Object System.Xml.NameTable))
    $namespaceManager.AddNamespace("ns", "http://schemas.microsoft.com/windows/2004/02/mit/task")

    # Get a list of all XML files in the folder
    $xmlFiles = Get-ChildItem -Path $XMLTasks -Filter "*.xml"

    Unregister-ScheduledTask -TaskName "Reporting_SRV" -TaskPath "\" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "Send_Reporting_HTML" -TaskPath "\" -Confirm:$false -ErrorAction SilentlyContinue

    Write-Host 'Creating Scheduled Tasks...'

    # Loop through each XML file and import the scheduled task
    foreach ($xmlFile in $xmlFiles) {
        $xmlContent = Get-Content $xmlFile.FullName -Raw

        # Load the XML content
        $xml = [xml]$xmlContent

        # Extract the URI from the XML content
        $taskName = $xml.SelectSingleNode("//ns:Task/ns:RegistrationInfo/ns:URI", $namespaceManager).InnerText

        Register-ScheduledTask -TaskName $taskName -Xml $xmlContent -User $UserCredential.UserName -Password $UserCredential.GetNetworkCredential().Password
    }
}

Write-Host "Waiting for credentials... (Please enter domain before username if not local account)"
# Prompt the user for credentials
$UserCredential = $host.ui.PromptForCredential("Need credentials", "Please enter your user name and password.", "", "NetBiosUserName")
Write-Host
# Call the function and pass the credentials
CreateXMLTasks -UserCredential $UserCredential

function Loading #Loading
{
for ($i = 1; $i -le 100; $i++ ) {
    Write-Progress -Activity "in Progress" -Status "$i% Complete:" -PercentComplete $i
    Start-Sleep -Milliseconds 5
}
}

