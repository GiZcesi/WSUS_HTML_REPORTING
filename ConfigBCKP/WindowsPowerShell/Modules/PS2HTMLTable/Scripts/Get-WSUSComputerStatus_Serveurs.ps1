#Requires -Modules @{ModuleName="PS2HTMLTable";ModuleVersion="1.0.0.0"}
#Requires -PSEdition Desktop

[CmdletBinding()]
param (
    [string]$ComputerName = "DefaultServer",
    [int32]$WSUSPort = 123456,
    [switch]$UseSSL = "ValueSSL",
    $SMTPEncoding =[System.Text.Encoding]::UTF8

)


process {
  
function Get-LocalTime ($UTCTime) {
    # Create a time zone object for UTC+1
    $TimeZone = [System.TimeZoneInfo]::CreateCustomTimeZone('UTC+1', [System.TimeSpan]::FromHours(1), '(UTC+01:00) UTC+1', 'UTC+1')

    # Convert the UTC time to the local time in the UTC+1 time zone
    [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TimeZone)
}

    # Test for connection
    if ($null -eq $WSUS) {
        # No connection detected, load assembly. Requires RSAT to be installed if run remotely.
        Add-Type -Path 'C:\Program Files\Update Services\Api\Microsoft.UpdateServices.Administration.dll' | Out-Null
    }

    # Connect to WSUS server
    try {
        $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($ComputerName, $UseSSL, $WSUSPort)
    } catch {
        throw $_
    }

    $StateInstall = 'Installed' #Choisis le state des install (Failed, Pending, Installed,etc) 
    $computers = @()
    $computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
    $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope 
    $updateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::$StateInstall  
    $Group = $WSUS.GetComputerTargetGroups() | Where-Object {$_.Name -like "Tous les ordinateurs"}
    
    #Update Scope fonction de temps (a delete /commenter si prendre toute les majs)
    # Set the start and end dates for the update scope to one month ago
    
    $startDate = (Get-Date).AddMonths(-1).AddDays(-(Get-Date).Day + 1)
    $endDate = (Get-Date).AddMonths(0).AddDays(-(Get-Date).Day)
    $updateScope.FromCreationDate = $startDate
    $updateScope.ToCreationDate = $endDate

    Write-Host "Date de Début: $startDate"
    Write-Host "Date de Fin: $endDate"
     
$computers = $Group.GetTotalSummaryPerComputerTarget() | ForEach-Object {
  
  $kbnumbers = New-Object System.Collections.ArrayList
  $IDKB = New-Object System.Collections.ArrayList
  $computer = $WSUS.GetComputerTarget($_.ComputerTargetId)
  $baseUrl = "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid="
  
  if ($computer -ne $null -and $computer.OSDescription -ne "Windows 10 Pro" -and $computer.OSDescription -ne "Windows 7 Professionnel" -and $computer.OSDescription -ne "Windows 10 Pro for Workstations") {
    $computer.GetUpdateInstallationInfoPerUpdate($UpdateScope) | Where-Object {$_.GetUpdate().LegacyName -ne ""} | ForEach-Object { #Delete le where object pour choper toute les maj avec office 
      $title = $_.GetUpdate().Title
      
      $kbnumber = ($_.GetUpdate().LegacyName -split '[-_]')[0]
      if ($kbnumber -ne "" -and $kbnumber -ne $null) { #A delete si ajouter maj office au reporting
        if (-not $kbnumber.StartsWith("KB")) {
          $kbnumber = "KB" + $kbnumber
        }
        $IDKB = $_.UpdateId
        $title = $_.GetUpdate().Title
        $values = @()
        $values += "<a href='$baseUrl$IDKB'>$kbnumber</a>" #Modification possible a la place de >$kbnumber</a> // exemple >$title</a> pour mettre le titre de la maj
        $kbnumbers.AddRange($values)
        #$kbnumbers.Add($kbnumber)
        #$kbnumbers.Add($kbnumber) //Ancienne fonction
        #$kbnumbers.Add($title) //POUR TITRE UPDATE
        Write-Host "Ajout de $kbnumber au reporting"
       Write-Host  "Ajout de l'id $IDKB au reporting"
      }
   }

    #https://www.catalog.update.microsoft.com/Search.aspx?q=
    
    if ($computer -ne $null) {    
    #Write-Host "Adding computer $($computer.FullDomainName) to report file..."
    #Write-Host $kbnumber
        
        
        
        [PSCustomObject]@{
        
        "Nom de l'ordinateur"   = $computer.FullDomainName
        "Adresse IP"             = $computer.IPAddress.ToString()
        "Installée(s)"           = $_.InstalledCount
        "Non installée(s)"       = $_.NotInstalledCount
        "À effectuer"            = $_.DownloadedCount + $_.NotInstalledCount
        "Inconnue(s)"            = $_.UnknownCount
        "Échoué(s)"              = $_.FailedCount
        "En attente de reboot"   = $_.InstalledPendingRebootCount
        "Système d'exploitation" = $computer.OSDescription
        "Dernier contact"        = $(Get-LocalTime($Computer.LastSyncTime)).ToString("dd/MM/yyyy hh:mm:ss tt")
        "Dernier rapport d'état" = $(Get-LocalTime($Computer.LastReportedStatusTime)).ToString("dd/MM/yyyy hh:mm:ss tt")
        "KB $StateInstall du $($startDate.ToString("dd/MM/yyyy")) au $($endDate.ToString("dd/MM/yyyy")) (Hors Maj Office/Oth)"            = $kbnumbers
              
                }
            }
        }
    }

    # Define parameters array for the "Not Installed" column
    $paramsInstalled = @{
        # Column name
        Column = "Installée(s)"
        # Test criteria: Is value greater than or equal to Argument?
        ScriptBlock = {[double]$args[0] -ge [double]$args[1]}
        # CSS attribute to add if ScriptBlock is true
        CSSAttribute = "style"
    }
    $paramsNotInstalled = @{
        # Column name
        Column = "Non Installée(s)"
        # Test criteria: Is value greater than or equal to Argument?
        ScriptBlock = {[double]$args[0] -ge [double]$args[1]}
        # CSS attribute to add if ScriptBlock is true
        CSSAttribute = "style"
    }
    $paramsFailed = @{
        # Column name
        Column = "Échoué(s)"
        # Test criteria: Is value greater than Argument?
        ScriptBlock = {[double]$args[0] -gt [double]$args[1]}
        # CSS attribute to add if ScriptBlock is true
        CSSAttribute = "style"
    }
   $paramsNeeded = @{
    # Column name
    Column = "À Effectuer"
    # Test criteria: No test is needed for this column, so set ScriptBlock to $null
    ScriptBlock = {}
    # CSS attribute to add if ScriptBlock is true (not needed in this case)
    CSSAttribute = "style"
}
     $paramsUnknown = @{
    # Column name
    Column = "Inconnue(s)"
    # Test criteria: No test is needed for this column, so set ScriptBlock to $null
    ScriptBlock = {}
    # CSS attribute to add if ScriptBlock is true (not needed in this case)
    CSSAttribute = "style"
} 
    $paramsPendingReboot = @{
        # Column name
        Column = "En attente de reboot"
        # Test criteria: Is value greater than Argument?
        ScriptBlock = {[double]$args[0] -gt [double]$args[1]}
        # CSS attribute to add if ScriptBlock is true
        CSSAttribute = "style"
    }
    $paramsLastReportedStatusTime = @{
        # Column name
        Column = "Dernier rapport d'état"
        # Test criteria: Is date older than or equal to Argument?
        ScriptBlock = {[datetime]$args[0] -le [datetime]$args[1]}
        CSSAttribute = "style"
    }
    $paramsKBUpdates = @{
        # Column name
        Column = "KB Updates"
        # Test criteria: Is date older than or equal to Argument?
        ScriptBlock = {[double]$args[0]}
        CSSAttribute = "style"
        }

    # Create HTML document
    $HTML = New-HTMLHead -Title "$ComputerName - WSUS Status Report"
    $HTML += "<h3>Tous les serveurs ($($Computers.Count))</h3>"

    # Order Columns and create HTML table
 
    $HTMLTable = $Computers | Sort-Object @{Expression = "Nom de l'ordinateur";Descending = $false},@{Expression = "Non installée(s)";Descending = $false} -Descending | New-HTMLTable -HTMLDecode -SetAlternating
    $HTMLTable = 
    # Color Not Installed column red, orange, or yellow if their value is greater than or equal to 60, 40, or 15 respectively
    $HTMLTable = Add-HTMLTableColor -HTML $HTMLTable -Argument 15 -CSSAttributeValue "background-color:#f6ed60;" @paramsNotInstalled
    $HTMLTable = Add-HTMLTableColor -HTML $HTMLTable -Argument 40 -CSSAttributeValue "background-color:#feb74f;" @paramsNotInstalled
    $HTMLTable = Add-HTMLTableColor -HTML $HTMLTable -Argument 60 -CSSAttributeValue "background-color:#ed5e3c;" @paramsNotInstalled

    # Color Failed column if any updates failed
    $HTMLTable = Add-HTMLTableColor -HTML $HTMLTable -Argument 0 -CSSAttributeValue "background-color:#88AC76;" @paramsFailed

    # Color Pending Reboot column if any updates are pending reboot
    $HTMLTable = Add-HTMLTableColor -HTML $HTMLTable -Argument 0 -CSSAttributeValue "background-color:#70c3ed;" @paramsPendingReboot

    # Color Last Status Report column if a computer hasn't reported status in more than 7 or 30 days
    $HTMLTable = Add-HTMLTableColor -HTML $HTMLTable -Argument (Get-Date).AddDays(-7) -CSSAttributeValue "background-color:#9a6db0;" @paramsLastReportedStatusTime
    $HTMLTable = Add-HTMLTableColor -HTML $HTMLTable -Argument (Get-Date).AddDays(-30) -CSSAttributeValue "background-color:#c3add1;" @paramsLastReportedStatusTime

    # Add HTML Table to HTML and append legend
    $HTML += $HTMLTable
    $HTML += '<h4>Color Coding:</h4>'
    $HTML += '<ul>'
    $HTML += '<li>Non Installée(s) plus de 60 <span style="background-color:#ed5e3c;">Rouge</span></li>'
    $HTML += '<li>Non Installée(s) entre 40 et 59 <span style="background-color:#feb74f;">Orange</span></li>'
    $HTML += '<li>Non Installée(s) entre 15 et 39 <span style="background-color:#f6ed60;">Jaune</span></li>'
    $HTML += '<li>MAJ Échoué(s) <span style="background-color:#8fc975;">Vert</span></li>'
    $HTML += '<li>En attente de reboot <span style="background-color:#70c3ed;">Bleu</span></li>'
    $HTML += '<li>Dernier rapport d état de plus de 7 jours <span style="background-color:#9a6db0;">Mauve</span></li>'
    $HTML += '<li>Dernier rapport d état de plus de 30 jours <span style="background-color:#c3add1;">Mauve Claire</span></li>'
    $HTML += '</ul>'
    $HTML = $HTML | Close-HTML -Validate
        
        if ($SendEmail) {
        # Send HTML to recipient(s)
        try {
            $HTML
            Send-MailMessage -From $FromAddress -To $RecipientAddress -Subject "$ComputerName - WSUS Status Report" -Body $HTML -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -Encoding $SMTPEncoding
        } catch {
            throw $_
        }
    } else {
        $HTML
    }
}
