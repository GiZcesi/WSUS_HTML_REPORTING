###########################
$UserName = 'DefaultUsername' #  /!\ PLZ Change the Username to current user used for this reporting /!\
###########################


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

$choice = 1 #Forcing Choice 1 cause we want the server report
$currentDir = "C:\Users\$UserName\Documents\WSUS_Script\WindowsPowerShell\Modules\PS2HTMLTable\Scripts"
$currentDate = Get-Date -Format "yyyy-MM-dd"

# Set the script path based on the user's choice
if ($choice -eq 1) {
    $scriptPath = "$currentDir\Get-WSUSComputerStatus_Serveurs.ps1"
    $outputFile = "Reporting_serveurs_$currentDate.html"
}
elseif ($choice -eq 2) {
    $scriptPath = "$currentDir\Get-WSUSComputerStatus_Ordinateurs.ps1"
    $outputFile = "Reporting_ordinateurs_$currentDate.html"
}
else {
    #Write-Host "Choix Invalide."
    return
}

# Define the pipeline that should be run
$pipeline = "& { & '$scriptPath' | Out-File C:\Users\$UserName\Documents\WSUS_Script\Export\$outputFile }"

# Start the script with admin privileges, including the pipeline
Start-Process powershell.exe  -Verb runAs -ArgumentList "-Command $pipeline"
Start-Process "C:\Users\$UserName\Documents\WSUS_Script\Export\$outputFile"

