$ErrorActionPreference= 'silentlycontinue'

param([switch]$Elevated)

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}


if (-not (Get-Module -Name PS2HTMLTable -ErrorAction SilentlyContinue)) {
  Install-Module -Name PS2HTMLTable -AllowClobber -Force
}
if (-not (Get-Module -Name PSWriteHTML -ErrorAction SilentlyContinue)) {
  Install-Module -Name PSWriteHTML -AllowClobber -Force
}
