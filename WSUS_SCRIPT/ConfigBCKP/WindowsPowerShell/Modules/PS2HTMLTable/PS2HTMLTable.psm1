# Create PSScriptRoot variable for PowerShell Version 2
if (-not $PSScriptRoot) {
	$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
}

# Get public and private function definition files.
$Public  = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

# Dot source the files
foreach ($Import in @($Public + $Private)) {
	try {
		. $Import.FullName
	} catch {
		Write-Error -Message "Failed to import function $($Import.FullName): $_"
	}
}

Export-ModuleMember -Function ($Public | Select-Object -ExpandProperty BaseName)