#Make sure we are running Powershell 4 so we can unblock the files
if($PSVersionTable.PSVersion.Major -gt 4)
{
	Get-ChildItem -Path $PSScriptRoot | Unblock-File
}

#Run all the files so that we can get the functions into memory
Get-ChildItem -Path $PSScriptRoot\*.ps1 | Foreach-Object{ . $_.FullName }