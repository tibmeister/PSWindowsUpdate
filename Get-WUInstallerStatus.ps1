<#
.SYNOPSIS
    Shows Windows Update Installer status.
.DESCRIPTION
    Use Get-WUInstallerStatus to show Windows Update Installer status.
.PARAMETER Silent
    Get only status True/False without any more comments on screen.
.EXAMPLE
    Get-WUInstallerStatus
.EXAMPLE
    Get-WUInstallerStatus -Silent
.NOTES
    Author: Jody L. Whitlock
#>
Function Get-WUInstallerStatus
{	
	[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="Low")]
    Param
	(
		[Switch]$Silent
	)
	
	Begin
	{
		$User = [Security.Principal.WindowsIdentity]::GetCurrent()
		$Role = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

		if(!$Role)
		{
			Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."
            write-warning "Please re-run the script with elevated rights."	
            exit
		}		
	}
	
	Process
	{
        If ($pscmdlet.ShouldProcess($Env:COMPUTERNAME,"Check that Windows Installer is ready to install next updates")) 
		{	    
			$objInstaller=New-Object -ComObject "Microsoft.Update.Installer"
			
			Switch($objInstaller.IsBusy)
			{
				$true	{ If($Silent) {Return $true} Else {Write-Output "Installer is busy."}}
				$false	{ If($Silent) {Return $false} Else {Write-Output "Installer is ready."}}
			}
			
		}
	}
	
	End{}	
}