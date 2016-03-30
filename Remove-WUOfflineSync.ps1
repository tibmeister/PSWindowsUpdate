<#
.SYNOPSIS
    Unregisters offline scanner service.
.DESCRIPTION
    Use Remove-WUOfflineSync to unregister Windows Update offline scan file (wsusscan.cab or wsusscn2.cab) from current machine.
.EXAMPLE
    Remove-WUOfflineSync
.NOTES
    Author: Jody L. Whitlock
#>
Function Remove-WUOfflineSync
{
	[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
    Param()
	
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
	    $objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
	    
		$State = 1
	    Foreach ($objService in $objServiceManager.Services) 
	    {
	        If($objService.Name -eq "Offline Sync Service")
	        {
	           	If ($pscmdlet.ShouldProcess($Env:COMPUTERNAME,"Unregister Windows Update offline scan file")) 
				{
					Try
					{
						$objServiceManager.RemoveService($objService.ServiceID)
					}
					Catch
					{
			            If($_ -match "HRESULT: 0x80070005")
			            {
			                Write-Warning "Your security policy doesn't allow a non-administator identity to perform this task"
			            }
						Else
						{
							Write-Error $_
						}
						
			            Return
					}
	            }
				
				Get-WUServiceManager
	            $State = 0;    
				
	        }
	    }
	    
	    If($State)
	    {
	        Write-Warning "Offline Sync Service don't exist on current machine."
	    }
	}
	
	End{}
}