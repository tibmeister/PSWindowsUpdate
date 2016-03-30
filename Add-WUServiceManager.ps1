<#
.SYNOPSIS
    Register windows update service manager.
.DESCRIPTION
    Use Add-WUServiceManager to register new Windows Update Service Manager.
.PARAMETER ServiceID	
    An identifier for the service to be registered in GUID format 
.PARAMETER AddServiceFlag	
    A combination of AddServiceFlag values. 0x1 - asfAllowPendingRegistration, 0x2 - asfAllowOnlineRegistration, 0x4 - asfRegisterServiceWithAU
.PARAMETER authorizationCabPath	
    The path of the Microsoft signed local cabinet file (.cab) that has the information that is required for a service registration. If empty, the update agent searches for the authorization cabinet file (.cab) during service registration when a network connection is available.
.EXAMPLE
    Try to register Microsoft Update Service.

    PS H:\> Add-WUServiceManager -ServiceID "3bcb6f2a-58f3-477b-aabe-6c2ae3fa5b52"
.NOTES
    Author: Jody L. Whitlock
#>
Function Add-WUServiceManager 
{
    [OutputType('PSWindowsUpdate.WUServiceManager')]
	[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
    Param
    (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$ServiceID,
		[Int]$AddServiceFlag = 2,
		[String]$authorizationCabPath
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
        $objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
        Try
        {
            If ($pscmdlet.ShouldProcess($Env:COMPUTERNAME,"Register Windows Update Service Manager: $ServiceID")) 
			{
				
				$objService = $objServiceManager.AddService2($ServiceID,$AddServiceFlag,$authorizationCabPath)
				$objService.PSTypeNames.Clear()
				$objService.PSTypeNames.Add('PSWindowsUpdate.WUServiceManager')
			}
        }
        Catch 
        {
            If($_ -match "HRESULT: 0x80070005")
            {
                Write-Warning "Your security policy don't allow a non-administator identity to perform this task"
            }
			Else
			{
				Write-Error $_
			}
			
            Return
        }
		
        Return $objService	
	}
	End{}
}
