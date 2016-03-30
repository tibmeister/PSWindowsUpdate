<#
.SYNOPSIS
    Gets a list of Windows Update history
.DESCRIPTION
    Use function Get-WUHistory to get list of installed updates on current machine. It works similar like Get-Hotfix.
.PARAMETER ComputerName	
    Specify the name of the computer to the remote connection.
.PARAMETER Debuger	
    Debug mode.
.EXAMPLE  
    Get information about specific installed updates.

    $WUHistory = Get-WUHistory
    $WUHistory | Where-Object {$_.Title -match "KB2607047"} | Select-Object *
.NOTES
    Author: Jody L. Whitlock
#>
Function Get-WUHistory
{
	[OutputType('PSWindowsUpdate.WUHistory')]
	[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="Low")]
	Param
	(
		#Mode options
		[Switch]$Debuger,
		[parameter(ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true)]
		[String[]]$ComputerName	
	)

	Begin
	{
		If($PSBoundParameters['Debuger'])
		{
			$DebugPreference = "Continue"
		} 

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
		#region STAGE 0
		Write-Debug "STAGE 0: Prepare environment"
		######################################
		# Start STAGE 0: Prepare environment #
		######################################
		
		Write-Debug "Check if ComputerName in set"
		If($ComputerName -eq $null)
		{
			Write-Debug "Set ComputerName to localhost"
			[String[]]$ComputerName = $env:COMPUTERNAME
		} #End If $ComputerName -eq $null

		####################################
		# End STAGE 0: Prepare environment #
		####################################
		#endregion
		
		$UpdateCollection = @()
		Foreach($Computer in $ComputerName)
		{
			If(Test-Connection -ComputerName $Computer -Quiet)
			{
				#region STAGE 1
				Write-Debug "STAGE 1: Get history list"
				###################################
				# Start STAGE 1: Get history list #
				###################################
		
				If ($pscmdlet.ShouldProcess($Computer,"Get updates history")) 
				{
					Write-Verbose "Get updates history for $Computer"
					If($Computer -eq $env:COMPUTERNAME)
					{
						Write-Debug "Create Microsoft.Update.Session object for $Computer"
						$objSession = New-Object -ComObject "Microsoft.Update.Session" #Support local instance only
					} 
					Else
					{
						Write-Debug "Create Microsoft.Update.Session object for $Computer"
						$objSession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computer))
					}

					Write-Debug "Create Microsoft.Update.Session.Searcher object for $Computer"
					$objSearcher = $objSession.CreateUpdateSearcher()
					$TotalHistoryCount = $objSearcher.GetTotalHistoryCount()

					If($TotalHistoryCount -gt 0)
					{
						$objHistory = $objSearcher.QueryHistory(0, $TotalHistoryCount)
						$NumberOfUpdate = 1
						Foreach($obj in $objHistory)
						{
							Write-Progress -Activity "Get update histry for $Computer" -Status "[$NumberOfUpdate/$TotalHistoryCount] $($obj.Title)" -PercentComplete ([int]($NumberOfUpdate/$TotalHistoryCount * 100))

							Write-Debug "Get update histry: $($obj.Title)"
							Write-Debug "Convert KBArticleIDs"
							$matches = $null
							$obj.Title -match "KB(\d+)" | Out-Null
							
							If($matches -eq $null)
							{
								Add-Member -InputObject $obj -MemberType NoteProperty -Name KB -Value ""
							}
							Else
							{							
								Add-Member -InputObject $obj -MemberType NoteProperty -Name KB -Value ($matches[0])
							}
							
							Add-Member -InputObject $obj -MemberType NoteProperty -Name ComputerName -Value $Computer
							
							$obj.PSTypeNames.Clear()
							$obj.PSTypeNames.Add('PSWindowsUpdate.WUHistory')
						
							$UpdateCollection += $obj
							$NumberOfUpdate++
						}
						Write-Progress -Activity "Get update histry for $Computer" -Status "Completed" -Completed
					}
					Else
					{
						Write-Warning "Probably your history was cleared. Alternative please run 'Get-WUList -IsInstalled'"
					}
				}
				
				################################
				# End PASS 1: Get history list #
				################################
				#endregion
				
			}
		}	
		
		Return $UpdateCollection
	}

	End{}	
}