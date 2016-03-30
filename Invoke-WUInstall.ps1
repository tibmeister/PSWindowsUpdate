<#
.SYNOPSIS
    Invoke Get-WUInstall remotely.

.DESCRIPTION
    Use Invoke-WUInstall to invoke Windows Update install remotely. It uses the Task Scheduler because 
    CreateUpdateDownloader() and CreateUpdateInstaller() methods can't be called from a remote computer - E_ACCESSDENIED.
    
    Note:
    Because we do not have the ability to interact with the script at this time, is recommended use -AcceptAll with WUInstall filters in script block.

.PARAMETER ComputerName
    Specify computer name.

.PARAMETER TaskName
    Specify task name. Default is PSWindowsUpdate.
    
.PARAMETER Script
    Specify PowerShell script block that you what to run. Default is {ipmo PSWindowsUpdate; Get-WUInstall -AcceptAll | Out-File C:\PSWindowsUpdate.log}
    
.EXAMPLE
    PS C:\> $Script = {ipmo PSWindowsUpdate; Get-WUInstall -AcceptAll -AutoReboot | Out-File C:\PSWindowsUpdate.log}
    PS C:\> Invoke-WUInstall -ComputerName pc1.contoso.com -Script $Script
    
    
.NOTES
    Author: Jody L. Whitlock
#>
Function Invoke-WUInstall
{
	[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
	param
	(
		[Parameter(ValueFromPipeline=$True,
					ValueFromPipelineByPropertyName=$True)]
		[String[]]$ComputerName,
		[String]$TaskName = "PSWindowsUpdate",
		[ScriptBlock]$Script = {ipmo PSWindowsUpdate; Get-WUInstall -AcceptAll | Out-File C:\PSWindowsUpdate.log},
		[Switch]$OnlineUpdate
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
		
		$PSWUModule = Get-Module -Name PSWindowsUpdate -ListAvailable
		
		Write-Verbose "Create schedule service object"
		$Scheduler = New-Object -ComObject Schedule.Service
			
		$Task = $Scheduler.NewTask(0)

		$RegistrationInfo = $Task.RegistrationInfo
		$RegistrationInfo.Description = $TaskName
		$RegistrationInfo.Author = $User.Name

		$Settings = $Task.Settings
		$Settings.Enabled = $True
		$Settings.StartWhenAvailable = $True
		$Settings.Hidden = $False

		$Action = $Task.Actions.Create(0)
		$Action.Path = "powershell"
		$Action.Arguments = "-Command $Script"
		
		$Task.Principal.RunLevel = 1	
	}
	
	Process
	{
		ForEach($Computer in $ComputerName)
		{
			If ($pscmdlet.ShouldProcess($Computer,"Invoke WUInstall")) 
			{
				if(Test-Connection -ComputerName $Computer -Quiet)
				{
					Write-Verbose "Check PSWindowsUpdate module on $Computer"
					Try
					{
						$ModuleTest = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-Module -ListAvailable -Name PSWindowsUpdate} -ErrorAction Stop
					}
					Catch
					{
						Write-Warning "Can't access to machine $Computer. Try use: winrm qc"
						Continue
					} 
					$ModulStatus = $false
					
					if($ModuleTest -eq $null -or $ModuleTest.Version -lt $PSWUModule.Version)
					{
						if($OnlineUpdate)
						{
							Update-WUModule -ComputerName $Computer
						} 
						else
						{
							Update-WUModule -ComputerName $Computer	-LocalPSWUSource (Get-Module -ListAvailable -Name PSWindowsUpdate).ModuleBase
						} 
					} 
					
					#Sometimes can't connect at first time
					$Info = "Connect to scheduler and register task on $Computer"
					for ($i=1; $i -le 3; $i++)
					{
						$Info += "."
						Write-Verbose $Info
						Try
						{
							$Scheduler.Connect($Computer)
							Break
						} 
						Catch
						{
							if($i -ge 3)
							{
								Write-Error "Can't connect to Schedule service on $Computer" -ErrorAction Stop
							} 
							else
							{
								sleep -Seconds 1
							} 
						} 					
					} 
					
					$RootFolder = $Scheduler.GetFolder("\")
					$SendFlag = 1
					if($Scheduler.GetRunningTasks(0) | Where-Object {$_.Name -eq $TaskName})
					{
						$CurrentTask = $RootFolder.GetTask($TaskName)
						$Title = "Task $TaskName is curretly running: $($CurrentTask.Definition.Actions | Select-Object -exp Path) $($CurrentTask.Definition.Actions | Select-Object -exp Arguments)"
						$Message = "What do you want to do?"

						$ChoiceContiniue = New-Object System.Management.Automation.Host.ChoiceDescription "&Continue Current Task"
						$ChoiceStart = New-Object System.Management.Automation.Host.ChoiceDescription "Stop and Start &New Task"
						$ChoiceStop = New-Object System.Management.Automation.Host.ChoiceDescription "&Stop Task"
						$Options = [System.Management.Automation.Host.ChoiceDescription[]]($ChoiceContiniue, $ChoiceStart, $ChoiceStop)
						$SendFlag = $host.ui.PromptForChoice($Title, $Message, $Options, 0)
					
						if($SendFlag -ge 1)
						{
							($RootFolder.GetTask($TaskName)).Stop(0)
						} 	
						
					} 
						
					if($SendFlag -eq 1)
					{
						$RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, "SYSTEM", $Null, 1) | Out-Null
						$RootFolder.GetTask($TaskName).Run(0) | Out-Null
					} 
					
				} 
				else
				{
					Write-Warning "Machine $Computer is not responding."
				} 
			} 
		} 
		Write-Verbose "Invoke-WUInstall complete."
	}
	
	End {}

}