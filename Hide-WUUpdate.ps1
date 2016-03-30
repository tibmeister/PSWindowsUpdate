<#
.SYNOPSIS
    Get a list of available updates meeting the criteria and try to hide/unhide it.
.DESCRIPTION
    Use Hide-WUUpdate to get list of available updates meeting specific criteria. In next step script try to hide (or unhide) updates.
    There are two types of filtering update: Pre-search criteria, Post-search criteria.
    - Pre-search works on server side, like example: ( IsInstalled = 0 and IsHidden = 0 and CategoryIds contains '0fa1201d-4330-4fa8-8ae9-b877473b6441' )
    - Post-search work on client side after downloading the pre-filtered list of updates, like example $KBArticleID -match $Update.KBArticleIDs
Pre-search
    Status list:
    D - IsDownloaded, I - IsInstalled, M - IsMandatory, H - IsHidden, U - IsUninstallable, B - IsBeta
.PARAMETER UpdateType
    Pre-search criteria. Finds updates of a specific type, such as 'Driver' and 'Software'. Default value contains all updates.
.PARAMETER UpdateID
    Pre-search criteria. Finds updates of a specific UUID (or sets of UUIDs), such as '12345678-9abc-def0-1234-56789abcdef0'.
.PARAMETER RevisionNumber
    Pre-search criteria. Finds updates of a specific RevisionNumber, such as '100'. This criterion must be combined with the UpdateID param.
.PARAMETER CategoryIDs
    Pre-search criteria. Finds updates that belong to a specified category (or sets of UUIDs), such as '0fa1201d-4330-4fa8-8ae9-b877473b6441'.
.PARAMETER IsInstalled
    Pre-search criteria. Finds updates that are installed on the destination computer.
.PARAMETER IsHidden
    Pre-search criteria. Finds updates that are marked as hidden on the destination computer.
.PARAMETER IsNotHidden
    Pre-search criteria. Finds updates that are not marked as hidden on the destination computer. Overwrite IsHidden param.
.PARAMETER Criteria
    Pre-search criteria. Set own string that specifies the search criteria.
.PARAMETER ShowSearchCriteria
    Show choosen search criteria. Only works for Pre-search criteria.
.PARAMETER Category
    Post-search criteria. Finds updates that contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...
.PARAMETER KBArticleID
    Post-search criteria. Finds updates that contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.
.PARAMETER Title
    Post-search criteria. Finds updates that match part of title, such as ''
.PARAMETER NotCategory
    Post-search criteria. Finds updates that not contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...
.PARAMETER NotKBArticleID
    Post-search criteria. Finds updates that not contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.
.PARAMETER NotTitle
    Post-search criteria. Finds updates that not match part of title.
.PARAMETER IgnoreUserInput
    Post-search criteria. Finds updates that the installation or uninstallation of an update can't prompt for user input.
.PARAMETER IgnoreRebootRequired
    Post-search criteria. Finds updates that specifies the restart behavior that not occurs when you install or uninstall the update.
.PARAMETER ServiceID
    Set ServiceID to change the default source of Windows Updates. It overwrite ServerSelection parameter value.
.PARAMETER WindowsUpdate
    Set Windows Update Server as source. Default update configs are taken from computer policy.
.PARAMETER MicrosoftUpdate
    Set Microsoft Update Server as source. Default update configs are taken from computer policy.
.PARAMETER HideStatus
    Status used in script. Default is $True = hide update.
.PARAMETER ComputerName	
    Specify the name of the computer to the remote connection.
.PARAMETER Debuger	
    Debug mode.
.EXAMPLE
    Hide-WUList -MicrosoftUpdate
.EXAMPLE
    Hide-WUUpdate -Title 'Windows Malicious*' -HideStatus:$false
.NOTES
    Author: Jody L. Whitlock
#>
Function Hide-WUUpdate
{
	[OutputType('PSWindowsUpdate.WUList')]
	[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]	
	Param
	(
		#Pre-search criteria
		[ValidateSet("Driver", "Software")]
		[String]$UpdateType = "",
		[String[]]$UpdateID,
		[Int]$RevisionNumber,
		[String[]]$CategoryIDs,
		[Switch]$IsInstalled,
		[Switch]$IsHidden,
		[Switch]$IsNotHidden,
		[String]$Criteria,
		[Switch]$ShowSearchCriteria,		
		
		#Post-search criteria
		[String[]]$Category="",
		[String[]]$KBArticleID,
		[String]$Title,
		
		[String[]]$NotCategory="",
		[String[]]$NotKBArticleID,
		[String]$NotTitle,	
		
		[Alias("Silent")]
		[Switch]$IgnoreUserInput,
		[Switch]$IgnoreRebootRequired,
		
		#Connection options
		[String]$ServiceID,
		[Switch]$WindowsUpdate,
		[Switch]$MicrosoftUpdate,
		[Switch]$HideStatus = $true,
		
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
		Write-Debug "STAGE 0: Prepare environment"
		######################################
		# Start STAGE 0: Prepare environment #
		######################################
		
		Write-Debug "Check if ComputerName in set"
		If($ComputerName -eq $null)
		{
			Write-Debug "Set ComputerName to localhost"
			[String[]]$ComputerName = $env:COMPUTERNAME
		}
		
		####################################			
		# End STAGE 0: Prepare environment #
		####################################
		
		$UpdateCollection = @()
		Foreach($Computer in $ComputerName)
		{
			If(Test-Connection -ComputerName $Computer -Quiet)
			{
				Write-Debug "STAGE 1: Get updates list"
				###################################
				# Start STAGE 1: Get updates list #
				###################################			

				If($Computer -eq $env:COMPUTERNAME)
				{
					Write-Debug "Create Microsoft.Update.ServiceManager object"
					$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" #Support local instance only
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

				If($WindowsUpdate)
				{
					Write-Debug "Set source of updates to Windows Update"
					$objSearcher.ServerSelection = 2
					$serviceName = "Windows Update"
				}
				ElseIf($MicrosoftUpdate)
				{
					Write-Debug "Set source of updates to Microsoft Update"
					$serviceName = $null
					Foreach ($objService in $objServiceManager.Services) 
					{
						If($objService.Name -eq "Microsoft Update")
						{
							$objSearcher.ServerSelection = 3
							$objSearcher.ServiceID = $objService.ServiceID
							$serviceName = $objService.Name
							Break
						}
					}
					
					If(-not $serviceName)
					{
						Write-Warning "Can't find registered service Microsoft Update. Use Get-WUServiceManager to get registered service."
						Return
					}
				}
				ElseIf($Computer -eq $env:COMPUTERNAME)
				{
					Foreach ($objService in $objServiceManager.Services) 
					{
						If($ServiceID)
						{
							If($objService.ServiceID -eq $ServiceID)
							{
								$objSearcher.ServiceID = $ServiceID
								$objSearcher.ServerSelection = 3
								$serviceName = $objService.Name
								Break
							}
						}
						Else
						{
							If($objService.IsDefaultAUService -eq $True)
							{
								$serviceName = $objService.Name
								Break
							}
						}
					}
				}
				ElseIf($ServiceID)
				{
					$objSearcher.ServiceID = $ServiceID
					$objSearcher.ServerSelection = 3
					$serviceName = $ServiceID
				}
				Else
				{
					$serviceName = "default (for $Computer) Windows Update"
				}
				Write-Debug "Set source of updates to $serviceName"
				
				Write-Verbose "Connecting to $serviceName server. Please wait..."
				Try
				{
					$search = ""
					If($Criteria)
					{
						$search = $Criteria
					}
					Else
					{
						If($IsInstalled) 
						{
							$search = "IsInstalled = 1"
							Write-Debug "Set Pre-search criteria: IsInstalled = 1"
						}
						Else
						{
							$search = "IsInstalled = 0"	
							Write-Debug "Set Pre-search criteria: IsInstalled = 0"
						}
						
						If($UpdateType -ne "")
						{
							Write-Debug "Set Pre-search criteria: Type = $UpdateType"
							$search += " and Type = '$UpdateType'"
						}					
						
						If($UpdateID)
						{
							Write-Debug "Set Pre-search criteria: UpdateID = '$([string]::join(", ", $UpdateID))'"
							$tmp = $search
							$search = ""
							$LoopCount = 0
							Foreach($ID in $UpdateID)
							{
								If($LoopCount -gt 0)
								{
									$search += " or "
								}
								If($RevisionNumber)
								{
									Write-Debug "Set Pre-search criteria: RevisionNumber = '$RevisionNumber'"	
									$search += "($tmp and UpdateID = '$ID' and RevisionNumber = $RevisionNumber)"
								}
								Else
								{
									$search += "($tmp and UpdateID = '$ID')"
								}
                                
								$LoopCount++
							}
						}

						If($CategoryIDs)
						{
							Write-Debug "Set Pre-search criteria: CategoryIDs = '$([string]::join(", ", $CategoryIDs))'"
							$tmp = $search
							$search = ""
							$LoopCount =0
							Foreach($ID in $CategoryIDs)
							{
								If($LoopCount -gt 0)
								{
									$search += " or "
								}
                                
								$search += "($tmp and CategoryIDs contains '$ID')"
								$LoopCount++
							}
						}
						
						If($IsNotHidden) 
						{
							Write-Debug "Set Pre-search criteria: IsHidden = 0"
							$search += " and IsHidden = 0"	
						}
						ElseIf($IsHidden) 
						{
							Write-Debug "Set Pre-search criteria: IsHidden = 1"
							$search += " and IsHidden = 1"	
						}

						#Don't know why every update has RebootRequired=false which is not always true
						If($IgnoreRebootRequired) 
						{
							Write-Debug "Set Pre-search criteria: RebootRequired = 0"
							$search += " and RebootRequired = 0"	
						}
					}
					
					Write-Debug "Search criteria is: $search"
					
					If($ShowSearchCriteria)
					{
						Write-Output $search
					}
			
					$objResults = $objSearcher.Search($search)
				}
				Catch
				{
					If($_ -match "HRESULT: 0x80072EE2")
					{
						Write-Warning "Probably you don't have connection to Windows Update server"
					}
					Return
				}

				$NumberOfUpdate = 1
				$PreFoundUpdatesToDownload = $objResults.Updates.count
				Write-Verbose "Found [$PreFoundUpdatesToDownload] Updates in Pre-search criteria"				
				
				If($PreFoundUpdatesToDownload -eq 0)
				{
					Continue
				} 
				
				Foreach($Update in $objResults.Updates)
				{	
					$UpdateAccess = $true
					Write-Progress -Activity "Post-search updates for $Computer" -Status "[$NumberOfUpdate/$PreFoundUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$PreFoundUpdatesToDownload * 100))
					Write-Debug "Set Post-search criteria: $($Update.Title)"
					
					If($Category -ne "")
					{
						$UpdateCategories = $Update.Categories | Select-Object Name
						Write-Debug "Set Post-search criteria: Categories = '$([string]::join(", ", $Category))'"	
						Foreach($Cat in $Category)
						{
							If(!($UpdateCategories -match $Cat))
							{
								Write-Debug "UpdateAccess: false"
								$UpdateAccess = $false
							}
							Else
							{
								$UpdateAccess = $true
								Break
							}
						}	
					}

					If($NotCategory -ne "" -and $UpdateAccess -eq $true)
					{
						$UpdateCategories = $Update.Categories | Select-Object Name
						Write-Debug "Set Post-search criteria: NotCategories = '$([string]::join(", ", $NotCategory))'"	
						Foreach($Cat in $NotCategory)
						{
							If($UpdateCategories -match $Cat)
							{
								Write-Debug "UpdateAccess: false"
								$UpdateAccess = $false
								Break
							}
						}	
					}					
					
					If($KBArticleID -ne $null -and $UpdateAccess -eq $true)
					{
						Write-Debug "Set Post-search criteria: KBArticleIDs = '$([string]::join(", ", $KBArticleID))'"
						If(!($KBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs))
						{
							Write-Debug "UpdateAccess: false"
							$UpdateAccess = $false
						}								
					}

					If($NotKBArticleID -ne $null -and $UpdateAccess -eq $true)
					{
						Write-Debug "Set Post-search criteria: NotKBArticleIDs = '$([string]::join(", ", $NotKBArticleID))'"
						If($NotKBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs)
						{
							Write-Debug "UpdateAccess: false"
							$UpdateAccess = $false
						}					
					}
					
					If($Title -and $UpdateAccess -eq $true)
					{
						Write-Debug "Set Post-search criteria: Title = '$Title'"
						If($Update.Title -notmatch $Title)
						{
							Write-Debug "UpdateAccess: false"
							$UpdateAccess = $false
						}
					}

					If($NotTitle -and $UpdateAccess -eq $true)
					{
						Write-Debug "Set Post-search criteria: NotTitle = '$NotTitle'"
						If($Update.Title -match $NotTitle)
						{
							Write-Debug "UpdateAccess: false"
							$UpdateAccess = $false
						}
					}
					
					If($IgnoreUserInput -and $UpdateAccess -eq $true)
					{
						Write-Debug "Set Post-search criteria: CanRequestUserInput"
						If($Update.InstallationBehavior.CanRequestUserInput -eq $true)
						{
							Write-Debug "UpdateAccess: false"
							$UpdateAccess = $false
						}
					}

					If($IgnoreRebootRequired -and $UpdateAccess -eq $true) 
					{
						Write-Debug "Set Post-search criteria: RebootBehavior"
						If($Update.InstallationBehavior.RebootBehavior -ne 0)
						{
							Write-Debug "UpdateAccess: false"
							$UpdateAccess = $false
						}	
					}

					If($UpdateAccess -eq $true)
					{
						Write-Debug "Convert size"
						Switch($Update.MaxDownloadSize)
						{
							{[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
							{[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }  
							{[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }    
							{[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
							default { $size = $_+"B" }
						}
					
						Write-Debug "Convert KBArticleIDs"
						If($Update.KBArticleIDs -ne "")    
						{
							$KB = "KB"+$Update.KBArticleIDs
						}
						Else 
						{
							$KB = ""
						}
						
						if($Update.IsHidden -ne $HideStatus)
						{
							if($HideStatus)
							{
								$StatusName = "Hide"
							}
							else
							{
								$StatusName = "Unhide"
							}
							
							If($pscmdlet.ShouldProcess($Computer,"$StatusName $($Update.Title)?")) 
							{
								Try
								{
									$Update.IsHidden = $HideStatus
								}
								Catch
								{
									Write-Warning "You haven't privileges to make this. Try start an eleated Windows PowerShell console."
								}
								
							}
						}
						
						$Status = ""
				        If($Update.IsDownloaded)    {$Status += "D"} else {$status += "-"}
				        If($Update.IsInstalled)     {$Status += "I"} else {$status += "-"}
				        If($Update.IsMandatory)     {$Status += "M"} else {$status += "-"}
				        If($Update.IsHidden)        {$Status += "H"} else {$status += "-"}
				        If($Update.IsUninstallable) {$Status += "U"} else {$status += "-"}
				        If($Update.IsBeta)          {$Status += "B"} else {$status += "-"} 
		
						Add-Member -InputObject $Update -MemberType NoteProperty -Name ComputerName -Value $Computer
						Add-Member -InputObject $Update -MemberType NoteProperty -Name KB -Value $KB
						Add-Member -InputObject $Update -MemberType NoteProperty -Name Size -Value $size
						Add-Member -InputObject $Update -MemberType NoteProperty -Name Status -Value $Status
					
						$Update.PSTypeNames.Clear()
						$Update.PSTypeNames.Add('PSWindowsUpdate.WUList')
						$UpdateCollection += $Update
					}
					
					$NumberOfUpdate++
				}				
				Write-Progress -Activity "Post-search updates for $Computer" -Status "Completed" -Completed
				
				$FoundUpdatesToDownload = $UpdateCollection.count
				Write-Verbose "Found [$FoundUpdatesToDownload] Updates in Post-search criteria"
				
				#################################
				# End STAGE 1: Get updates list #
				#################################
				
			}
		}

		Return $UpdateCollection
	}
	
	End{}		
}