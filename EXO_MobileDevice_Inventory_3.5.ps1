#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Exchange Online Device partnership inventory, dependent on EXOv2 module being installed 
#			(https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
#  EXO_MobileDevice_Inventory_3.1.ps1
#  
#  Created by: Austin McCollum 2/11/2018 austinmc@microsoft.com
#  Updated by: Garrin Thompson 7/23/2020 garrint@microsoft.com *** "Borrowed" a few 
#	quality-of-life functions from Start-RobustCloudCommand.ps1 and added EXOv2 connection
#
#########################################################################################
# This script enumerates all devices in Office 365 and reports on many properties of the
#   device/application and the mailbox owner.
#
# $ResultsList is an array of hashtables, because deviceIDs may not be
#   unique in an environment. For instance when a device is configured with
#   two separate mailboxes in the same org, the same deviceID will appear twice.
#   Hashtables require uniqueness of the key so that's why the array of Hashtable data 
#   structure was chosen.
#
# The devices can be sorted by a variety of properties like "LastPolicyUpdate" to determine 
#   stale partnerships or outdated devices needing to be removed.
# 
# The DisplayName of the user's CAS mailbox is recorded for importing with the 
#   Set-CasMailbox commandlet to configure allowedDeviceIDs. This is especially useful in 
#   scenarios where a migration to ABQ framework requires "grandfathering" in all or some
#   of the existing partnerships.
#
# Get-CasMailbox is run efficiently with the -HasActiveSyncDevicePartnership filter 
#########################################################################################

# Writes output to a log file with a time date stamp
Function Write-Log {
	Param ([string]$string)
	$NonInteractive = 1
	# Get the current date
	[string]$date = Get-Date -Format G
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	# If NonInteractive true then supress host output
	If (!($NonInteractive)){
		( "[" + $date + "] - " + $string) | Write-Host
	}
}

# Sleeps X seconds and displays a progress bar
Function Start-SleepWithProgress {
	Param([int]$sleeptime)
	# Loop Number of seconds you want to sleep
	For ($i=0;$i -le $sleeptime;$i++){
		$timeleft = ($sleeptime - $i);
		# Progress bar showing progress of the sleep
		Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		# Sleep 1 second
		start-sleep 1
	}
	Write-Progress -Completed -Activity "Sleeping"
}

# Setup a new O365 Powershell Session using RobustCloudCommand concepts to help maintain the session
Function New-CleanO365Session {
	 #Prompt for UPN used to login to EXO 
	Write-log ("Removing all PS Sessions")

	# Destroy any outstanding PS Session
	Get-PSSession | Remove-PSSession -Confirm:$false
	
	# Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
	[System.GC]::Collect()
	
	# Sleep 10s to allow the sessions to tear down fully
	Write-Log ("Sleeping 10 seconds to clear existing PS sessions")
	Start-Sleep -Seconds 10

	# Clear out all errors
	$Error.Clear()
	
	# Create the session
	Write-Log ("Creating new PS Session")
		#OLD BasicAuth method create session
			#$Exchangesession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
	# Check for an error while creating the session
		If ($Error.Count -gt 0){
			Write-log ("[ERROR] - Error while setting up session")
			Write-log ($Error)
			# Increment our error count so we abort after so many attempts to set up the session
			$ErrorCount++
			# If we have failed to setup the session > 3 times then we need to abort because we are in a failure state
			If ($ErrorCount -gt 3){
				Write-log ("[ERROR] - Failed to setup session after multiple tries")
				Write-log ("[ERROR] - Aborting Script")
				exit		
			}	
			# If we are not aborting then sleep 60s in the hope that the issue is transient
			Write-log ("Sleeping 60s then trying again...standby")
			Start-SleepWithProgress -sleeptime 60
			
			# Attempt to set up the sesion again
			New-CleanO365Session
		}
	
	# If the session setup worked then we need to set $errorcount to 0
	else {
		$ErrorCount = 0
	}
	# Import the PS session/connect to EXO
		$null = Connect-ExchangeOnline -UserPrincipalName $EXOLogonUPN -DelegatedOrganization $EXOtenant -ShowProgress:$false -ShowBanner:$false
	# Set the Start time for the current session
		Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy; Resets it every "$ResetSeconds" number of seconds (14.5 mins) either way 
Function Test-O365Session {
	# Get the time that we are working on this object to use later in testing
	$ObjectTime = Get-Date
	# Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession
	# Make sure we found a session
	If ($SessionInfo -eq $null) { 
		Write-log ("[ERROR] - No Session Found")
		Write-log ("Recreating Session")
		New-CleanO365Session
	}	
	# Make sure it is in an opened state If not log and recreate
	elseif ($SessionInfo.State -ne "Opened"){
		Write-log ("[ERROR] - Session not in Open State")
		Write-log ($SessionInfo | fl | Out-String )
		Write-log ("Recreating Session")
		New-CleanO365Session
	}
	# If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
		Write-log ("Rebuilding Connection")
		
		# Estimate the throttle delay needed since the last session rebuild
		# Amount of time the session was allowed to run * our activethrottle value
		# Divide by 2 to account for network time, script delays, and a fudge factor
		# Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		[int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		# If the delay is >15s then sleep that amount for throttle to recover
		If ($DelayinSeconds -gt 0){
			Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
			Start-SleepWithProgress -SleepTime $DelayinSeconds
		}
		# If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
		else {
			Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		}
		# new O365 session and reset our object processed count
		New-CleanO365Session
	}
	else {
		# If session is active and it hasn't been open too long then do nothing and keep going
	}
	# If we have a manual throttle value then sleep for that many milliseconds
	If ($ManualThrottle -gt 0){
		Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
		Start-Sleep -Milliseconds $ManualThrottle
	}
}

#------------------v
#ScriptSetupSection
#------------------v

#Set Variables
	$logfilename = '\EXO_LongRunnerScript_Execution_logfile_'
	$outputfilename = '\EXO_LongRunnerScript_Output_'
	$execpol = get-executionpolicy
	Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force  #this is just for the session running this script
	Write-Host;$EXOLogonUPN=Read-host "Type in UPN for account that will execute this script";$EXOtenant=Read-host "Type in your tenant domain name (eg <domain>.onmicrosoft.com)";write-host "...pleasewait...connecting to EXO..."
	$SmtpCreds = (get-credential -Message "Provide EXO account Pasword" -UserName "$EXOLogonUPN")
	# Set $OutputFolder to Current PowerShell Directory
	[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))
	$outputFolder = [IO.Directory]::GetCurrentDirectory()
	$logFile = $outputFolder + $logfilename + (Get-Date).Ticks + ".txt"
	$OutputFile= $outputfolder + $outputfilename + (Get-Date).Ticks + ".csv"
	[int]$ManualThrottle=0
	[double]$ActiveThrottle=.25
	[int]$ResetSeconds=870
# Setup our first session to O365
	$ErrorCount = 0
	New-CleanO365Session
	Write-Log ("Connected to Exchange Online")
	write-host;write-host -ForegroundColor Green "...Connected to Exchange Online as $EXOLogonUPN";write-host
# Get when we started the script for estimating time to completion
	$ScriptStartTime = Get-Date
	$startDate = Get-Date
	write-progress -id 1 -activity "Beginning..." -PercentComplete (1) -Status "initializing variables"
# Clear the error log so that sending errors to file relate only to this run of the script
	$error.clear()

#-------------------------v
#Start CUSTOM CODE Section
#-------------------------v

# Create Arrays
	# Prepare a new array for the first set of Objects you want to get from EXO and set the attributes list
		[System.Collections.ArrayList]$Object1Users = New-Object System.Collections.ArrayList($null)
		$Object1Users | Select distinguishedname,displayname,id,primarysmtpaddress,activesyncmailboxpolicy,activesyncsuppressreadreceipt,activesyncdebuglogging,activesyncallowedids,activesyncblockeddeviceids
		$Object1Users.Clear()

	# Create a new array for the second set of Objects you want to get from EXO and set the attributes list
		[System.Collections.ArrayList]$Object1List = New-Object System.Collections.ArrayList($null)
		$Object1List | Select friendlyname,deviceid,DeviceOS,DeviceModel,DeviceUseragent,devicetype,FirstSyncTime,WhenChangedUTC,identity,clientversion,clienttype,ismanaged,DeviceAccessState,DeviceAccessStateReason
		$Object1List.Clear()

# Get Data from EXO and populate arrays
	# Fill in $Object1Users array with data from EXO cmds
		Write-Progress -Id 1 -Activity "Getting all EXO users with Devices" -PercentComplete (10) -Status "Get-CasMailbox -ResultSize Unlimited"
		$Object1Users = Invoke-Command -Session (Get-PSSession) -ScriptBlock {
			Get-CASMailbox -RecalculateHasActiveSyncDevicePartnership -ResultSize unlimited -Filter {HasActiveSyncDevicePartnership -eq "True"} | Select-Object -Property distinguishedname,displayname,id,primarysmtpaddress,activesyncmailboxpolicy,activesyncsuppressreadreceipt,activesyncdebuglogging,activesyncallowedids,activesyncblockeddeviceids
			}
	# Fill in the $Object1List array with data from EXO cmds
		$mobiledevices = $null
		ForEach ($dvcuser in $Object1Users) {
			$mobiledevices = Get-MobileDevice -Mailbox $dvcuser.id.name | Select-Object -Property friendlyname,deviceid,DeviceOS,DeviceModel,DeviceUseragent,devicetype,FirstSyncTime,WhenChangedUTC,identity,clientversion,clienttype,ismanaged,DeviceAccessState,DeviceAccessStateReason
			$Object1List += $mobiledevices
		}
		write-progress -id 1 -Activity "Getting all EXO Devices" -PercentComplete (5) -Status "Get-MobileDevice is running"

	# Measure and record the time it takes to enumerate objects from Exchange Online 
		$progressActions = $Object1List.count
		$invokeEndDate = Get-Date
		$invokeElapsedTime = $invokeEndDate - $startDate
		#Update Log
			Write-Log ("Starting device collection");Write-Log ("Number of Devices found in Exchange Online: " + ($progressActions));Write-Log ("Time to run Invoke command for Device retrieval: " + ($($invokeElapsedTime)))
		#Update Screen
			write-host -foregroundcolor Cyan "Starting device collection";;sleep 2;write-host "-------------------------------------------------"
			Write-Host -NoNewline "Total Devices found for users with a device:      ";Write-Host -ForegroundColor Green $progressActions
			Write-Host -NoNewline "Time to run Invoke command for Device retrieval:  ";write-host -ForegroundColor Yellow "$($invokeElapsedTime)"
		$casMailboxUnlimitedEndDate = Get-Date
		$casMailboxUnlimitedElapsedTime = $casMailboxUnlimitedEndDate - $invokeEndDate
		#Update Log
			Write-Log ("Number of Users with Devices in Exchange Online: " + $($Object1Users.count));Write-Log ("Time for User retrieval via Get-CASMailbox run: " + $($casMailboxUnlimitedElapsedTime))
		#Update Screen
			Write-Host -NoNewline "Number of Users with Devices in Exchange Online:  ";Write-Host -ForegroundColor Green "$($Object1Users.count)"
			Write-Host -NoNewline "Time to retrieve User info via Get-CasMailbox:    ";write-host -ForegroundColor Yellow "$($casMailboxUnlimitedElapsedTime)"

# Set a counter and some variables to use for periodic write/flush and reporting for loop to create Hashtable
	$currentProgress = 1
	[TimeSpan]$caseCheckTotalTime=0
	# report counter
		$c = 0
	# running counter
		$i = 0
	# Set the number of objects to cycle before writing to disk and sending stats, i'd consider 5000 max
		$statLimit = 1000
	# Get the total number of devices, which we use in some stat calculations
		$t = $Object1List.count
	# Set some timedate variables for the stats report
		$loopStartTime = Get-Date
		$loopCurrentTime = Get-Date

#  Create a new array for use in a Hashtable containing calculated properties indexed by a property from the both lists
	[System.Collections.ArrayList]$ResultsList = New-Object System.Collections.ArrayList($null)
	$ResultsList.Clear()

#  Now from the two arrays, let's create the output data...this is a big LOOP
	ForEach ($Object1 in $Object1List) {
		#Check that we still have a valid EXO PS session
			Test-O365Session
		# Total up the running count 
			$i++
		# Dump the $ResultsList to CSV at every $statLimit number of objects (defined above); also send status e-mail with some metrics at each dump.
			If (++$c -eq $statLimit) {
				# Moved this from the bottom of the script, and added -Append parameter
					$ResultsList | select DisplayName,User,UserId,PrimarySMTPAddress,FriendlyName,UserAgent,FirstSyncTime,LastPolicyUpdate,DeviceOS,ClientProtocolVersion,ClientType,DeviceModel,DeviceId,AccessState,AccessReason,ActivesyncSuppressReadReceipt,ActivesyncDebugLogging,Managed,DistinguishedName,RemovalId | export-csv -path $OutputFile -notypeinformation -Append
					$loopLastTime = $loopCurrentTime
					$loopCurrentTime = Get-Date
					$currentRate = $statLimit/($loopCurrentTime-$loopLastTime).TotalHours
					$avgRate = $i/($loopCurrentTime-$loopStartTime).TotalHours
				# Send a status email each time we write $statimit number of objects to the file (requires $SmtpCreds to be defined)
					$old_ErrorActionPreference = $ErrorActionPreference
					$ErrorActionPreference = 'SilentlyContinue'
					Send-MailMessage -From "$EXOLogonUPN" -To "$EXOLogonUPN" -Subject "$OutputFile : Progress" -Body "$OutputFile PROGRESS report`n`nCurrentTime: $loopCurrentTime`nStartTime: $loopStartTime`n`nCounter: $i out of $t devices, at a current rate of $currentRate per hour.`n`nBased on the overall average rate, we will be done in $($(1/($avgRate*24)*($t-$i)) - $((Get-Date).TotalDays)) days on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i))))." -SmtpServer 'smtp.office365.com' -Port:25 -UseSsl:$true -BodyAsHtml:$false -Credential:$SmtpCreds
					$ErrorActionPreference = $old_ErrorActionPreference
				# Update Log
					Write-Log ("Counter: $i out of $t devices at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))")
				# Update Screen
					Write-host "Counter: $i out of $t devices at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))" 
				# Clear StatLimit and $ResultsLost for next run
					$c = 0
					$ResultsList.Clear()
			}
		# Parse the MobileDevice.Identity 
			#$userIndex = $Object1.Identity.parent.split("/")[3]
			$userIndex = $Object1.Identity.split("\")[0]
		# Update Progress
			Write-Progress -Id 1 -Activity "Getting all device partnerships from " -PercentComplete (5 + ($currentProgress/$progressActions * 90)) -Status "Enumerating a device for user $($userIndex)"
		#  If the CASmailbox user ONLY has a REST partnership with OutlookMobile for iOS/Android, the HasActiveSyncDevicePartnership will be false. In this case, we need to make a new call to EXO...
			If($Object1.ClientType) 
				{
					# Powershell v4 allows super efficient handy reference of the array by an object value using the .where() method
					# I haven't tested this method with over 1000 users, so test here if efficiency results falter
					$UserObject1 = $Object1Users.where({[string]$_.id -eq "$userIndex"})
					#$UserObject1 = $Object1Users.where({$_.id -eq '$userIndex'})
				}
			Else 
				{
					$caseCheckStartDate = Get-Date
						If($userindex){
							# This could potentially be an expensive call if $userindex is null, then get-casmailbox is calling EXO powershell for default limit of results for a blank identity
							$UserObject1 = Get-CASMailbox -Identity $userIndex | Select-Object -Property distinguishedname,displayname,id,primarysmtpaddress,activesyncsuppressreadreceipt,activesyncdebuglogging
						}
						else {
						# Write-Output "Could not find CASmailbox information for this device $Object1" | Out-File $debugoutput -Append
							Write-Log ("Could not find CASmailbox information for this device $Object1")
						}
					[timespan]$caseCheckEndTime = (Get-Date) - $caseCheckStartDate
					$caseCheckTotalTime += $caseCheckEndTime
				}
			
		# Update this pivotal index Hashtable which prevents the need to make more timely calls to EXO; Using shorthand notation for add-member
			$line = @{
				# First Include the User info
				User=$userIndex
				DisplayName=$UserObject1.DisplayName
				PrimarySmtpAddress=$UserObject1.PrimarySmtpAddress
				UserId=$UserObject1.Id
				ActivesyncSuppressReadReceipt=$UserObject1.activesyncsuppressreadreceipt
				ActivesyncDebugLogging=$UserObject1.activesyncdebuglogging
				DistinguishedName=$UserObject1.distinguishedname
				# Now include the Obect info
				FriendlyName=$Object1.friendlyname
				DeviceID=$Object1.deviceid
				DeviceOS=$Object1.DeviceOS
				DeviceModel=$Object1.DeviceModel
				UserAgent=$Object1.DeviceUserAgent
				FirstSyncTime=$Object1.FirstSyncTime
				LastPolicyUpdate=$Object1.WhenChangedUTC
				ClientProtocolVersion=$Object1.clientversion
				ClientType=$Object1.clienttype
				Managed=$Object1.ismanaged
				AccessState=$Object1.DeviceAccessState
				AccessReason=$Object1.DeviceAccessStateReason
				RemovalId=$Object1.Identity
					# NOTE: The RemovalID has some caveats before it can be re-used for the remove-mobiledevice commandlet. 
					# When exporting this value to a .csv file, there is a special character called, "section sign" § that gets 
					# converted to a '?' so we adjust for that with a regex in the "ABQ_remove.ps1" script example end of this script.
			}
		# Since arraylist.add returns highest index, we need a way to ignore that value with Out-Null
			$ResultsList.Add((New-Object PSobject -property $line)) | Out-Null
			$currentProgress++
	}
# Update Log
	Write-Log ("Time to re-run Get-CasMailbox for REST devices: " + $($caseCheckTotalTime))
# Update Screen
	Write-Host -NoNewLine "Time to re-run Get-CasMailbox for REST devices:   ";write-host -ForegroundColor Yellow "$($caseCheckTotalTime)"

# Disconnect from EXO and cleanup the PS session
	Get-PSSession | Remove-PSSession -Confirm:$false -ErrorAction silentlycontinue

# Create the Output File (report) using the attributes created in the Hashtable by exporting to CSV
	# Update Progress
		write-progress -id 1 -activity "Creating Output Report" -PercentComplete (96) -Status "$outputFolder"
	# Create Report
		$ResultsList | select DisplayName,User,UserId,PrimarySMTPAddress,FriendlyName,UserAgent,FirstSyncTime,LastPolicyUpdate,DeviceOS,ClientProtocolVersion,ClientType,DeviceModel,deviceid,AccessState,AccessReason,ActivesyncSuppressReadReceipt,ActivesyncDebugLogging,Managed,DistinguishedName,RemovalId | export-csv -path $OutputFile -notypeinformation -Append

# Separately capture any PowerShell errors and output to an errorfile
	$errfilename = $outputfolder + $logfilename + "_ERRORs_" + (Get-Date).Ticks + ".txt" 
	write-progress -id 1 -activity "Error logging" -PercentComplete (99) -Status "$errfilename"
	ForEach ($err in $error) {  
		$logdata = $null 
		$logdata = $err 
		If ($logdata) 
			{ 
				out-file -filepath $errfilename -Inputobject $logData -Append 
			} 
	}
#Clean Up and Show Completion in session and logs
	# Update Progress	
		write-progress -id 1 -activity "Complete" -PercentComplete (100) -Status "Success!"
	# Update Log
		$endDate = Get-Date
		$elapsedTime = $endDate - $startDate
		Write-Log ("Report started at: " + $($startDate));Write-Log ("Report ended at: " + $($endDate));Write-Log ("Total Elapsed Time: " + $($elapsedTime)); Write-Log ("Device Collection Completed!")
	# Update Screen
		write-host;Write-Host -NoNewLine "Report started at    ";write-host -ForegroundColor Yellow "$($startDate)"
		Write-Host -NoNewLine "Report ended at      ";write-host -ForegroundColor Yellow "$($endDate)"
		Write-Host -NoNewLine "Total Elapsed Time:   ";write-host -ForegroundColor Yellow "$($elapsedTime)"
		Write-host "-------------------------------------------------";write-host -foregroundcolor Cyan "Device collection Complete!";write-host;write-host -foregroundcolor Green "...The EXOMobileDeviceInventory CSV and log were created in $outputFolder";write-host;write-host;sleep 1

#------------------------v
#End CUSTOM CODE Section
#------------------------v
