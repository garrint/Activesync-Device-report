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
# Exchange Online Device partnership inventory, dependent on EXOv2 module being installed (https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
#  EXO_MobileDevice_Inventory_3.1.ps1
#  
#  Created by: Austin McCollum 2/11/2018 austinmc@microsoft.com
#  Updated by: Garrin Thompson 7/23/2020 garrint@microsoft.com *** "Borrowed" a few quality-of-life functions from Start-RobustCloudCommand.ps1 and added EXOv2 connection
#
#########################################################################################
# This script enumerates all devices in Office 365 and reports on many properties of the
#   device/application and the mailbox owner.
#
# $deviceList is an array of hashtables, because deviceIDs may not be
#   unique in an environment. For instance when a device is configured with
#   two separate mailboxes in the same org, the same deviceID will appear twice.
#   Hashtables require uniqueness of the key so that's why the array of hashtable data 
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
	if (!($NonInteractive)){
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

# Setup a new O365 Powershell Session using RobustCloudCommand concepts
Function New-CleanO365Session {
	 #Prompt for UPN used to login to EXO (or provide directly in line 138)
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
	
	#	OLD BasicAuth method create session
	#	$Exchangesession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection

	# Check for an error while creating the session
	if ($Error.Count -gt 0){
		Write-log ("[ERROR] - Error while setting up session")
		Write-log ($Error)
		# Increment our error count so we abort after so many attempts to set up the session
		$ErrorCount++
		# if we have failed to setup the session > 3 times then we need to abort because we are in a failure state
		if ($ErrorCount -gt 3){
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
		#OLD Basic Auth Method import session
		#$null = Import-PSSession $Exchangesession -AllowClobber
	$null = Connect-ExchangeOnline -UserPrincipalName $EXOLogonUPN -ShowProgress:$false -ShowBanner:$false
	# Set the Start time for the current session
	Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy; Goes ahead and resets it every "$ResetSeconds" number of seconds (14.5 mins) either way 
Function Test-O365Session {
	# Get the time that we are working on this object to use later in testing
	$ObjectTime = Get-Date
	# Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession
	# Make sure we found a session
	if ($SessionInfo -eq $null) { 
		Write-log ("[ERROR] - No Session Found")
		Write-log ("Recreating Session")
		New-CleanO365Session
	}	
	# Make sure it is in an opened state if not log and recreate
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
		if ($DelayinSeconds -gt 0){
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
	if ($ManualThrottle -gt 0){
		Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
		Start-Sleep -Milliseconds $ManualThrottle
	}
}

#---------------------v
#MAINCODESTARTSBELOW
#---------------------v
$execpol = get-executionpolicy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
Write-Host; $EXOLogonUPN=Read-host "Type in UPN for account that will execute this script"; write-host "...pleasewait...connecting to EXO..."
$SmtpCreds = (get-credential -Message "Provide EXO account Pasword" -UserName "$EXOLogonUPN")
# Set $OutputFolder to Current PowerShell Directory
[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))
$outputFolder = [IO.Directory]::GetCurrentDirectory()
$logFile = $outputFolder + "\EXOMobileDevice_logfile_" + (Get-Date).Ticks + ".txt"
[int]$ManualThrottle=0
[double]$ActiveThrottle=.25
[int]$ResetSeconds=870

# Setup our first session to O365
$ErrorCount = 0
New-CleanO365Session
Write-Log ("Connected to Exchange Online")
write-host;write-host -ForegroundColor Green "...Connected to Exchange Online as $EXOLogonUPN"
# Get when we started the script for estimating time to completion
$ScriptStartTime = Get-Date
$startDate = Get-Date
write-progress -id 1 -activity "Beginning..." -PercentComplete (1) -Status "initializing variables"
# Clear the error log so that sending errors to file relate only to this run of the script
$error.clear()

# Get mailbox attributes for users with device partnerships from your tenant
Write-Progress -Id 1 -Activity "Getting all EXO users with Devices" -PercentComplete (10) -Status "Get-CasMailbox -ResultSize Unlimited"
#$mobileDeviceUsers = Import-Csv .\mobileDeviceUsers.csv #should only be used in place of the command below if CSV being imported has same attributes outputed as command below 
$mobileDeviceUsers = Invoke-Command -Session (Get-PSSession) -ScriptBlock {
	Get-CASMailbox -RecalculateHasActiveSyncDevicePartnership -ResultSize unlimited -Filter {HasActiveSyncDevicePartnership -eq "True"} | Select-Object -Property distinguishedname,displayname,id,primarysmtpaddress,activesyncmailboxpolicy,activesyncsuppressreadreceipt,activesyncdebuglogging,activesyncallowedids,activesyncblockeddeviceids
	}
# Get all mobiledevices from users with Devices
write-progress -id 1 -Activity "Getting all EXO Devices" -PercentComplete (5) -Status "Get-MobileDevice is running"

[System.Collections.ArrayList]$mobiledeviceslist = New-Object System.Collections.ArrayList($null)
$mobiledeviceslist | Select friendlyname,deviceid,DeviceOS,DeviceModel,DeviceUseragent,devicetype,FirstSyncTime,WhenChangedUTC,identity,clientversion,clienttype,ismanaged,DeviceAccessState,DeviceAccessStateReason
$mobiledevices = $null
$mobiledeviceslist.Clear()
foreach ($dvcuser in $mobileDeviceUsers) {
        $mobiledevices = Get-MobileDevice -Mailbox $dvcuser.id.name | Select-Object -Property friendlyname,deviceid,DeviceOS,DeviceModel,DeviceUseragent,devicetype,FirstSyncTime,WhenChangedUTC,identity,clientversion,clienttype,ismanaged,DeviceAccessState,DeviceAccessStateReason
        $mobiledeviceslist += $mobiledevices
	}
# Measure the time the Invoke Command call takes to enumerate devices from Exchange Online
$progressActions = $mobiledeviceslist.count
$invokeEndDate = Get-Date
$invokeElapsedTime = $invokeEndDate - $startDate
Write-Log ("Starting device collection");Write-Log ("Number of Devices found in Exchange Online: " + ($progressActions));Write-Log ("Time to run Invoke command for Device retrieval: " + ($($invokeElapsedTime)))
write-host -foregroundcolor Cyan "Starting device collection";;sleep 2;write-host "-------------------------------------------------"
Write-Host -NoNewline "Total Devices found for users with a device:      ";Write-Host -ForegroundColor Green $progressActions
Write-Host -NoNewline "Time to run Invoke command for Device retrieval:  ";write-host -ForegroundColor Yellow "$($invokeElapsedTime)"

# Measure the time the get-casmailbox cmd takes to grab info for users with devices
$casMailboxUnlimitedEndDate = Get-Date
$casMailboxUnlimitedElapsedTime = $casMailboxUnlimitedEndDate - $invokeEndDate
Write-Log ("Number of Users with Devices in Exchange Online: " + $($mobiledeviceslist.count));Write-Log ("Time for User retrieval via Get-CASMailbox run: " + $($casMailboxUnlimitedElapsedTime))
Write-Host -NoNewline "Number of Users with Devices in Exchange Online:  ";Write-Host -ForegroundColor Green "$($mobiledeviceusers.count)"
Write-Host -NoNewline "Time to retrieve User info via Get-CasMailbox:    ";write-host -ForegroundColor Yellow "$($casMailboxUnlimitedElapsedTime)"

#  Now from the two arrays of hashtables, let's create a new array of hashtables containing calculated properties indexed by a property from the device list
#  This is a BIG LOOP!!
[System.Collections.ArrayList]$deviceList = New-Object System.Collections.ArrayList($null)
$currentProgress = 1
[TimeSpan]$caseCheckTotalTime=0
# Set a really simple counter and some variables to use for periodic write/flush and reporting 
# report counter
$c = 0
# running counter
$i = 0
# Set the number of objects to cycle before writing to disk and sending stats, i'd consider 5000 max
$statLimit = 500
# Get the total number of devices, which we use in some stat calculations
$t = $mobiledeviceslist.count
# Set some timedate variables for the stats report
$loopStartTime = Get-Date
$loopCurrentTime = Get-Date
# Moved this from the bottom of the script to generate the output file ahead of time
$devicesOutput= $outputfolder + "\EXO Mobile Device Inventory_" + (Get-Date).Ticks + ".csv"

foreach ($mobileDevice in $mobiledeviceslist) {
    Test-O365Session
    # Total running count 
    $i++
    # Dump the $deviceList to CSV every $statLimit devices; also send status e-mail with some metrics at each dump.
    if (++$c -eq $statLimit) {
        # Moved this from the bottom of the script, and added -Append parameter
        $deviceList | select DisplayName,User,UserId,PrimarySMTPAddress,FriendlyName,UserAgent,FirstSyncTime,LastPolicyUpdate,DeviceOS,ClientProtocolVersion,ClientType,DeviceModel,DeviceId,AccessState,AccessReason,ActivesyncSuppressReadReceipt,ActivesyncDebugLogging,Managed,DistinguishedName,RemovalId | export-csv -path $devicesoutput -notypeinformation -Append
        $loopLastTime = $loopCurrentTime
        $loopCurrentTime = Get-Date
        $currentRate = $statLimit/($loopCurrentTime-$loopLastTime).TotalHours
        $avgRate = $i/($loopCurrentTime-$loopStartTime).TotalHours
	#SEND STATUS EMAIL each time we write 500 devices to file (requires $SmtpCreds on line 191)
		$old_ErrorActionPreference = $ErrorActionPreference
		$ErrorActionPreference = 'SilentlyContinue'
		Send-MailMessage -From "$EXOLogonUPN" -To "$EXOLogonUPN" -Subject "EXOMobileDeviceInventory : Progress" -Body "EXOMobileDeviceInventory PROGRESS report`n`nCurrentTime: $loopCurrentTime`nStartTime: $loopStartTime`n`nCounter: $i out of $t devices, at a current rate of $currentRate per hour.`n`nBased on the overall average rate, we will be done in $($(1/($avgRate*24)*($t-$i)) - $((Get-Date).TotalDays)) days on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i))))." -SmtpServer 'smtp.office365.com' -Port:25 -UseSsl:$true -BodyAsHtml:$false -Credential:$SmtpCreds
		$ErrorActionPreference = $old_ErrorActionPreference
	Write-Log ("Counter: $i out of $t devices at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))")
	Write-host "Counter: $i out of $t devices at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))" 
	$c = 0
        $deviceList.Clear()
		}
	# This is a pivotal index call that makes this whole thing work much faster by correlating the list of unique mobile devices
	# to a CASmailbox instead of having to make expensive 'Get-MobileDeviceStatistics' calls from EXO.
	#
	# The MobileDevice.Identity has a consistent pattern in the EXO directory, containing the mobile user's casmailbox id 
	#$userIndex = $mobileDevice.Identity.parent.split("/")[3]
	$userIndex = $mobileDevice.Identity.split("\")[0]

	Write-Progress -Id 1 -Activity "Getting all device partnerships from " -PercentComplete (5 + ($currentProgress/$progressActions * 90)) -Status "Enumerating a device for user $($userIndex)"
	#  UPDATE: In some cases, if a CASmailbox user ONLY has a REST partnership with Outlook for iOS / Android, the HasActiveSyncDevicePartnership will be false.
	#  In this case, we need to make a new call to EXO...
	if($mobiledevice.ClientType)
	    {
	    # Powershell v4 allows super efficient handy reference of the array by an object value using the .where() method
	    # I haven't tested this method with over 1000 users, so test here if efficiency results falter
	    $mobileUser = $mobileDeviceUsers.where({[string]$_.id -eq "$userIndex"})
	    #$mobileUser = $mobiledeviceusers.where({$_.id -eq '$userIndex'})
	    }
	Else 
	    {
		$caseCheckStartDate = Get-Date
		    if($userindex){
			    # This could potentially be an expensive call if $userindex is null, then get-casmailbox is calling EXO powershell for default limit of results for a blank identity
			    $mobileUser = Get-CASMailbox -Identity $userIndex | Select-Object -Property distinguishedname,displayname,id,primarysmtpaddress,activesyncsuppressreadreceipt,activesyncdebuglogging
		    }
		    else {
			   # Write-Output "Could not find CASmailbox information for this device $mobileDevice" | Out-File $debugoutput -Append
			    Write-Log ("Could not find CASmailbox information for this device $mobileDevice")
		    }
	    [timespan]$caseCheckEndTime = (Get-Date) - $caseCheckStartDate
	    $caseCheckTotalTime += $caseCheckEndTime
	    }
	# Setting the hashtable created from both get-mobiledevice and get-casmailbox, using shorthand notation for add-member
	$line = @{
		User=$userIndex
		DisplayName=$mobileUser.DisplayName
		PrimarySmtpAddress=$mobileUser.PrimarySmtpAddress
		UserId=$mobileUser.Id
		ActivesyncSuppressReadReceipt=$mobileUser.activesyncsuppressreadreceipt
		ActivesyncDebugLogging=$mobileUser.activesyncdebuglogging
		DistinguishedName=$mobileUser.distinguishedname
		# Now including the MobileDevice information
		FriendlyName=$mobileDevice.friendlyname
		DeviceID=$mobileDevice.deviceid
		DeviceOS=$mobileDevice.DeviceOS
		DeviceModel=$mobileDevice.DeviceModel
		UserAgent=$mobileDevice.DeviceUserAgent
		FirstSyncTime=$mobileDevice.FirstSyncTime
		LastPolicyUpdate=$mobileDevice.WhenChangedUTC
		ClientProtocolVersion=$mobileDevice.clientversion
		ClientType=$mobileDevice.clienttype
		Managed=$mobileDevice.ismanaged
		AccessState=$mobileDevice.DeviceAccessState
		AccessReason=$mobileDevice.DeviceAccessStateReason
		# NOTE: The RemovalID has some caveats before it can be re-used for the remove-mobiledevice commandlet. 
			# When exporting this value to a .csv file, there is a special character called, "section sign" § that gets 
			# converted to a '?' so we adjust for that with a regex in the "ABQ_remove.ps1" script example end of this script.
		RemovalId=$mobiledevice.Identity
		}
	# out-null since arraylist.add returns highest index so this is a way to ignore that value
	$deviceList.Add((New-Object PSobject -property $line)) | Out-Null
	$currentProgress++
}
Write-Log ("Time to re-run Get-CasMailbox for REST devices: " + $($caseCheckTotalTime))
Write-Host -NoNewLine "Time to re-run Get-CasMailbox for REST devices:   ";write-host -ForegroundColor Yellow "$($caseCheckTotalTime)"

# We've got all the data we need. Hopefully the calls to Exchange Online above are quick. Let's cleanup the session now.
Get-PSSession | Remove-PSSession -Confirm:$false -ErrorAction silentlycontinue

# Now to put all that info into a spreadsheet. 
write-progress -id 1 -activity "Creating spreadsheet" -PercentComplete (96) -Status "$outputFolder"
$deviceList | select DisplayName,User,UserId,PrimarySMTPAddress,FriendlyName,UserAgent,FirstSyncTime,LastPolicyUpdate,DeviceOS,ClientProtocolVersion,ClientType,DeviceModel,deviceid,AccessState,AccessReason,ActivesyncSuppressReadReceipt,ActivesyncDebugLogging,Managed,DistinguishedName,RemovalId | export-csv -path $devicesoutput -notypeinformation -Append

# Capture any PS errors and output to a file
	$errfilename = $outputfolder + "\EXO Mobile Device Errorlog_" + (Get-Date).Ticks + ".txt" 
	write-progress -id 1 -activity "Error logging" -PercentComplete (99) -Status "$errfilename"
	foreach ($err in $error) 
	{  
	    $logdata = $null 
	    $logdata = $err 
	    if ($logdata) 
	    { 
		out-file -filepath $errfilename -Inputobject $logData -Append 
	    } 
	}
write-progress -id 1 -activity "Complete" -PercentComplete (100) -Status "Success!"
$endDate = Get-Date
$elapsedTime = $endDate - $startDate
Write-Log ("Report started at: " + $($startDate));Write-Log ("Report ended at: " + $($endDate));Write-Log ("Total Elapsed Time: " + $($elapsedTime)); Write-Log ("Device Collection Completed!")
write-host;Write-Host -NoNewLine "Report started at    ";write-host -ForegroundColor Yellow "$($startDate)"
Write-Host -NoNewLine "Report ended at      ";write-host -ForegroundColor Yellow "$($endDate)"
Write-Host -NoNewLine "Total Elapsed Time:   ";write-host -ForegroundColor Yellow "$($elapsedTime)"
Write-host "-------------------------------------------------";write-host -foregroundcolor Cyan "Device collection Complete!";write-host -foregroundcolor Green "...The EXO_MobileDevice_Inventory.CSV and log were created in $outputFolder";write-host;write-host;sleep 1

#ENDSCRIPT

#################################################################################################################
#
# Running Get-MobileDeviceStatistics in Exchange Online is very expensive. If you have less than 100 devices, you could 
#  add the following lines and get a bit more information. If you need this information, run a separate script with 
#  the output of this script to append what you need.
#
# [string]$temp="$($mobiledevices[0].Identity.Parent)/$($mobiledevices[0].Identity.Name)"
# $mbxDeviceStatistics = Get-MobileDeviceStatistics -id $temp 
#
# Then add the following section to the $line loop
#
#   LastPolicyUpdateTime=$mbxDeviceStatistics.LastPolicyUpdateTime
#	LastSyncAttemptTime=$mbxdevicestats.LastSyncAttemptTime
#   LastPingHeartBeat=$mbxdevicestats.LastPingHeartBeat
#   Status=$mbxdevicestats.Status
#   StatusNote=$mbxdevicestats.StatusNote
#   DeviceAccessControlRule=$mbxdevicestats.DeviceAccessControlRule
#   DevicePolicyApplied=$mbxdevicestats.DevicePolicyApplied
#   DevicePolicyApplicationStatus=$mbxdevicestats.DevicePolicyApplicationStatus
#   FoldersSynced=$mbxdevicestats.NumberOfFoldersSynced
#
# Then you would need to add those values to the reports at the end.
#
#################################################################################################################

#################################################################################################################
#
# ABQ_allow.ps1
#
# A one line command takes all the entries from a CSV under the heading "username"
#  to add to the CASmailbox activesync allowlist all the entries under "deviceID". 
# This is useful in grandfathering an appropriate list of Activesync Device partnerships
#  before implementing an ABQ framework of quarantine or block.
#
# The @{Add notation used here ensures multiple deviceIDs can be added to the multi value property
#  http://blogs.technet.com/b/dstrome/archive/2011/05/29/multivalued-properties-in-exchange-2010.aspx
#
#command to set allowlist
#import-csv c:\allowedDevices.csv | foreach{set-EXOcasmailbox -identity $_.dn -ActiveSyncAllowedDeviceIDs @{Add=$_.deviceid}}
#
#################################################################################################################

#################################################################################################################
#
# ABQ_remove.ps1
#
# A two line command takes all the entries from a CSV to remove the partnership and remove allowlist entries
# 
# This is useful for forcing a list of devices to be re-evaluated on a new or updated ABQ framework.
#   Devices already connected should not see any impact if they are allowed in the ABQ rules defined
#   and will re-establish their partnership. Devices no longer allowed will get the expected behavior
#   of a blocked or quarantined device depending on the ABQ configuration defined.
#
#commands to remove partnerships and remove the deviceID from the allowlist if there
# $removaldevices=import-csv c:\removalDevices.csv 
# foreach($device in $removaldevices){remove-mobiledevice –identity ($device.removalid -replace "\?","§") -confirm:$false;if((get-casmailbox -identity $device.dn).allowedDeviceIDs){set-casmailbox -identity $device.dn -ActiveSyncAllowedDeviceIDs @{Remove=$device.deviceid_AD}}
#
#################################################################################################################

#################################################################################################################
#
# Device_remove.ps1
#
# A two line command takes all the entries from a CSV to remove the partnership and remove allowlist entries
# 
# This is useful for forcing a list of devices to be re-evaluated on a new or updated ABQ framework.
#   Devices already connected should not see any impact if they are allowed in the ABQ rules defined
#   and will re-establish their partnership. Devices no longer allowed will get the expected behavior
#   of a blocked or quarantined device depending on the ABQ configuration defined.
#
# $csv = Import-Csv "<Path to EXO Mobile Device Inventory CSV file>"
#	Foreach ($dvc in $csv) {
#		If ($null -eq ($dvc.LastPolicyUpdate)) {
#			Write-Log ("Skipping...No recorded LastPolicyUpdate time for" + $dvc.DeviceID)
#		}
#		ElseIf ((date).adddays(-90) -gt $dvc.LastPolicyUpdate) {
#			Remove-MobileDevice -Identity ($dvc.RemovalID -replace "\?","§") -confirm:$false -verbose -WhatIf
#			If ((get-casmailbox -identity $dvc.dn).allowedDeviceIDs) {set-casmailbox -identity $dvc.dn -ActiveSyncAllowedDeviceIDs @{Remove=$dvc.deviceid_AD} -confirm:$false -verbose -WhatIf}
#			Write-Log ("Removing MobileDevice:" + $dvc.DeviceID)
#		}
#	}
#
#################################################################################################################
#End
