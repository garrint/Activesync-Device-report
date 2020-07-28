# EXO Mobile Device Inventory
Create a device report for Exchange Online users via powershell.

Exchange Online Device partnership inventory
EXO_MobileDevice_Inventory_3.1.ps1

  Created by: Austin McCollum 2/11/2018 austinmc@microsoft.com
  Updated by: Garrin Thompson 7/23/2020 garrint@microsoft.com (ver 3.1) 
  *** "Borrowed" a few quality-of-life functions from Start-RobustCloudCommand.ps1

NOTE: The 3.1 script is dependent on having the EXOv2 module installed (https://www.powershellgallery.com/packages/ExchangeOnlineManagement).  If you dont have the EXOv2 module installed yet, the 2.5.1(EXOv1-module) script will use the v1 ADAL module instead, but version 2.5 still requests with resultsize unlmited (which may not complete in large env).

 This script enumerates the devices of Exchange Online mailboxes that have ActiveSyncDevicePartnerships and reports on many properties of the device/application and the mailbox owner.  Here’s what my script does:

  1.	First, it asks you for the UPN you’ll be using to run the script (this is so the connection to the EXOv2 module will be able to reconnect when session is lost.  Then it asks you to provide the password for the UPN provided to store the credential for use in sending an email to the UPN provided with progress (the email may not work, but I added the same progress info to the PowerShell output so you can see how it’s going every 500 devices)
  2.	Once connected to EXO, it invokes a query for Get-CASMailbox against EXO with a filter for the attribute HasActiveSyncDevicePartnership and only returns users that do have one.  
  3.	Once all users that have a device partnership are stored in a variable, it runs Get-MobileDevice against only the users that have devices and stores the results in a hashtable that is written to the CSV report every 500 obtained.  At this same time it writes to the CSV (every 500), it adds a progress line to the PowerShell output indicating how many have been collected so far out of the total found and an estimated completion date/time based on throughput calculations and tries to send an email to the UPN provided with the same info so you don’t have to keep an eye on the PowerShell screen.

 $deviceList is an array used in a hashtable, because deviceIDs may not be unique in an environment. For instance when a device is configured with two separate mailboxes in the same org, the same deviceID will appear twice.  Hashtables require uniqueness of the key so that's why the array of hashtable data structure was chosen.

 The devices can be sorted by a variety of properties like "LastPolicyUpdate" to determine stale partnerships or outdated devices needing to be removed.  I do not pull LastSuccessSync from Get-MobileDeviceStatistics because that call is too expensive (takes way too long).
 
 The DisplayName of the user's CAS mailbox is recorded for importing with the Set-CasMailbox commandlet to configure allowedDeviceIDs. This is especially useful in scenarios where a migration to ABQ framework requires "grandfathering" in all or some of the existing partnerships.

Creates an actionable CSV report in the PowerShell current directory which is enumerated when the script is executed.
