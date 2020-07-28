# Activesync-Device-report
Creating activesync device reports for Office 365 using powershell

Exchange Online Device partnership inventory
EXO_MobileDevice_Inventory_<ver>.ps1

  Created by: Austin McCollum 2/11/2018 austinmc@microsoft.com
  Updated by: Garrin Thompson 7/23/2020 garrint@microsoft.com (ver 3.1) 
  *** "Borrowed" a few quality-of-life functions from Start-RobustCloudCommand.ps1

NOTE: The 3.1 script is dependent on having the EXOv2 module installed (https://www.powershellgallery.com/packages/ExchangeOnlineManagement).  If you dont have the EXOv2 module installed yet, the 2.5.1(EXOv1-module) script will use the v1 ADAL module instead.

 This script enumerates the devices of Exchange Online mailboxes that have ActiveSyncDevicePartnerships and reports on many properties of the device/application and the mailbox owner.

 $deviceList is an array of hashtables, because deviceIDs may not be
   unique in an environment. For instance when a device is configured with
   two separate mailboxes in the same org, the same deviceID will appear twice.
   Hashtables require uniqueness of the key so that's why the array of hashtable data 
   structure was chosen.

 The devices can be sorted by a variety of properties like "LastActivity" ("LastPolicyUpdate" in ver 3.1) to determine stale partnerships or outdated devices needing to be removed.  We do not pull LastSuccessSync from Get-MobileDeviceStatistics because that call is too expensive (takes way too long).
 
 The DisplayName of the user's CAS mailbox is recorded for importing with the 
   Set-CasMailbox commandlet to configure allowedDeviceIDs. This is especially useful in 
   scenarios where a migration to ABQ framework requires "grandfathering" in all or some
   of the existing partnerships.

 Get-CasMailbox is run efficiently with the -HasActiveSyncDevicePartnership filter 

Creates an actionable CSV report in the PowerShell current directory which is enumerated when the script is executed.
