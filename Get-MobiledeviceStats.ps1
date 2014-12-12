#########################################################################################
# COMPANY: CDW								                                            #
# NAME: Get-MobiledeviceStats.ps1                                                       #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  12/12/2014                                                                     #
# EMAIL: Dean.Sesko@S3.CDW.com                                                          #
#                                                                                       #
# COMMENT:  Script to get Primary SMTP and All devices for each user in the Org         #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 12/12/2014 Initial Version.                                                       #
#                                                                                       #
#########################################################################################



$MBXTable = New-Object System.Data.DataTable
$MBXTable.Columns.Add("PrimarySMTPAddress") | Out-Null
$MBXTable.Columns.Add("DeviceFriendlyName") | Out-Null
$MBXTable.Columns.Add("DeviceOS") | Out-Null
$MBXTable.Columns.Add("LastSyncAttemptTime") | Out-Null
$MBXTable.Columns.Add("LastSuccessSync") | Out-Null

$UserList = Get-CASMailbox -Filter { hasactivesyncdevicepartnership -eq $true -and -not displayname -like "CAS_{*" } | Get-Mailbox
foreach ($user in $UserList){
	$devices = Get-MobileDeviceStatistics -Mailbox $user
	foreach ($dev in $devices){
	$row = $MBXTable.NewRow()
	$row["PrimarySMTPAddress"] = $user.PrimarySMTPAddress
	$row["DeviceFriendlyName"] = $dev.DeviceFriendlyName
	$row["DeviceOS"] = $dev.DeviceOS
	$row["LastSyncAttemptTime"] = $dev.LastSyncAttemptTime
	$row["LastSuccessSync"] = $dev.LastSuccessSync
	
	$MBXTable.Rows.Add($row)
	
    }
}



$MBXTable | Export-Csv Devices.csv -NoTypeInformation