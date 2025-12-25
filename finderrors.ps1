###############################################################################
# Find COM errors
# Author: Kit Menke
# Version 1.0 11/6/2016
###############################################################################

# Notes:
# Get-EventLog doesn't quite work I guess:
# https://stackoverflow.com/questions/31396903/get-eventlog-valid-message-missing-for-some-event-log-sources#
# Get-EventLog Application -EntryType Error -Source "DistributedCOM"
# The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID
#$logs = Get-EventLog -LogName "System" -EntryType Error -Source "DCOM" -Newest 1 -Message "The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID*"

$EVT_MSG = "Local Activation"
# Search for System event log ERROR entries starting with the specified EVT_MSG
# Level 2 is error, 3 is warning

$logSeen = [System.Collections.Generic.HashSet[string]]::new()
$checkArray = @()
$fixArray = @()


$logEntries = Get-WinEvent -FilterHashTable @{LogName='System'; Id=10016} | Where-Object { $_.Message -like "*$EVT_MSG*" } 
if (!$logEntries) {
  Write-Host "No event log entries found."
  exit 1
}

foreach ($logEntry in $logEntries) {

 
  #Write-Host ($logEntry.Properties | Format-List | Out-String)

  # Get CLSID and APPID from the event log entry
  # which we'll use to look up keys in the registry
  $CLSID = $logEntry.Properties[3].Value
  $APPID = $logEntry.Properties[4].Value
  $USERDOMAIN = $logEntry.Properties[5].Value
  $USERNAME = $logEntry.Properties[6].Value
  $USERSID = $logEntry.Properties[7].Value

  $SIG = "$CLSID-$APPID-$USERDOMAIN-$USERNAME-$USERSID"

  $notSeen = $logSeen.Add($SIG)

 #Write-Host "SIGNATURE=$SIG => Not seen=$notSeen"

  if($notSeen){
    Write-Host "Found an event log entry not seen before : SIGNATURE=$SIG => Not seen=$notSeen"
    Write-Host ($logEntry | Format-List | Out-String)



    Write-Host "CLSID=$CLSID"
    Write-Host "APPID=$APPID"
    Write-Host "USERDOMAIN=$USERDOMAIN"
    Write-Host "USERNAME=$USERNAME"
    Write-Host "USERSID=$USERSID"

    $cmdCheck = ".\checkerrors.ps1 ""$APPID"" ""$USERSID"""
    $checkArray += $cmdCheck
    Write-Host $cmdCheck
    $cmdFix = ".\fixerrors.ps1 ""$APPID"" ""$CLSID"" ""$USERDOMAIN"" ""$USERNAME"""
    $fixArray += $cmdFix
    Write-Host $cmdFix 
  }
}

Write-Host "----------------"
Write-Host "Summary"
Write-Host "----------------"
$logSeen | ForEach-Object { Write-Host $_ }
Write-Host "----------------"
$checkArray | ForEach-Object { Write-Host $_ }
Write-Host "----------------"
$fixArray | ForEach-Object { Write-Host $_ }