$synctime=(Get-MsolCompanyInformation).LastDirSyncTime
$timezoneoffset=([TimeZoneInfo]::Local).BaseUtcOffset.TotalHours
$localsynctime=$synctime.addhours($timezoneoffset)
Write-host "Last Directory Sync Time - $($localsynctime)" -ForegroundColor black -BackgroundColor cyan
$nextsynctime=$localsynctime.addMinutes(30)
Write-host "Next Directory Sync Time - $($nextsynctime)" -ForegroundColor black -BackgroundColor green
$remainingtime=New-TimeSpan -Start (Get-Date) -End $nextsynctime
Write-host "Time remaining until sync - $($remainingtime.Minutes) Minutes" -ForegroundColor red -BackgroundColor darkblue