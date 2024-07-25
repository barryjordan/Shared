# Purge-EmailMessages.PS1
# A script to purge messages from Exchange Online using a compliance search and a purge action applied to 
# the results of that search. The search can be targeted to the members of a distribution group.
# ---------------------------------------

# A check to verify a connection to Exchange Online
$Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
If (!($Status)) {
  Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

# Connect to the compliance endpoint
Connect-IPPSSession

# Some information to identify the messages we want to purge
$Sender = "<enter sender email address>"
# $Subject = "Special Offer for you"
# Note - if the subject contains a special character like a $ sign, make sure that you escape that character otherwise it won't
# be used in the search
$ComplianceSearch = "<Enter a Search Name>"
# Date range for the search - make this as precise as possible
$StartDate = "1-Nov-2010"
$EndDate = "1-Nov-2020"
$Start = (Get-Date $StartDate).ToString('yyyy-MM-dd')   
$End = (Get-Date $EndDate).ToString('yyyy-MM-dd')
# $ContentQuery = '(Received:' + $Start + '..' + $End +') AND (From:' + $BadSender + ')'
$ContentQuery = '(c:c)(from='+ Sender +')(received='+ $Start +'..'+ $End +')'

# Search is being targeted to the members of a distribution group
$GroupName = "<enter distribution group address>"

Clear-Host
If (Get-ComplianceSearch -Identity $ComplianceSearch -ErrorAction SilentlyContinue ) {
   Write-Host "Cleaning up old search"
    Try {
      $Status = Remove-ComplianceSearch -Identity $ComplianceSearch -Confirm:$False  
    } 
    Catch {
       Write-Host "Unable to clean up the old search" ; break 
    }
}

Write-Host "Starting Compliance Search..."
New-ComplianceSearch -Name $ComplianceSearch -ContentMatchQuery $ContentQuery -ExchangeLocation $GroupName -AllowNotFoundExchangeLocationsEnabled $True | Out-Null
Start-ComplianceSearch -Identity $ComplianceSearch | Out-Null
[int]$Seconds = 10
Start-Sleep -Seconds $Seconds
# Loop until the search finishes
While ((Get-ComplianceSearch -Identity $ComplianceSearch).Status -ne "Completed") {
    Write-Host ("Still searching after {0} seconds..." -f $Seconds)
    $Seconds = $Seconds + 10
    Start-Sleep -Seconds $Seconds
}

[int]$ItemsFound = (Get-ComplianceSearch -Identity $ComplianceSearch).Items

If ($ItemsFound -gt 0) {
   $Stats = Get-ComplianceSearch -Identity $ComplianceSearch | Select-Object -Expand SearchStatistics | Convertfrom-JSON
   $Data = $Stats.ExchangeBinding.Sources | Where-Object {$_.ContentItems -gt 0}
   # Get the count of messages in the mailbox with the maximum search results.
   [int]$MaxCount = $Data.ContentItems.Get(0)
   Write-Host ""
   Write-Host "Total Items found matching query:" $ItemsFound 
   Write-Host ""
   Write-Host "Items found in the following mailboxes"
   Write-Host "--------------------------------------"
   Foreach ($D in $Data)  {
        Write-Host ("{0} has {1} items of size {2}" -f $D.Name, $D.ContentItems, $D.ContentSize)
   }
   Write-Host " "
   [int]$Iterations = 0; [int]$ItemsProcessed = 0
   While ($ItemsProcessed -lt $MaxCount) {
       $Iterations++
       Write-Host ("Deleting items...({0})" -f $Iterations)
       New-ComplianceSearchAction -SearchName $ComplianceSearch -Purge -PurgeType HardDelete -Confirm:$False | Out-Null
       $SearchActionName = $ComplianceSearch + "_Purge"
       While ((Get-ComplianceSearchAction -Identity $SearchActionName).Status -ne "Completed") { # Let the search action complete
           Start-Sleep -Seconds 5 }
       $ItemsProcessed = 10 + $ItemsProcessed # Can remove a maximum of 10 items per mailbox
	   # Remove the search action so we can recreate it
       Remove-ComplianceSearchAction -Identity $SearchActionName -Confirm:$False -ErrorAction SilentlyContinue }
} Else {
    Write-Host "The search didn't find any items..." 
}

Write-Host "Completed!"