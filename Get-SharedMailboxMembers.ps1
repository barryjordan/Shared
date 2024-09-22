# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Function to check if connected to Exchange Online
function Is-ConnectedToExchangeOnline {
    try {
        # Attempt to get a mailbox as a connection check
        Get-Mailbox -ResultSize 1 | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Check if the connection is active
if (-not (Is-ConnectedToExchangeOnline)) {
    Write-Host "You are not connected to Exchange Online. Please connect." -ForegroundColor Yellow -BackgroundColor Black

    # Attempt to connect to Exchange Online
    try {
        Connect-ExchangeOnline -ShowProgress $true
        Write-Host "Connected to Exchange Online successfully."
    } catch {
        Write-Host "Failed to connect to Exchange Online. Exiting script." -ForegroundColor Red
        exit
    }
}

# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a new SaveFileDialog instance
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$saveFileDialog.Title = "Save Output CSV File"
$saveFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('MyDocuments')

# Show the dialog and check if the user clicked OK
if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $outputCsvPath = $saveFileDialog.FileName

    # Get all shared mailboxes and their permissions
    $mailboxes = Get-Mailbox -Filter {RecipientTypeDetails -eq "SharedMailbox"} -ResultSize Unlimited

    # Create an array to hold results
    $results = @()

    # Loop through each shared mailbox to get permissions
    foreach ($mailbox in $mailboxes) {
        $permissions = Get-MailboxPermission -Identity $mailbox.DistinguishedName | Where-Object { $_.AccessRights -eq "FullAccess" }

        foreach ($permission in $permissions) {
            			
			$results += [PSCustomObject]@{
                SharedMailbox = $mailbox.PrimarySmtpAddress
                User          = $permission.User
				AccessRights  = $permission.AccessRights
            }
        }
    }

    # Export the results to a new CSV file
    $results | Export-Csv -Path $outputCsvPath -NoTypeInformation

    Write-Host "Export completed successfully. The data has been saved to: $outputCsvPath" -ForegroundColor White -BackgroundColor Blue
} else {
    Write-Host "User canceled the operation."
}

# Disconnect from Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false