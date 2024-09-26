<#
This script creates a new M365 FIC user account licensed with Office 365 A1 for Faculty.
The script should be used in tandem with a CSV file as input with the following header;
FirstName,LastName,UserPrincipalName,MailNickname,PrimaryEmail
==============================================================
The new user is created and licensed, and the assigned randomly generated password is printed to screen
#>


# Install the Microsoft Graph PowerShell module if not already installed
# Install-Module Microsoft.Graph -Force -AllowClobber

# Connect to Microsoft Graph with necessary permissions
Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All"

# Group ID to which users will be added (replace with your actual group ID)
$groupId = "< enter group id here >"

$datetag = (Get-Date -Format yyyyMMdd-hhmm)
$logFilePath = ".\NewUser-"+$datetag+".log" <#adjust log file location if needed #>
if (!(Test-Path $logFilePath)){
    $NewLogFile = New-Item $logFilePath -Force -ItemType File
    }
else { <# nothing to do #> }

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Select a CSV file with a list of users to be created"
	$OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

# Function to generate a random password
Function Get-RandomPassword {
    param (
        [int]$length = 9
    )

    # Define character sets as arrays of individual characters
    $letters = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray()
    $numbers = '123456789'.ToCharArray()
    $specialChars = '@#$%!'.ToCharArray()

    # Create a password with at least one letter, one number, and at most two special characters
    $passwordChars = @()
    $passwordChars += $letters | Get-Random -Count 1
    $passwordChars += $numbers | Get-Random -Count 1

    # Add random characters ensuring at most two special characters
    for ($i = 2; $i -lt $length; $i++) {
        if ($i -lt ($length - 2) -and (Get-Random -Minimum 0 -Maximum 2) -eq 0) {
            # Add a special character (at most 2)
            if (($passwordChars | Where-Object {$_ -match '[^a-zA-Z0-9]'}).Count -lt 2) {
                $passwordChars += $specialChars | Get-Random -Count 1
            } else {
                # Add a letter or number if special chars limit reached
                $passwordChars += ($letters + $numbers) | Get-Random -Count 1
            }
        } else {
            # Add a letter or number
            $passwordChars += ($letters + $numbers) | Get-Random -Count 1
        }
    }

    # Shuffle the password characters and join them into a string
    return -join ($passwordChars | Get-Random -Count $length)
}

$inputfile = Get-FileName "C:\"

if ($inputfile){
    $users = Import-CSV -Path $inputfile

    foreach ($user in $users) {
        # Create display name in the format "LAST NAME, First Name" with LAST NAME in uppercase
        $displayName = "$($user.LastName.ToUpper()), $($user.FirstName)"
        
        # Generate a random password for the user
        $randomPassword = Get-RandomPassword
    
        # Create user parameters
        $userParams = @{
            DisplayName       = $displayName
            GivenName         = $user.FirstName
            Surname           = $user.LastName
            UserPrincipalName = $user.UserPrincipalName
            MailNickname      = $user.MailNickname
            AccountEnabled    = $true
            PasswordProfile   = @{
                ForceChangePasswordNextSignIn = $true
                Password                      = $randomPassword
            }
            CompanyName       = "FIC"
            UsageLocation     = "AG"
        }
    
        try {
            # Create the user account
            $newUser = New-MgUser @userParams -ErrorAction Stop
            Write-Host "New User Created..." -ForegroundColor Yellow
            $createTime = Get-Date -Format hh:mm:ss
    
            # Add user to the specified group for licensing
            New-MgGroupMember -GroupId $groupId -DirectoryObjectId $newUser.Id -ErrorAction Stop
            Write-Host "User Added to licensing group..." -ForegroundColor Yellow
            
            #Pausing the script for 50 seconds to allow the base account creation to be fully completed
            for ($i=50; $i -gt 1; $i--){
            Write-Progress -Activity "Base account being created" -SecondsRemaining $i -Status "Applying license..."
            Start-Sleep 1}
            
            
            # Set primary email address
            Update-MgUser -UserId $newUser.Id -Mail $user.PrimaryEmail -ErrorAction Stop
                    
            # Log success message and output password to screen
            Write-Host "Successfully created user: $displayName with Password: $randomPassword" -ForegroundColor White -BackgroundColor Blue
            Add-Content -path $logFilePath -Value "$($createTime) - Successfully created user: $displayName"
            Write-Host "Log file created: $($NewLogFile)"
            
        } catch {
            # Log error message and output error to screen
            $ErrorMessage = $_.Exception.message
            $ErrorMessage
            Write-Host "Failed to create user: $displayName. Error: $($ErrorMessage)" -ForegroundColor Red
            Add-Content -path $logFilePath -Value "Failed to create user: $displayName. Error: $($ErrorMessage)"
            Write-Host "Log file created: $($NewLogFile)"
        }
    }

}else{
    Exit
}


# Disconnect from Microsoft Graph
Disconnect-MgGraph
