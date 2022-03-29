# Add-DistributionGroupBulk
# v1.1
# Written by Elliott Nash
# Created 11/03/2022

# This source code is distributed under the terms of the Bad Code License.
# You are forbidden from distributing software containing this code to end
# users, because it is bad


# Get name of mailbox
$DistList = Read-Host -Prompt "Enter the name of the mailbox"

# Check mailbox exists
if (Get-DistributionGroup -Identity $DistList -ErrorAction SilentlyContinue) {
    Write-Host "Mailbox found. Continuing..." -ForegroundColor Green
}
else {
    Write-Host "Mailbox not found" -ForegroundColor Red
    Exit
}

# Get name of .csv
$DistCSV = Read-Host -Prompt "Enter the name of the CSV"

# Check CSV exists
if (Test-Path -Path $DistCSV) {
    Write-Host "File found. Continuing..." -ForegroundColor Green
}
else {
    Write-Host "File not found" -ForegroundColor Red
    Exit
}

# Confirm to avoid errors
Write-Host "This action wil remove everyone from $DistList and replace them with the users in $DistCSV Press Y to confirm, or N to cancel" -ForegroundColor Red 
$Confirm = Read-Host
if ($Confirm -eq "Y") {
    # Removes all existing users from mailbox
    # This is to ensure only the necessary people are members
    Write-Host "Removing users" -ForegroundColor Yellow
    $List = Get-DistributionGroupMember -Identity $DistLIst
    foreach ($DistUser in $List) {
        Write-Host $DistUser
        Remove-DistributionGroupMember -Identity $DistList -Member $DistUser -Confirm: $False
    }
    }
    Write-Host "Users removed" -ForegroundColor Green

    # Add users
    Write-Host "Adding users" -ForegroundColor Yellow
        # Reads the CSV to iterate through one user at a time
        # The iteration is to be able to flag errors and close gracefully
        $CsvRead = Import-Csv $DistCSV
        foreach ( $Line in $CsvRead ) {
            try {
                Add-DistributionGroupMember -Identity $DistList -Member $( $Line.Name ) -ErrorAction Stop
            }
            catch {
                $ErrorMsg = $_
                # If the heading of the CSV hasn't been set, it will throw an error.
                # This error can be resolved by checking the first line of the CSV is "Name,"
                if ($ErrorMsg -like "*Cannot validate argument on parameter 'Member'. The argument is null. Provide a valid value for the argument, and then try running the command again.*") {
                    Write-Host "ERROR: Please ensure the first item on the CSV is 'Name,'" -ForegroundColor Red
                    Exit
                }
                else {
                    Write-Host $ErrorMsg
                }
            }
        }

    Write-Host "Users added" -ForegroundColor Green
}
else {
  Write-Host "Exiting" -ForegroundColor Red
  Exit
}

# CHANGELOG:
#  * Tested and fixed removing all users
#  * Added changelog 

# TODO:
#  * More error handling
#  * Refactor into subroutines
#  * Run from shell as command with arguments