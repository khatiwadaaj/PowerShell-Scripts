# Import the Microsoft.Graph module
Import-Module Microsoft.Graph.Users

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Mail.Read"

# Get the users
$users = Get-MgUser -All

# Prepare an array for storing the results
$result = @()

foreach ($user in $users) {
    $userId = $user.Id
    $displayName = $user.DisplayName
    $upn = $user.UserPrincipalName

    # Get the mailbox size
    $mailbox = Get-MgUserMailboxSetting -UserId $userId
    
    $mailboxSize = if ($mailbox -ne $null) { $mailbox.Quota } else { "Unknown" }

    # Add to result array
    $result += [PSCustomObject]@{
        DisplayName = $displayName
        UPN = $upn
        MailboxSize = $mailboxSize
    }
}

# Export the results to a CSV file
$result | Export-Csv -Path "c:\Office365Users.csv" -NoTypeInformation

Write-Host "User information has been exported to c:\Office365Users.csv"