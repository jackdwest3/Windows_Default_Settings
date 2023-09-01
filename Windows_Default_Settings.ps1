# Import the Syncro PowerShell module
Import-Module $env:SyncroModule

# Create a Syncro ticket for setting the registry key
$ticketResult = Create-Syncro-Ticket -Subject "Set Registry Key" -IssueType "Script Execution" -Status "New"

# Check if the ticket was created successfully
if ($ticketResult -and $ticketResult.ticket) {
    $ticketId = $ticketResult.ticket.id

    # Specify the registry path and key details
    $registryPath = "HKCU:\Microsoft\Windows\CurrentVersion\Feeds"
    $registryKeyName = "ShellFeedsTaskbarViewMode"
    $desiredValue = "2"

    # Check if the registry key exists and its current value
    $keyExists = Test-Path -Path $registryPath
    if ($keyExists) {
        $currentValue = (Get-ItemProperty -Path $registryPath -Name $registryKeyName).$registryKeyName
    }

    # Check if the key exists and if it's already set to the desired value
    if ($keyExists -and $currentValue -eq $desiredValue) {
        Write-Host "Registry key is already set to the desired value."
        # Add note to the ticket
        $note = "Registry key is already set to the desired value."
    }
    else {
        # Create the registry key and value or update the value
        if (-not $keyExists) {
            New-Item -Path $registryPath -Force
        }
        New-ItemProperty -Path $registryPath -Name $registryKeyName -Value $desiredValue -Force

        Write-Host "Registry key and value created or updated."
        # Add note to the ticket
        $note = "Registry key and value created or updated."
    }

    # Add time entry with note to the ticket
    Create-Syncro-Ticket-TimerEntry -TicketIdOrNumber $ticketId -StartTime (Get-Date).ToString("o") -DurationMinutes 5 -Notes $note -UserIdOrEmail "your.user.email@here.com" -ChargeTime "false"

    # Log activity on the asset
    Log-Activity -Message "Checked and set registry key" -EventName "Registry Key Set" -TicketIdOrNumber $ticketId
}
else {
    Write-Host "Failed to create Syncro ticket for setting registry key"
}
