# Import the Syncro PowerShell module
Import-Module $env:SyncroModule

# Intialize varibles
$note = ""
$userEmail = "jack@westcomputers.com"

# Create a Syncro ticket for setting the registry key
$ticketResult = Create-Syncro-Ticket -Subject "Set Registry Key" -IssueType "Script Execution" -Status "New"

# Check if the ticket was created successfully
if ($ticketResult -and $ticketResult.ticket) {
    $ticketId = $ticketResult.ticket.id

    ########### Windows News and Interest Taskbars ###########
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
        Write-Host "Windows News and Interest Taskbars is turned off."
        # Add note to the ticket
        $note += "`nWindows News and Interest Taskbars is turned off."
    }
    else {
        # Create the registry key and value or update the value
        if (-not $keyExists) {
            New-Item -Path $registryPath -Force
        }
        New-ItemProperty -Path $registryPath -Name $registryKeyName -Value $desiredValue -Force

        Write-Host "Windows News and Interest Taskbars is turned on > off."
        # Add note to the ticket
        $note += "`nWindows News and Interest Taskbars is turned on > off."
        # Log activity on the asset
        Log-Activity -Message "Windows News and Interest Taskbars is turned on > off." -EventName "Registry Key Set" -TicketIdOrNumber $ticketId
    }
    ##########################################################

    ############ Windows SysMain Service Disable #############
    # Check and handle the SysMain service
    $sysMainService = Get-Service -Name SysMain -ErrorAction SilentlyContinue

    if ($sysMainService -and $sysMainService.Status -ne "Stopped") {
        # Stop the SysMain service if it's not already stopped
        Stop-Service -Name SysMain -Force

        Write-Host "SysMain service stopped."

        # Disable the SysMain service from starting
        Set-Service -Name SysMain -StartupType Disabled

        Write-Host "SysMain service disabled from starting."

        # Add note to the ticket
        $note += "`nSysMain service stopped and disabled from starting."
        # Log activity on the asset
        Log-Activity -Message "SysMain service disabled from startup." -EventName "Service and Registry Actions" -TicketIdOrNumber $ticketId
    }

    # Add time entry with note to the ticket
    Create-Syncro-Ticket-TimerEntry -TicketIdOrNumber $ticketId -StartTime (Get-Date).ToString("o") -DurationMinutes 5 -Notes $note -UserIdOrEmail $userEmail -ChargeTime "false"
}
else {
    # Let know that failed to create ticket in terminal
    Write-Host "Failed to create Syncro ticket."
}
