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

    ############ Disable Windows Transparency ###############
    # Specify the registry path and key details for EnableTransparency
    $transparencyRegistryPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    $transparencyRegistryKeyName = "EnableTransparency"
    $desiredTransparencyValue = "0"  # Disable transparency

    $transparencyKeyExists = Test-Path -Path $transparencyRegistryPath
    if ($transparencyKeyExists) {
        $currentTransparencyValue = (Get-ItemProperty -Path $transparencyRegistryPath -Name $transparencyRegistryKeyName).$transparencyRegistryKeyName
    }

    # Check if the key exists and if it's already set to the desired value for EnableTransparency
    if ($transparencyKeyExists -and $currentTransparencyValue -eq $desiredTransparencyValue) {
        Write-Host "Transparency registry key is already set to disabled."
        # Add note to the ticket
        $note += "`nTransparency registry key is already set to disabled."
    }
    else {
        # Create the registry key and value or update the value for EnableTransparency
        if (-not $transparencyKeyExists) {
            New-Item -Path $transparencyRegistryPath -Force
        }
        New-ItemProperty -Path $transparencyRegistryPath -Name $transparencyRegistryKeyName -Value $desiredTransparencyValue -Force

        Write-Host "Transparency registry key and value created or updated."
        # Add note to the ticket
        $note += "`nTransparency registry key and value created or updated."
        # Log activity on the asset
        Log-Activity -Message "Windows Transparancy disabled." -EventName "Registry Actions" -TicketIdOrNumber $ticketId
    }

    ############ Adjust Visual Effects for Performance ############
    $menuShowDelay = (Get-ItemProperty -Path "HKCU:\Control Panel\Desktop").MenuShowDelay
    if ($menuShowDelay -ge 400) {
        Write-Host "Adjusting visual effects for performance..."
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "DragFullWindows" -Type String -Value 0
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "MenuShowDelay" -Type String -Value 200
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Type Binary -Value ([byte[]](144,18,3,128,16,0,0,0))
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Type String -Value 0
        Set-ItemProperty -Path "HKCU:\Control Panel\Keyboard" -Name "KeyboardDelay" -Type DWord -Value 0
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ListviewAlphaSelect" -Type DWord -Value 0
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ListviewShadow" -Type DWord -Value 0
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAnimations" -Type DWord -Value 0
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFXSetting" -Type DWord -Value 3
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\DWM" -Name "EnableAeroPeek" -Type DWord -Value 0
        Write-Host "Visual effects adjusted for performance."
        # Add note to the ticket
        $note += "`nWindows visual effects tuned for performance."
        # Log activity on the asset
        Log-Activity -Message "Windows visual effects tuned for performance." -EventName "Registry Actions" -TicketIdOrNumber $ticketId
    }
    else {
        Write-Host "Menu  ShowDelay is less than 400. No visual tweaks applied."
        # Add note to the ticket
        $note += "`nMenu  ShowDelay is less than 400. No visual tweaks applied."
    }

    ############ Enable Always-Visible Scroll Bars ############
    # Specify the registry path and key details for DynamicScrollbars
    $scrollbarRegistryPath = "HKCU:\Control Panel\Accessibility"
    $scrollbarRegistryKeyName = "DynamicScrollbars"
    $desiredScrollbarValue = "0"  # Enable always-visible scroll bars

    $scrollbarKeyExists = Test-Path -Path $scrollbarRegistryPath
    if ($scrollbarKeyExists) {
        $currentScrollbarValue = (Get-ItemProperty -Path $scrollbarRegistryPath -Name $scrollbarRegistryKeyName).$scrollbarRegistryKeyName
    }

    # Check if the key exists and if it's already set to the desired value for DynamicScrollbars
    if ($scrollbarKeyExists -and $currentScrollbarValue -eq $desiredScrollbarValue) {
        Write-Host "Always-visible scroll bars are already enabled."
        # Add note to the ticket
        $note += "`nAlways-visible scroll bars are already enabled."
    }
    else {
        # Create the registry key and value or update the value for DynamicScrollbars
        if (-not $scrollbarKeyExists) {
            New-Item -Path $scrollbarRegistryPath -Force
        }
        New-ItemProperty -Path $scrollbarRegistryPath -Name $scrollbarRegistryKeyName -Value $desiredScrollbarValue -Force
        Write-Host "Always-visible scroll bars enabled."
        # Add note to the ticket
        $note += "`nAlways-visible scroll bars enabled."
        # Log activity on the asset
        Log-Activity -Message "Always-visible scroll bars enabled." -EventName "Registry Actions" -TicketIdOrNumber $ticketId
    }
    ##########################################################
    # Add time entry with combined notes to the ticket
    Create-Syncro-Ticket-TimerEntry -TicketIdOrNumber $ticketId -StartTime (Get-Date).ToString("o") -DurationMinutes 10 -Notes $notes -UserIdOrEmail $userEmail -ChargeTime "false"
}
else {
    # Let know that failed to create ticket in terminal
    Write-Host "Failed to create Syncro ticket."
}