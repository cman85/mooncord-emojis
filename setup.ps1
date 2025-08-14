# PowerShell script to create "Pull Emoji" scheduled task
# Run this script as Administrator

# Task basic information
$TaskName = "Pull Emoji"
$TaskPath = "\"

try {
    # Check if task already exists and remove it
    if (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue) {
        Write-Host "Task '$TaskName' already exists. Removing existing task..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false
    }

    # Create Task Scheduler COM object
    $TaskService = New-Object -ComObject Schedule.Service
    $TaskService.Connect()
    $RootFolder = $TaskService.GetFolder("\")

    # Create task definition
    $TaskDefinition = $TaskService.NewTask(0)
    
    # Set general task properties
    $TaskDefinition.RegistrationInfo.Description = "Pull Emoji"
    $TaskDefinition.RegistrationInfo.Author = $env:USERNAME
    $TaskDefinition.Settings.Enabled = $true
    $TaskDefinition.Settings.Hidden = $false
    $TaskDefinition.Settings.AllowDemandStart = $true
    $TaskDefinition.Settings.DisallowStartIfOnBatteries = $true
    $TaskDefinition.Settings.StopIfGoingOnBatteries = $true
    $TaskDefinition.Settings.WakeToRun = $false
    $TaskDefinition.Settings.RunOnlyIfNetworkAvailable = $false
    $TaskDefinition.Settings.RunOnlyIfIdle = $false
    $TaskDefinition.Settings.ExecutionTimeLimit = "PT2M"  # 2 minutes
    $TaskDefinition.Settings.Priority = 7
    $TaskDefinition.Settings.RestartCount = 0
    $TaskDefinition.Settings.MultipleInstances = 0  # Ignore new instance
    
    # Set idle settings
    $TaskDefinition.Settings.IdleSettings.StopOnIdleEnd = $false
    $TaskDefinition.Settings.IdleSettings.RestartOnIdle = $false
    $TaskDefinition.Settings.IdleSettings.IdleDuration = "PT10M"  # 10 minutes
    $TaskDefinition.Settings.IdleSettings.WaitTimeout = "PT1H"   # 1 hour

    # Create trigger (One time at 8/14/2025 3:05 PM with repetition every 5 minutes)
    $Trigger = $TaskDefinition.Triggers.Create(1)  # TASK_TRIGGER_TIME = 1
    $Trigger.StartBoundary = "2025-08-14T15:05:00"
    # No EndBoundary set - task will run indefinitely
    $Trigger.Enabled = $true
    
    # Set repetition (every 5 minutes, indefinitely)
    $Trigger.Repetition.Interval = "PT5M"     # Every 5 minutes  
    $Trigger.Repetition.Duration = ""         # Indefinitely
    $Trigger.Repetition.StopAtDurationEnd = $false

    # Create action (Start conhost.exe with git pull command)
    $Action = $TaskDefinition.Actions.Create(0)  # TASK_ACTION_EXEC = 0
    $Action.Path = "C:\Windows\System32\conhost.exe"
    $Action.Arguments = "--headless powershell.exe -WindowStyle Hidden -NoProfile -NonInteractive /c `"git pull`""
    $Action.WorkingDirectory = "$env:USERPROFILE\.runelite\emojis"

    # Set principal (run only when user is logged on)
    $TaskDefinition.Principal.UserId = $env:USERNAME
    $TaskDefinition.Principal.LogonType = 3  # TASK_LOGON_INTERACTIVE_TOKEN = 3
    $TaskDefinition.Principal.RunLevel = 0   # TASK_RUNLEVEL_LUA = 0 (standard user)

    # Register the task
    $RegisteredTask = $RootFolder.RegisterTaskDefinition(
        $TaskName,
        $TaskDefinition,
        6,      # TASK_CREATE_OR_UPDATE = 6
        $null,  # User (use current user)
        $null,  # Password (not needed for interactive logon)
        3       # TASK_LOGON_INTERACTIVE_TOKEN = 3
    )

    Write-Host "Successfully created scheduled task '$TaskName'" -ForegroundColor Green
    
    # Display task information
    Write-Host "`nTask Details:" -ForegroundColor Cyan
    Write-Host "Name: $($RegisteredTask.Name)" -ForegroundColor White
    Write-Host "Path: $($RegisteredTask.Path)" -ForegroundColor White
    Write-Host "State: $($RegisteredTask.State)" -ForegroundColor White
    Write-Host "Enabled: $($RegisteredTask.Enabled)" -ForegroundColor White
    
    # Get next run time
    try {
        $NextRunTime = $RegisteredTask.NextRunTime
        Write-Host "Next Run Time: $NextRunTime" -ForegroundColor White
    } catch {
        Write-Host "Next Run Time: Not scheduled (task may be disabled or conditions not met)" -ForegroundColor Yellow
    }
    
    # Show trigger details
    Write-Host "`nTrigger Details:" -ForegroundColor Cyan
    Write-Host "Start: 8/14/2025 3:05:00 PM" -ForegroundColor White
    Write-Host "Repeat: Every 5 minutes indefinitely" -ForegroundColor White
    Write-Host "Expire: Never (runs indefinitely)" -ForegroundColor White
    
    # Show power conditions
    Write-Host "`nPower Conditions:" -ForegroundColor Cyan
    Write-Host "Start only if on AC power: True" -ForegroundColor White
    Write-Host "Stop if going on batteries: True" -ForegroundColor White
    Write-Host "Wake computer to run: False" -ForegroundColor White
    
    # Note about the working directory
    Write-Host "`nIMPORTANT NOTE:" -ForegroundColor Yellow
    Write-Host "Make sure the working directory '$env:USERPROFILE\.runelite\emojis' exists and contains a git repository." -ForegroundColor Yellow
    Write-Host "The task will fail if this directory doesn't exist or isn't a valid git repository." -ForegroundColor Yellow

} catch {
    Write-Error "Failed to create scheduled task: $($_.Exception.Message)"
    Write-Host "Error details: $($_.Exception)" -ForegroundColor Red
    exit 1
} finally {
    # Clean up COM objects
    if ($TaskService) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($TaskService) | Out-Null
    }
}

# Optional: Show the created task in Task Scheduler
$ShowInTaskScheduler = Read-Host "`nWould you like to open Task Scheduler to view the created task? (y/n)"
if ($ShowInTaskScheduler -eq 'y' -or $ShowInTaskScheduler -eq 'Y') {
    Start-Process "taskschd.msc"
}
