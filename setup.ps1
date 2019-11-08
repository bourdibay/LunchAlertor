# This script creates a new entry in the Task Scheduler to run the script before noon.

$scriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$taskName = "Lunch alertor"

# Delete existing task, if any
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false

# Create the action
$action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument '-windowstyle hidden -ExecutionPolicy ByPass -File .\bootstrap.ps1 -startHour 12 -endHour 15' -WorkingDirectory $scriptDir

# Create trigger 1: trigger at noon
$triggerNoon =  New-ScheduledTaskTrigger -Daily -At 12pm

# Create trigger 2: trigger at unlock
# from https://stackoverflow.com/questions/53704188/syntax-for-execute-on-workstation-unlock
$stateChangeTrigger = Get-CimClass `
    -Namespace ROOT\Microsoft\Windows\TaskScheduler `
    -ClassName MSFT_TaskSessionStateChangeTrigger

$triggerUnlock = New-CimInstance `
    -CimClass $stateChangeTrigger `
    -Property @{
        StateChange = 8  # TASK_SESSION_STATE_CHANGE_TYPE.TASK_SESSION_UNLOCK (taskschd.h)
    } `
    -ClientOnly

# Create an array of triggers to have a task with 2 triggers
$triggers = @(
   $triggerNoon,
   $triggerUnlock
)

# Create the task
Register-ScheduledTask -Action $action -Trigger $triggers -TaskName $taskName -Description "Alert when there is a meeting at lunch time"
