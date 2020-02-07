using namespace System.Security.Principal

param (
    [switch] $Run,
    [switch] $Uninstall
)

$ErrorActionPreference = "Stop"

$UpdateActiveHours = {
    $key = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
    $hour = (get-date).hour
    if ($hour -ge 18 -or $hour -lt 6) {
        Set-ItemProperty $key ActiveHoursStart 18
        Set-ItemProperty $key ActiveHoursEnd 12
    } 
    else {
        Set-ItemProperty $key ActiveHoursStart 6
        Set-ItemProperty $key ActiveHoursEnd 0
    }
    Set-ItemProperty $key IsActiveHoursEnabled 1
}

function GetTask {
    return Get-ScheduledTask | Where-Object { $_.TaskName -eq $TaskName -and $_.TaskPath -eq $TaskPath }
}

function RunTask {
    $task = GetTask
    if ($task) {
        $task | Start-ScheduledTask
        Write-Output "$TaskPath$TaskName executed"
    }
    else {
        Write-Output "$TaskPath$TaskName not found"
    }
}

function InstallTask {
    $command = EncodeCommand $UpdateActiveHours
    $arguments = "-NoLogo -NonInteractive -WindowStyle Hidden"
    $options = @{
        Force = $true
        Action = New-ScheduledTaskAction `
            -Execute "powershell" `
            -Argument "$arguments -EncodedCommand $command"
        Principal = New-ScheduledTaskPrincipal `
            -UserId "LOCALSERVICE" `
            -LogonType ServiceAccount `
            -RunLevel Highest
        Settings = New-ScheduledTaskSettingsSet `
            -WakeToRun `
            -StartWhenAvailable `
            -DontStopOnIdleEnd `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries
        Trigger = @(
            New-ScheduledTaskTrigger_AtWakeUp
            New-ScheduledTaskTrigger -AtStartup
            New-ScheduledTaskTrigger -Daily -At "06:00"
            New-ScheduledTaskTrigger -Daily -At "18:00"
        )
    }
    Register-ScheduledTask $TaskName $TaskPath @options | Out-Null
    Write-Output "$TaskPath$TaskName installed"
    RunTask
}

function UninstallTask {
    $task = GetTask
    if ($task) {
        $task | Unregister-ScheduledTask -Confirm:$false
        $pathName = $TaskPath.Replace("\", "")
        $scheduleObject = New-Object -ComObject Schedule.Service
        $scheduleObject.Connect()
        $rootFolder = $scheduleObject.GetFolder("\")
        $rootFolder.DeleteFolder("$pathName", $null)
        Write-Output "$TaskPath$TaskName uninstalled"
    }
    else {
        Write-Output "$TaskPath$TaskName not found"
    }
}

function New-ScheduledTaskTrigger_AtWakeUp {
    $name = "MSFT_TaskEventTrigger"
    $namespace = "Root/Microsoft/Windows/TaskScheduler:MSFT_TaskEventTrigger"
    $class = Get-CimClass -ClassName $name -Namespace $namespace
    $trigger = New-CimInstance -CimClass $class -ClientOnly
    $trigger.Enabled = $true
    $trigger.Subscription = "<QueryList><Query Id=""0"" Path=""System""><Select Path=""System"">*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]</Select></Query></QueryList>"
    return $trigger
}

function EncodeCommand {
    param ([string] $command)
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($command)
    $encodedCommand = [Convert]::ToBase64String($bytes)
    return $encodedCommand
}

if (-not ([WindowsPrincipal][WindowsIdentity]::GetCurrent()).IsInRole([WindowsBuiltInRole]::Administrator)) {
    Write-Output "Run as administrator"
    exit
}

$TaskName = "Update"
$TaskPath = "\Win10ActiveHours\"

if ($Run.IsPresent) {
    RunTask
}
elseif ($Uninstall.IsPresent) {
    UninstallTask
}
else {
    InstallTask
}

exit
