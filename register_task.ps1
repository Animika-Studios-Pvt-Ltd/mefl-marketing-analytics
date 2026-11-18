# Script to configure Windows Scheduled Task for MEFL Data Sync
$TaskName = "MEFL_Marketing_Sync"
$PythonExe = "C:\Python314\python.exe"
$ScriptPath = "C:\MEFL_data1\sync_campaign_data.py"
$ExcelPath = "C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\MEFL_Analysis folder\MEFL_Analysis.xlsx"
$WorkingDir = "C:\MEFL_data1"

Write-Host "Configuring Scheduled Task: $TaskName"

# 1. Action: Execute Python with script path and excel path
$Action = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "`"$ScriptPath`" --excel `"$ExcelPath`"" `
    -WorkingDirectory $WorkingDir

# 2. Trigger: Repeat every 15 minutes indefinitely
$Trigger = New-ScheduledTaskTrigger `
    -Once `
    -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes 15)

# 3. Settings:
# - MultipleInstances IgnoreNew: Prevents overlapping runs if previous is still running
# - ExecutionTimeLimit: Terminate if hung after 10 minutes
# - StartWhenAvailable: Run if scheduled start missed
# - DisallowStartIfOnBatteries $false: Run even when laptop is on battery power
# - StopIfGoingOnBatteries $false: Don't kill task when switching to battery
$Settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries

# 4. Register Task under current user
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Synchronizes Mother Earth Farmlands marketing campaign data from Excel/CSV to PostgreSQL every 15 minutes." `
    -Force

Write-Host "Task registered successfully."
