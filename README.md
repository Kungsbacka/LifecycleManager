# Lifecycle Manager

## Description
Lifecycle Manager creates, deletes and updates user accounts in Active Directory.

## Dependencies
* Assemblies: [Kungsbacka.DS](https://github.com/Kungsbacka/Kungsbacka.DS), [Kungsbacka.AccountTasks](https://github.com/Kungsbacka/Kungsbacka.AccountTasks), [Kungsbacka.CommonExtensions](https://github.com/Kungsbacka/Kungsbacka.CommonExtensions), [EPPlus](https://epplus.codeplex.com/) and [Newtonsoft Json](http://www.newtonsoft.com/json)
* Databases: MetaDirectory and ADEvents

## Deploying
1. Create a service account (preferably a Managed Service Account) with the appropriate permissions (see below)
2. Create a folder on a server and copy/clone LifecycleManager to the folder
3. Create a subfolder called lib and copy DLLs for the assemblies above to the folder (Kungsbacka.DS.dll, Kungsbacka.AccountTasks.dll, Kungsbacka.CommonExtensions.dll, EPPlus.dll and Newtonsoft.Json.dll).
4. Rename Config.example.ps1 to Config.ps1 and update it with settings for your environment.
5. Register a new event source on the server: [System.Diagnostics.EventLog]::CreateEventSource('LifecycleManager', 'Application')
6. Register a scheduled task (see below)

## Scheduled task
This is a template script for creating a scheduled task that runs Lifecycle Manager
```powershell
Register-ScheduledTask `
    -TaskName 'LifecycleManager' `
    -TaskPath '\' `
    -Description 'Creates, deletes and updates user accounts in Active Directory.' `
    -Principal (New-ScheduledTaskPrincipal -UserId '<service account>' -LogonType Password) `
    -Trigger (New-ScheduledTaskTrigger -At 02:00 -Daily) `
    -Action (New-ScheduledTaskAction `
        -Execute 'powershell.exe' `
        -Argument '-Command "<path to Dispatcher.ps1>"' `
        -WorkingDirectory '<script folder path>') `
    -Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable)
```

## Permissions
The following permissions are needed for the account running the script:
* Read/write to the script folder. Reports are created temporarily in this folder before they are sent.
* Manage users in Active Directory
* Access databases (MetaDirectory and ADEvents)
* Start SQL Agent job for Active Directory import

## Additional information
This solution is tailored specifically for Kungsbacka municipality. Schema for the two databases are not included here, but may get published later.

## TODO
Improve logging
Better documentation
