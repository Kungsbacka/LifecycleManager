# Lifecycle Manager

## Description
Lifecycle Manager creates, deletes and updates user accounts in Active Directory.

## Dependencies
* Assemblies: Kungsbacka.DS, Kungsbacka.AccountTasks, Kungsbacka.CommonExtensions, [EPPlus](https://epplus.codeplex.com/) and [Newtonsoft Json](http://www.newtonsoft.com/json)
* Databases: MetaDirectory and ADEvents

## Deploying
1. Create a folder on a server and a service account (preferably a gMSA) with read/write permission to the folder.
2. Copy all ps1 files to the folder.
3. Rename Config.example.ps1 to Config.ps1 and update it with settings for your environment.
4. Copy DLLs for the assemblies above to the folder: Kungsbacka.DS.dll, Kungsbacka.AccountTasks.dll, Kungsbacka.CommonExtensions.dll, EPPlus.dll and Newtonsoft.Json.dll.
5. Set up a new scheduled task on the server with the new service account that runs Dispatcher.ps1 in a PowerShell session.

## Additional information
This solution is tailored specifically for Kungsbacka municipality. Schemas for the two databases are not included here, but may get published later. Missing are also Kungsbacka.DS, Kungsbacka.AccountTasks and Kungsbacka.CommonExtensions. These will be moved to GitHub in the near future.

## TODO
Improve logging
