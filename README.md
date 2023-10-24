# Lifecycle Manager

## Description

Lifecycle Manager manages the Active Directory user account lifecycle. It can perform the following tasks:

* Create a new account (Create-Account)
* Update account properties (Update-Account)
* Delete an account (Delete-Account)
* Set/remove expiration date (Expire-Account/Unexpire-Account)
* Move an account to a different location (Move-Account)
* Instruct Resource Manager to remove or restore Microsoft 365 licenses (Remove-MsolLicense/Restore-MsolLicense).

Information on what tasks to perform is fetched from the MetaDirectory database and processed by the dispatcher.

## Dependencies

* Assemblies: [Kungsbacka.DS](https://github.com/Kungsbacka/Kungsbacka.DS), [Kungsbacka.AccountTasks](https://github.com/Kungsbacka/Kungsbacka.AccountTasks), [Kungsbacka.CommonExtensions](https://github.com/Kungsbacka/Kungsbacka.CommonExtensions), [EPPlus](https://github.com/EPPlusSoftware/EPPlus) and [Newtonsoft Json](https://www.newtonsoft.com/json)
* Databases: MetaDirectory and ADEvents (logging)

## Deploying

1. Create a service account (preferably a Managed Service Account) with the appropriate permissions (see below)
2. Create a folder on a server and copy/clone LifecycleManager to the folder
3. Create a sub folder called lib and copy DLLs for the assemblies above to the folder (Kungsbacka.DS.dll, Kungsbacka.AccountTasks.dll, Kungsbacka.CommonExtensions.dll, EPPlus.dll and Newtonsoft.Json.dll).
4. Rename Config.example.ps1 to Config.ps1 and update the file with settings for your environment.
5. Register a scheduled task (see below)

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

* Read and write permissions in the script folder. Reports are created temporarily in this folder before they are sent.
* Manage users in Active Directory (create, update and remove)
* Read and write permissions in databases MetaDirectory and ADEvents
* Start SQL Agent job for Active Directory import

## Log Database

```SQL
CREATE TABLE [dbo].[LmBatch] (
    [id] [int] IDENTITY(1,1) NOT NULL,
    [started] [datetime] NOT NULL,
    [ended] [datetime] NULL,
    CONSTRAINT [PK_ApBatch] PRIMARY KEY CLUSTERED ([id])
);

CREATE TABLE [dbo].[LmLog](
    [id] [int] IDENTITY(1,1) NOT NULL,
    [batchId] [int] NOT NULL,
    [task] [nvarchar](50) NOT NULL,
    [completed] [datetime] NOT NULL,
    [employeeNumber] [nvarchar](512) NOT NULL,
    [objectGUID] [uniqueidentifier] NULL,
    [taskJson] [nvarchar](2000) NULL,
    [error] [nvarchar](2000) NULL,
    CONSTRAINT [PK_LmLog] PRIMARY KEY CLUSTERED ([id])
);

CREATE PROCEDURE [dbo].[spLmEndBatch]
    @batchId int
AS
BEGIN

    UPDATE
        dbo.LmBatch
    SET
        ended = GETDATE()
    WHERE
        id = @batchId
    AND
        ended IS NULL

    IF @@ROWCOUNT = 0
        THROW 50000, 'Failed to end batch', 1
END;

CREATE PROCEDURE [dbo].[spLmGetBatch]
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @batchId int;

    SELECT @batchId = (
        SELECT MAX(id) FROM dbo.LmBatch
        WHERE ended IS NULL
    );

    IF @batchId IS NULL
    BEGIN
        INSERT INTO dbo.LmBatch (started)
        VALUES (GETDATE());
        SELECT @batchId = SCOPE_IDENTITY();
    END

    SELECT @batchId;
END;

CREATE PROCEDURE [dbo].[spLmNewLogEntry]
      @batchId int
    , @task nvarchar(50)
    , @employeeNumber nvarchar(512)
    , @objectGUID uniqueidentifier = NULL
    , @taskJson nvarchar(2000) = NULL
    , @error nvarchar(2000) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    IF NOT EXISTS (SELECT 1 FROM dbo.LmBatch WHERE id = @batchId AND ended IS NULL)
        THROW 50000, 'No open batch with requested ID exists', 1

    INSERT INTO
        dbo.LmLog (
            [batchId]
          , [task]
          , [completed]
          , [employeeNumber]
          , [objectGUID]
          , [taskJson]
          , [error]
    )
    VALUES (
        @batchId
      , @task
      , GETDATE()
      , @employeeNumber
      , @objectGUID
      , @taskJson
      , @error
    )
END;
```

## Additional information

This solution is tailored specifically for Kungsbacka municipality. Schema for the MetaDirectory database is not publicly available. 
