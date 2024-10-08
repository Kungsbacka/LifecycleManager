﻿# Make all errors terminating errors
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Common.ps1"
. "$PSScriptRoot\Tasks.ps1"
. "$PSScriptRoot\Secure.ps1"
. "$PSScriptRoot\Store.ps1"
. "$PSScriptRoot\Report.ps1"
. "$PSScriptRoot\Logger.ps1"
. "$PSScriptRoot\SqlAgent.ps1"

$taskLimits = Get-Limits
$includedAccountTypes = Get-IncludedAccountTypes
$tasks = Get-PendingTask -TaskName All | Where-Object accountType -in @($includedAccountTypes)
$taskGroups = $tasks | Group-Object -Property task -NoElement
if ($taskLimits.DisableLimits) {
    Enable-Limits
}
else {
    foreach ($group in $taskGroups)
    {
        if (-not $taskLimits.ContainsKey($group.Name))
        {
            throw "Configuration does not contain a limit for task '$($group.Name)'."
            return
        }
        $limit = $taskLimits[$group.Name]
        if ($group.Count -gt $limit)
        {
            throw "Task '$($group.Name)' has a configured limit of $limit, but there are $($group.Count) tasks pending."
            return
        }
    }
}
$batchId = New-LogBatch
foreach ($task in $tasks)
{
    $params = @{
        BatchId = $batchId
        TaskName = $task.Task
        EmployeeNumber = $task.EmployeeNumber
        ObjectGuid = $task.ObjectGuid
        TaskObject = $task
    }
    try
    {
        switch ($task.Task)
        {
            Expire
            {
                $task | Expire-Account
            }
            Unexpire
            {
                $task | Unexpire-Account
            }
            RemoveLicense
            {
                $task | Remove-MsolLicense
            }
            RestoreLicense
            {
                $task | Restore-MsolLicense
            }
            Delete
            {
                $task | Delete-Account
            }
            Update
            {
                $task | Update-Account -NoRename # | Store-UpdatedAccount # LmAccount does not yet exist, will be created when we start to rename
            }
            Move
            {
                $task | Move-Account
            }
            Create
            {
                $task | Create-Account | Store-NewAccount
            }
        }
        New-LogEntry @params
    }
    catch
    {
        $params.ErrorObject = $_
        New-LogEntry @params
        $_ | Write-LmEventLog
    }
}
try
{
    if ($Script:Config.AllAccountsRecipient)
    {
        Send-NewAccountReport -Recipient $Script:Config.AllAccountsRecipient -DoNotMarkReported
    }
    Send-NewAccountReport
}
catch
{
    New-LogEntry -TaskName Report -BatchId $batchId -ErrorObject $_
    $_ | Write-LmEventLog
}

# Sleep to let domain controllers sync changes before full AD import
Start-Sleep -Seconds 60

try
{
    Start-ActiveDirectoryImportJob
}
catch
{
    New-LogEntry -TaskName ADImport -BatchId $batchId -ErrorObject $_
    $_ | Write-LmEventLog
}
Close-LogBatch -BatchId $batchId
