# Make all errors terminating errors
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Common.ps1"
. "$PSScriptRoot\Tasks.ps1"
. "$PSScriptRoot\Secure.ps1"
. "$PSScriptRoot\Store.ps1"
. "$PSScriptRoot\Report.ps1"
. "$PSScriptRoot\Logger.ps1"
. "$PSScriptRoot\SqlAgent.ps1"

$taskLimits = Get-Limits
$tasks = Get-PendingTask -TaskName All | Where-Object accountType -in $Script:Config.IncludedAccountTypes
$taskGroups = $tasks | Group-Object -Property task -NoElement
if (-not $taskLimits.DisableLimits)
{
    foreach ($group in $taskGroups)
    {
        if (-not $taskLimits.ContainsKey($group.Name))
        {
            throw "Configuration does not contain a limit for task '$($group.Name)'."
        }
        $limit = $taskLimits[$group.Name]
        if ($group.Count -gt $limit)
        {
            throw "Task '$($group.Name)' has a configured limit of $limit, but there are $($group.Count) tasks pending."
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
    Send-NewAccountReport
}
catch
{
    New-LogEntry -TaskName Report -BatchId $batchId -ErrorObject $_
    $_ | Write-LmEventLog
}
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
