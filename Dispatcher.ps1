# Make all errors terminating errors
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Common.ps1"
. "$PSScriptRoot\Tasks.ps1"
. "$PSScriptRoot\Secure.ps1"
. "$PSScriptRoot\Store.ps1"
. "$PSScriptRoot\Report.ps1"
. "$PSScriptRoot\Logger.ps1"
. "$PSScriptRoot\SqlAgent.ps1"
$tasks = @()
$tasks += Get-PendingTask -TaskName Unexpire
$tasks += Get-PendingTask -TaskName Expire
$tasks += Get-PendingTask -TaskName RemoveLicense
$tasks += Get-PendingTask -TaskName RestoreLicense
$tasks += Get-PendingTask -TaskName Delete
$tasks += Get-PendingTask -TaskName Update
$tasks += Get-PendingTask -TaskName Move
$tasks += Get-PendingTask -TaskName Create | where accountType -eq 'Elev' # Only create student accounts for now
$taskGroups = $tasks | Group-Object -Property task -NoElement
foreach ($group in $taskGroups) {
    if (-not $Script:Config.Limits.ContainsKey($group.Name))
    {
        throw "Configuration does not contain a limit for task '$($group.Name)'."
    }
    $limit = $Script:Config.Limits[$group.Name]
    if ($group.Count -gt $limit) {
        throw "Task '$($group.Name)' has a configured limit of $limit, but there are $($group.Count) tasks pending."
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
