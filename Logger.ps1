# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Common.ps1"

function New-LogBatch
{
    $params = @{
        Database = 'ADEvents'
        Procedure = 'dbo.spLmGetBatch'
        Scalar = $true
    }
    Invoke-StoredProcedure @params
}

function Close-LogBatch
{
    param
    (
        [Parameter(Mandatory=$true)]
        [int]
        $BatchId
    )
    $params = @{
        Database = 'ADEvents'
        Procedure = 'dbo.spLmEndBatch'
        Parameters = @{
            BatchId = $BatchId
        }
    }
    Invoke-StoredProcedure @params
}

function New-LogEntry
{
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateSet('Expire', 'Unexpire', 'RemoveLicense', 'RestoreLicense', 'Delete', 'Update', 'Move', 'Create', 'Report', 'ADImport')]
        [string]
        $TaskName,
        [Parameter(Mandatory=$true)]
        [int]
        $BatchId,
        [string]
        $EmployeeNumber,
        [string]
        $ObjectGuid,
        [object]
        $TaskObject,
        [object]
        $ErrorObject
    )
    process
    {
        $spParams = @{
            Task = $TaskName
            BatchId = $BatchId
            EmployeeNumber = $EmployeeNumber
        }
        if ($ObjectGuid)
        {
            $spParams.ObjectGuid = $ObjectGuid
        }
        if ($TaskObject)
        {
            $spParams.TaskJson = $TaskObject | Serialize-DataRow | ConvertTo-NewtonsoftJson
        }
        if ($ErrorObject)
        {
            $spParams.Error = $ErrorObject.Exception.ToString()
        }
        $params = @{
            Database = 'ADEvents'
            Procedure = 'dbo.spLmNewLogEntry'
            Parameters = $spParams 
        }
        Invoke-StoredProcedure @params
    }
}

function Serialize-DataRow
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseApprovedVerbs", "")]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [System.Data.DataRow]
        $DataRow
    )
    $out = @{}
    foreach ($col in $DataRow.Table.Columns)
    {
        $name = $col.ColumnName
        $value = $DataRow."$name"
        if ($value -isnot [DBNull])
        {
            $out."$name" = $value
        }
    }
    $out | Write-Output
}
