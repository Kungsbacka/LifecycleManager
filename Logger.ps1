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
        [Parameter(Mandatory = $true)]
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
        [Parameter(Mandatory)]
        [ValidateSet('Expire', 'Unexpire', 'Delete', 'Update', 'Create', 'Report')]
        [string]
        $TaskName,
        [Parameter(Mandatory)]
        [int]
        $BatchId,
        [Parameter(Mandatory)]
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
            $spParams.ErrorJson = $ErrorObject | ConvertTo-NewtonsoftJson
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
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
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
