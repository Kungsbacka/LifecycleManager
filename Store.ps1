# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

function Store-NewAccount
{
    param
    (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $ObjectGuid,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $EmployeeNumber,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $AccountPassword,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $AccountType,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $GivenName,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Surname,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Department,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Office,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Title
    )
    $query = 'INSERT INTO dbo.LmNewAccount (created, objectGUID, employeeNumber, userPrincipalName, sAMAccountName, accountPassword, accountType, givenName, sn, department, physicalDeliveryOfficeName, title) VALUES (GETDATE(), @guid, @empno, @upn, @sam, @pw, @type, @gn, @sn, @dept, @office, @title)'
    $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
    [void]$cmd.Parameters.AddWithValue('@guid', $ObjectGuid)
    [void]$cmd.Parameters.AddWithValue('@empno', $EmployeeNumber)
    [void]$cmd.Parameters.AddWithValue('@upn', $UserPrincipalName)
    [void]$cmd.Parameters.AddWithValue('@sam', $SamAccountName)
    [void]$cmd.Parameters.AddWithValue('@pw',  $AccountPassword)
    [void]$cmd.Parameters.AddWithValue('@type',  $AccountType)
    [void]$cmd.Parameters.AddWithValue('@gn',  $GivenName)
    [void]$cmd.Parameters.AddWithValue('@sn',  $Surname)
    [void]$cmd.Parameters.AddWithValue('@dept',  $Department)
    [void]$cmd.Parameters.AddWithValue('@office',  $Office)
    [void]$cmd.Parameters.AddWithValue('@title',  $Title)
    [void]$cmd.ExecuteNonQuery()
}

function Get-PendingTask
{
    param
    (
        [Parameter(Position = 0)]
        [ValidateSet('Expire','Unexpire','Delete','Create','Update','Move','RemoveLicense','RestoreLicense','All')]
        [string]
        $TaskName = 'All',
        [Parameter(Position = 1)]
        [ValidateSet('Current','Stored')]
        [string]
        $TaskSource = 'Current'
    )

    if ($TaskSource -eq 'Current')
    {
        InternalGetCurrentTask -TaskName $TaskName
    }
    else
    {
        InternalGetStoredTask -TaskName $TaskName
    }
}

function InternalGetCurrentTask
{
    param
    (
        [string]$TaskName
    )
    $query = 'SELECT * FROM dbo.ufLmGetPendingLifecycleTask(@task, 1)' # 1 = Do not include renames
    $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
    if ($TaskName -eq 'All')
    {
        [void]$cmd.Parameters.AddWithValue('@task', [DBNull]::Value)
    }
    else
    {
        [void]$cmd.Parameters.AddWithValue('@task', $TaskName)
    }
    $reader = $cmd.ExecuteReader()
    $table = New-Object 'System.Data.DataTable'
    if ($reader.HasRows)
    {
        $table.Load($reader)
    }
    Write-Output $table.Rows
}

function InternalGetStoredTask
{
    param
    (
        [string]$TaskName
    )
    $query =
        'SELECT [task],[objectGUID],[path],[employeeNumber],[employeeType],[msDScloudExtensionAttribute10],[departmentNumber],[department],[givenName],[initials],[manager],[physicalDeliveryOfficeName],[sn],[telephoneNumber],[title],[accountType] ' +
        'FROM dbo.LmPendingTaskView'
    if ($TaskName -ne 'All')
    {
        $query += " WHERE [task] = '$TaskName'"
    }
    $query += ' ORDER BY [sortOrder] ASC'
    $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
    $reader = $cmd.ExecuteReader()
    $table = New-Object 'System.Data.DataTable'
    if ($reader.HasRows)
    {
        $table.Load($reader)
    }
    Write-Output $table.Rows
}

