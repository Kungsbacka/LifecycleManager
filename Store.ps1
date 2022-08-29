# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

function Store-NewAccount
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseApprovedVerbs", "")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
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

function Get-Limits
{
    if ($null -ne $Script:Config.Limits -and $null -ne $Script:Config.DisableLimits) {
        $limits = InternalGetLimitsFromConfig
        $source = 'configuration file'
    }
    else {
        $limits = InternalGetLimitsFromDatabase
        $source = 'database'
    }
    $mandatory = @('DisableLimits', 'Create', 'Expire', 'Unexpire', 'RemoveLicense', 'RestoreLicense', 'Delete', 'Update', 'Move')
    foreach ($name in $mandatory) {
        if ($null -eq $limits[$name] -or ($limits[$name] -is [int] -and $limits[$name] -eq 0)) {
            throw "Configuration for limits is not complete in $source. Missing or misconfigured entry: $name"
            return
        }
    }
    Write-Output $limits
}

function Enable-Limits
{
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text "UPDATE dbo.LmConfig SET [value]='false' WHERE [name]='Limit.DisableLimits'"
        $null = $cmd.ExecuteNonQuery()
    }
    finally {
        if ($cmd) {
            $cmd.Dispose()
        }
    }
}


function InternalGetLimitsFromConfig
{
    $limits = $Script:Config.Limits
    $limits['DisableLimits'] = $Script:Config.DisableLimits
    Write-Output $limits
}

function InternalGetLimitsFromDatabase
{
    $query = "SELECT [name],[value] FROM dbo.LmConfig WHERE [name] LIKE 'Limit.%'"
    $limits = @{
        DisableLimits = $null
        Create = 0
        Expire = 0
        Unexpire = 0
        RemoveLicense = 0
        RestoreLicense = 0
        Delete = 0
        Update = 0
        Move = 0
    }
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $name = $reader['name']
            $value = $reader['value']
            $name = $name.Split('.')[1]
            if ($null -eq $name) {
                throw 'Invalid configuration entry in database.'
                return
            }
            if ($name -eq 'DisableLimits') {
                $limits.DisableLimits = [bool]::Parse($value)
            }
            else {
                $limits[$name] = [int]$value
            }
        }
    }
    finally {
        if ($cmd) {
            $cmd.Dispose()
        }
        if ($reader) {
            $reader.Dispose()
        }
    }
    Write-Output $limits
}

function InternalGetCurrentTask
{
    param
    (
        [string]$TaskName
    )
    $query = 'SELECT * FROM dbo.ufLmGetPendingLifecycleTask(@task, 1)' # 1 = Do not include renames
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
        $cmd.CommandTimeout = 120
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
    finally {
        if ($cmd) {
            $cmd.Dispose()
        }
        if ($reader) {
            $reader.Dispose()
        }
    }
}

function InternalGetStoredTask
{
    param
    (
        [string]$TaskName
    )
    $query =
        'SELECT [task],[objectGUID],[path],[employeeNumber],[employeeType],[msDScloudExtensionAttribute9],[msDScloudExtensionAttribute10],[departmentNumber],[department],[givenName],[initials],[manager],[physicalDeliveryOfficeName],[sn],[telephoneNumber],[title],[accountType] ' +
        'FROM dbo.LmPendingTaskView'
    if ($TaskName -ne 'All')
    {
        $query += " WHERE [task] = '$TaskName'"
    }
    $query += ' ORDER BY [sortOrder] ASC'
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
        $reader = $cmd.ExecuteReader()
        $table = New-Object 'System.Data.DataTable'
        if ($reader.HasRows)
        {
            $table.Load($reader)
        }
        Write-Output $table.Rows
    }
    finally {
        if ($cmd) {
            $cmd.Dispose()
        }
        if ($reader) {
            $reader.Dispose()
        }
    }
}

