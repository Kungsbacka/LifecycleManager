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
        [Parameter(Mandatory=$false)]
        [ValidateSet('Expire','Unexpire','Update','Create','Delete','Move','RemoveLicense','RestoreLicense','ChangeLicense')]
        [string[]]
        $TaskName,
        [Parameter(Mandatory=$false)]
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
        if (-not $TaskName) {
            $TaskName = @('Expire','Unexpire','Update','Create','Delete','Move','RemoveLicense','RestoreLicense','ChangeLicense')
        }
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

function Get-IncludedAccountTypes
{
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text "SELECT [value] FROM dbo.LmConfig WHERE [name]='IncludedAccountTypes'"
        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $accountTypes = $rdr.GetString(0)
        }
        $accountTypes -split ','
    }
    finally {
        if ($rdr) {
            $rdr.Dispose()
        }
        if ($cmd) {
            $cmd.Dispose()
        }
    }
}

function Get-AccountTypes
{
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text "SELECT [accountType] FROM dbo.LmAccountConfiguration"
        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $rdr.GetString(0)
        }
    }
    finally {
        if ($rdr) {
            $rdr.Dispose()
        }
        if ($cmd) {
            $cmd.Dispose()
        }
    }
}

function Get-AccountConfiguration
{
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [ArgumentCompleter(
            {
                param ($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                Get-AccountTypes
            }
        )]
        [ValidateScript({$_ -in (Get-AccountTypes)})]
        [string]
        $AccountType
   )

    $columns = @(
        'accountType'
        'accountTypeDescription'
        'upnSuffix'
        'defaultLocation'
        'accountNamePrefix'
        'accountTasks'
        'hasMail'
        'isManaged'
        'requireEmployeeNumber'
        'requireDescription'
        'requirePhoneNumber'
        'requireDepartment'
        'requireOffice'
        'requireEmailAddress'
        'requireCompany'
        'cannotChangePassword'
        'changePasswordAtLogon'
        'passwordNeverExpires'
        'expiresOnCreation'
        'linkUpnAndSam'
        'addBirthYearAsSamPrefix'
        'defaultGroups'
    )

    $query = "SELECT [$($columns -join '],[')] FROM dbo.LmAccountConfiguration WHERE [accountType] = @accountType"

    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
        [void]$cmd.Parameters.AddWithValue('@accountType', $AccountType.ToString())
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $h = @{}
            foreach ($c in $columns) {
                $key = $c.Substring(0, 1).ToUpper() + $c.Substring(1)
                if ($reader[$c] -is [DBNull]) {
                    $h[$key] = $null
                }
                else {
                    $h[$key] = $reader[$c]
                }
            }
            return [pscustomobject]$h
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
        [string[]]$TaskName
    )
    $query = 'SELECT * FROM dbo.ufLmGetPendingLifecycleTask(@task, 1) ORDER BY taskPriority ASC' # 1 = Do not include name changes
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
        $taskTable = New-Object 'System.Data.DataTable'
        [void]$taskTable.Columns.Add('value', [string])
        foreach ($name in $TaskName)
        {
            [void]$taskTable.Rows.Add($name)
        }
        $param = $cmd.Parameters.Add('@task', [System.Data.SqlDbType]::Structured)
        $param.TypeName = 'dbo.NvarcharTable'
        $param.Value = $taskTable
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
        [string[]]$TaskName
    )
    $query = 'SELECT a.* FROM dbo.LmPendingTaskView a INNER JOIN @task b ON a.task = b.value ORDER BY taskPriority ASC'
    try {
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -Text $query
        $taskTable = New-Object 'System.Data.DataTable'
        [void]$taskTable.Columns.Add('value', [string])        
        foreach ($name in $TaskName)
        {
            [void]$taskTable.Rows.Add($name)
        }
        $param = $cmd.Parameters.Add('@task', [System.Data.SqlDbType]::Structured)
        $param.TypeName = 'dbo.NvarcharTable'
        $param.Value = $taskTable
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
