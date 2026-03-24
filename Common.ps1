Add-Type -Path "$PSScriptRoot\lib\Kungsbacka.DS.dll"
. "$PSScriptRoot\Config.ps1"

function ConvertTo-NewtonsoftJson
{
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [object]
        $InputObject,
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=1)]
        [Newtonsoft.Json.Formatting]
        $Formatting = [Newtonsoft.Json.Formatting]::None
    )
    [Newtonsoft.Json.JsonConvert]::SerializeObject($InputObject, $Formatting) | Write-Output
}

function Invoke-StoredProcedure
{
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateSet('MetaDirectory', 'ADEvents')]
        [string]
        $Database,
        [Parameter(Mandatory=$true)]
        [string]
        $Procedure,
        [object]
        $Parameters,
        [switch]
        $Scalar
    )
    $cmd = Get-SqlCommand -Database $Database -CommandType StoredProcedure -CommandText $Procedure
    if ($Parameters)
    {
        foreach ($key in $Parameters.Keys)
        {
            [void]$cmd.Parameters.AddWithValue($key, $Parameters[$key])
        }
    }
    if ($Scalar)
    {
        $cmd.ExecuteScalar()
    }
    else
    {
        [void]$cmd.ExecuteNonQuery()
    }
}

function Get-SqlCommand
{
    param
    (
        # Database for connection
        [Parameter(Mandatory=$true)]
        [ValidateSet('MetaDirectory', 'ADEvents')]
        [string]
        $Database,
        # Command text
        [Parameter(Mandatory=$true)]
        [Alias('Text')]
        $CommandText,
        # Command type (Text is default)
        [Alias('Type')]
        [System.Data.CommandType]
        $CommandType = [System.Data.CommandType]::Text
    )
    $conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
    if ($Database -eq 'MetaDirectory')
    {
        $conn.ConnectionString = $Script:Config.MetaDirectoryConnectionString
    }
    elseif ($Database -eq 'ADEvents')
    {
        $conn.ConnectionString = $Script:Config.ADEventsConnectionString
    }
    $conn.Open()
    $cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
    $cmd.Connection = $conn
    $cmd.CommandType = $CommandType
    $cmd.CommandText = $CommandText
    $cmd | Write-Output
}

function Write-LmEventLog
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [System.Management.Automation.ErrorRecord]
        $ErrorRecord
    )
    $param = @{
        LogName = 'Application'
        Source = 'LifecycleManager'
        EventId = 1
        EntryType = 'Error'
    }
    $sb = New-Object -TypeName 'System.Text.StringBuilder'
    [void]$sb.AppendLine($ErrorRecord.Exception.Message)
    [void]$sb.AppendLine()
    if ($ErrorRecord.TargetObject)
    {
        [void]$sb.Append("Target object: ")
        [void]$sb.AppendLine($ErrorRecord.TargetObject.ToString())
        [void]$sb.AppendLine()
    }
    [void]$sb.AppendLine("Script stack trace:")
    [void]$sb.AppendLine($ErrorRecord.ScriptStackTrace)
    $param.Message = $sb.ToString()
    Write-EventLog @param
}

function Get-LicenseGroup
{
    param (
        [Parameter(Mandatory=$true, ParameterSetName='ByDistinguishedName')]
        [string]
        $DistinguishedName,
        [Parameter(Mandatory=$true, ParameterSetName='ByGuid')]
        [Guid]
        $Guid,
        [Parameter(Mandatory=$true, ParameterSetName='All')]
        [switch]
        $All
    )
    if ($Script:CachedLicenseGroupsByDn -and $Script:CachedLicenseGroupsByGuid)
    {
        $groupsByDn = $Script:CachedLicenseGroupsByDn
        $groupsByGuid = $Script:CachedLicenseGroupsByGuid
    }
    else
    {
        $groupsFromAd = [Kungsbacka.DS.DSFactory]::GetLicenseGroups()
        $groupsByDn = @{}
        $groupsByGuid = @{}
        foreach ($group in $groupsFromAd)
        {
            $groupsByDn[$group.DistinguishedName] = $group
            $groupsByGuid[$group.Guid] = $group
        }
        $Script:CachedLicenseGroupsByDn = $groupsByDn
        $Script:CachedLicenseGroupsByGuid = $groupsByGuid
    }

    if ($PSBoundParameters.ContainsKey('DistinguishedName'))
    {
        Write-Output -InputObject $groupsByDn[$DistinguishedName]
        return
    }

    if ($PSBoundParameters.ContainsKey('Guid'))
    {
        Write-Output -InputObject $groupsByGuid[$Guid]
        return
    }

    Write-Output -InputObject $groupsByDn.Values
}

function Get-UserInfo
{
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    $params = @{
        Identity = $Identity
        Properties = @(
            'ExtensionAttribute11'
            'MailNickname'
            'MsExchRemoteRecipientType'
            'ProxyAddresses'
            'Department'
            'MemberOf'
            'msDS-cloudExtensionAttribute1'
        )
    }
    if ($Credential)
    {
        $params.Credential = $Credential
    }
    try
    {
        $user = Get-ADUser @params -ErrorAction 'Stop'
    }
    catch
    {
        return
    }
    $licenseGroup = $null
    foreach ($group in $user.MemberOf)
    {
        if ($group -like 'CN=G-Licens-A-*')
        {
            $licenseGroup = Get-LicenseGroup -DistinguishedName $group
            break
        }
    }
    $isLicensed = $null -ne $licenseGroup

    if ($null -eq $licenseGroup -and $user.'msDS-cloudExtensionAttribute1')
    {
        $haveStashedLicense = $true
        $temp = $user.'msDS-cloudExtensionAttribute1' -split ','
        foreach ($guid in $temp)
        {
            $licenseGroup = Get-LicenseGroup -Guid $guid
            if ($licenseGroup -and $licenseGroup.Category -eq 'A')
            {
                break
            }
        }
    }
    $mailEnabledLicense = ($licenseGroup -and $licenseGroup.MailEnabled)
    $mailEnabledUser = ($user.MailNickname -and $user.ExtensionAttribute11 -and $user.MsExchRemoteRecipientType -and $user.ProxyAddresses -like 'SMTP:*')
    $userInfo = [pscustomobject]@{
        MailEnabled = ($mailEnabledLicense -and $mailEnabledUser)
        AccountType =
            if ($user.UserPrincipalName -like '*@elev.kungsbacka.se') {'Elev'}
            elseif ($user.Department -in ('Förskola & Grundskola', 'Gymnasium & Arbetsmarknad')) {'Skolpersonal'}
            else {'Personal'}
        Licensed = $isLicensed
        Stashed = $haveStashedLicense
    }
    Write-Output -InputObject $userInfo
}

class HashtableComparer : System.Collections.Generic.IEqualityComparer[hashtable] {
    [bool] Equals([hashtable]$x, [hashtable]$y) {
        if ($x.Count -ne $y.Count) { return $false }
        foreach ($key in $x.Keys) {
            if (-not $y.ContainsKey($key)) { return $false }
            if ($x[$key] -ne $y[$key]) { return $false }
        }
        return $true
    }

    [int] GetHashCode([hashtable]$obj) {
        [long]$hash = 17
        foreach ($key in $obj.Keys | Sort-Object) {
            $hash = $hash * 31 + $key.GetHashCode()
            if ($null -ne $obj[$key]) {
                $hash = $hash * 31 + $obj[$key].GetHashCode()
            }
        }
        return [int]($hash -band 0x7FFFFFFF)
    }
}
