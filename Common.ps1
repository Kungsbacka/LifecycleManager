# "Popular high-performance JSON framework for .NET"
# http://www.newtonsoft.com/json
Add-Type -Path "$PSScriptRoot\lib\Newtonsoft.Json.dll"
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
        Properties = @('ExtensionAttribute11','MailNickname','MsExchRemoteRecipientType','ProxyAddresses', 'Department','MemberOf')
    }
    if ($null -ne $Credential)
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
    foreach($group in $user.MemberOf)
    {
        if ($group -like 'CN=G-Licens-A-*')
        {
            $licenseGroup = [Kungsbacka.DS.DSFactory]::GetLicenseGroup($group)
            break
        }
    }
    $mailEnabledLicense = ($null -ne $licenseGroup -and $licenseGroup.MailEnabled)
    $mailEnabledUser = ($user.MailNickname -and $user.ExtensionAttribute11 -and $user.MsExchRemoteRecipientType -and $user.ProxyAddresses -like 'SMTP:*')
    $userInfo = [pscustomobject]@{
        MailEnabled = ($mailEnabledLicense -and $mailEnabledUser)
        AccountType =
            if ($user.UserPrincipalName -like '*@elev.kungsbacka.se') {'Elev'}
            elseif ($user.Department -in ('Förskola & Grundskola', 'Gymnasium & Arbetsmarknad')) {'Skolpersonal'}
            else {'Personal'}
    }
    Write-Output -InputObject $userInfo
}
