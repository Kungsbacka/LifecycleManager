# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"

function Start-ActiveDirectoryImportJob
{
    [void][reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")
    [void][reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo")
    $sqlServer = New-Object -TypeName 'Microsoft.SqlServer.Management.Smo.Server' -ArgumentList @($Script:Config.SqlAgentServer)
    $sqlServer.JobServer.Jobs[$Script:Config.SqlAgentJob].Start()
}
