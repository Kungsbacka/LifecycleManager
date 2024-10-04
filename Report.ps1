# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\Common.ps1"

# "EPPlus is a .net library that reads and writes Excel 2007/2010 files using the Open Office Xml format (xlsx)."
# http://epplus.codeplex.com/
Add-Type -Path "$PSScriptRoot\lib\EPPlus.dll"

function Send-NewAccountReport
{
    param
    (
        # Override recipients for all reports
        [Parameter()]
        [string]
        $Recipient,
        # If present, reported accounts are not marked as reported in the database
        [Parameter()]
        [switch]
        $DoNotMarkReported
    )
    
    # To be able to run this function as a gMSA where the environment can be restricted,
    # we write the attachments to a location we know exists and should be writable.
    if ($Script:PSScriptRoot)
    {
        $reportPath = $Script:PSScriptRoot
    }
    else
    {
        $reportPath = Convert-Path -Path .
    }
    $reportPath = Join-Path -Path $reportPath -ChildPath 'elevkonton.xlsx'
    $query = 'SELECT * FROM dbo.LmNewAccountView'
    $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -CommandText $query
    $reader = $cmd.ExecuteReader()
    $table = New-Object 'System.Data.DataTable'
    if (-not $reader.HasRows)
    {
        return
    }
    $table.Load($reader)
    # Use SmtpClient instead of Send-MailMessage since the latter
    # always tries to authenticate with default credentials. A gMSA is
    # not allowed to authenticate to our Exchange SMTP receive connector.
    $smtpClient = New-Object -TypeName 'System.Net.Mail.SmtpClient'
    $smtpClient.UseDefaultCredentials = $false
    $smtpClient.Host = $Script:Config.SmtpServer
    $table.Rows | Group-Object -Property Mottagare | ForEach-Object -Process {
        $_.Group | Export-NewAccountReport -Path $reportPath
        $msg = New-Object -TypeName 'System.Net.Mail.MailMessage'
        $msg.BodyEncoding = [System.Text.Encoding]::UTF8
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
        $msg.From = $Script:Config.SmtpFrom
        $msg.Subject = $Script:Config.SmtpSubject
        $msg.Body = $Script:Config.SmtpBody
        $attachment = New-Object 'System.Net.Mail.Attachment' -ArgumentList @($reportPath)
        $msg.Attachments.Add($attachment)
        if ($Recipient)
        {
            $msg.To.Add($Recipient)
        }
        else
        {
            $_.Name -split ';' | ForEach-Object -Process {
                $msg.To.Add($_)
            }
        }
        $smtpClient.Send($msg)
        $attachment.Dispose()
        $msg.Dispose()
    }
    if (-not $DoNotMarkReported)
    {
        $query = 'UPDATE dbo.LmNewAccount SET reported = 1 WHERE reported = 0'
        $cmd = Get-SqlCommand -Database MetaDirectory -Type Text -CommandText $query
        [void]$cmd.ExecuteNonQuery()
    }
    Remove-Item -Path $reportPath -Confirm:$false
}

function Export-NewAccountReport
{
    # False positive. $i is both declared *and* used
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
    param
    (
        # Object array to export
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [object[]]
        $InputObject,
        [Parameter(Mandatory=$true)]
        $Path
    )

    begin
    {
        $fileInfo = [System.IO.FileInfo]($Path)
        if ($fileInfo.Exists)
        {
            $fileInfo.Delete()
        }
        $package = [OfficeOpenXml.ExcelPackage]($fileInfo)
        $worksheet = $package.Workbook.Worksheets.Add("Nya konton")
        @(
            'Förnamn'
            'Efternamn'
            'Skola, klass'
            'Kontonamn'
            'E-postadress'
            'Lösenord'
        ) | ForEach-Object -Begin {$i = 1} -Process {$worksheet.Cells[1, $i++].Value = $_}
        $worksheet.Cells[1, 1, 1, $worksheet.Cells.Columns].Style.Font.Bold = $true
        $row = 1
    }

    process
    {
        foreach ($item in $InputObject)
        {
            @(
                $($item['Förnamn'])
                $($item['Efternamn'])
                $($item['Skola & klass'])
                $($item['Kontonamn'])
                $($item['E-postadress'])
                $($item['Lösenord'])
            ) | ForEach-Object -Begin {$row++; $i = 1} -Process {$worksheet.Cells[$row, $i++].Value = $_}
        }
    }

    end
    {
        $worksheet.Cells.AutoFitColumns()
        $package.Save()
        $worksheet.Dispose()
        $package.Dispose()
    }
}
