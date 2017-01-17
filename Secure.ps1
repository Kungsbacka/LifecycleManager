<#
    This script uses AES in combination with DAPI to protect data. To use
    the script you first need to generate a new AES key and store it in
    a safe place. You then take a copy of the key and encrypt it using DAPI
    as the user who is going to run the script. Finally you store the encrypted
    key in the configuration file (Config.ps1).
    
    AES key
    -------
    The key is 32 bytes and can be generated using the steps below. If a new
    key is generated, all data enctypted with the old key must be re-encrypted
    with the new key before the old key is discarded. Use Update-PasswordEncryptionKey
    to re-encrypt all passwords in the database.
    
    1. Generate key:
       $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
       $key = New-Object byte[](32)
       $rng.GetBytes($key)
        
    2. Convert to base64:
       [System.Convert]::ToBase64String($key)
    
    3. Store base64 encoded key in a safe place.
    
    4. Encrypt the key as the user (service account) that is going to run the
       script. Start a PowerShell prompt as the user and run the code below:
       $key = [System.Convert]::FromBase64String('[base64 encoded key]')
       $keyString = [System.Text.Encoding]::Unicode.GetString($key)
       $secureKey = ConvertTo-SecureString -String $keyString -AsPlainText -Force
       $secureKey | ConvertFrom-SecureString

    5. Copy the result to the configuration file (Config.ps1)
#>

# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\Common.ps1"

function Protect-String
{
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [string]
        $PlainTextString
    )
    $secureKey = ConvertTo-SecureString -String $Script:Config.EncryptionKey
    $secureString = ConvertTo-SecureString -String $PlainTextString -AsPlainText -Force
    ConvertFrom-SecureString -SecureString $secureString -SecureKey $SecureKey
}

function Unprotect-String
{
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [string]
        $EncryptedString
    )
    $secureKey = ConvertTo-SecureString -String $Script:Config.EncryptionKey
    $secureString = ConvertTo-SecureString -String $EncryptedString -SecureKey $secureKey
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) 
}

# Follow the steps at the top of this file to generate a new base64 encoded
# encryption key. Supply the old and the new key to this function to first
# decrypt and then encrypt all passwords in the database. The function uses
# a transaction to make sure all passwords gets re-encrypted with the new key
# before commiting the changes to the database.
function Update-PasswordEncryptionKey
{
    param
    (
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory, Position = 0)]
        [string]
        $OldBase64Key,
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory, Position = 1)]
        [string]
        $NewBase64Key
    )
    $bytes = [Convert]::FromBase64String($OldBase64Key)
    $keyString = [System.Text.Encoding]::Unicode.GetString($bytes)
    $secureOldKey = ConvertTo-SecureString -String $keyString -AsPlainText -Force
    $bytes = [Convert]::FromBase64String($NewBase64Key)
    $keyString = [System.Text.Encoding]::Unicode.GetString($bytes)
    $secureNewKey = ConvertTo-SecureString -String $keyString -AsPlainText -Force
    $cmd = Get-SqlCommand -Database MetaDirectory -CommandText 'SELECT id,accountPassword FROM dbo.LmNewAccount WHERE accountPassword IS NOT NULL'
    $rst = $cmd.ExecuteReader()
    $data = [System.Collections.ArrayList]@()
    while ($rst.Read())
    {
        [void]$data.Add([pscustomobject]@{
            Id = $rst['id']
            Password = $rst['accountPassword']
        })
    }
    $rst.Dispose()
    $cmd.Dispose()
    $cmd = Get-SqlCommand -Database MetaDirectory -CommandText 'UPDATE dbo.LmNewAccount SET accountPassword=@pwd WHERE id=@id'
    $cmd.Transaction = $cmd.Connection.BeginTransaction()
    foreach ($item in $data)
    {
        # Decrypt with old key
        $secureString = ConvertTo-SecureString -String $item.Password -SecureKey $secureOldKey -ErrorAction Stop
        # Encrypt with new key
        $newEncryptedPassword = ConvertFrom-SecureString -SecureString $secureString -SecureKey $secureNewKey -ErrorAction Stop
        # Store changes
        $cmd.Parameters.Clear()
        [void]$cmd.Parameters.AddWithValue('@pwd', $newEncryptedPassword)
        [void]$cmd.Parameters.AddWithValue('@id', $item.Id)
        [void]$cmd.ExecuteNonQuery()
    }
    $cmd.Transaction.Commit()
}
