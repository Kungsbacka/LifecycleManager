# Make all errors terminating errors
$ErrorActionPreference = 'Stop'
Import-Module -Name 'ActiveDirectory'
Add-Type -Path "$PSScriptRoot\lib\Kungsbacka.CommonExtensions.dll"
Add-Type -Path "$PSScriptRoot\lib\Kungsbacka.AccountConfiguration.dll"
Add-Type -Path "$PSScriptRoot\lib\Kungsbacka.AccountTasks.dll"
Add-Type -Path "$PSScriptRoot\lib\Kungsbacka.DS.dll"
$Script:AccountNamesFactory = New-Object -TypeName 'Kungsbacka.DS.AccountNamesFactory'

function Expire-Account
{
    param
    (
        # Identity is passed unmodified to AD cmdlets
        [Alias('ObjectGuid')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        # Credentials passed on to AD cmdlets.
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    process
    {
        $params = @{
            Identity = $Identity
            Properties = @('AccountExpirationDate')
        }
        if ($null -ne $Credential)
        {
            $params.Credential = $Credential
        }
        $user = Get-ADUser @params
        # Don't touch the account if an expiration date is set and it's in the future
        if ($user.AccountExpirationDate -gt (Get-Date))
        {
            return
        }
        $params = @{
            Identity = $Identity
            DateTime = [DateTime]::Today
        }
        if ($null -ne $Credential)
        {
            $params.Credential = $Credential
        }
        Set-ADAccountExpiration @params
    }
}

function Unexpire-Account
{
    param
    (
        # Identity is passed unmodified to AD cmdlets
        [Alias('ObjectGuid')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    process
    {
        $params = @{
            Identity = $Identity
        }
        if ($null -ne $Credential)
        {
            $params.Credential = $Credential
        }
        Clear-ADAccountExpiration @params
    }
}

function Create-Account
{
    param
    (
        # Fristname
        [Alias('Firstname')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $GivenName,
        # Lastname
        [Alias('Lastname', 'Sn')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Surname,
        # Account type
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [ArgumentCompleter(
            {
                param ($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                [Enum]::GetNames([Kungsbacka.AccountConfiguration.AccountType])
            }
        )]
        [ValidateScript(
            {
                $_ -in [Enum]::GetNames([Kungsbacka.AccountConfiguration.AccountType])
            }
        )]
        [string]
        $AccountType,
        # Personnummer. Optional, but if AccountConfiguration for
        # the account type has RequireEmployeeNumber = true, the
        # function will throw if no employee number is supplied.
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $EmployeeNumber,
        # Location where the account is created in AD
        # Overrides location from AccountConfiguration
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $Path,
        # Office
        [Alias('PhysicalDeliveryOfficeName')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $Office,
        # Department
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $Department,
        # Title
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $Title,
        # Manager (DN)
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $Manager,
        # Department number
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $DepartmentNumber,
        # Mobile phone
        # employeeType attribute is used instead of mobilePhone
        # because employeeType can be flagged confidential.
        [Alias('EmployeeType')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $MobilePhone,
        # Telephone number
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $TelephoneNumber,
        # Extension attribute used for skola.
        [Alias('MsDsCloudExtensionAttribute10')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        [string]
        $Skola,
        # Account source
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet("Elevregister", "Personalsystem", "None")]
        [string]
        $AccountSource,
        # Credentials passed on to AD cmdlets.
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    process
    {
        $accountConfig = [Kungsbacka.AccountConfiguration.AccountConfiguration]::GetConfiguration($AccountType)
        $empNo = $null
        if ($accountConfig.AddBirthYearAsSamPrefix)
        {
            $empNo = $EmployeeNumber
        }
        $names =  $Script:AccountNamesFactory.GetNames(
            $GivenName,
            $Surname,
            $accountConfig.UpnSuffix,
            $empNo,
            $false,
            $accountConfig.SamPrefix,
            $accountConfig.LinkUpnAndSam
        )
        $password = [Kungsbacka.DS.PasswordGenerator]::GenerateReadablePassword()
        $params = @{
            Name = $names.CommonName
            DisplayName = $names.DisplayName
            GivenName = $names.FirstName
            Surname = $names.LastName
            SamAccountName = $names.SamAccountName
            UserPrincipalName = $names.UserPrincipalName
            AccountPassword = ($password | ConvertTo-SecureString -AsPlainText -Force)
            CannotChangePassword = $accountConfig.CannotChangePassword
            ChangePasswordAtLogon = $accountConfig.ChangePasswordAtLogon
            PasswordNeverExpires = $accountConfig.PasswordNeverExpires
            Enabled = $true
            Path = $accountConfig.DefaultLocation
            OtherAttributes = @{}
            PassThru = $true
        }
        if ($Path) # Override default location
        {
            $params.Path = $Path
        }
        if ($EmployeeNumber)
        {
            $params.EmployeeNumber = $EmployeeNumber
        }
        elseif ($accountConfig.RequireEmployeeNumber)
        {
            throw 'This account type requires an employee number, but a value for parameter EmployeeNumber was not supplied.'
        }
        if ($AccountSource -eq 'Elevregister')
        {
            $params.OtherAttributes.Add('gidNumber', 2)
        }
        elseif ($AccountSource -eq 'Personalsystem')
        {
            $params.OtherAttributes.Add('gidNumber', 2)
        }
        $tasks = $accountConfig.AccountTasks
        if ($tasks.Count -gt 0)
        {
            $json =  ConvertTo-NewtonsoftJson -InputObject $tasks
            $params.OtherAttributes.Add('carLicense', $json)
        }
        $optionalParameters = @(
            'Office'
            'Department'
            'Title'
            'Manager'
            'DepartmentNumber'
            'MobilePhone'
            'TelephoneNumber'
            'Skola'
        )
        foreach ($key in $optionalParameters)
        {
            if ($PSBoundParameters.ContainsKey($key) -and $PSBoundParameters[$key] -isnot [DBNull] -and $PSBoundParameters[$key].Length -gt 0)
            {
                $value = [string]$PSBoundParameters[$key]
                if ($key -eq 'Skola')
                {
                    $params.OtherAttributes.Add('msDS-cloudExtensionAttribute10', $value)
                }
                elseif ($key -eq 'MobilePhone')
                {
                    $params.OtherAttributes.Add('employeeType', $value)
                }
                elseif ($key -eq 'TelephoneNumber')
                {
                    $params.OtherAttributes.Add('telephoneNumber', $value)
                }
                elseif ($key -eq 'DepartmentNumber')
                {
                    $params.OtherAttributes.Add('departmentNumber', $value)
                }
                else
                {
                    $params.Add($key, $value)
                }
            }
        }
        if ($null -ne $Credential)
        {
            $params.Credential = $Credential
        }
        $newAccount = New-ADUser @params
        [pscustomobject]@{
            ObjectGuid = $newAccount.ObjectGuid
            EmployeeNumber = $EmployeeNumber
            SamAccountName = $names.SamAccountName
            UserPrincipalName = $names.UserPrincipalName
            AccountPassword = $password
            AccountType = $AccountType
            GivenName = $GivenName
            Surname = $Surname
            Department = $Department
            Office = $Office
            Title = $Title
        }
    }
}

<#
.SYNOPSIS
    Uses Set-ADUser to update an account in Active Directory
.DESCRIPTION
    Updates an account i Active Directory. The reason not giving the
    the parameters a type is because the supplied value comes from a
    database query and can contain DBNull.
    AllowNull attribute makes PSSharper happy.
.NOTES
    https://blogs.technet.microsoft.com/exchange/2005/01/10/fun-with-changing-e-mail-addresses/
#>
function Update-Account
{
    param
    (
        # Identity is passed unmodified to AD cmdlets
        [Alias('ObjectGuid')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        # Updates department attribute
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Department,
        # Updates physicalDeliveryOfficeName attribute
        [Alias('PhysicalDeliveryOfficeName')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Office,
        # Updates departmentNumber attribute
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $DepartmentNumber,
        # Updates manager attribute
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Manager,
        # Updates givenName attribute. Attributes displayName and commonName are also updated
        # and the account gets a new userPrincipalName and mail address. The old mail address is
        # stored as secondary
        [Alias('FirstName')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $GivenName,
        # Updates sn attribute. Attributes displayName and cn (common name) are also updated
        # and the account gets a new userPrincipalName and mail address. The old mail address is
        # stored as secondary
        [Alias('LastName', 'Sn')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Surname,
        # Updates initials attribute
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Initials,
        # Updates telephoneNumber attribute
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $TelephoneNumber,
        # Updates employeeType attribute. For privacy reasons this attribute contains the mobile
        # phone number. This is one of a few attributes that can be flaged as confidential in
        # Active Directory.
        [Alias('EmployeeType')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $MobilePhone,
        # Updates title attribute
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Title,
        # Moves the account to the OU specified by Path. As a side affect the cn (common name)
        # can also change to avoid a naming conflict.
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Path,
        # Updates msDS-cloudExtensionAttribute10 attribute
        [Alias('MsDsCloudExtensionAttribute10')]
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [AllowNull()]
        $Skola,
        # Do not try to move account
        [switch]
        $NoMove,
        # Do not try to rename account. This disables all updates to givenName, sn (surname),
        # displayName, cn (common name), userPrincipalName and SMTP address.
        [switch]
        $NoRename,
        # Credentials passed on to AD cmdlets.
        [pscredential]
        [System.Management.Automation.Credential()]
        [AllowNull()]
        $Credential
    )
    process
    {
        $params = @{
            Identity = $Identity
        }
        if ($null -ne $Credential)
        {
            $params.Credential = $Credential
        }
        foreach ($parameter in $PSBoundParameters.GetEnumerator())
        {
            $name = $parameter.Key
            $value = $parameter.Value
            # Renaming and moving are done as separate steps.
            if ($name -in @('Identity', 'Credential', 'GivenName', 'Surname', 'Path', 'NoRename', 'NoMove'))
            {
                continue
            }
            if ($null -eq $value -or $value -is [DBNull])
            {
                continue
            }
            if ($name -eq 'Office')
            {
                $name = 'physicalDeliveryOfficeName'
            }
            elseif ($name -eq 'MobilePhone')
            {
                $name = 'employeeType'
            }
            elseif ($name -eq 'Skola')
            {
                $name = 'msDS-cloudExtensionAttribute10'
            }
            if ($value -or $value -eq 0)
            {
                if (-not $params['Replace'])
                {
                    $params['Replace'] = @{}
                }
                $params['Replace'].Add($name, $value)
            }
            else
            {
                $params['Clear'] += @($name)
            }
        }
        if ($params.ContainsKey('Replace') -or $params.ContainsKey('Clear'))
        {
            Set-ADUser @params
        }
        # Move account
        if (-not $NoMove)
        {
            if ($Path -isnot [DBNull] -and $Path.Length -gt 0)
            {
                $params = @{
                    Identity = $Identity
                    TargetPath = $Path
                }
                if ($null -ne $Credential)
                {
                    $params.Credential = $Credential
                }
                Move-ADObject @params
            }
        }
        # Rename account if surname or given name changed
        #   - Update GivenName, Surname and DisplayName
        #   - Create new unique UPN, CN and SMTP address
        #   - Check if user already has the new SMTP address as primary or secondary and act accordingly
        if ($NoRename)
        {
            return
        }
        $givenNameChanged = $GivenName -isnot [DBNull] -and $GivenName.Length -gt 0
        $surnameChanged = $Surname -isnot [DBNull] -and $Surname.Length -gt 0
        if ($surnameChanged -or $givenNameChanged)
        {
            $params = @{
                Identity = $Identity
                Properties = @('ProxyAddresses', 'Mail', 'EmployeeNumber')
            }
            if ($null -ne $Credential)
            {
                $params.Credential = $Credential
            }
            $targetUser = Get-ADUser @params
            # Sanity check before we start renaming
            $currentPrimarySmtp = $targetUser.ProxyAddresses -clike 'SMTP:*' | ForEach-Object {$_.Substring(5)}
            $upnDomain = ($targetUser.UserPrincipalName -split '@')[1]
            if ($upnDomain -notin @('kungsbacka.se','elev.kungsbacka.se'))
            {
                Write-Error -Message "UPN domain '$upnDomain' is not handled by this cmdlet." -TargetObject $Identity
            }
            if ($null -eq $currentPrimarySmtp -or ($upnDomain -eq 'kungsbacka.se' -and $currentPrimarySmtp -ne $targetUser.UserPrincipalName))
            {
                Write-Error -Message 'Primary SMTP address is either missing och not equal to UserPrincipalName.' -TargetObject $Identity
            }
            elseif ($upnDomain -eq 'elev.kungsbacka.se' -and $currentPrimarySmtp -notlike '*@kungsbackakommun.mail.onmicrosoft.com')
            {
                Write-Error -Message 'Primary SMTP address for students should have kungsbackakommun.mail.onmicrosoft.com as domain.' -TargetObject $Identity
            }
            # Build new parameters for Set-ADUser further down...
            $params = @{
                Identity = $Identity
            }
            if ($null -ne $Credential)
            {
                $params.Credential = $Credential
            }
            # Get new names to base the renaming on
            if ($givenNameChanged)
            {
                $newGivenName = $GivenName
                $params.GivenName = [Kungsbacka.DS.AccountNamesFactory]::GetName($GivenName)
            }
            else
            {
                $newGivenName = $targetUser.GivenName
            }
            if ($surnameChanged)
            {
                $newSurname = $Surname
                $params.Surname = [Kungsbacka.DS.AccountNamesFactory]::GetName($Surname)
            }
            else
            {
                $newSurname = $targetUser.Surname
            }
            if ($upnDomain -eq 'kungsbacka.se')
            {
                try
                {
                    $names = $Script:AccountNamesFactory.GetNames($newGivenName, $newSurname, $upnDomain, $true)
                }
                catch
                {
                    Write-Error -Exception $_.Exception -TargetObject $Identity
                }
            }
            else # elev.kungsbacka.se
            {
                try
                {
                    $names = $Script:AccountNamesFactory.GetNames($newGivenName, $newSurname, $upnDomain, $targetUser.EmployeeNumber, $true)
                }
                catch
                {
                    Write-Error -Exception $_.Exception -TargetObject $Identity
                }
            }
            $params.DisplayName = $names.DisplayName
            $useNewNames = $false
            $newUserPrincipalName = $null
            $newSmtpWithoutSuffix = [Kungsbacka.DS.AccountNamesFactory]::GetUpnNamePart($newGivenName, $newSurname)
            # Select first address that has the same base name (name w/o suffix) as the new address.
            $existingSmtp = $targetUser.ProxyAddresses -match "^smtp:$newSmtpWithoutSuffix[^d]*@$upnDomain`$" |
                Sort-Object | Select-Object -First 1
            if ($upnDomain -eq 'kungsbacka.se')
            {
                if ($null -ne $existingSmtp)
                {
                    # Address is in proxyAddresses, but is secondary -> promote to primary
                    if ($existingSmtp -clike 'smtp:*')
                    {
                        $smtp = $existingSmtp.Substring(5)
                        $params.Remove = @{ProxyAddresses = @(
                            'SMTP:' + $currentPrimarySmtp
                            'smtp:' + $smtp
                        )}
                        $params.Add = @{ProxyAddresses = @(
                            'smtp:' + $currentPrimarySmtp
                            'SMTP:' + $smtp
                        )}
                        $params.UserPrincipalName = $smtp
                        $params.EmailAddress = $smtp
                        $newUserPrincipalName = $smtp
                    }
                }
                else
                {
                    # Address is not in proxyAddresses -> add new address and demote primary to secondary
                    $params.Remove = @{ProxyAddresses = 'SMTP:' + $currentPrimarySmtp}
                    $params.Add = @{ProxyAddresses = @(
                        'smtp:' + $currentPrimarySmtp
                        'SMTP:' + $names.UserPrincipalName
                    )}
                    $params.UserPrincipalName = $names.UserPrincipalName
                    $params.EmailAddress = $names.UserPrincipalName
                    $newUserPrincipalName = $names.UserPrincipalName
                    $useNewNames = $true
                }
            }
            else # elev.kungsbacka.se
            {
                if ($null -ne $existingSmtp)
                {
                    # elev.kungsbacka.se exists in Exchange Online and the primary SMTP address
                    # is samaccountname@kungsbackakommun.mail.onmicrosoft.com. This doesn't change
                    # when the user gets a new SMTP address.
                    $smtp = $existingSmtp.Substring(5)
                    if ($smtp -ne $targetUser.UserPrincipalName)
                    {
                        $params.UserPrincipalName = $smtp
                        $params.EmailAddress = $smtp
                        $newUserPrincipalName = $smtp
                    }
                }
                else
                {
                    $params.Add = @{ProxyAddresses = @('smtp:' + $names.UserPrincipalName)}
                    $params.UserPrincipalName = $names.UserPrincipalName
                    $params.EmailAddress = $names.UserPrincipalName
                    $newUserPrincipalName = $names.UserPrincipalName
                    $useNewNames = $true
                }
            }
            Set-ADUser @params
            # Moving on to common name...
            $params = @{
                Identity = $Identity
                NewName = $names.CommonName
            }
            if ($null -ne $Credential)
            {
                $params.Credential = $Credential
            }
            if (-not $useNewNames)
            {
                $newCommonNameWithoutSuffix = [Kungsbacka.DS.AccountNamesFactory]::GetCommonName($newGivenName, $newSurname)
                $conflictingNames = Get-ADUser -Filter "cn -eq '$newCommonNameWithoutSuffix'"
                if ($null -eq $conflictingNames)
                {
                    $params.NewName = $newCommonNameWithoutSuffix
                }
            }
            Rename-ADObject @params
            if ($newUserPrincipalName)
            {
                [pscustomobject]@{
                    ObjectGuid = $targetUser.ObjectGuid
                    UserPrincipalName = $newUserPrincipalName
                    EmployeeNumber = $targetUser.EmployeeNumber
                } | Write-Output
            }
        }
    }
}

function Move-Account
{
    param
    (
        [Alias('ObjectGuid')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Path,
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    $params = @{
        Identity = $Identity
        TargetPath = $Path
    }
    if ($null -ne $Credential)
    {
        $params.Credential = $Credential
    }
    Move-ADObject @params
}

function Delete-Account
{
    param
    (
        [Alias('ObjectGuid')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    $params = @{
        Identity = $Identity
        Properties = @('AccountExpirationDate', 'EmployeeNumber')
    }
    if ($null -ne $Credential)
    {
        $params.Credential = $Credential
    }
    $user = Get-ADUser @params
    if ((Get-Date).AddDays(-90) -lt $user.AccountExpirationDate)
    {
        Write-Error -Message 'Account expiration date is less than 90 days ago.' -TargetObject $Identity
    }
    $params = @{
        Identity = $Identity
        Recursive = $true
        Confirm = $false
    }
    if ($null -ne $Credential)
    {
        $params.Credential = $Credential
    }
    Remove-ADObject @params
}

function Remove-MsolLicense
{
    param
    (
        # Identity is passed unmodified to AD cmdlets
        [Alias('ObjectGuid')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        # Credentials passed on to AD cmdlets.
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    process
    {
        $userInfo = Get-UserInfo -Identity $Identity -Credential $Credential
        if ($userInfo.MailEnabled)
        {
            $task = New-Object 'Kungsbacka.AccountTasks.MicrosoftOnlinePostExpireTask'
        }
        else
        {
            $task = New-Object 'Kungsbacka.AccountTasks.MsolRemoveAllLicenseGroupTask'
        }
        $params = @{
            Identity = $Identity
            Replace = @{'carLicense'="[$($task.ToJson())]"}
        }
        if ($null -ne $Credential)
        {
            $params.Credential = $Credential
        }
        Set-ADUser @params
    }
}

function Restore-MsolLicense
{
    param
    (
        # Identity is passed unmodified to AD cmdlets
        [Alias('ObjectGuid')]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Identity,
        # Credentials passed on to AD cmdlets.
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    process
    {
        $userInfo = Get-UserInfo -Identity $Identity -Credential $Credential
        if ($userInfo.MailEnabled)
        {
            switch ($userInfo.AccountType)
            {
                'Elev'         { $mailboxType = 'Student'  }
                'Skolpersonal' { $mailboxType = 'Faculty'  }
                default        { $mailboxType = 'Employee' }
            }
            $task = New-Object 'Kungsbacka.AccountTasks.MicrosoftOnlineRestoreTask' -ArgumentList @($mailboxType)
        }
        else
        {
            $task = New-Object 'Kungsbacka.AccountTasks.MsolRestoreLicenseGroupTask'
        }
        $params = @{
            Identity = $Identity
            Replace = @{'carLicense'="[$($task.ToJson())]"}
        }
        if ($null -ne $Credential)
        {
            $params.Credential = $Credential
        }
        Set-ADUser @params
    }
}