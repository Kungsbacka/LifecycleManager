﻿# Configuration
$Config = @{
    # See Secure.ps1 for more details
    EncryptionKey = ''
    # This is where new accounts are stored. Reporting also relies on this database.
    MetaDirectoryConnectionString = ''
    # All logging is done to this database. Critical errors are also written to the event log.
    ADEventsConnectionString = ''
    # Server that runs the SSIS package for Active Directory import
    SqlAgentServer = ''
    # Name of the SQL Agent job for Active Directory import
    SqlAgentJob = 'Import Active Directory'
    # Configuration used for sending reports
    SmtpServer = ''
    SmtpFrom = ''
    SmtpSubject = ''
    SmtpBody = ''
    # Used to determine which accounts belongs to employees
    EmployeeUpnSuffix = 'example.com'
    # Turns limits on or off
    DisableLimits = $false
    # Sets a limit on how many tasks is processed each run. This limits the effect of bad input.
    Limits = @{
        Create = 100
        Expire = 50
        Unexpire = 50
        RemoveLicense = 50
        RestoreLicense = 50
        Delete = 50
        Update = 500
        Move = 100
    }
}
