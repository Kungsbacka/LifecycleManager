# Configuration
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
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
    # Recipient that gets a report with all new accounts (can be omitted)
    AllAccountsRecipient = ''
    # Used to determine which accounts belongs to employees
    EmployeeUpnSuffix = 'example.com'
    
    # Configuration for limits and included account types are now moved to the database
}
