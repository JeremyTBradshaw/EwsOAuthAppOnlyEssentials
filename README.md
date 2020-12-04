# EwsOAuthAppOnlyEssentials
A PowerShell module providing functions for easy App-Only use with EWS in Exchange Online.

At this time, there is just the one function, making it easy to get an access token for EWS OAuth use, particularly in App-Only fashion using certificate credentials.  For instructions on setting up your App Registration in Azure AD, visit the link below:

https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth

With that out of the way, use this module to simplify the process of getting tokens.  For sample use cases, check out my two scripts which use EWS Managed API (2.2) and support either OAuth (pairing well with this module) or Basic authentication:

- [Get-MailboxLargeItems.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxLargeItems.ps1)
- [New-LargeItemsSearchFolder.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/New-LargeItemsSearchFolder.ps1)

## Functions

### New-EwsAccessToken

This function is used to get an access token.  Access tokens last for 1 hour, so keep this in mind in long-running scripts.

Parameters | Description
---------: | :-----------
ApplicationId | The app's ApplicationId (a.k.a. ClientId)
TenantId | The directory/tenant ID (Guid)
Certificate<br />(Option 1) | Use `$Certificate`, where `$Certificate = Get-ChildItem Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317`
CertificateStorePath<br/>(Option 2) | E.g. 'Cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
JWTExpMinutes | In case of a poorly-synced clock, use this to adjust the expiry of the JWT that is the client assertion sent in the request.  Max. value is 10.

**Example 1**

```powershell
$EwsTKParams = @{

    ApplicationId = '4ba21eca-462c-46cd-b296-9467232638a4'
    TenantId      = 'c7bdcf5c-7a22-44f0-8240-146ababc5858'
    Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\F046351F8B17FA1755F4A567C175BEA1FC86A1EC'
}
$EwsToken = New-EwsAccessToken @EwsTKParams

.\Get-MailboxLargeItems.ps1 -AccessToken $EwsToken -MailboxSmtpAddress Larry.Iceberg@jb365.ca
```
**Example 2**

```powershell
$EwsTKParams = @{

    ApplicationId        = '4ba21eca-462c-46cd-b296-9467232638a4'
    TenantId             = 'c7bdcf5c-7a22-44f0-8240-146ababc5858'
    CertificateStorePath = 'Cert:\CurrentUser\My\F046351F8B17FA1755F4A567C175BEA1FC86A1EC'
}
$EwsToken = New-EwsAccessToken @EwsTKParams

.\New-LargeItemsSearchFolder.ps1 -AccessToken $EwsToken -MailboxListCSV .\Desktop\Users.csv -Archive -WhatIf
```
