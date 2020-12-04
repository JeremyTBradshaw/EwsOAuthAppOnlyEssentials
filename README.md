# EwsOAuthAppOnlyEssentials
A PowerShell module providing functions for easy App-Only use with EWS in Exchange Online.

At this time, there are just two functions, making it easy to get an access token for EWS OAuth use, particularly in App-Only fashion using certificate credentials.  I plan to soon port the functions `Add-MSGraphApplicationKeyCredential` and `Remove-MSGraphApplicationKeyCredential` over to this module.  Beyond that, I'm debating whether or not to distribute the EWS Managed API 2.2 DLL file with this module, and then start including some extra functions (e.g. `New-EwsClient`).

For instructions on setting up your App Registration in Azure AD, visit the link below:

https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth

If you're interested in learning about 'certificate credentials', you can get a _brief_ description [here](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials).  That page describes the basics of a client assertion, in the form of a JSON Web Token.  How that client assertion is packaged for delivery to Azure AD when requesting an access token is a more advanced topic.  With the `New-EwsOAuthAccessToken` function, and also with my [other module's](https://github.com/JeremyTBradshaw/MSGraphAppOnlyEssentials) function `New-MSGraphAccessToken`, I'm mimicking in PowerShell manually what is normally being done behind closed doors with the Microsoft Identity Client libraries.  That is, I'm performing the same steps of using the *Microsoft.IdentityModel.JsonWebTokens* library to create a signed assertion for use with the *.WithSignedAssertion()* method of the *ConfidentialClientApplicationBuilder* class.  To see an official Microsoft example of doing the same, but using C#, check out [this awesome, official wiki page](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Client-Assertions).

*With that out of the way*, use this module to simplify the process of getting tokens.  For sample use cases, check out my two scripts which use EWS Managed API (2.2) and support either OAuth (pairing well with this module) or Basic authentication:

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

### New-SelfSignedEwsOAuthApplicationCertificate

This function is simply a wrapper for New-SelfSignedCertificate with some base settings to ensure the certificate will work for App-Only authentication with Azure AD.  It uses the Microsoft Enhanced RSA and AES Cryptographic Provider and the SHA-256 hashing algorithm, ensuring the certificate will be able to do everything it might need with the other functions.  The other functions in this module (current/future) also insist on this provider.  I will soon begin enforcing the SHA-256 hashing algorigthm in the `New-EwsAccessToken` function (as well as the `New-MSGraphAccessToken` function from MSGraphAppOnlyEssentials).

Parameters | Description
---------: | :----------
DnsName | Any FQDN of choice.  E.g. 20201204.ewsclient.jb365.ca
FriendlyName | "jb365 EWS Client 2020-12-04"
CertStoreLocation | Maps directly to the same parameter of New-SelfSignedCertificate.  Default is 'cert:\CurrentUser\My'.  Any valid location where write access is available will work (e.g. 'cert:\LocalMachine\My', when scheduling tasks using local SYSTEM account).
NotAfter | Default is 90 days.  Supply a [datetime] like this `(Get-Date).AddDays(7)`.  The shorter the better, because applications can add new certificates for themselves using the addKey method, so we can easily roll these often, programmatically.
KeySpec | **Signature**, KeyExchange.  Recommendation: don't change this unless there is a reason.

**Example 1**

```powershell
New-SelfSignedEwsOAuthApplicationCertificate -DnsName "ewsclient.jb365.ca" -FriendlyName "jb365 EWS Client ($($date))"
```

**Example 2**

```powershell
$date = [datetime]::Now.ToString('yyyyMMdd')
$newCertParams = @{
    DnsName           = "$($date).EwsClient.jb365.ca"
    FriendlyName      = "jb365 EWS client ($($date))"
    CertStoreLocation = 'Cert:\LocalMachine\My'
    NotAfter          = (Get-Date).AddDays(7)
}
New-SelfSignedEwsOAuthApplicationCertificate @newCertParams
```