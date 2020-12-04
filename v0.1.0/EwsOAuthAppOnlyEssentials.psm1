#Requires -Version 5.1
using namespace System
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

<# v0.0.1 (incomplete and unpublished, but working great!) #>

function New-EwsAccessToken {

    [CmdletBinding(
        DefaultParameterSetName = 'Certificate'
    )]
    param (
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [Alias('ClientId')]
        [Guid]$ApplicationId,

        [Parameter(
            Mandatory,
            ParameterSetName = 'Certificate',
            HelpMessage = 'E.g. Use $Certificate, where `$Certificate = Get-ChildItem cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [X509Certificate2]$Certificate,

        [Parameter(
            Mandatory,
            ParameterSetName = 'CertificateStorePath',
            HelpMessage = 'E.g. cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317; E.g. cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [ValidateScript(
            {
                if (Test-Path -Path $_) { $true } else {
                
                    throw "An example proper path would be 'cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'."
                }
            }
        )]
        [string]$CertificateStorePath,

        [ValidateRange(1, 10)]
        [int16]$JWTExpMinutes = 2
    )

    if ($PSCmdlet.ParameterSetName -eq 'CertificateStorePath') {

        try {
            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        catch { throw $_ }
    }
    else { $Script:Certificate = $Certificate }

    if (-not (Test-CertificateProvider -Certificate $Script:Certificate)) {

        $ErrorMessage = "The supplied certificate does not use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'.  " +
        "For best luck, use a certificate generated using New-SelfSignedEwsOAuthApplicationCertificate."

        throw $ErrorMessage
    }

    $NowUTC = [datetime]::UtcNow

    $JWTHeader = @{

        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-Base64UrlFriendly -String ([Convert]::ToBase64String($Script:Certificate.GetCertHash()))
    }

    $JWTClaims = @{

        aud = "https://login.microsoftonline.com/$TenantId/oauth2/token"
        exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes) -UFormat '%s') -replace '\..*'
        iss = $ApplicationId.Guid
        jti = [Guid]::NewGuid()
        nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
        sub = $ApplicationId.Guid
    }

    $EncodedJWTHeader = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTHeader))
    )
    
    $EncodedJWTClaims = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTClaims))
    )

    $JWT = ConvertTo-Base64UrlFriendly -String ($EncodedJWTHeader + '.' + $EncodedJWTClaims)

    $Signature = ConvertTo-Base64UrlFriendly -String ([Convert]::ToBase64String(
        
            $Script:Certificate.PrivateKey.SignData(
            
                [Text.Encoding]::UTF8.GetBytes($JWT),
                [HashAlgorithmName]::SHA256,
                [RSASignaturePadding]::Pkcs1
            )
        )
    )

    $JWT = $JWT + '.' + $Signature

    $Body = @{

        client_id             = $ApplicationId
        client_assertion      = $JWT
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        scope                 = 'https://outlook.office365.com/.default'
        grant_type            = "client_credentials"
    }

    $TokenRequestParams = @{

        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        Body        = $Body
        Headers     = @{ Authorization = "Bearer $($JWT)" }
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }

    try {
        Invoke-RestMethod @TokenRequestParams
    }
    catch { throw $_ }
}

function New-SelfSignedEwsOAuthApplicationCertificate {
    [CmdletBinding()]
    param (
        [ValidatePattern('(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+[a-zA-Z]{2,63}$)')]
        [string]$DnsName,    

        [Parameter(Mandatory)]
        [string]$FriendlyName,

        [ValidateScript(
            {
                if (Test-Path -Path $_) { $true } else {

                    throw "An example proper location would be 'cert:\CurrentUser\My'."
                }
            }
        )]
        [string]$CertStoreLocation = 'cert:\CurrentUser\My',

        [datetime]$NotAfter = [datetime]::Now.AddDays(90),

        [ValidateSet('Signature', 'KeyExchange')]
        [string]$KeySpec = 'Signature'
    )

    $NewCertParams = @{

        DnsName           = $DnsName
        FriendlyName      = $FriendlyName
        CertStoreLocation = $CertStoreLocation
        NotAfter          = $NotAfter
        KeyExportPolicy   = 'Exportable'
        KeySpec           = $KeySpec
        Provider          = 'Microsoft Enhanced RSA and AES Cryptographic Provider'
        HashAlgorithm     = 'SHA256'
        ErrorAction       = 'Stop'
    }

    try {
        New-SelfSignedCertificate @NewCertParams
    }
    catch { throw $_ }
}