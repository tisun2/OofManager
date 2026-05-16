$local:ErrorActionPreference = "Stop"

Add-Type -Path (Join-Path (Split-Path $script:MyInvocation.MyCommand.Path) "Microsoft.Identity.Client.dll")

function Get-JwtTokenClaims
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$JwtToken
    )

    $tokenSplit = $JwtToken.Split(".")
    $claimsSegment = $tokenSplit[1].Replace(" ", "+").Replace("-", "+").Replace('_', '/');
    
    $mod = $claimsSegment.Length % 4
    if ($mod -gt 0)
    {
        $paddingCount = 4 - $mod;
        for ($i = 0; $i -lt $paddingCount; $i++)
        {
            $claimsSegment += "="
        }
    }

    $decodedClaimsSegment = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($claimsSegment))

    return ConvertFrom-Json $decodedClaimsSegment
}


function Get-DefaultAudienceForEndPoint
{
    [CmdletBinding()]
    Param(
        [string] $Endpoint
    )

    $audienceMapping = @{
        "prod" = "https://service.powerapps.com/";
        "preview" = "https://service.powerapps.com/";
        "tip1"= "https://service.powerapps.com/";
        "tip2"= "https://service.powerapps.com/";
        "usgov"= "https://gov.service.powerapps.us/";
        "usgovhigh"= "https://high.service.powerapps.us/";
        "dod" = "https://service.apps.appsplatform.us/";
        "china" = "https://service.powerapps.cn/";
    }

    if ($null -ne $audienceMapping[$Endpoint])
    {
        return $audienceMapping[$Endpoint];
    }

    Write-Verbose "Unknown endpoint $Endpoint. Using https://service.powerapps.com/ as a default";
    return "https://service.powerapps.com/";
}

function RemoveHttpHttps {
    param (
        [string] $inputString
    )

    foreach($removalTarget in @('http://','https://'))
    {
        if($inputString.StartsWith($removalTarget))
        {
            $inputString = $inputString.Remove(0, $removalTarget.Length)
        }
    }

    return $inputString
}

function Await-Task {
    param (
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        $task
    )

    process {
        while (-not $task.AsyncWaitHandle.WaitOne(200)) { }
        $task.GetAwaiter().GetResult()
    }
}

function Add-PowerAppsAccount
{
    <#
    .SYNOPSIS
    Add PowerApps account.
    .DESCRIPTION
    The Add-PowerAppsAccount cmdlet logins the user or application account and save login information to cache. 
    Use 'Get-Help Add-PowerAppsAccount -Detailed' for descriptions of the parameters and example usages.
    .PARAMETER Audience
    The service audience which is used for login.
    .PARAMETER Endpoint
    The serivce endpoint which to call. The value can be "prod", "preview", "tip1", "tip2", "usgov", "dod", "usgovhigh", or "china".
    Can't be used if providing endpoint overrides
    .PARAMETER Username
    The user name used for login.
    .PARAMETER Password
    The password for the user.
    .PARAMETER TenantID
    The tenant Id of the user or application.
    .PARAMETER CertificateThumbprint
    The certificate thumbprint of the application.
    .PARAMETER ClientSecret
    The client secret of the application.
    .PARAMETER ApplicationId
    The application Id.
    .PARAMETER UseSystemBrowser
    Uses system browser rather than embedded web view for login. This is useful for scenarios where 2FA/yubikey login is required.
    .PARAMETER AudienceOverride
    Must be provided if giving endpoint overrides; this audience will be used for all subsequent auth calls, ignoring normally derived audiences
    .PARAMETER AuthBaseUriOverride
    Must be provided if giving endpoint overrides    
    .PARAMETER FlowEndpointOverride
    Must be provided if giving endpoint overrides
    .PARAMETER PowerAppsEndpointOverride
    Must be provided if giving endpoint overrides
    .PARAMETER BapEndpointOverride
    Must be provided if giving endpoint overrides
    .PARAMETER GraphEndpointOverride
    Must be provided if giving endpoint overrides
    .PARAMETER CdsOneEndpointOverride
    Can be provided if giving endpoint overrides
    .PARAMETER PvaEndpointOverride
    Can be provided if giving endpoint overrides
    .EXAMPLE
    Add-PowerAppsAccount
    Login to "prod" endpoint.
    .EXAMPLE
    Add-PowerAppsAccount -Endpoint "prod" -Username "username@test.onmicrosoft.com" -Password "password"
    Login to "prod" for user "username@test.onmicrosoft.com" by using password "password"
    .EXAMPLE
    Add-PowerAppsAccount `
      -Endpoint "tip1" `
      -TenantID 1a1fbe33-1ff4-45b2-90e8-4628a5112345 `
      -ClientSecret ABCDE]NO_8:YDLp0J4o-:?=K9cmipuF@ `
      -ApplicationId abcdebd6-e62c-4f68-ab74-b046579473ad
    Login to "tip1" for application abcdebd6-e62c-4f68-ab74-b046579473ad in tenant 1a1fbe33-1ff4-45b2-90e8-4628a5112345 by using client secret.
    .EXAMPLE
    Add-PowerAppsAccount `
      -Endpoint "tip1" `
      -TenantID 1a1fbe33-1ff4-45b2-90e8-4628a5112345 `
      -CertificateThumbprint 12345137C1B2D4FED804DB353D9A8A18465C8027 `
      -ApplicationId 08627eb8-8eba-4a9a-8c49-548266012345
    Login to "tip1" for application 08627eb8-8eba-4a9a-8c49-548266012345 in tenant 1a1fbe33-1ff4-45b2-90e8-4628a5112345 by using certificate.
    .EXAMPLE
    Add-PowerAppsAccount `
      -AudienceOverride:  "https://service.powerapps.com/" `
      -AuthBaseUriOverride: "https://login.microsoftonline.com" `
      -BapEndpointOverride:  "api.bap.microsoft.com" `
      -CdsOneEndpointOverride:  "api.cds.microsoft.com" `
      -FlowEndpointOverride:  "api.flow.microsoft.com" `
      -GraphEndpointOverride:  "graph.windows.net" `
      -PowerAppsEndpointOverride:  "api.powerapps.com" `
      -PvaEndpointOverride:  "powerva.microsoft.com"
    Login to an environment with the provided endpoints (examples above are for 'PROD')
    .EXAMPLE
    $Inputs | Add-PowerAppsAccount
    Login to an environment with the endpionts stored in a PS Custom Object variable; where its content is defined as:
    $Inputs = [pscustomobject]@{ `
      "AudienceOverride" = "https://service.powerapps.com/"; `
      "AuthBaseUriOverride" = "https://login.microsoftonline.com"; `
      "BapEndpointOverride" = "api.bap.microsoft.com"; `
      "CdsOneEndpointOverride" = "api.cds.microsoft.com"; `
      "FlowEndpointOverride" = "api.flow.microsoft.com"; `
      "GraphEndpointOverride" = "graph.windows.net"; `
      "PowerAppsEndpointOverride" = "api.powerapps.com"; `
      "PvaEndpointOverride" = "powerva.microsoft.com" } 
    .EXAMPLE
    Get-Content -Raw ".\OverrideEndpoints.json" | ConvertFrom-Json | Add-PowerAppsAccount
    Login to an environment with the endpoints stored in 'OverrideEndpoints.json'; where its content is of the form:
    {	
      "AudienceOverride":  "https://service.powerapps.com/",
      "AuthBaseUriOverride": "https://login.microsoftonline.com",
      "BapEndpointOverride":  "api.bap.microsoft.com",
      "CdsOneEndpointOverride":  "api.cds.microsoft.com",
      "FlowEndpointOverride":  "api.flow.microsoft.com",
      "GraphEndpointOverride":  "graph.windows.net",
      "PowerAppsEndpointOverride":  "api.powerapps.com",
      "PvaEndpointOverride":  "powerva.microsoft.com"
    }
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName="DerivedEndpoints")]
        [string] $Audience = "https://service.powerapps.com/",

        [Parameter(Mandatory = $false, ParameterSetName="DerivedEndpoints")]
        [ValidateSet("prod","preview","tip1", "tip2", "usgov", "usgovhigh", "dod", "china")]
        [string]$Endpoint = "prod",

        [string]$Username = $null,

        [SecureString]$Password = $null,

        [string]$TenantID = $null,

        [string]$CertificateThumbprint = $null,

        [string]$ClientSecret = $null,

        [string]$ApplicationId = $null,

        [bool] $UseSystemBrowser = $false,

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $AudienceOverride = [System.Environment]::GetEnvironmentVariable('AUDIENCE_OVERRIDE'),

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $AuthBaseUriOverride = [System.Environment]::GetEnvironmentVariable('AUTH_BASE_URI_OVERRIDE'),

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $FlowEndpointOverride = [System.Environment]::GetEnvironmentVariable('FLOW_ENDPOINT_OVERRIDE'),

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $PowerAppsEndpointOverride = [System.Environment]::GetEnvironmentVariable('POWERAPPS_ENDPOINT_OVERRIDE'),

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $BapEndpointOverride = [System.Environment]::GetEnvironmentVariable('BAP_ENDPOINT_OVERRIDE'),

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $GraphEndpointOverride = [System.Environment]::GetEnvironmentVariable('GRAPH_ENDPOINT_OVERRIDE'),

        [Parameter(Mandatory = $false,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $CdsOneEndpointOverride = "unsupported",

        [Parameter(Mandatory = $false,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $PvaEndpointOverride = "unsupported"
    )

    if ($Audience -eq "https://service.powerapps.com/" -and -not $PSBoundParameters.ContainsKey('AudienceOverride') -and -not $AudienceOverride)
    {
        # It's the default audience - we should remap based on endpoint as needed
        $Audience = Get-DefaultAudienceForEndPoint($Endpoint)
        $PSBoundParameters['Audience'] = $Audience

    }
    
    if ($AudienceOverride -and -not $PSBoundParameters.ContainsKey('AudienceOverride') ) {
        $envVarsEndpointsAsDict = @{
            AudienceOverride = $AudienceOverride
            AuthBaseUriOverride = $AuthBaseUriOverride
            FlowEndpointOverride = $FlowEndpointOverride
            PowerAppsEndpointOverride = $PowerAppsEndpointOverride
            BapEndpointOverride = $BapEndpointOverride
            GraphEndpointOverride = $GraphEndpointOverride
            CdsOneEndpointOverride = $CdsOneEndpointOverride
            PvaEndpointOverride = $PvaEndpointOverride
        }
    }

    $global:currentSession = $null

    if ($envVarsEndpointsAsDict) {
        Add-PowerAppsAccountInternal @envVarsEndpointsAsDict
    } else {
        Add-PowerAppsAccountInternal @PSBoundParameters
    }
}

function Add-PowerAppsAccountInternal
{
    param
    (
        [Parameter(Mandatory = $false, ParameterSetName="DerivedEndpoints")]
        [string] $Audience = "https://service.powerapps.com/",

        [Parameter(Mandatory = $false, ParameterSetName="DerivedEndpoints")]
        [ValidateSet("prod","preview","tip1", "tip2", "usgov", "usgovhigh", "dod", "china")]
        [string]$Endpoint = "prod",

        [string]$Username = $null,

        [SecureString]$Password = $null,

        [string]$TenantID = $null,

        [string]$CertificateThumbprint = $null,

        [string]$ClientSecret = $null,

        [string]$ApplicationId = $null,

        [bool] $UseSystemBrowser = $false,

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $AudienceOverride,

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $AuthBaseUriOverride,

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $FlowEndpointOverride,

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $PowerAppsEndpointOverride,

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $BapEndpointOverride,

        [Parameter(Mandatory = $true,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $GraphEndpointOverride,

        [Parameter(Mandatory = $false,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $CdsOneEndpointOverride = "unsupported",

        [Parameter(Mandatory = $false,ParameterSetName="ProvidedEndpoints", ValueFromPipelineByPropertyName)]
        [string] $PvaEndpointOverride = "unsupported"
    )

    $InputEndpoint = $Endpoint

    #Enforce format requirements for endpoint overrides and force their usage
    if ($PSBoundParameters.ContainsKey('AudienceOverride') -and $PSBoundParameters.ContainsKey('AuthBaseUriOverride') -and $PSBoundParameters.ContainsKey('FlowEndpointOverride') -and $PSBoundParameters.ContainsKey('PowerAppsEndpointOverride') -and $PSBoundParameters.ContainsKey('BapEndpointOverride') -and $PSBoundParameters.ContainsKey('GraphEndpointOverride'))
    {
        $InputEndpoint = 'override'

        Write-Verbose "Overrides were passed and will be used in place of any derived audiences or endpoints; run this command again to change configured overrides"
        #Ensure exactly 1 trailing '/'
        $AudienceOverride = $AudienceOverride.TrimEnd('/') + '/';
        $Audience = $AudienceOverride

        #Ensure no trailing '/'
        $AuthBaseUriOverride = $AuthBaseUriOverride.TrimEnd('/');

        #Ensure no leading 'http://' or 'https://' and no trailing '/'
        $FlowEndpointOverride = RemoveHttpHttps $FlowEndpointOverride.TrimEnd('/')
        $PowerAppsEndpointOverride = RemoveHttpHttps $PowerAppsEndpointOverride.TrimEnd('/')
        $BapEndpointOverride = RemoveHttpHttps $BapEndpointOverride.TrimEnd('/')
        $GraphEndpointOverride = RemoveHttpHttps $GraphEndpointOverride.TrimEnd('/')
        $CdsOneEndpointOverride = RemoveHttpHttps $CdsOneEndpointOverride.TrimEnd('/')
        $PvaEndpointOverride = RemoveHttpHttps $PvaEndpointOverride.TrimEnd('/')
    }
    elseif ($global:currentSession.audienceOverride -ne $null -and $global:currentSession.audienceOverride -ne '')
    {
        Write-Debug "Provided Audience '$Audience' is being replaced with previously provided override value '$($global:currentSession.audienceOverride)'"
        $Audience = $global:currentSession.audienceOverride
    }

    [string[]]$scopes = "$Audience/.default"
    Write-Debug "Using endpoint, $Endpoint"
    Write-Debug "Using audience, $Audience"

    if ([string]::IsNullOrWhiteSpace($ApplicationId)) {
        if (($Endpoint -eq "tip1")  -or ($Endpoint -eq "tip2")) {
           $ApplicationId = "66575185-e05b-476b-ab0a-8b574e1bacbd"
        } elseif (($Endpoint -eq "prod") -or ($Endpoint -eq "preview")) {
           $ApplicationId = "689e5960-2e49-4505-98d8-369236220fc6"
        } elseif ($Endpoint -eq "usgov") {
           $ApplicationId = "689e5960-2e49-4505-98d8-369236220fc6"
        } elseif ($Endpoint -eq "usgovhigh") {
           $ApplicationId = "0d74f23a-a499-466e-b3fb-9b08a1afb5c7"
        } elseif ($Endpoint -eq "dod") {
           $ApplicationId = "800ede4f-8bf0-4947-ac19-bfa70e9b0940"
        } elseif ($Endpoint -eq "china") {
           $ApplicationId = "689e5960-2e49-4505-98d8-369236220fc6"
        } else {
            $ApplicationId = "689e5960-2e49-4505-98d8-369236220fc6"
        }
    }

    Write-Debug "Using appId, $ApplicationId"

    [Microsoft.Identity.Client.IClientApplicationBase]$clientBase = $null
    [Microsoft.Identity.Client.AuthenticationResult]$authResult = $null

    if ($global:currentSession.loggedIn -eq $true -and $global:currentSession.recursed -ne $true)
    {
        Write-Debug "Already logged in, checking for token for resource $Audience"
        $authResult = $null
        
        if ($global:currentSession.resourceTokens[$Audience] -ne $null)
        {
            if ($global:currentSession.resourceTokens[$Audience].accessToken -ne $null -and `
                $global:currentSession.resourceTokens[$Audience].expiresOn -ne $null -and `
                $global:currentSession.resourceTokens[$Audience].expiresOn -gt (Get-Date))
            {
                Write-Debug "Token found and value, returning for audience $Audience"
                return
            }
            else
            {
                 # Already logged in with an account, silently asking for a token from MSAL which should refresh
                try
                {
                    Write-Debug "Already logged in, silently requesting token from MSAL"
                    $authResult = $global:currentSession.msalClientApp.AcquireTokenSilent($scopes, $global:currentSession.msalAccount).ExecuteAsync() | Await-Task
                }
                catch [Microsoft.Identity.Client.MsalUiRequiredException] 
                {
                    Write-Debug ('{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message)
                }
            }
        }

        if ($authResult -eq $null)
        {
            Write-Debug "No token found, reseting audience and recursing: $Audience"
            # Reset the current audience values and call Add-PowerAppsAccount again
            $global:currentSession.resourceTokens[$Audience] = $null
            $global:currentSession.recursed = $true

            $PSBoundParameters['Audience'] = $Audience

            # the override endpoint is set automatically when required params are passed
            if ($global:currentSession.endpoint -ne 'override')
            {
                $PSBoundParameters['Endpoint'] = $global:currentSession.endpoint
            }

            $PSBoundParameters['Username'] = $global:currentSession.username
            $PSBoundParameters['Password'] = $global:currentSession.password
            $PSBoundParameters['TenantID'] = $global:currentSession.InitialTenantId
            $PSBoundParameters['CertificateThumbprint'] = $global:currentSession.certificateThumbprint
            $PSBoundParameters['ClientSecret'] = $global:currentSession.clientSecret
            $PSBoundParameters['ApplicationId'] = $global:currentSession.applicationId

            Add-PowerAppsAccountInternal @PSBoundParameters
            $global:currentSession.recursed = $false

            # Afer recursing we can early return
            return
        }
    }
    else
    {
        [string] $jwtTokenForClaims = $null

        if ($InputEndpoint -ne "override")
        {
            [Microsoft.Identity.Client.AzureCloudInstance] $authBaseUri =
                switch ($InputEndpoint)
                    {
                        "usgov"     { [Microsoft.Identity.Client.AzureCloudInstance]::AzurePublic }
                        "usgovhigh" { [Microsoft.Identity.Client.AzureCloudInstance]::AzureUsGovernment }
                        "dod"       { [Microsoft.Identity.Client.AzureCloudInstance]::AzureUsGovernment }
                        "china"     { [Microsoft.Identity.Client.AzureCloudInstance]::AzureChina }
                        default     { [Microsoft.Identity.Client.AzureCloudInstance]::AzurePublic }
                    };
        }
        else
        {
            [string] $authBaseUri = $AuthBaseUriOverride
        }

        if ($Username -ne $null -and $Password -ne $null)
        {
            $authUriWithAudience = $AuthBaseUriOverride + "/organizations/"
            [Microsoft.Identity.Client.AadAuthorityAudience] $aadAuthAudience = [Microsoft.Identity.Client.AadAuthorityAudience]::AzureAdMultipleOrgs
        }
        else
        {
            $authUriWithAudience = $AuthBaseUriOverride + "/common/"
            [Microsoft.Identity.Client.AadAuthorityAudience] $aadAuthAudience = [Microsoft.Identity.Client.AadAuthorityAudience]::AzureAdAndPersonalMicrosoftAccount
        }

        Write-Debug "Using $Audience : $ApplicationId : $aadAuthAudience : $authUriWithAudience"

        if (![string]::IsNullOrWhiteSpace($TenantID) -and `
            (![string]::IsNullOrWhiteSpace($ClientSecret) -or ![string]::IsNullOrWhiteSpace($CertificateThumbprint)))
        {
            $options = New-Object -TypeName Microsoft.Identity.Client.ConfidentialClientApplicationOptions
            $options.ClientId = $ApplicationId
            $options.TenantId = $TenantID

            [Microsoft.Identity.Client.IConfidentialClientApplication ]$ConfidentialClientApplication = $null

            if (![string]::IsNullOrWhiteSpace($CertificateThumbprint))
            {
                Write-Debug "Using certificate for token acquisition"
                $clientCertificate = [Array] (Get-ChildItem -path Cert:*$CertificateThumbprint -Recurse)
                if ($clientCertificate.Length -gt 1)
                {
                    Write-Debug "Multiple instances of the certificate found"
                    $matchingCertificate = $null
                    foreach ($certificateInstance in $clientCertificate)
                    {
                        if ($null -ne $certificateInstance.PrivateKey)
                        {
                            Write-Debug "Found certificate instance with associated private key"
                            $matchingCertificate = $certificateInstance
                            break
                        }
                    }
                    if ($null -eq $matchingCertificate)
                    {
                        throw "Could not find an instance of a certificate with associated private key for thumbprint $CertificateThumbprint"
                    }
                    $clientCertificate = $matchingCertificate
                }
                elseif ($clientCertificate.Length -eq 1)
                {
                    Write-Debug "A single instance of the certificate was found"
                    $clientCertificate = $clientCertificate[0]
                }
                else
                {
                    throw "Could not find an instance of a certificate with thumbprint $CertificateThumbprint"
                }
                $ConfidentialClientApplication = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ApplicationId).WithCertificate($clientCertificate).WithAuthority($authBaseUri, $TenantID, $true).Build()
            }
            else
            {
                Write-Debug "Using clientSecret for token acquisition"
                $ConfidentialClientApplication = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ApplicationId).WithClientSecret($ClientSecret).WithAuthority($authBaseUri, $TenantID, $true).Build()
            }

            $authResult = $ConfidentialClientApplication.AcquireTokenForClient($scopes).ExecuteAsync() | Await-Task
            $clientBase = $ConfidentialClientApplication
        }
        else
        {
            if ($InputEndpoint -eq "override")
            {
                if ($UseSystemBrowser)
                {
                    $PublicClientApplication = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ApplicationId).WithAuthority($authUriWithAudience, $true).WithRedirectUri("http://localhost").Build()
                }
                else
                {
                    $PublicClientApplication = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ApplicationId).WithAuthority($authUriWithAudience, $true).WithDefaultRedirectUri().Build()
                }
            }
            else
            {
                if ($UseSystemBrowser)
                {
                    $PublicClientApplication = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ApplicationId).WithAuthority($authBaseUri, $aadAuthAudience, $true).WithRedirectUri("http://localhost").Build()
                }
                else
                {
                    $PublicClientApplication = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ApplicationId).WithAuthority($authBaseUri, $aadAuthAudience, $true).WithDefaultRedirectUri().Build()
                }
            }

            if ($Username -ne $null -and $Password -ne $null)
            {
                Write-Debug "Using username, password"
                $authResult = $PublicClientApplication.AcquireTokenByUsernamePassword($scopes, $UserName, $Password).ExecuteAsync() | Await-Task
            }
            else
            {
                Write-Debug "Using interactive login"
                if ($UseSystemBrowser)
                {
                    $authResult = $PublicClientApplication.AcquireTokenInteractive($scopes).WithUseEmbeddedWebView($false).ExecuteAsync() | Await-Task
                }
                else
                {
                    $authResult = $PublicClientApplication.AcquireTokenInteractive($scopes).ExecuteAsync() | Await-Task
                }
            }

            $clientBase = $PublicClientApplication
        }
    }

    if ($authResult -ne $null)
    {
        if (![string]::IsNullOrWhiteSpace($authResult.IdToken))
        {
            $jwtTokenForClaims = $authResult.IdToken
        }
        else
        {
            $jwtTokenForClaims = $authResult.AccessToken
        }

        $claims = Get-JwtTokenClaims -JwtToken $jwtTokenForClaims

        if ($global:currentSession.loggedIn -eq $true)
        {
           Write-Debug "Adding new audience '$Audience' to resourceToken map. Expires $($authResult.ExpiresOn)"
            # addition of a new token for a new audience
            $global:currentSession.resourceTokens[$Audience] = @{
                accessToken = $authResult.AccessToken;
                expiresOn = $authResult.ExpiresOn;
            };

            if ($AudienceOverride -ne $null -and $AudienceOverride -ne '')
            {
                Write-Debug "A new audience override '$AudienceOverride' was provided and is in use, instead of the previous audience override '$($global:currentSession.audienceOverride)', for all token aquisitions"
                $global:currentSession.audienceOverride = $AudienceOverride
            }
        }
        else
        {
            Write-Debug "Adding first audience '$Audience' to resourceToken map. Expires $($authResult.ExpiresOn)"
            $global:currentSession = @{
                audienceOverride = $AudienceOverride;
                loggedIn = $true;
                recursed = $false;
                endpoint = $InputEndpoint;
                msalClientApp = $clientBase;
                msalAccount = $authResult.Account;
                upn = $claims.upn;
                InitialTenantId = $TenantID;
                tenantId = $claims.tid;
                userId = $claims.oid;
                applicationId = $ApplicationId;
                username = $Username;
                password = $Password;
                certificateThumbprint = $CertificateThumbprint;
                clientSecret = $ClientSecret;
                resourceTokens = @{
                    $Audience = @{
                        accessToken = $authResult.AccessToken;
                        expiresOn = $authResult.ExpiresOn;
                    }
                };
                selectedEnvironment = "~default";
                flowEndpoint = 
                    switch ($InputEndpoint)
                    {
                        "override"  { $FlowEndpointOverride }
                        "prod"      { "api.flow.microsoft.com" }
                        "usgov"     { "gov.api.flow.microsoft.us" }
                        "usgovhigh" { "high.api.flow.microsoft.us" }
                        "dod"       { "api.flow.appsplatform.us" }
                        "china"     { "api.powerautomate.cn" }
                        "preview"   { "preview.api.flow.microsoft.com" }
                        "tip1"      { "tip1.api.flow.microsoft.com"}
                        "tip2"      { "tip2.api.flow.microsoft.com" }
                        default     { throw "Unsupported endpoint '$InputEndpoint'"}
                    };
                powerAppsEndpoint = 
                    switch ($InputEndpoint)
                    {
                        "override"  { $PowerAppsEndpointOverride }
                        "prod"      { "api.powerapps.com" }
                        "usgov"     { "gov.api.powerapps.us" }
                        "usgovhigh" { "high.api.powerapps.us" }
                        "dod"       { "api.apps.appsplatform.us" }
                        "china"     { "api.powerapps.cn" }
                        "preview"   { "preview.api.powerapps.com" }
                        "tip1"      { "tip1.api.powerapps.com"}
                        "tip2"      { "tip2.api.powerapps.com" }
                        default     { throw "Unsupported endpoint '$InputEndpoint'"}
                    };            
                bapEndpoint = 
                    switch ($InputEndpoint)
                    {
                        "override"  { $BapEndpointOverride }
                        "prod"      { "api.bap.microsoft.com" }
                        "usgov"     { "gov.api.bap.microsoft.us" }
                        "usgovhigh" { "high.api.bap.microsoft.us" }
                        "dod"       { "api.bap.appsplatform.us" }
                        "china"     { "api.bap.partner.microsoftonline.cn" }
                        "preview"   { "preview.api.bap.microsoft.com" }
                        "tip1"      { "tip1.api.bap.microsoft.com"}
                        "tip2"      { "tip2.api.bap.microsoft.com" }
                        default     { throw "Unsupported endpoint '$InputEndpoint'"}
                    };      
                graphEndpoint = 
                    switch ($InputEndpoint)
                    {
                        "override"  { $GraphEndpointOverride }
                        "prod"      { "graph.windows.net" }
                        "usgov"     { "graph.windows.net" }
                        "usgovhigh" { "graph.windows.net" }
                        "dod"       { "graph.windows.net" }
                        "china"     { "graph.windows.net" }
                        "preview"   { "graph.windows.net" }
                        "tip1"      { "graph.windows.net"}
                        "tip2"      { "graph.windows.net" }
                        default     { throw "Unsupported endpoint '$InputEndpoint'"}
                    };
                msGraphEndpoint = 
                    switch ($InputEndpoint)
                    {
                        "override"  { $GraphEndpointOverride }
                        "prod"      { "graph.microsoft.com" }
                        "usgov"     { "graph.microsoft.com" }
                        "usgovhigh" { "graph.microsoft.us" }
                        "dod"       { "dod-graph.microsoft.us" }
                        "china"     { "microsoftgraph.chinacloudapi.cn" }
                        "preview"   { "graph.microsoft.com" }
                        "tip1"      { "graph.microsoft.com"}
                        "tip2"      { "graph.microsoft.com" }
                        default     { throw "Unsupported endpoint '$InputEndpoint'"}
                    };
                cdsOneEndpoint = 
                    switch ($InputEndpoint)
                    {
                        "override"  { $CdsOneEndpointOverride }
                        "prod"      { "api.cds.microsoft.com" }
                        "usgov"     { "gov.api.cds.microsoft.us" }
                        "usgovhigh" { "high.api.cds.microsoft.us" }
                        "dod"       { "dod.gov.api.cds.microsoft.us" }
                        "china"     { "unsupported" }
                        "preview"   { "preview.api.cds.microsoft.com" }
                        "tip1"      { "tip1.api.cds.microsoft.com"}
                        "tip2"      { "tip2.api.cds.microsoft.com" }
                        default     { throw "Unsupported endpoint '$InputEndpoint'"}
                    };
                pvaEndpoint = 
                    switch ($InputEndpoint)
                    {
                        "override"  { $PvaEndpointOverride }
                        "prod"      { "powerva.microsoft.com" }
                        "usgov"     { "gcc.api.powerva.microsoft.us" }
                        "usgovhigh" { "high.api.powerva.microsoft.us" }
                        "dod"       { "powerva.api.appsplatform.us" }
                        "china"     { "unsupported" }
                        "preview"   { "bots.sdf.customercareintelligence.net" }
                        "tip1"       { "bots.ppe.customercareintelligence.net"}
                        "tip2"       { "bots.int.customercareintelligence.net"}
                        default     { throw "Unsupported endpoint '$InputEndpoint'"}
                    };
            };
        }
    }
}


function Test-PowerAppsAccount
{
    <#
    .SYNOPSIS
    Test PowerApps account.
    .DESCRIPTION
    The Test-PowerAppsAccount cmdlet checks cache and calls Add-PowerAppsAccount if user account is not in cache.
    Use Get-Help Test-PowerAppsAccount -Examples for more detail.
    .EXAMPLE
    Test-PowerAppsAccount
    Check if user account is cached.
    #>
    [CmdletBinding()]
    param
    (
    )

    if (-not $global:currentSession -or $global:currentSession.loggedIn -ne $true)
    {
        Add-PowerAppsAccountInternal
    }
}

function Remove-PowerAppsAccount
{
    <#
    .SYNOPSIS
    Remove PowerApps account.
    .DESCRIPTION
    The Remove-PowerAppsAccount cmdlet removes the user or application login information from cache.
    Use Get-Help Remove-PowerAppsAccount -Examples for more detail.
    .EXAMPLE
    Remove-PowerAppsAccount
    Removes the login information from cache.
    #>
    [CmdletBinding()]
    param
    (
    )

    if ($global:currentSession -ne $null -and $global:currentSession.upn -ne $null)
    {
        Write-Verbose "Logging out $($global:currentSession.upn)"
    }
    else
    {
        Write-Verbose "No user logged in"
    }

    $global:currentSession = @{
        loggedIn = $false;
    };
}

function Get-JwtToken
{
    <#
    .SYNOPSIS
    Get user login token.
    .DESCRIPTION
    The Get-JwtToken cmdlet get the user or application login information from cache. It will call Add-PowerAppsAccount if login token expired.
    Use Get-Help Get-JwtToken -Examples for more detail.
    .EXAMPLE
    Get-JwtToken "https://service.powerapps.com/"
    Get login token for PowerApps "prod".
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string] $Audience
    )

    if ($global:currentSession -eq $null)
    {
        $global:currentSession = @{
            loggedIn = $false;
        };
    }
    elseif($global:currentSession.audienceOverride -ne $null -and $global:currentSession.audienceOverride -ne '')
    {
        Write-Verbose "The provided audience '$Audience' will be ignored in place of the AudienceOverride '$($global:currentSession.audienceOverride)' provided in the most recent call to Add-PowerAppsAccount"
        $Audience = $global:currentSession.audienceOverride
    }

    Add-PowerAppsAccountInternal -Audience $Audience

    return $global:currentSession.resourceTokens[$Audience].accessToken;
}

function Invoke-OAuthDialog
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string] $ConsentLinkUri
    )

    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width=440; Height=640 }
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width=420; Height=600; Url=$ConsentLinkUri }
    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri        
        if ($Global:uri -match "error=[^&]*|code=[^&]*")
        {
            $form.Close()
        }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)

    $output = @{}

    foreach($key in $queryOutput.Keys)
    {
        $output["$key"] = $queryOutput[$key]
    }
    
    return $output
}

# see graph explorer https://developer.microsoft.com/en-us/graph/graph-explorer
function Get-TenantDetailsFromGraph
{
    <#
    .SYNOPSIS
    Get my organization tenant details from graph.
    .DESCRIPTION
    The Get-TenantDetailsFromGraph function calls graph and gets my organization tenant details. 
    Use Get-Help Get-TenantDetailsFromGraph -Examples for more detail.
    .PARAMETER GraphApiVersion
    Graph version to call. The default version is "v1.0".
    .EXAMPLE
    Get-TenantDetailsFromGraph
    Get my organization tenant details from graph by calling graph service in version v1.0.
    #>
    param
    (
        [string]$GraphApiVersion = "v1.0"
    )

    process 
    {

        #Write-Host "Checking Microsoft.Graph module installation"
        # Check if Microsoft.Graph module is installed
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            #Write-Host "Microsoft.Graph module not found. Installing Microsoft.Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }

        #Write-Host "Importing Microsoft.Graph Module"

        Import-Module Microsoft.Graph -ErrorAction stop

        Connect-MgGraph -Scopes "User.Read.All"

        $graphResponse = Get-MgOrganization -Property "id,countryLetterCode,preferredLanguage,displayName,verifiedDomains"
        
        if ($graphResponse -ne $null)
        {
            CreateTenantObject -TenantObj $graphResponse
        }
        else
        {
            return $graphResponse
        }
    }
}

#Returns users or groups from Graph
# see graph explorer https://developer.microsoft.com/en-us/graph/graph-explorer
function Get-UsersOrGroupsFromGraph {
    <#
    .SYNOPSIS
    Returns users or groups from Graph.

    .DESCRIPTION
    The Get-UsersOrGroupsFromGraph function calls Graph and gets users or groups from Graph. 
    Use Get-Help Get-UsersOrGroupsFromGraph -Examples for more detail.

    .PARAMETER ObjectId
    User object Id.

    .PARAMETER SearchString
    Search string.

    .PARAMETER GraphApiVersion
    Graph version to call. The default version is "v1.0".

    .EXAMPLE
    Get-UsersOrGroupsFromGraph -ObjectId "12345ba9-805f-43f8-98f7-34fa34aa51a7"
    Get user with user object Id "12345ba9-805f-43f8-98f7-34fa34aa51a7" from Graph by calling Graph service in version v1.0.

    .EXAMPLE
    Get-UsersOrGroupsFromGraph -SearchString "gfd"
    Get users whose UserPrincipalName starts with "gfd" from Graph by calling Graph service in version v1.0.
    #>
    [CmdletBinding(DefaultParameterSetName = "Id")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Id")]
        [string]$ObjectId,

        [Parameter(Mandatory = $true, ParameterSetName = "Search")]
        [string]$SearchString,

        [Parameter(Mandatory = $false, ParameterSetName = "Search")]
        [Parameter(Mandatory = $false, ParameterSetName = "Id")]
        [string]$GraphApiVersion = "v1.0"
    )

    Process {
        
        #Write-Host "Checking Microsoft.Graph module installation"

        # Check if Microsoft.Graph module is installed
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            #Write-Host "Microsoft.Graph module not found. Installing Microsoft.Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }

        #Write-Host "Importing Microsoft.Graph Module"

        Import-Module Microsoft.Graph -ErrorAction stop


        Connect-MgGraph -Scopes "User.ReadBasic.All", "User.Read.All", "Group.Read.All"

        if (-not [string]::IsNullOrWhiteSpace($ObjectId)) {
            # Fetch user details using Microsoft Graph SDK
            $userGraphResponse = Get-MgUser -UserId $ObjectId -Property "assignedPlans,assignedLicenses,id,displayName,mail,userPrincipalName" -ErrorAction SilentlyContinue

            if ($userGraphResponse -ne $null){
                CreateUserObject -UserObj $userGraphResponse
            }

            # Fetch group details using Microsoft Graph SDK
            $groupGraphResponse = Get-MgGroup -GroupId $ObjectId -Property "id,displayName,mail" -ErrorAction SilentlyContinue

            if ($groupGraphResponse -ne $null) {
                CreateGroupObject -GroupObj $groupGraphResponse
            }
        } else {
            # Search for users using Microsoft Graph SDK
            $userGraphResponse = Get-MgUser -Filter "startswith(userPrincipalName,'$SearchString') or startswith(displayName,'$SearchString')" -Property "assignedPlans,assignedLicenses,id,displayName,mail,userPrincipalName" -ErrorAction SilentlyContinue

            if ($userGraphResponse -ne $null) {
                foreach ($user in $userGraphResponse) {
                    CreateUserObject -UserObj $user
                }
            }

            # Search for groups using Microsoft Graph SDK
            $groupsGraphResponse = Get-MgGroup -Filter "startswith(displayName,'$SearchString')" -Property "id,displayName,mail" -ErrorAction SilentlyContinue

            if ($groupsGraphResponse -ne $null) {
                foreach ($group in $groupsGraphResponse) {
                    CreateGroupObject -GroupObj $group
                }
            }
        }
    }
}


function CreateUserObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$UserObj
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectType -Value "User" `
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectId -Value $UserObj.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name UserPrincipalName -Value $UserObj.userPrincipalName `
        | Add-Member -PassThru -MemberType NoteProperty -Name Mail -Value $UserObj.mail `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $UserObj.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name AssignedLicenses -Value $UserObj.assignedLicenses `
        | Add-Member -PassThru -MemberType NoteProperty -Name AssignedPlans -Value $UserObj.assignedLicenses `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $UserObj;
}

function CreateGroupObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$GroupObj
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectType -Value "Group" `
        | Add-Member -PassThru -MemberType NoteProperty -Name Objectd -Value $GroupObj.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name Mail -Value $GroupObj.mail `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $GroupObj.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $GroupObj;
}


function CreateTenantObject
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$TenantObj
    )

    return New-Object -TypeName PSObject `
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectType -Value 'Organization' `
        | Add-Member -PassThru -MemberType NoteProperty -Name TenantId -Value $TenantObj.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name Country -Value $TenantObj.countryLetterCode `
        | Add-Member -PassThru -MemberType NoteProperty -Name Language -Value $TenantObj.preferredLanguage `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $TenantObj.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name Domains -Value $TenantObj.verifiedDomains `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $TenantObj;
}
# SIG # Begin signature block
# MIIoRwYJKoZIhvcNAQcCoIIoODCCKDQCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBRBHL+wPpo5QZQ
# 5j5+AeChO72aQwnW60nehfxdcidtXqCCDYUwggYDMIID66ADAgECAhMzAAAEhJji
# EuB4ozFdAAAAAASEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjUwNjE5MTgyMTM1WhcNMjYwNjE3MTgyMTM1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDtekqMKDnzfsyc1T1QpHfFtr+rkir8ldzLPKmMXbRDouVXAsvBfd6E82tPj4Yz
# aSluGDQoX3NpMKooKeVFjjNRq37yyT/h1QTLMB8dpmsZ/70UM+U/sYxvt1PWWxLj
# MNIXqzB8PjG6i7H2YFgk4YOhfGSekvnzW13dLAtfjD0wiwREPvCNlilRz7XoFde5
# KO01eFiWeteh48qUOqUaAkIznC4XB3sFd1LWUmupXHK05QfJSmnei9qZJBYTt8Zh
# ArGDh7nQn+Y1jOA3oBiCUJ4n1CMaWdDhrgdMuu026oWAbfC3prqkUn8LWp28H+2S
# LetNG5KQZZwvy3Zcn7+PQGl5AgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUBN/0b6Fh6nMdE4FAxYG9kWCpbYUw
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwNTM2MjAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# AGLQps1XU4RTcoDIDLP6QG3NnRE3p/WSMp61Cs8Z+JUv3xJWGtBzYmCINmHVFv6i
# 8pYF/e79FNK6P1oKjduxqHSicBdg8Mj0k8kDFA/0eU26bPBRQUIaiWrhsDOrXWdL
# m7Zmu516oQoUWcINs4jBfjDEVV4bmgQYfe+4/MUJwQJ9h6mfE+kcCP4HlP4ChIQB
# UHoSymakcTBvZw+Qst7sbdt5KnQKkSEN01CzPG1awClCI6zLKf/vKIwnqHw/+Wvc
# Ar7gwKlWNmLwTNi807r9rWsXQep1Q8YMkIuGmZ0a1qCd3GuOkSRznz2/0ojeZVYh
# ZyohCQi1Bs+xfRkv/fy0HfV3mNyO22dFUvHzBZgqE5FbGjmUnrSr1x8lCrK+s4A+
# bOGp2IejOphWoZEPGOco/HEznZ5Lk6w6W+E2Jy3PHoFE0Y8TtkSE4/80Y2lBJhLj
# 27d8ueJ8IdQhSpL/WzTjjnuYH7Dx5o9pWdIGSaFNYuSqOYxrVW7N4AEQVRDZeqDc
# fqPG3O6r5SNsxXbd71DCIQURtUKss53ON+vrlV0rjiKBIdwvMNLQ9zK0jy77owDy
# XXoYkQxakN2uFIBO1UNAvCYXjs4rw3SRmBX9qiZ5ENxcn/pLMkiyb68QdwHUXz+1
# fI6ea3/jjpNPz6Dlc/RMcXIWeMMkhup/XEbwu73U+uz/MIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGhgwghoUAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAASEmOIS4HijMV0AAAAA
# BIQwDQYJYIZIAWUDBAIBBQCggaAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIOEk
# srmFSVRMoX2hGSm2FJ1a8JAbAqV1ahKywR9Faz5NMDQGCisGAQQBgjcCAQwxJjAk
# oBKAEABUAGUAcwB0AFMAaQBnAG6hDoAMaHR0cDovL3Rlc3QgMA0GCSqGSIb3DQEB
# AQUABIIBAB/jco8tRgqBrpAvzSeHqk0854J2RxfeEgDfT3A9gfPBOQGIQf2EPnVt
# czsyza2Xe0ZRwmopSSwMZ+OZh1chj66Y3z/cGvmXrJ4QIgpRZHjP2IyUdmEWK+N9
# F7yBx1pOIAsnFe66oW5k3y1jpwygIHK4/fNNuBhm9rnRXhceNUvrmPYb1I7PHQkX
# FbYTIZtvM1WMG/95XMY6gU3R3Rl6qPuzhYSvAvNaWvw8M76fMCgEW/Xe6bqzstXw
# ml9qKoiI9a6F33u8gcgQah6Dmwt82MvHviieo/a7Ec5yLUcay8ONnP9avR6IndUg
# z/7diMDpSyjFQwwyokhVQw7x/+s5urWhghewMIIXrAYKKwYBBAGCNwMDATGCF5ww
# gheYBgkqhkiG9w0BBwKggheJMIIXhQIBAzEPMA0GCWCGSAFlAwQCAQUAMIIBWgYL
# KoZIhvcNAQkQAQSgggFJBIIBRTCCAUECAQEGCisGAQQBhFkKAwEwMTANBglghkgB
# ZQMEAgEFAAQgvyd+oKBR6rlu6CgM98Dv6j3oMP6+uMxZhrbVx+pmTD4CBmm8Uqsf
# ehgTMjAyNjAzMjMxODM4NTAuNjIyWjAEgAIB9KCB2aSB1jCB0zELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IEly
# ZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBF
# U046MkQxQS0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2WgghH+MIIHKDCCBRCgAwIBAgITMwAAAhLRCAY8yhhPqgABAAACEjAN
# BgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
# NTA4MTQxODQ4MTVaFw0yNjExMTMxODQ4MTVaMIHTMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBP
# cGVyYXRpb25zIExpbWl0ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjoyRDFB
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAK9M06A5KVkLbGXpEtHF
# dKrg3SkkxpKL7wWmR9DgCBItDtwnDu+yYl/HOBKavbomx1WWdVvy6LnxNe9r5EVz
# vkGDbVlqKxAidgHGUNdG6QJbZIWTYl5VSfC90M4SoK165jJJtifv4PNVNtyT3DuM
# FxxH7aJ098KXf//d+q45sMTJuzZG7MoGyX/uAFQTDa+GjD0IQXe+qHdCjAelq78h
# BjjlNPPdzhbn0uRA3suJ+OFoGsSRNFZ79/zjr7jsOSqdSj6o42Cfi/csC7gTLXA3
# y6TUDv4dXhRKDK24hDOu0znzZV38Ww8+DJbGLy4qMYHdl83stUMa1dfoviclQyTI
# knvYCjrs6YkEBRNfQ1D4LIncoy081xIUlSwZUaK9HglX+4AukX5PDWN6ztrIIDi+
# /b1ORbgyk4f7CDrXFB3hwuNowRgfrX3SgtSjgUflJTfWjs4PJqVDSNhYKkL4q1T/
# aaW3jFH76dsAPb6Mk4kVrw4MwsaPMZSdZ7HGExyEK5pBfY/wmtCA5rfH7zp+uJ55
# SThlGGWBzAtZdMbYJGExNRElKqpGsCpO4qm8XZy8snEvnUfs2sT7nTAy60Bc/JYH
# 8vaG5NA/Bwtnc4VqWDZ+YXZKDxM4AqkFDqfL90I/7HeGp8rqXuSqApIwATj98oUk
# OFvWfg6yZ4TP627Fwu6E79unAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUUgW5UsW+
# XGCmDsL3f1X6Fzn7t94wHwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIw
# XwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
# cy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3Js
# MGwGCCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENB
# JTIwMjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEF
# BQcDCDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAF50gJJl+/UX
# O5cUtcDqM1ye1dKLuQ57VXDiaZSJA47xDfASEKYeFPoEji2blUXj/8/WHVIlJHD8
# C9TzNuOf5BF9RHoKcTXPBwwOoSgh4NACzRHLsZxjpYCi6JF31Hq9Q+YYlvDwATPU
# fp3orXk4D2mkZSRbxk2L6LXNLqhohEuYTEIS/ETRYaUvUXCFh7Z0BhJ63TlLwgIb
# OBXlYmsHJi0yr/tfO9HPzHkx4tEA3Xkfu/1oOKoZCdpNYhByXZmH/KyFdDUQWXHU
# U03R3nt+Ulz+Z3jQnwJIwyLmcnEbneo8zywjS4jWMxlwbMMycoI+BHkjkU+8DL2h
# binkQF6Fwt8MGL4CLCNc59wxmOtuWPsmUovFDUjR7q4t+mb/WvkOLycA6WCt+ktE
# wdqX+8S4oh99p5O2Cu90YPfFun2diDbs2M2exoYL3335I+BFF4NRNBH32NaRKpG6
# Q0z+4fwwarc6D17MsNjFIfu8r1nKtgRmUrnGugmOl+IqDLnOT1qbJrzjpYuwETQO
# wG/JCnQnNoDQy2nIJbMHmRHPf1UAeoZbP2+ipN9p5MLhxMSpWnqElaygeVPcZadf
# nPCf+xiY+EcOwIkPLXKflpn8g/CsV9kJSmw4uElI54Jb+Ote0fPmv3A1icmjLfNu
# /Vp+39sjHnTe5HxiEOUmY+ukXYXZWTqvMIIHcTCCBVmgAwIBAgITMwAAABXF52ue
# AptJmQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgz
# MjI1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxO
# dcjKNVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQ
# GOhgfWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq
# /XJprx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVW
# Te/dvI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7
# mka97aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De
# +JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM
# 9ituqBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEz
# OUyOArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2
# ZItboKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqv
# UAV6bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q
# 4i6tAgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcV
# AgQWBBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXS
# ZacbUzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcC
# ARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRv
# cnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1
# AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaA
# FNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8y
# MDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAt
# MDYtMjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8
# qW/qXBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7p
# Zmc6U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2C
# DPVtI1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BA
# ljis9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJ
# eBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1
# MMU0sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz
# 138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1
# V1QJsWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLB
# gqJ7Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0l
# lOZ0dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFx
# BmoQtB1VM1izoXBm8qGCA1kwggJBAgEBMIIBAaGB2aSB1jCB0zELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IEly
# ZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBF
# U046MkQxQS0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2WiIwoBATAHBgUrDgMCGgMVAOVRwa+IdNBDe41HUN90hPqm5P/AoIGD
# MIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwDQYJKoZIhvcNAQEL
# BQACBQDta24DMCIYDzIwMjYwMzIzMDc0NTA3WhgPMjAyNjAzMjQwNzQ1MDdaMHcw
# PQYKKwYBBAGEWQoEATEvMC0wCgIFAO1rbgMCAQAwCgIBAAICCkUCAf8wBwIBAAIC
# EkwwCgIFAO1sv4MCAQAwNgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAK
# MAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQsFAAOCAQEAPJncCtnP
# z2z57wk/7t1aO3wFf5lEoo9nZivoR6IVjDb/q+OGilwQ7wzHOW3CZxbtzVcRiULb
# VjI1fqEDVs8ABCbZ1HPqneqboth3Zsl+mXQfquzY7+Fle8zvjbAO1QwqPkCYwe8z
# Xzi+1TDSfnDmC3VKnygXfyjAP5N6wd0IBzzpH1+QJn8YY4bglY4fXmbeeFHp3R+C
# H+ae3DnJ4tajnZ+tLqOJUQfRq1o0qEeQlXEqcvcDatHRAy7v6qEPeMoRa4i3tjeS
# 5tgifjLPY/DcE7jW7O2pl31BkQOd7FwTOgYLfut+lJ87kOLcrZmeKAuj1XxyLh9V
# uPE/UbhHAxlzFTGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAACEtEIBjzKGE+qAAEAAAISMA0GCWCGSAFlAwQCAQUAoIIBSjAa
# BgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIAk5FuQ7
# 7ZbMIjsufJ36szsPxLxnTxJjQqazCYA0qhpqMIH6BgsqhkiG9w0BCRACLzGB6jCB
# 5zCB5DCBvQQgc/l+Rrzu1p4JJx+AWXmFPgcwS9ScuaCfMFZUzSMX2IMwgZgwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAhLRCAY8yhhPqgAB
# AAACEjAiBCAramPNGqWuhTTZYz/Xfs8omOua5R8WgxyodUpLDo10fDANBgkqhkiG
# 9w0BAQsFAASCAgAy2zEvQgoxTNKTZq4de+srfOHs1zerlOsfcFPnfw9WtN9gy5Nh
# 34hq9JbH8liKxBUEI/5qhwLBapf+6SFfgaIt+OiWmGFMw+K+j7qkR2egByRlXnBX
# EwjwiFH0u7n4v7+elmHpt13NJyjz+4J68E3eyuvrhtVCUrgSQZcW/FT8OW/g0ZRM
# 4jNSfp8I5EzzONm4sp2ZSP7wMUpVNRQQvJZyq7OEQfyqqMeqn0BdAPKNdADhOyJU
# AKZh/G8UW1K8XblGtZPUAXCp/rUmEUqOFsbZzI76ViVfqeWe/j8U5qR8ykAOnA8q
# kybzwaPcxilBtAMArm2nYglz8jMIqlwcRs4fj5k531D1GDtLyKiEM2p16SGdTBZt
# FZEDY0DOu0JmfIKhNtEgIfqDBhhaFDotcWdZVoiKDToopnXjujmKfLI6KHq6aOTh
# cv9nIsuDsrSV6Cd9onrGnBKuBzVcOtQoch1KOSfuwioDZZefZ/qg9EdF1U+QA5Zx
# 4UumZLdBi+LpyefXvo20sDgAZzIZwe4niFkZWX5ZdTpxQrfsINk2ZN9MExL7FYt8
# jnkhZ5TlLkaMpf7nqD3bv4ve+fTjdgtoJbeK2iJN4HSPCMyP4xEFaGRJOTFJoNXp
# MgLu7i2UWXQgIMQbMJBvzFPYjerWWQVWWlSCL0LdjwyqfH2XbkH03NN4tw==
# SIG # End signature block
