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

function Register-OofManagerMsalTokenCache {
    param(
        [Microsoft.Identity.Client.ITokenCache] $TokenCache
    )

    if ($null -eq $TokenCache) { return }

    try {
        $cacheDir = Join-Path $env:LOCALAPPDATA "OofManager"
        $cacheFile = Join-Path $cacheDir "powerapps-msal-cache.bin"

        $helperType = [System.Management.Automation.PSTypeName]'OofManager.PowerApps.MsalTokenCacheHelper'
        if ($null -eq $helperType.Type) {
            $msalPath = Join-Path (Split-Path $script:MyInvocation.MyCommand.Path) "Microsoft.Identity.Client.dll"
            Add-Type -ReferencedAssemblies $msalPath, 'System.Security.dll' -TypeDefinition @'
using System;
using System.IO;
using System.Security.Cryptography;
using Microsoft.Identity.Client;

namespace OofManager.PowerApps
{
    public static class MsalTokenCacheHelper
    {
        public static void Register(ITokenCache tokenCache, string cacheFile)
        {
            if (tokenCache == null || string.IsNullOrWhiteSpace(cacheFile))
            {
                return;
            }

            tokenCache.SetBeforeAccess(args =>
            {
                try
                {
                    if (!File.Exists(cacheFile))
                    {
                        return;
                    }

                    var protectedBytes = File.ReadAllBytes(cacheFile);
                    if (protectedBytes.Length == 0)
                    {
                        return;
                    }

                    var bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
                    args.TokenCache.DeserializeMsalV3(bytes);
                }
                catch
                {
                }
            });

            tokenCache.SetAfterAccess(args =>
            {
                try
                {
                    if (!args.HasStateChanged)
                    {
                        return;
                    }

                    var directory = Path.GetDirectoryName(cacheFile);
                    if (!string.IsNullOrWhiteSpace(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    var bytes = args.TokenCache.SerializeMsalV3();
                    var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
                    File.WriteAllBytes(cacheFile, protectedBytes);
                }
                catch
                {
                }
            });
        }
    }
}
'@ -ErrorAction Stop
        }

        [OofManager.PowerApps.MsalTokenCacheHelper]::Register($TokenCache, $cacheFile)
    } catch {
        Write-Debug "OofManager token cache setup failed: $($_.Exception.Message)"
    }
}

function Add-PowerAppsAccount
{
    <#
    .SYNOPSIS
    Add PowerApps account.
    .DESCRIPTION
    The Add-PowerAppsAccount cmdlet logins the user or application account and save login information to cache. 
    Use Get-Help Add-PowerAppsAccount -Examples for more detail.
    .PARAMETER Audience
    The service audience which is used for login.
    .PARAMETER Endpoint
    The serivce endpoint which to call. The value can be "prod", "preview", "tip1", "tip2", "usgov", "dod", "usgovhigh", or "china".
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
    #>
    [CmdletBinding()]
    param
    (
        [string] $Audience = "https://service.powerapps.com/",

        [Parameter(Mandatory = $false)]
        [ValidateSet("prod","preview","tip1", "tip2", "usgov", "usgovhigh", "dod", "china")]
        [string]$Endpoint = "prod",

        [string]$Username = $null,

        [SecureString]$Password = $null,

        [string]$TenantID = $null,

        [string]$CertificateThumbprint = $null,

        [string]$ClientSecret = $null,

        [string]$ApplicationId = $null
    )

    if ($Audience -eq "https://service.powerapps.com/")
    {
        # It's the default audience - we should remap based on endpoint as needed
        $Audience = Get-DefaultAudienceForEndPoint($Endpoint)
    }
    $global:currentSession = $null
    Add-PowerAppsAccountInternal -Audience $Audience -Endpoint $Endpoint -Username $Username -Password $Password -TenantID $TenantID -CertificateThumbprint $CertificateThumbprint -ClientSecret $ClientSecret -ApplicationId $ApplicationId
}


function Add-PowerAppsAccountInternal
{
    param
    (
        [string] $Audience = "https://service.powerapps.com/",

        [Parameter(Mandatory = $false)]
        [ValidateSet("prod","preview","tip1", "tip2", "usgov", "usgovhigh", "dod", "china")]
        [string]$Endpoint = "prod",

        [string]$Username = $null,

        [SecureString]$Password = $null,

        [string]$TenantID = $null,

        [string]$CertificateThumbprint = $null,

        [string]$ClientSecret = $null,

        [string]$ApplicationId = $null
    )

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
                Write-Debug "Token found and value, returning"
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

            Add-PowerAppsAccountInternal -Audience $Audience -Endpoint $global:currentSession.endpoint -Username $global:currentSession.username -Password $global:currentSession.password -TenantID $global:currentSession.InitialTenantId -CertificateThumbprint $global:currentSession.certificateThumbprint -ClientSecret $global:currentSession.clientSecret -ApplicationId $global:currentSession.applicationId
            $global:currentSession.recursed = $false

            # Afer recursing we can early return
            return
        }
    }
    else
    {
        [string] $jwtTokenForClaims = $null
        [Microsoft.Identity.Client.AzureCloudInstance] $authBaseUri =
            switch ($Endpoint)
                {
                    "usgov"     { [Microsoft.Identity.Client.AzureCloudInstance]::AzurePublic }
                    "usgovhigh" { [Microsoft.Identity.Client.AzureCloudInstance]::AzureUsGovernment }
                    "dod"       { [Microsoft.Identity.Client.AzureCloudInstance]::AzureUsGovernment }
                    "china"     { [Microsoft.Identity.Client.AzureCloudInstance]::AzureChina }
                    default     { [Microsoft.Identity.Client.AzureCloudInstance]::AzurePublic }
                };

        [Microsoft.Identity.Client.AadAuthorityAudience] $aadAuthAudience = [Microsoft.Identity.Client.AadAuthorityAudience]::AzureAdAndPersonalMicrosoftAccount
        if ($Username -ne $null -and $Password -ne $null)
        {
            $aadAuthAudience = [Microsoft.Identity.Client.AadAuthorityAudience]::AzureAdMultipleOrgs
        }

        Write-Debug "Using $aadAuthAudience : $Audience : $ApplicationId"

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
                $clientCertificate = Get-Item -Path Cert:\CurrentUser\My\$CertificateThumbprint
                $ConfidentialClientApplication = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder ]::Create($ApplicationId).WithCertificate($clientCertificate).Build()
            }
            else
            {
                Write-Debug "Using clientSecret for token acquisition"
                $ConfidentialClientApplication = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder ]::Create($ApplicationId).WithClientSecret($ClientSecret).Build()
            }

            $authResult = $ConfidentialClientApplication.AcquireTokenForClient($scopes).WithAuthority($authBaseuri, $TenantID, $true).ExecuteAsync() | Await-Task
            $clientBase = $ConfidentialClientApplication
        }
        else 
        {
            [Microsoft.Identity.Client.IPublicClientApplication]$PublicClientApplication = $null
            $PublicClientApplication = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ApplicationId).WithAuthority($authBaseuri, $aadAuthAudience, $true).WithDefaultRedirectUri().Build()
            Register-OofManagerMsalTokenCache -TokenCache $PublicClientApplication.UserTokenCache

            if ($Username -ne $null -and $Password -ne $null)
            {
                Write-Debug "Using username, password"
                $authResult = $PublicClientApplication.AcquireTokenByUsernamePassword($scopes, $UserName, $Password).ExecuteAsync() | Await-Task
            }
            else
            {
                try
                {
                    $accounts = @($PublicClientApplication.GetAccountsAsync() | Await-Task)
                    $account = $null
                    if ($Username)
                    {
                        $account = $accounts | Where-Object { $_.Username -ieq $Username } | Select-Object -First 1
                    }
                    if ($null -eq $account -and $accounts.Count -gt 0)
                    {
                        $account = $accounts[0]
                    }

                    if ($null -ne $account)
                    {
                        Write-Debug "Using cached MSAL account"
                        $authResult = $PublicClientApplication.AcquireTokenSilent($scopes, $account).ExecuteAsync() | Await-Task
                    }
                }
                catch [Microsoft.Identity.Client.MsalUiRequiredException]
                {
                    Write-Debug "Cached MSAL account requires interactive login"
                }
                catch
                {
                    Write-Debug "Cached MSAL account failed: $($_.Exception.Message)"
                }

                if ($null -eq $authResult)
                {
                    Write-Debug "Using interactive login"
                    $interactive = $PublicClientApplication.AcquireTokenInteractive($scopes)
                    if ($Username)
                    {
                        $interactive = $interactive.WithLoginHint($Username)
                    }
                    $authResult = $interactive.ExecuteAsync() | Await-Task
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
           Write-Debug "Adding new audience to resourceToken map. Expires $authResult.ExpiresOn"
            # addition of a new token for a new audience
            $global:currentSession.resourceTokens[$Audience] = @{
                accessToken = $authResult.AccessToken;
                expiresOn = $authResult.ExpiresOn;
            };
        }
        else
        {
            Write-Debug "Adding first audience to resourceToken map. Expires $authResult.ExpiresOn"
            $global:currentSession = @{
                loggedIn = $true;
                recursed = $false;
                endpoint = $Endpoint;
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
                    switch ($Endpoint)
                    {
                        "prod"      { "api.flow.microsoft.com" }
                        "usgov"     { "gov.api.flow.microsoft.us" }
                        "usgovhigh" { "high.api.flow.microsoft.us" }
                        "dod"       { "api.flow.appsplatform.us" }
                        "china"     { "api.powerautomate.cn" }
                        "preview"   { "preview.api.flow.microsoft.com" }
                        "tip1"      { "tip1.api.flow.microsoft.com"}
                        "tip2"      { "tip2.api.flow.microsoft.com" }
                        default     { throw "Unsupported endpoint '$Endpoint'"}
                    };
                powerAppsEndpoint = 
                    switch ($Endpoint)
                    {
                        "prod"      { "api.powerapps.com" }
                        "usgov"     { "gov.api.powerapps.us" }
                        "usgovhigh" { "high.api.powerapps.us" }
                        "dod"       { "api.apps.appsplatform.us" }
                        "china"     { "api.powerapps.cn" }
                        "preview"   { "preview.api.powerapps.com" }
                        "tip1"      { "tip1.api.powerapps.com"}
                        "tip2"      { "tip2.api.powerapps.com" }
                        default     { throw "Unsupported endpoint '$Endpoint'"}
                    };            
                bapEndpoint = 
                    switch ($Endpoint)
                    {
                        "prod"      { "api.bap.microsoft.com" }
                        "usgov"     { "gov.api.bap.microsoft.us" }
                        "usgovhigh" { "high.api.bap.microsoft.us" }
                        "dod"       { "api.bap.appsplatform.us" }
                        "china"     { "api.bap.partner.microsoftonline.cn" }
                        "preview"   { "preview.api.bap.microsoft.com" }
                        "tip1"      { "tip1.api.bap.microsoft.com"}
                        "tip2"      { "tip2.api.bap.microsoft.com" }
                        default     { throw "Unsupported endpoint '$Endpoint'"}
                    };      
                graphEndpoint = 
                    switch ($Endpoint)
                    {
                        "prod"      { "graph.windows.net" }
                        "usgov"     { "graph.windows.net" }
                        "usgovhigh" { "graph.windows.net" }
                        "dod"       { "graph.windows.net" }
                        "china"     { "graph.windows.net" }
                        "preview"   { "graph.windows.net" }
                        "tip1"      { "graph.windows.net"}
                        "tip2"      { "graph.windows.net" }
                        default     { throw "Unsupported endpoint '$Endpoint'"}
                    };
                msGraphEndpoint = 
                    switch ($Endpoint)
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
                        default     { throw "Unsupported endpoint '$Endpoint'"}
                    };
                cdsOneEndpoint = 
                    switch ($Endpoint)
                    {
                        "prod"      { "api.cds.microsoft.com" }
                        "usgov"     { "gov.api.cds.microsoft.us" }
                        "usgovhigh" { "high.api.cds.microsoft.us" }
                        "dod"       { "dod.gov.api.cds.microsoft.us" }
                        "china"     { "unsupported" }
                        "preview"   { "preview.api.cds.microsoft.com" }
                        "tip1"      { "tip1.api.cds.microsoft.com"}
                        "tip2"      { "tip2.api.cds.microsoft.com" }
                        default     { throw "Unsupported endpoint '$Endpoint'"}
                    };
                pvaEndpoint = 
                    switch ($Endpoint)
                    {
                        "prod"      { "powerva.microsoft.com" }
                        "usgov"     { "gcc.api.powerva.microsoft.us" }
                        "usgovhigh" { "high.api.powerva.microsoft.us" }
                        "dod"       { "powerva.api.appsplatform.us" }
                        "china"     { "unsupported" }
                        "preview"   { "bots.sdf.customercareintelligence.net" }
                        "tip1"       { "bots.ppe.customercareintelligence.net"}
                        "tip2"       { "bots.int.customercareintelligence.net"}
                        default     { throw "Unsupported endpoint '$Endpoint'"}
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

        # Write-Host "Checking Microsoft.Graph module installation"
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

            if ($userGraphResponse -ne $null) {
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
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectType -Value 'User' `
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
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectType -Value 'Group' `
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
        | Add-Member -PassThru -MemberType NoteProperty -Name ObjectType -Value "Organization" `
        | Add-Member -PassThru -MemberType NoteProperty -Name TenantId -Value $TenantObj.id `
        | Add-Member -PassThru -MemberType NoteProperty -Name Country -Value $TenantObj.countryLetterCode `
        | Add-Member -PassThru -MemberType NoteProperty -Name Language -Value $TenantObj.preferredLanguage `
        | Add-Member -PassThru -MemberType NoteProperty -Name DisplayName -Value $TenantObj.displayName `
        | Add-Member -PassThru -MemberType NoteProperty -Name Domains -Value $TenantObj.verifiedDomains `
        | Add-Member -PassThru -MemberType NoteProperty -Name Internal -Value $TenantObj;
}
# SIG # Begin signature block
# MIIoJgYJKoZIhvcNAQcCoIIoFzCCKBMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJkg75Ma6NL9uC
# liy3yEaETpdBnNtDq3RYFtAG/1nkSKCCDYUwggYDMIID66ADAgECAhMzAAAEhJji
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
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGfcwghnzAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAASEmOIS4HijMV0AAAAA
# BIQwDQYJYIZIAWUDBAIBBQCggYIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# LwYJKoZIhvcNAQkEMSIEINFyM5WwmMPq3oFxkajdNxhB7gy49Wz0RteVHR4wZceN
# MDQGCisGAQQBgjcCAQwxJjAkoBKAEABUAGUAcwB0AFMAaQBnAG6hDoAMaHR0cDov
# L3Rlc3QgMA0GCSqGSIb3DQEBAQUABIIBALK4qldxWzoJUjcyshEROTs9ZsvFDuBP
# EC+QDXPCTxjmWwx03PtOIt9rUye0gqL1KCjz4nYKqv6cPECzCeomgk24zborCYKP
# DYq59yRQCcqqCrDhZ2vftNQ0N1av712/a0f1P2SfX82DrMrsxgZvdyopVBd+SfhT
# Ezv3874g9u40KclkLNVJ23IwU1xucrhHnCQ35DFDV+fd5LaBGgq52BJefPUfzavy
# 6FDRXNYjaGWfZyR2+EhQJQdU8WlLdtviN6LZ3YaoQP/ni+Kbjdzc+U4N3e+Bc1d7
# RdA6yhTkwCHTKdjcxGGcUKg5I0wexL1m5UfUOldn35V7HH173fTRGwahghetMIIX
# qQYKKwYBBAGCNwMDATGCF5kwgheVBgkqhkiG9w0BBwKggheGMIIXggIBAzEPMA0G
# CWCGSAFlAwQCAQUAMIIBWgYLKoZIhvcNAQkQAQSgggFJBIIBRTCCAUECAQEGCisG
# AQQBhFkKAwEwMTANBglghkgBZQMEAgEFAAQgXdNe8UadUZWMLGM+31a1wv9zsP0g
# MBo2Td2gfsc7dfkCBmijrYOStxgTMjAyNTA5MTUxNTU1MzIuMDI0WjAEgAIB9KCB
# 2aSB1jCB0zELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsG
# A1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMScwJQYD
# VQQLEx5uU2hpZWxkIFRTUyBFU046MkQxQS0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghH7MIIHKDCCBRCgAwIBAgITMwAA
# Af1z+WhazQxh7QABAAAB/TANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1T
# dGFtcCBQQ0EgMjAxMDAeFw0yNDA3MjUxODMxMTZaFw0yNTEwMjIxODMxMTZaMIHT
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRN
# aWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJzAlBgNVBAsTHm5T
# aGllbGQgVFNTIEVTTjoyRDFBLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAKFlrPg/jruCY2J0R0XnbtDExWMzSRFT5yC83NNkd6m57o74WYJIafqf5cpm
# C85EMhts6cWHHk4yBex4kFm7ehVtwEZAa7YSVM9OWZyqXBd9ZaVBG/IFF4g9sSKa
# PGDPkg9EvoUz9UwgP8Ht/MmdwRLZmbXFZ2i0afwL7KoPuSiNCsOkwyaSsEy5dFVt
# P9t7CopHlg0px0Hk6aztMyJv27WoEmJt1f/M15X8cu7PxFRXUoJRxrFKvBGbqVDv
# F2x88+7VEcog95DsTZ8OaMdXmV/3P15luB+m+MjZmRdME2bsN+8gNTySjskkq161
# hIfh+vvlm+vtZbTAj6DCR1LTz9wp9AjXDb6z8ibQ2nKo5yE6y867B3Ti6o7B9tvW
# ZL53ZNCKsQQ2YDKGPhH+33xUT9qT5KxdRfSHAZGM/IS/kI1/ruMuFKquFLU+1UZ7
# Kr0f8f/kCxNKXEhIf1xNcNX3KeiZqvEZxxF4pMnDCzf2vymMaUj9xXxWy2bn/qiK
# 8hS9IBA8rWqRp9TjY1ZIiqVT9rqlSGI+FYgo8uaS1HHjHqoioGKoaZlBwhNlrLCy
# 4XUAR3aZdvPpPmWOpuHTxZxKBnCR7jHCGZ8OHDsIsaI0Tq/jau9XCY+0OC9F8D77
# kx0LdKB+0SjEIJrMuwlQ+7+eXToXR13WLMjuvXQHSvp1pcmHAgMBAAGjggFJMIIB
# RTAdBgNVHQ4EFgQU6QzFwOGVvPsi9vt7wOkZlO6BCqQwHwYDVR0jBBgwFoAUn6cV
# XQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUy
# MFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQ
# aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQl
# MjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcN
# AQELBQADggIBAGPYWF/k+QJgq2Bmh/ek3UeU+dvzzThu8kmHqKb+H5Zw1kC4QZa2
# rwIPqY5Tb+V0l2ayhr/HuLOXSeVnYXwvcsBUKuE5l51Hrz17Zbm2ZPtNgVyuv9t4
# TNE0irNipYWIqs20XvEGzHylxA7bzKB0mU+6/sCNiII2EMJGvtz/VV4BEcLuOv3M
# 8/CEf2avrzuedtyZXerLFbs7PbsCKyYX3GAY+dJl1kQXDIc2oy41g4HIodA7spD3
# AaaEy5Ti/C6V6KKp6/kC2BOAaVHqdyckjGHz89oXzi94NNlhH7DsafADW3HYqjN9
# XZt70oXhJJoxwNs7jPk4J+I+Z/gJ8uyDg2EJCKzVYS3TC9PXrtXSD4aduJRbZ1k2
# DWhUznzKhWtwG/CgyonJqdALYUTWVYNATwC+fPgdFHKARis0vY7HMDk7tSZjZYrD
# ipFVFZEieRaP3LXw0j3Qk1WiF1xe5eNJNXDP19jtCXQEve0+/JWI7cPz8m7s1+bI
# cQYf0akz7wsgISMQVSnzf4X7OAiKBWqlidK//EgdQhrMsiHD3xIDKPHHqtcOWaNC
# X58hYuhrqPs9yzxZf3sUGkbmxK7AFE38gWOf+ZYsr4wIMg2JxAfLxzu3OxYNrRne
# YRoGLPgDqFsduPl3MsaVJAGow4ZMvQ5fvCWU47bOgXE/bGE5jqHZP0oCMIIHcTCC
# BVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWlj
# cm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMw
# MTgyMjI1WhcNMzAwOTMwMTgzMjI1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOThpkzntHIhC3mi
# y9ckeb0O1YLT/e6cBwfSqWxOdcjKNVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+
# Slr+uDZnhUYjDLWNE893MsAQGOhgfWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3
# oAo4bo3t1w/YJlN8OWECesSq/XJprx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+
# tuhiJdxqD89d9P6OU8/W7IVWTe/dvI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0
# hyTD4MmPfrVUj9z6BVWYbWg7mka97aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLN
# ueKNiOSWrAFKu75xqRdbZ2De+JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZ
# nkXftnIv231fgLrbqn427DZM9ituqBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n
# 6Jl8P0zbr17C89XYcz1DTsEzOUyOArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC
# 4jMYctenIPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vc
# G9H9stQcxWv2XFJRXRLbJbqvUAV6bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtF
# tvUeh17aj54WcmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEE
# BQIDAQABMCMGCSsGAQQBgjcVAgQWBBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNV
# HQ4EFgQUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3
# TIN9AQEwQTA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9Eb2NzL1JlcG9zaXRvcnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkG
# CSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8E
# BTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRP
# ME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEww
# SgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMv
# TWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCd
# VX38Kq3hLB9nATEkW+Geckv8qW/qXBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQ
# dTltuw8x5MKP+2zRoZQYIu7pZmc6U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnu
# e99qb74py27YP0h1AdkY3m2CDPVtI1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYo
# VSfQJL1AoL8ZthISEV09J+BAljis9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlC
# GVJ1ijbCHcNhcy4sa3tuPywJeBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZ
# lvSP9pEB9s7GdP32THJvEKt1MMU0sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/
# ZPkkvnNtyo4JvbMBV0lUZNlz138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtq
# RRFHqfG3rsjoiV5PndLQTHa1V1QJsWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+
# y/g75LcVv7TOPqUxUYS8vwLBgqJ7Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgk
# NWcr4A245oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqK
# Oghif9lwY1NNje6CbaUFEMFxBmoQtB1VM1izoXBm8qGCA1YwggI+AgEBMIIBAaGB
# 2aSB1jCB0zELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsG
# A1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMScwJQYD
# VQQLEx5uU2hpZWxkIFRTUyBFU046MkQxQS0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMVAKI9FrVV
# UFDUiqKra44p0QLAVHaDoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTAwDQYJKoZIhvcNAQELBQACBQDscmqlMCIYDzIwMjUwOTE1MTAzNjIxWhgP
# MjAyNTA5MTYxMDM2MjFaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAOxyaqUCAQAw
# BwIBAAICA+owBwIBAAICEqIwCgIFAOxzvCUCAQAwNgYKKwYBBAGEWQoEAjEoMCYw
# DAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0B
# AQsFAAOCAQEAaegPcvkFRdLJmgA1Y/u+UDetGyg0SaabGyugZgJqedGtuTg3iZTr
# YQQS/ccFdkVxxSqFghkFzbc61gUmcwbZb1B6zgU4k0JRNUdNxLgdSL5Umt1+kVg/
# L4HniLsZ2Ar+AzujGmX13+l/teT3GQ1cI50wQhMU6QRowxCKkubAEcrMhodowFkz
# jZQX2rt4Rr9TrrTC9857H6+lExazjVYcgMYbPYBNkdyO9E8ieZdyTXGpaAmmsXFY
# XXOxpGyBJciCmI/SsYc8M5L7FfEj2oUmxzsNcKndWDrGIxJbymYU6N3XHQhqZbrh
# Ajd+9utUIR5VGcpUO7S6hF54Bb96oMnAAzGCBA0wggQJAgEBMIGTMHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB/XP5aFrNDGHtAAEAAAH9MA0GCWCG
# SAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwLwYJKoZI
# hvcNAQkEMSIEIJkFs2JFNP4ZHajEaXpZaKgGeOrGCt23y8kToeYqFbCzMIH6Bgsq
# hkiG9w0BCRACLzGB6jCB5zCB5DCBvQQggChIDclKMLyH8f3g32ErqR5HhdaehhcI
# ygbPJUQeDUcwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAIT
# MwAAAf1z+WhazQxh7QABAAAB/TAiBCBzhk9DXHMehCZ6ucBkhm0+dDjftnT37wZ9
# Z5wv40tCJDANBgkqhkiG9w0BAQsFAASCAgA/70Yc51U00s24gzGhd3/OseWS6zo/
# Ij2W4nUVMifiDiBQyFG4HRCwQAqur+/i9IIt3Oo6yxUgBISrnoMfMABczSfHiA8H
# eUnc0OGyGq9423pjLu74VkgF9EHGkMc780y4FAUfD7rUMdf2aFtGdD9CxzVvgdRj
# 1OZ4qS7v42Mcu4US8utAxremjw64jqBChC+iCQYR2+LtEf7840sAdZljT1t612W9
# Ry/NgxSpYGuz6ccnWMLhHNbI7ZjVz7H0kAHCZoAIVn+Adzd6RmxT7sABuhqGd2XS
# xHQycuisaPTdSwlzNDyMmHL0ISFn5iAbtIc7J8gYkfiQvo/E5nVQT5AspxXxAyXq
# LOfINSBnj+1T0Z2/haVk87xitiN6L6EM9Pa77BpVco3b1URVt25Vxe8bE6iMVpqq
# 2scplyWdmAAGFghsD+WPZfCIxQAuNexO1Ozk9ea+XtyAmsH3g5ECSjMvrcdIJKdL
# wsZrrGglTxoyhfRLMLptcLRW0mZs8gpNXEzpQxWWzXwoKE0dA4ymFH3G4iDlA7cW
# UAgb4khjhnj8mphMBvDVF3icOqTJfXKnhYNgJYqGUJZETzuP+JF7ABCsVaG3kSSY
# Y7K/7NlbTc1UY4LWj5pq98n1hcqCHSblnY84VY1sfeHeQwzJGXQ/R6LZRyODG5b0
# qS4kX1mNaFPdtw==
# SIG # End signature block
