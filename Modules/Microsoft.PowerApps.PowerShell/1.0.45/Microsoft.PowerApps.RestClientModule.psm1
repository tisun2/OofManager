$local:ErrorActionPreference = "Stop"
function Get-AudienceForHostName
{
    [CmdletBinding()]
    Param(
        [string] $Uri
    )

    $hostMapping = @{
        "management.azure.com" = "https://management.azure.com/";
        "api.powerapps.com" = "https://service.powerapps.com/";
        "api.apps.appsplatform.us" = "https://service.apps.appsplatform.us/";
        "tip1.api.powerapps.com" = "https://service.powerapps.com/";
        "tip2.api.powerapps.com" = "https://service.powerapps.com/";
        "graph.windows.net" = "https://graph.windows.net/";
        "graph.microsoft.com" = "https://graph.microsoft.com/"; #tip1, tip2, and prod, gccmod
        "graph.microsoft.us" = "https://graph.microsoft.us/"; #high
        "dod-graph.microsoft.us" = "https://dod-graph.microsoft.us/"; #dod    
        "microsoftgraph.chinacloudapi.cn" = "https://microsoftgraph.chinacloudapi.cn/"; #china
        "api.bap.microsoft.com" = "https://service.powerapps.com/";
        "tip1.api.bap.microsoft.com" = "https://service.powerapps.com/";
        "tip2.api.bap.microsoft.com" = "https://service.powerapps.com/";
        "api.flow.microsoft.com" = "https://service.flow.microsoft.com/";
        "api.flow.appsplatform.us" = "https://service.flow.appsplatform.us/";
        "tip1.api.flow.microsoft.com" = "https://service.flow.microsoft.com/";
        "tip2.api.flow.microsoft.com" = "https://service.test.powerautomate.com/";
        "gov.api.bap.microsoft.us" = "https://gov.service.powerapps.us/";
        "high.api.bap.microsoft.us" = "https://high.service.powerapps.us/";
        "api.bap.appsplatform.us" = "https://service.apps.appsplatform.us/";
        "gov.api.powerapps.us" = "https://gov.service.powerapps.us/";
        "high.api.powerapps.us" = "https://high.service.powerapps.us/";
        "gov.api.flow.microsoft.us" = "https://gov.service.flow.microsoft.us/";
        "high.api.flow.microsoft.us" = "https://high.service.flow.microsoft.us/";
        "bots.int.customercareintelligence.net" = "a522f059-bb65-47c0-8934-7db6e5286414";
        "bots.ppe.customercareintelligence.net" = "a522f059-bb65-47c0-8934-7db6e5286414";
        "powerva.microsoft.com" = "96ff4394-9197-43aa-b393-6a41652e21f8";
        "gcc.api.powerva.microsoft.us" = "9315aedd-209b-43b3-b149-2abff6a95d59";
        "high.api.powerva.microsoft.us" = "69c6e40c-465f-4154-987d-da5cba10734e";
        "powerva.api.appsplatform.us" = "bd4a9f18-e349-4c74-a6b7-65dd465ea9ab";
    }

    $uriObject = New-Object System.Uri($Uri)
    $hostName = $uriObject.Host

    if ($null -ne $hostMapping[$hostName])
    {
        return $hostMapping[$hostName];
    }

    foreach ($host in $hostMapping.GetEnumerator()) 
    {
        if ($hostName -match $host.Key)
        {
            return $hostMapping[$host.Key]
        }
    }

    Write-Verbose "Unknown host $hostName. Using https://management.azure.com/ as a default";
    return "https://management.azure.com/";
}

function Invoke-Request(
    [CmdletBinding()]

    [Parameter(Mandatory=$True)]
    [string] $Uri,

    [Parameter(Mandatory=$True)]
    [string] $Method,

    [object] $Body = $null,

    [Hashtable] $Headers = @{},

    [switch] $ParseContent,

    [switch] $ThrowOnFailure
)
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $audience = Get-AudienceForHostName -Uri $Uri
    $token = Get-JwtToken -Audience $audience
    $Headers["Authorization"] = "Bearer $token";
    $Headers["User-Agent"] = "PowerShell cmdlets 1.0";

    try {
        if ($null -eq $Body -or $Body -eq "")
        {
            $response = Invoke-WebRequest -Uri $Uri -Headers $Headers -Method $Method -UseBasicParsing
        }
        else
        {
            $jsonBody = ConvertTo-Json $Body -Depth 20
            $response = Invoke-WebRequest -Uri $Uri -Headers $Headers -Method $Method -ContentType "application/json; charset=utf-8" -Body $jsonBody -UseBasicParsing
        }

        if ($ParseContent)
        {
            if ($response.Content)
            {
                return ConvertFrom-JsonWithErrorHandling -JsonString $response.Content;
            }
        }

        return $response
    } catch {
        $response = $_.Exception.Response
        if ($_.ErrorDetails)
        {
            $errorResponse = ConvertFrom-JsonWithErrorHandling -JsonString $_.ErrorDetails;
            $code = $response.StatusCode.value__
            $message = $errorResponse.Error.Message
            Write-Verbose "Status Code: '$code'. Message: '$message'"

            $response = New-Object -TypeName PSObject `
                | Add-Member -PassThru -MemberType NoteProperty -Name StatusCode -Value $response.StatusCode.value__ `
                | Add-Member -PassThru -MemberType NoteProperty -Name StatusDescription -Value $response.StatusDescription `
                | Add-Member -PassThru -MemberType NoteProperty -Name Headers -Value $response.Headers `
                | Add-Member -PassThru -MemberType NoteProperty -Name Error -Value $errorResponse.Error `
                | Add-Member -PassThru -MemberType NoteProperty -Name Message -Value $message `
                | Add-Member -PassThru -MemberType NoteProperty -Name Internal -value $response;
        }
        if ($ThrowOnFailure)
        {
            throw;
        }
        else
        {
            return $response
        }
    }
}

function InvokeApi
{
    <#
    .SYNOPSIS
    Invoke an API.
    .DESCRIPTION
    The InvokeApi cmdlet invokes an API based on input parameters.
    Use Get-Help InvokeApi -Examples for more detail.
    .PARAMETER Method
    The http request method.
    .PARAMETER Route
    The http URL.
    .PARAMETER Body
    The http request body.
    .PARAMETER ThrowOnFailure
    Throw exception on failure if it is true.
    .PARAMETER ApiVersion
    The service API version.
    .EXAMPLE
    InvokeApi -Method GET -Route $uri -Body $body -ThrowOnFailure
    Call $uri API as GET method with $body input and throw exception on failure.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Route,

        [Parameter(Mandatory = $false)]
        [object]$Body = $null,

        [Parameter(Mandatory = $false)]
        [switch]$ThrowOnFailure,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )

    Test-PowerAppsAccount;

    $uri = $Route `
        | ReplaceMacro -Macro "{apiVersion}"  -Value $ApiVersion `
        | ReplaceMacro -Macro "{flowEndpoint}" -Value $global:currentSession.flowEndpoint `
        | ReplaceMacro -Macro "{powerAppsEndpoint}" -Value $global:currentSession.powerAppsEndpoint `
        | ReplaceMacro -Macro "{bapEndpoint}" -Value $global:currentSession.bapEndpoint `
        | ReplaceMacro -Macro "{graphEndpoint}" -Value $global:currentSession.graphEndpoint `
        | ReplaceMacro -Macro "{msGraphEndpoint}" -Value $global:currentSession.msGraphEndpoint `
        | ReplaceMacro -Macro "{pvaEndpoint}" -Value $global:currentSession.pvaEndpoint `
        | ReplaceMacro -Macro "{cdsOneEndpoint}" -Value $global:currentSession.cdsOneEndpoint;

    Write-Verbose $uri

    If($ThrowOnFailure)
    {
        $result = Invoke-Request `
        -Uri $uri `
        -Method $Method `
        -Body $body `
        -ParseContent `
        -ThrowOnFailure `
        -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
    }
    else {
        $result = Invoke-Request `
        -Uri $uri `
        -Method $Method `
        -Body $body `
        -ParseContent `
        -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
    }

    if($result.nextLink)
    {
        $nextLink = $result.nextLink
        $resultValue = $result.value

        while($nextLink)
        {
            If($ThrowOnFailure)
            {
                $nextResult = Invoke-Request `
                -Uri $nextLinkuri `
                -Method $Method `
                -Body $body `
                -ParseContent `
                -ThrowOnFailure `
                -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
            }
            else {
                $nextResult = Invoke-Request `
                -Uri $nextLink `
                -Method $Method `
                -Body $body `
                -ParseContent `
                -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
            }

            $nextLink = $nextResult.nextLink
            $resultValue = $resultValue + $nextResult.value
        }

        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name value -Value $resultValue `
    }

    return $result;
}

function InvokeApiNoParseContent
{
    <#
    .SYNOPSIS
    Invoke an API without parsing return content.
    .DESCRIPTION
    The InvokeApiNoParseContent cmdlet invokes an API based on input parameters without parsing return content.
    Use Get-Help InvokeApiNoParseContent -Examples for more detail.
    .PARAMETER Method
    The http request method.
    .PARAMETER Route
    The http URL.
    .PARAMETER Body
    The http request body.
    .PARAMETER ThrowOnFailure
    Throw exception on failure if it is true.
    .PARAMETER ApiVersion
    The service API version.
    .EXAMPLE
    InvokeApiNoParseContent -Method PUT -Route $uri -Body $body -ThrowOnFailure
    Call $uri API as PUT method with $body input and throw exception on failure.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Route,

        [Parameter(Mandatory = $false)]
        [object]$Body = $null,

        [Parameter(Mandatory = $false)]
        [switch]$ThrowOnFailure,

        [Parameter(Mandatory = $false)]
        [string]$ApiVersion = "2016-11-01"
    )

    Test-PowerAppsAccount;

    $uri = $Route `
        | ReplaceMacro -Macro "{apiVersion}"  -Value $ApiVersion `
        | ReplaceMacro -Macro "{flowEndpoint}" -Value $global:currentSession.flowEndpoint `
        | ReplaceMacro -Macro "{powerAppsEndpoint}" -Value $global:currentSession.powerAppsEndpoint `
        | ReplaceMacro -Macro "{bapEndpoint}" -Value $global:currentSession.bapEndpoint `
        | ReplaceMacro -Macro "{graphEndpoint}" -Value $global:currentSession.graphEndpoint `
        | ReplaceMacro -Macro "{msGraphEndpoint}" -Value $global:currentSession.msGraphEndpoint `
        | ReplaceMacro -Macro "{pvaEndpoint}" -Value $global:currentSession.pvaEndpoint `
        | ReplaceMacro -Macro "{cdsOneEndpoint}" -Value $global:currentSession.cdsOneEndpoint;

    Write-Verbose $uri

    If($ThrowOnFailure)
    {
        $result = Invoke-Request `
        -Uri $uri `
        -Method $Method `
        -Body $body `
        -ThrowOnFailure `
        -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
    }
    else {
        $result = Invoke-Request `
        -Uri $uri `
        -Method $Method `
        -Body $body `
        -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
    }

    if($result.nextLink)
    {
        $nextLink = $result.nextLink
        $resultValue = $result.value

        while($nextLink)
        {
            If($ThrowOnFailure)
            {
                $nextResult = Invoke-Request `
                -Uri $nextLinkuri `
                -Method $Method `
                -Body $body `
                -ThrowOnFailure `
                -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
            }
            else {
                $nextResult = Invoke-Request `
                -Uri $nextLink `
                -Method $Method `
                -Body $body `
                -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
            }

            $nextLink = $nextResult.nextLink
            $resultValue = $resultValue + $nextResult.value
        }

        return New-Object -TypeName PSObject `
            | Add-Member -PassThru -MemberType NoteProperty -Name value -Value $resultValue `
    }

    return $result;
}

function ReplaceMacro
{
    <#
    .SYNOPSIS
    Replace macro to the specified value.
    .DESCRIPTION
    The ReplaceMacro cmdlet replace macro in input string with the specified value.
    Use Get-Help ReplaceMacro -Examples for more detail.
    .PARAMETER Input
    The input string.
    .PARAMETER Macro
    The macro to be replaced.
    .PARAMETER Value
    The value for the replacement.
    .EXAMPLE
    ReplaceMacro -Macro "{apiVersion}"  -Value $ApiVersion
    Replace {apiVersion} to $ApiVersion.
    #>
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$Input,

        [Parameter(Mandatory = $true)]
        [string]$Macro,

        [Parameter(Mandatory = $false)]
        [string]$Value
    )

    return $Input.Replace($Macro, $Value)
}


function BuildFilterPattern
{
    param
    (
        [Parameter(Mandatory = $false)]
        [object]$Filter
    )

    if ($null -eq $Filter -or $Filter.Length -eq 0)
    {
        return New-Object System.Management.Automation.WildcardPattern "*"
    }
    else
    {
        return New-Object System.Management.Automation.WildcardPattern @($Filter,"IgnoreCase")
    }
}

function ConvertFrom-JsonWithErrorHandling
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$JsonString
    )

    try {
        return ConvertFrom-Json $JsonString
    } catch {
        Write-Verbose "Invalid JSON string: '$JsonString', falling back to .NET deserialization."

        # try to de-serialize the json string by using .Net json serializer
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
        return (New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer -Property @{MaxJsonLength=67108864}).DeserializeObject($JsonString)
    }
}

function ResolveEnvironment
{
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$OverrideId
    )

    if (-not [string]::IsNullOrWhiteSpace($OverrideId))
    {
        return $OverrideId;
    }
    elseif ($global:currentSession.selectedEnvironment)
    {
        return $global:currentSession.selectedEnvironment;
    }

    return "~default";
}


function Select-CurrentEnvironment
{
 <#
 .SYNOPSIS
 Sets the current environment for listing powerapps, flows, and other environment resources
 .DESCRIPTION
 The Select-CurrentEnvironment cmdlet sets the current environment in which commands will
 execute when an environment is not specified. Use Get-Help Select-CurrentEnvironment -Examples
 for more detail.
 .PARAMETER EnvironmentName
 Environment identifier (not display name).
 .PARAMETER Default
 Shortcut to specify the default tenant environment
 .EXAMPLE
 Select-CurrentEnvironment -EnvironmentName 3c2f7648-ad60-4871-91cb-b77d7ef3c239
 Select environment 3c2f7648-ad60-4871-91cb-b77d7ef3c239 as the current environment. Cmdlets invoked
 after running this command will operate against this environment.
 .EXAMPLE
 Select-CurrentEnvironment ~default
 Select the default environment. Cmdlets invoked after running this will operate against the default
 environment.
 #>
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName=$true, ParameterSetName = "Name")]
        [String]$EnvironmentName,

        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [Switch]$Default
    )

    Test-PowerAppsAccount;

    if ($Default)
    {
        $global:currentSession.selectedEnvironment = "~default";
    }
    else
    {
        $global:currentSession.selectedEnvironment = $EnvironmentName;
    }
}
# SIG # Begin signature block
# MIIoAAYJKoZIhvcNAQcCoIIn8TCCJ+0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBiBx34J4ZKWoRA
# lGP0mzh9c7xzFk8ccy0iI69b9odKOqCCDXYwggX0MIID3KADAgECAhMzAAAEhV6Z
# 7A5ZL83XAAAAAASFMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjUwNjE5MTgyMTM3WhcNMjYwNjE3MTgyMTM3WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDASkh1cpvuUqfbqxele7LCSHEamVNBfFE4uY1FkGsAdUF/vnjpE1dnAD9vMOqy
# 5ZO49ILhP4jiP/P2Pn9ao+5TDtKmcQ+pZdzbG7t43yRXJC3nXvTGQroodPi9USQi
# 9rI+0gwuXRKBII7L+k3kMkKLmFrsWUjzgXVCLYa6ZH7BCALAcJWZTwWPoiT4HpqQ
# hJcYLB7pfetAVCeBEVZD8itKQ6QA5/LQR+9X6dlSj4Vxta4JnpxvgSrkjXCz+tlJ
# 67ABZ551lw23RWU1uyfgCfEFhBfiyPR2WSjskPl9ap6qrf8fNQ1sGYun2p4JdXxe
# UAKf1hVa/3TQXjvPTiRXCnJPAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUuCZyGiCuLYE0aU7j5TFqY05kko0w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwNTM1OTAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBACjmqAp2Ci4sTHZci+qk
# tEAKsFk5HNVGKyWR2rFGXsd7cggZ04H5U4SV0fAL6fOE9dLvt4I7HBHLhpGdE5Uj
# Ly4NxLTG2bDAkeAVmxmd2uKWVGKym1aarDxXfv3GCN4mRX+Pn4c+py3S/6Kkt5eS
# DAIIsrzKw3Kh2SW1hCwXX/k1v4b+NH1Fjl+i/xPJspXCFuZB4aC5FLT5fgbRKqns
# WeAdn8DsrYQhT3QXLt6Nv3/dMzv7G/Cdpbdcoul8FYl+t3dmXM+SIClC3l2ae0wO
# lNrQ42yQEycuPU5OoqLT85jsZ7+4CaScfFINlO7l7Y7r/xauqHbSPQ1r3oIC+e71
# 5s2G3ClZa3y99aYx2lnXYe1srcrIx8NAXTViiypXVn9ZGmEkfNcfDiqGQwkml5z9
# nm3pWiBZ69adaBBbAFEjyJG4y0a76bel/4sDCVvaZzLM3TFbxVO9BQrjZRtbJZbk
# C3XArpLqZSfx53SuYdddxPX8pvcqFuEu8wcUeD05t9xNbJ4TtdAECJlEi0vvBxlm
# M5tzFXy2qZeqPMXHSQYqPgZ9jvScZ6NwznFD0+33kbzyhOSz/WuGbAu4cHZG8gKn
# lQVT4uA2Diex9DMs2WHiokNknYlLoUeWXW1QrJLpqO82TLyKTbBM/oZHAdIc0kzo
# STro9b3+vjn2809D0+SOOCVZMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGeAwghncAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAASFXpnsDlkvzdcAAAAABIUwDQYJYIZIAWUDBAIB
# BQCggYIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJKoZIhvcNAQkEMSIE
# INR7tQG25O1/Zdppwk9DR5d9pkctntWrHL9ojf10AAzKMDQGCisGAQQBgjcCAQwx
# JjAkoBKAEABUAGUAcwB0AFMAaQBnAG6hDoAMaHR0cDovL3Rlc3QgMA0GCSqGSIb3
# DQEBAQUABIIBAG6CwcOVrQevbRxV/fPR/skp0DpVbN2dQbmv4eKj+1R/UIAAbsIc
# bjmgD2ovQfTKSer5fqdWN8uj3QJ6DFqgs+2G4EGmJkkTK8OZsSEYnpGtHYOKmZz3
# 9v9QjrvCr63/OH8gmI/7Zqwuu4233XIs9lEmytd6DRhAQSZylkaHnv2hNQ4vwrg6
# oj1CL4dUprjZavYlFFOHZUa/02WZmZMAVsFdwO0TPFKG/HbGGtoivMyKgZg8Hfoc
# 3W/TEbwt2vor5JWkz7rkKBPuzjBfcV3ct39c8jT7aOionpvguizYJztQwkzryLzW
# lucEDaNRkYaXspAV/m+DadZB128/ZWTtIUihgheWMIIXkgYKKwYBBAGCNwMDATGC
# F4Iwghd+BgkqhkiG9w0BBwKgghdvMIIXawIBAzEPMA0GCWCGSAFlAwQCAQUAMIIB
# UQYLKoZIhvcNAQkQAQSgggFABIIBPDCCATgCAQEGCisGAQQBhFkKAwEwMTANBglg
# hkgBZQMEAgEFAAQgpFKTzPYNcANWRewG/bobdemF97Gn2tDajvMyR12+5RcCBmij
# oQeiPhgSMjAyNTA5MTUxNTU1NDUuMzlaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQg
# QW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046N0Yw
# MC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2WgghHtMIIHIDCCBQigAwIBAgITMwAAAgbXvFE4mCPsLAABAAACBjANBgkqhkiG
# 9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYw
# JAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yNTAxMzAx
# OTQyNTBaFw0yNjA0MjIxOTQyNTBaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRp
# b25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046N0YwMC0wNUUwLUQ5NDcxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDpRIWbIM3Rlr397cjHaYx85l7I+ZVWGMCBCM91
# 1BpU6+IGWCqksqgqefZFEjKzNVDYC9YcgITAz276NGgvECm4ZfNv/FPwcaSDz7xb
# DbsOoxbwQoHUNRro+x5ubZhT6WJeU97F06+vDjAw/Yt1vWOgRTqmP/dNr9oqIbE5
# oCLYdH3wI/noYmsJVc7966n+B7UAGAWU2se3Lz+xdxnNsNX4CR6zIMVJTSezP/2S
# TNcxJTu9k2sl7/vzOhxJhCQ38rdaEoqhGHrXrmVkEhSv+S00DMJc1OIXxqfbwPjM
# qEVp7K3kmczCkbum1BOIJ2wuDAbKuJelpteNZj/S58NSQw6khfuJAluqHK3igkS/
# Oux49qTP+rU+PQeNuD+GtrCopFucRmanQvxISGNoxnBq3UeDTqphm6aI7GMHtFD6
# DOjJlllH1gVWXPTyivf+4tN8TmO6yIgB4uP00bH9jn/dyyxSjxPQ2nGvZtgtqnvq
# 3h3TRjRnkc+e1XB1uatDa1zUcS7r3iodTpyATe2hgkVX3m4DhRzI6A4SJ6fbJM9i
# sLH8AGKcymisKzYupAeFSTJ10JEFa6MjHQYYohoCF77R0CCwMNjvE4XfLHu+qKPY
# 8GQfsZdigQ9clUAiydFmVt61hytoxZP7LmXbzjD0VecyzZoL4Equ1XszBsulAr5L
# d2KwcwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFO0wsLKdDGpT97cx3Iymyo/SBm4S
# MB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29m
# dCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRg
# MF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0
# MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/
# BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQB23GZOfe9ThTUvD29i4t6lDpxJhpVR
# Mme+UbyZhBFCZhoGTtjDdphAArU2Q61WYg3YVcl2RdJm5PUbZ2bA77zk+qtLxC+3
# dNxVsTcdtxPDSSWgwBHxTj6pCmoDNXolAYsWpvHQFCHDqEfAiBxX1dmaXbiTP1d0
# XffvgR6dshUcqaH/mFfjDZAxLU1s6HcVgCvBQJlJ7xEG5jFKdtqapKWcbUHwTVqX
# QGbIlHVClNJ3yqW6Z3UJH/CFcYiLV/e68urTmGtiZxGSYb4SBSPArTrTYeHOlQIj
# /7loVWmfWX2y4AGV/D+MzyZMyvFw4VyL0Vgq96EzQKyteiVeBaVEjxQKo3AcPULR
# F4Uzz98P2tCM5XbFZ3Qoj9PLg3rgFXr0oJEhfh2tqUrhTJd13+i4/fek9zWicosh
# lwXgFu002ZWBVzASEFuqED48qyulZ/2jGJBcta+Fdk2loP2K3oSj4PQQe1MzzVZO
# 52AXO42MHlhm3SHo3/RhQ+I1A0Ny+9uAehkQH6LrxkrVNvZG4f0PAKMbqUcXG7xz
# nKJ0x0HYr5ayWGbHKZRcObU+/34ZpL9NrXOedVDXmSd2ylKSl/vvi1QwNJqXJl/+
# gJkQEetqmHAUFQkFtemi8MUXQG2w/RDHXXwWAjE+qIDZLQ/k4z2Z216tWaR6RDKH
# GkweCoDtQtzkHTCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJ
# KoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0
# eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25
# PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsH
# FPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTa
# mDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc
# 6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF
# 50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpG
# dc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOm
# TTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi
# 0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU
# 2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSF
# F5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCC
# AdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6C
# kTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1Ud
# IARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUE
# DDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8E
# BAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2U
# kFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5j
# b20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmww
# WgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkq
# hkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaT
# lz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYu
# nKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f
# 8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVC
# s/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzs
# kYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzH
# VG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+k
# KNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+
# CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAo
# GokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEz
# fbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKh
# ggNQMIICOAIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9u
# czEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjdGMDAtMDVFMC1EOTQ3MSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoD
# FQAEa0f118XHM/VNdqKBs4QXxNnN96CBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA7HJd7TAiGA8yMDI1MDkxNTA5
# NDIwNVoYDzIwMjUwOTE2MDk0MjA1WjB3MD0GCisGAQQBhFkKBAExLzAtMAoCBQDs
# cl3tAgEAMAoCAQACAg4MAgH/MAcCAQACAhKNMAoCBQDsc69tAgEAMDYGCisGAQQB
# hFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAw
# DQYJKoZIhvcNAQELBQADggEBAAVzxLcfxyh0NTTH2cp9GVm28JlRDMOT9xWL9T/k
# Cy1fKFm357v4I3F/QxBP5fcudNQpt//5SAdz4ZfKIbJD+/JmDZQFdmdNdwbIWOpj
# qXgu7+yOp6KpyYId6d9PIjl+r4i/t6abvrsEcwTmVSswehCyISJCjUOQj6z9xJyK
# YgUEG0O8FhOjsjcNTatKge3fCZar11KgiMGJp0fi9uwBnAAA8J5Gnmvigkyj04FP
# 0gpWkfyzrZxZRfqgfQryb/G8wyuUKXnLvga7YS/5eiZCwaYRnBNHLlTxKyHbboSK
# 6Bo7H6IuBccteeAWEc+3ZR5aZYWrZTcQy41EIm7nHYKqxTExggQNMIIECQIBATCB
# kzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAgbXvFE4mCPsLAAB
# AAACBjANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJ
# EAEEMC8GCSqGSIb3DQEJBDEiBCDRo6LoAJZasMYyiXwoTeYMf+1gOw5deN7JRLNV
# jKzTMzCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIODo9ZSIkZ6dVtKT+E/u
# Zx2WAy7KiXM5R1JIOhNJf0vSMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
# UENBIDIwMTACEzMAAAIG17xROJgj7CwAAQAAAgYwIgQgwSdz/NZAViqky0cOQhh7
# 7B4Cq1qNbKwoj2epaOU1hccwDQYJKoZIhvcNAQELBQAEggIApRR7MZCCxxlC+1Pc
# i29nfQHE/JEbT9DMlDEPFAGNOWNfRsdDwNoniuT5x0eTNbpqyOlq00Y10GCIccS7
# JZKFeIx9qisbojJdLAJSyYbLrz1z7rzyTfl0tpHPGQJZh3DolGUqN6itwnRENCSk
# 9vMYvanSp41n5DIDvD8IbKPbYSYqFNViEustd/Byr4pZoTNwthLPuEngtWHB7yJ2
# xxKkvfOg0TvaIOcqAFecC3fwGP3f9ffda9rL4zSAGX9fUnRp+EgJS3OXrAEpq1nS
# 8tbpkTRPjSzcopwzgXkIFgKeDuTcGfeksTO10C9yZlauk71O/QS5xNhoGyaDW/Ap
# k/Y86+cjSNoZ8TbtokxY1Vbp5iEmH1ANOI+SCP/JObNecUT6CfrME0tkygeXVyGC
# OilHE6UGZ5XiLSj9uda7pgmU/z+cCmpTgc7jEk9b/3TdLn+SZ1ybt9Sp0Yrcamc0
# z2fRPzMIOSVD1DY59OTvb+kK0yd2yR8qnukWAjUGAF7zeUqTvkYnt04tXteyUyPD
# Zm4bnVhaC/mFX38a55Kc4YpD6eHo7DwasVL6IM3MveR+e1DbZio936aEUzfKhpO/
# +g7bJPafnKSsyPjaLex0vGyU7Pksa7h6teNmsIYKdvw6keqG8UFrXv8YnoP51bH+
# NDN6LvEsrGPCPcBSOfhn/5LYmvg=
# SIG # End signature block
